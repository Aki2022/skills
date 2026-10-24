"""Transport mocks verify runner mechanics; they are not live model evidence."""
import copy
import importlib.util
import json
from pathlib import Path
import sys

import pytest

SCRIPT = Path(__file__).parents[1] / "classify.py"
spec = importlib.util.spec_from_file_location("design_classify", SCRIPT)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


def item(number=1):
    return {"id": f"dads-{number:06d}", "document": "components/button/index.md", "title": "Button", "headings": ["Usage"], "text": "Choose button style.", "source_url": "https://example.test/button", "classification": None}


def files(tmp_path, count=12):
    catalog, usage = tmp_path / "catalog.json", tmp_path / "usage.json"
    catalog.write_text(json.dumps({"schema_version": 1, "version": "test", "items": [item(n + 1) for n in range(count)]}))
    return catalog, usage


def answer(policy, choice="global0-lp1-product0-presentation0", confidence=0.9):
    keys = policy["outcomes"]
    return {"type": "choice", "choice": choice, "confidence": confidence, "probabilities": {key: float(key == choice) for key in keys}}


def responder(calls, usage_path=None):
    def mock(payload, key, timeout):
        if usage_path:
            ledger = json.loads(usage_path.read_text())
            assert ledger["attempted_judgments"] >= len(payload["questions"])
            assert ledger["batches"][-1]["status"] == "reserved"
        calls.append(copy.deepcopy(payload))
        policy = m.load_policy()
        return {"model": payload["model"], "answers": {qid: answer(policy) for qid in payload["questions"]}, "usage": {"input_tokens": 100, "output_tokens": 20}}
    return mock


def test_classification_exact_inputs_and_fixed_record():
    policy = m.load_policy()
    row = item()
    inp = m.classification_input(row)
    assert set(inp) == {"document", "title", "headings", "text", "source_url"}
    assert len(policy["outcomes"]) == 16
    result = m.classify_response(row, answer(policy), "jev-1.13.0", policy)
    assert result["grade"] == "nice-to-have"
    assert result["preset_must"] == ["lp"]
    assert result["model"] == "jev-1.13.0"
    assert result["threshold"] == 0.5
    assert len(result["input_hash"]) == len(result["policy_hash"]) == 64
    assert result["raw_answer"]["choice"] == "global0-lp1-product0-presentation0"
    assert result["uncertain"] is False
    result = m.classify_response(row, answer(policy, confidence=0.49), "jev-1.13.0", policy)
    assert result["grade"] == "must"
    assert result["preset_must"] == ["lp", "product-ui", "presentation"]
    assert result["uncertain"] is True


def test_reservation_batching_resume_dedup_and_usage(tmp_path, monkeypatch):
    monkeypatch.setenv("TYPESAFE_API_KEY", "test-secret")
    catalog, usage = files(tmp_path)
    calls = []
    result = m.run(catalog, usage, limit=10, _transport=responder(calls, usage))
    assert result["classified"] == 10
    assert len(calls) == 1
    payload = calls[0]
    assert len(payload["questions"]) == 10
    assert payload["state"]["items"]["dads-000001"] == m.classification_input(item())
    assert "items.dads-000001" in payload["questions"]["dads-000001"]["instructions"]["target"]
    assert "test-secret" not in usage.read_text() and "Choose button" not in usage.read_text()
    ledger = json.loads(usage.read_text())
    assert ledger["attempted_judgments"] == 10
    assert ledger["batches"][0]["usage_derived_usd"] == pytest.approx(100 * 0.042 / 1_000_000)
    assert m.run(catalog, usage, limit=50, _transport=responder(calls))["classified"] == 2
    complete = m.run(catalog, usage, limit=50, _transport=responder(calls))
    assert complete["status"] == "already-complete" and complete["classified"] == 0
    assert len(calls) == 2
    data = json.loads(catalog.read_text())
    data["items"][0]["text"] += " changed"
    catalog.write_text(json.dumps(data))
    assert m.run(catalog, usage, limit=50, _transport=responder(calls))["classified"] == 1
    assert json.loads(usage.read_text())["attempted_judgments"] == 13


def test_dry_run_never_reads_auth_calls_or_writes(tmp_path, monkeypatch):
    catalog, usage = files(tmp_path, 2)
    before = catalog.read_bytes()
    monkeypatch.setattr(m, "read_api_key", lambda: pytest.fail("auth read on dry-run"))
    result = m.run(catalog, usage, limit=50, dry_run=True, _transport=lambda *a: pytest.fail("network on dry-run"))
    assert result["status"] == "dry-run" and result["selected"] == 2
    assert catalog.read_bytes() == before and not usage.exists()


def test_shared_ceiling_and_failure_consumption(tmp_path, monkeypatch):
    monkeypatch.setenv("TYPESAFE_API_KEY", "test-secret")
    catalog, usage = files(tmp_path, 3)
    usage.write_text(json.dumps({"schema_version": 1, "attempted_judgments": 1999, "batches": []}))
    with pytest.raises(m.ClassificationError, match="budget"):
        m.run(catalog, usage, limit=3, _transport=lambda *a: pytest.fail("over budget call"))
    before = catalog.read_bytes()
    def fail(*args):
        raise RuntimeError("test-secret private response body")
    with pytest.raises(m.ClassificationError, match="request-failed") as error:
        m.run(catalog, usage, limit=1, _transport=fail)
    assert "test-secret" not in str(error.value)
    assert catalog.read_bytes() == before
    ledger = json.loads(usage.read_text())
    assert ledger["attempted_judgments"] == 2000
    assert ledger["batches"][-1]["status"] == "failed"
    assert "test-secret" not in usage.read_text()
    with pytest.raises(m.ClassificationError, match="budget"):
        m.run(catalog, usage, limit=1, max_judgments=9999, _transport=responder([]))


@pytest.mark.parametrize("kind", ["wrong_model", "missing_answer", "invalid_answer", "bad_usage", "not_json"])
def test_malformed_responses_are_atomic_and_sanitized(tmp_path, monkeypatch, kind):
    monkeypatch.setenv("TYPESAFE_API_KEY", "test-secret")
    catalog, usage = files(tmp_path, 2)
    before = catalog.read_bytes()
    def broken(payload, *args):
        response = responder([])(payload, *args)
        if kind == "wrong_model": response["model"] = "secret-response"
        if kind == "missing_answer": response["answers"].pop(next(iter(response["answers"])))
        if kind == "invalid_answer": response["answers"][next(iter(response["answers"]))]["choice"] = "secret-response"
        if kind == "bad_usage": response["usage"]["input_tokens"] = "secret-response"
        if kind == "not_json": return "secret-response"
        return response
    with pytest.raises(m.ClassificationError) as error:
        m.run(catalog, usage, limit=2, _transport=broken)
    assert "secret-response" not in str(error.value)
    assert catalog.read_bytes() == before
    assert json.loads(usage.read_text())["attempted_judgments"] == 2
    assert "secret-response" not in usage.read_text()


def test_text_packing_full_inputs_and_oversized_failure(tmp_path, monkeypatch):
    monkeypatch.setenv("TYPESAFE_API_KEY", "test-secret")
    catalog, usage = files(tmp_path, 3)
    data = json.loads(catalog.read_text())
    for row in data["items"]: row["text"] = "あ" * 7000
    catalog.write_text(json.dumps(data))
    calls = []
    result = m.run(catalog, usage, limit=3, _transport=responder(calls))
    assert result["classified"] == 3 and len(calls) == 2
    assert len(calls[0]["state"]["items"]["dads-000001"]["text"]) == 7000
    data["items"][0]["text"] = "x" * 16001
    catalog.write_text(json.dumps(data))
    with pytest.raises(m.ClassificationError, match="item-too-large"):
        m.run(catalog, usage, limit=3, _transport=responder(calls))
    assert len(calls) == 2


def test_safe_auth_precedence_and_failure(monkeypatch, tmp_path):
    monkeypatch.setenv("HOME", str(tmp_path))
    default = tmp_path / ".config/typesafe/api-key"
    default.parent.mkdir(parents=True)
    default.write_text("default-secret\n")
    explicit = tmp_path / "key"
    explicit.write_text("explicit-secret\n")
    monkeypatch.delenv("TYPESAFE_API_KEY", raising=False)
    monkeypatch.delenv("TYPESAFE_API_KEY_FILE", raising=False)
    assert m.read_api_key() == "default-secret"
    monkeypatch.setenv("TYPESAFE_API_KEY_FILE", str(explicit))
    assert m.read_api_key() == "explicit-secret"
    monkeypatch.setenv("TYPESAFE_API_KEY", "environment-secret")
    assert m.read_api_key() == "environment-secret"
    monkeypatch.delenv("TYPESAFE_API_KEY")
    monkeypatch.setenv("TYPESAFE_API_KEY_FILE", str(tmp_path / "missing-secret"))
    with pytest.raises(m.ClassificationError, match="auth-unavailable") as error:
        m.read_api_key()
    assert "missing-secret" not in str(error.value)


def test_zero_targets_and_duplicate_ids_fail(tmp_path):
    catalog, usage = files(tmp_path, 0)
    with pytest.raises(m.ClassificationError, match="no-targets"):
        m.run(catalog, usage, limit=50, dry_run=True)
    catalog, usage = files(tmp_path, 2)
    data = json.loads(catalog.read_text())
    data["items"][1]["id"] = data["items"][0]["id"]
    catalog.write_text(json.dumps(data))
    with pytest.raises(m.ClassificationError, match="invalid-catalog"):
        m.run(catalog, usage, limit=50, dry_run=True)


def test_fixed_endpoint_and_redirect_refusal():
    assert m.ENDPOINT == "https://api.typesafe.ai/v1/systemone"
    assert m.NoRedirect().redirect_request(None, None, 302, "redirect", {}, "https://evil.test") is None


def test_live_api_rounded_distribution_preserves_raw_probabilities():
    policy = m.load_policy()
    choice = 'global0-lp0-product0-presentation0'
    other = next(k for k in policy['outcomes'] if k != choice)
    raw = answer(policy,choice,0.8)
    raw['probabilities'] = {k:0.0 for k in policy['outcomes']}
    raw['probabilities'][choice] = 0.81
    raw['probabilities'][other] = 0.18
    record = m.classify_response(item(),raw,policy['model'],policy)
    assert record['raw_answer'] == raw
    raw['probabilities'][other] = 0.01
    with pytest.raises(m.ClassificationError):
        m.classify_response(item(),raw,policy['model'],policy)


def test_live_choice_need_not_be_argmax_and_uncertainty_stays_conservative():
    policy = m.load_policy()
    raw = answer(policy,confidence=0.35)
    keys = list(policy['outcomes'])
    raw['choice'] = keys[0]
    raw['probabilities'] = {k:0.0 for k in keys}
    raw['probabilities'].update({keys[0]:0.38, keys[1]:0.39, keys[2]:0.23})
    record = m.classify_response(item(),raw,policy['model'],policy)
    assert record['raw_answer'] == raw
    assert record['uncertain'] and record['grade'] == 'must'
