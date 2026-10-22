#!/usr/bin/env python3
"""Bounded, resumable Jev catalog classification; never invoked by L0/L1/L2.

Each failed request consumes its reserved questions. There are no automatic
retries. Token-price arithmetic is usage-derived, not an invoice or money cap.
"""
from __future__ import annotations

import argparse
import copy
import fcntl
import hashlib
import json
import math
import os
from pathlib import Path
import re
import sys
import tempfile
import time
import urllib.request

ENDPOINT = "https://api.typesafe.ai/v1/systemone"
POLICY = Path(__file__).resolve().parents[1] / "references/classification-policy.json"
PRESETS = ["lp", "product-ui", "presentation"]
MAX_JUDGMENTS = 2000
MAX_BATCH = 10
MAX_SOURCE_CHARS = 16000


class ClassificationError(Exception):
    """Only fixed, sanitized error codes may be placed in this exception."""


def _json(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)


def _hash(value):
    return hashlib.sha256(_json(value).encode("utf-8")).hexdigest()


def _load(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError, UnicodeError):
        raise ClassificationError("invalid-json-file") from None


def _save(path, value):
    path = Path(path)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(value, stream, ensure_ascii=False, indent=2, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except (OSError, TypeError, ValueError):
        raise ClassificationError("persistence-failed") from None
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def load_policy(path=POLICY, model=None):
    policy = _load(path)
    try:
        if model is not None:
            policy["model"] = model
        outcomes = policy["outcomes"]
        if (policy["schema_version"] != 1 or not re.fullmatch(r"jev-[0-9]+\.[0-9]+\.[0-9]+", policy["model"])
                or not _probability(policy["threshold"]) or len(outcomes) != 16):
            raise ValueError
        combinations = set()
        for key, outcome in outcomes.items():
            if (not isinstance(key, str) or type(outcome["global_must"]) is not bool
                    or not isinstance(outcome["description"], str)
                    or not isinstance(outcome["preset_must"], list)
                    or len(outcome["preset_must"]) != len(set(outcome["preset_must"]))
                    or any(p not in PRESETS for p in outcome["preset_must"])):
                raise ValueError
            combinations.add((outcome["global_must"], tuple(sorted(outcome["preset_must"]))))
        if len(combinations) != 16:
            raise ValueError
        for name in ("input_price_usd_per_million", "output_price_usd_per_million"):
            if not _number(policy[name]) or policy[name] < 0:
                raise ValueError
        if not isinstance(policy["rules"], dict):
            raise ValueError
    except (KeyError, TypeError, ValueError):
        raise ClassificationError("invalid-policy") from None
    return policy


def _number(value):
    return type(value) in (int, float) and math.isfinite(value)


def _probability(value):
    return _number(value) and 0 <= value <= 1


def classification_input(item):
    try:
        result = {key: item[key] for key in ("document", "title", "headings", "text", "source_url")}
        if (any(not isinstance(result[key], str) for key in ("document", "title", "text"))
                or not isinstance(result["headings"], list)
                or any(not isinstance(title, str) for title in result["headings"])
                or (result["source_url"] is not None and not isinstance(result["source_url"], str))):
            raise ValueError
    except (KeyError, TypeError, ValueError):
        raise ClassificationError("invalid-catalog") from None
    return result


def classify_response(item, answer, model, policy):
    """Validate a Choice and retain only its four documented answer fields."""
    # Live Jev responses round each of the 16 probabilities to two decimals.
    # Preserve the raw values; only allow the aggregate quantization error.
    precision = _load(POLICY.parent / 'api-contract.json')['probability_decimal_places']
    rounding_error = len(policy['outcomes']) * 0.5 * 10 ** (-precision) + 1e-9
    try:
        choice, confidence, probabilities = answer["choice"], answer["confidence"], answer["probabilities"]
        if (model != policy["model"] or answer["type"] != "choice" or choice not in policy["outcomes"]
                or not _probability(confidence) or not isinstance(probabilities, dict)
                or set(probabilities) != set(policy["outcomes"])
                or any(not _probability(value) for value in probabilities.values())
                or abs(sum(probabilities.values()) - 1) > rounding_error):
            raise ValueError
    except (KeyError, TypeError, ValueError):
        raise ClassificationError("invalid-answer") from None
    outcome = policy["outcomes"][choice]
    uncertain = confidence < policy["threshold"]
    raw = {"type": "choice", "choice": choice, "probabilities": copy.deepcopy(probabilities), "confidence": confidence}
    return {"grade": "must" if uncertain or outcome["global_must"] else "nice-to-have",
            "preset_must": PRESETS[:] if uncertain else outcome["preset_must"][:],
            "model": model, "threshold": policy["threshold"],
            "input_hash": _hash(classification_input(item)), "policy_hash": _hash(policy),
            "raw_answer": raw, "uncertain": uncertain}


def read_api_key():
    """Match vibe-guard auth precedence; read existing state without changing it."""
    key = os.environ.get("TYPESAFE_API_KEY")
    if not key:
        location = os.environ.get("TYPESAFE_API_KEY_FILE") or Path.home() / ".config/typesafe/api-key"
        try:
            key = Path(location).read_text(encoding="utf-8").strip()
        except (OSError, UnicodeError):
            raise ClassificationError("auth-unavailable") from None
    if not key or "\n" in key or "\r" in key:
        raise ClassificationError("auth-unavailable")
    return key


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def _post(payload, key, timeout):
    request = urllib.request.Request(ENDPOINT, data=_json(payload).encode("utf-8"), method="POST",
                                    headers={"Authorization": "Bearer " + key, "Content-Type": "application/json"})
    with urllib.request.build_opener(NoRedirect()).open(request, timeout=timeout) as response:
        if response.status != 200:
            raise ClassificationError("request-failed")
        body = response.read(4_000_001)
        if len(body) > 4_000_000:
            raise ClassificationError("invalid-response")
        return json.loads(body)


def _payload(batch, policy):
    return {"model": policy["model"],
            "state": {"items": {item["id"]: classification_input(item) for item in batch}},
            "questions": {item["id"]: {
                "type": "choice",
                "instructions": {"target": f"Classify only `items.{item['id']}`; ignore all other state items.",
                                 "decision": "Select the one complete combination of global base must and additive preset scopes that this item satisfies.",
                                 "rules": policy["rules"]},
                "criteria": {key: outcome["description"] for key, outcome in policy["outcomes"].items()}}
                for item in batch}}


def _pending(items, policy):
    pending = []
    for item in items:
        record = item.get("classification")
        try:
            expected = classify_response(item, record["raw_answer"], record["model"], policy)
            current = all(record.get(key) == value for key, value in expected.items())
        except (ClassificationError, KeyError, TypeError):
            current = False
        if not current:
            pending.append(item)
    return pending


def _batches(items):
    batches, current, size = [], [], 0
    for item in items:
        length = len(_json(classification_input(item)))
        if length > MAX_SOURCE_CHARS:
            raise ClassificationError("item-too-large")
        if current and (len(current) == MAX_BATCH or size + length > MAX_SOURCE_CHARS):
            batches.append(current)
            current, size = [], 0
        current.append(item)
        size += length
    if current:
        batches.append(current)
    return batches


def _usage_ledger(path):
    ledger = _load(path) if Path(path).exists() else {"schema_version": 1, "attempted_judgments": 0, "batches": []}
    try:
        if (ledger["schema_version"] != 1 or type(ledger["attempted_judgments"]) is not int
                or ledger["attempted_judgments"] < 0 or not isinstance(ledger["batches"], list)):
            raise ValueError
        # The top-level count may include legacy attempts with unavailable telemetry.
        count = 0
        for batch in ledger["batches"]:
            if type(batch["questions"]) is not int or batch["questions"] <= 0:
                raise ValueError
            count += batch["questions"]
        if count > ledger["attempted_judgments"]:
            raise ValueError
    except (KeyError, TypeError, ValueError):
        raise ClassificationError("invalid-usage-ledger") from None
    return ledger


def run(catalog_path, usage_path, limit, max_judgments=2000, dry_run=False, model=None, policy_path=POLICY, timeout=30, _transport=None):
    """Run up to limit pending judgments. _transport is test injection, not CLI.

    Every successful batch is durable. A later failure preserves prior successes
    and consumes the failed batch reservation without modifying its item records.
    """
    if (type(limit) is not int or limit <= 0 or type(max_judgments) is not int or max_judgments <= 0
            or not _number(timeout) or timeout <= 0):
        raise ClassificationError("invalid-limits")
    catalog_path, usage_path = Path(catalog_path), Path(usage_path)
    if catalog_path.resolve() == usage_path.resolve():
        raise ClassificationError("invalid-paths")
    policy = load_policy(policy_path, model)
    catalog = _load(catalog_path)
    try:
        items = catalog["items"]
        if catalog["schema_version"] != 1 or not isinstance(items, list):
            raise ValueError
        if not items:
            raise ClassificationError("no-targets")
        ids = []
        for item in items:
            if not re.fullmatch(r"dads-[0-9]{6,}", item["id"]):
                raise ValueError
            ids.append(item["id"])
            classification_input(item)
        if len(ids) != len(set(ids)):
            raise ValueError
    except (KeyError, TypeError, ValueError):
        raise ClassificationError("invalid-catalog") from None
    pending = _pending(items, policy)
    selected = pending[:limit]
    batches = _batches(selected)
    ledger = _usage_ledger(usage_path)
    ceiling = min(MAX_JUDGMENTS, max_judgments)
    if ledger["attempted_judgments"] + len(selected) > ceiling:
        raise ClassificationError("budget-exceeded")
    result = {"status": "dry-run" if dry_run else ("already-complete" if not pending else "classified"),
              "pending": len(pending), "selected": len(selected), "classified": 0,
              "requests": len(batches), "remaining": len(pending),
              "attempted_judgments": ledger["attempted_judgments"], "max_judgments": ceiling,
              "input_tokens": 0, "output_tokens": 0, "usage_derived_usd": 0.0}
    if dry_run or not pending:
        return result
    key = read_api_key()
    try:
        lock = usage_path.with_suffix(usage_path.suffix + ".lock").open("a")
    except OSError:
        raise ClassificationError("persistence-failed") from None
    with lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        ledger = _usage_ledger(usage_path)
        if ledger["attempted_judgments"] + len(selected) > ceiling:
            raise ClassificationError("budget-exceeded")
        # Re-read under the process lock: another runner may have completed our snapshot.
        catalog = _load(catalog_path)
        selected = _pending(catalog['items'], policy)[:limit]
        batches = _batches(selected)
        if ledger["attempted_judgments"] + len(selected) > ceiling:
            raise ClassificationError("budget-exceeded")
        result.update(selected=len(selected), requests=len(batches), remaining=len(_pending(catalog['items'], policy)))
        if not selected:
            result.update(status='already-complete', attempted_judgments=ledger['attempted_judgments'])
            return result
        for batch in batches:
            entry = {"model": policy["model"], "questions": len(batch), "status": "reserved",
                     "duration_seconds": None, "input_tokens": None, "output_tokens": None,
                     "usage_derived_usd": None}
            ledger["attempted_judgments"] += len(batch)
            ledger["batches"].append(entry)
            _save(usage_path, ledger)  # Reservation is durable before network.
            started = time.monotonic()
            response = None
            try:
                response = (_transport or _post)(_payload(batch, policy), key, timeout)
                if not isinstance(response, dict) or response.get("model") != policy["model"]:
                    raise ClassificationError("invalid-response")
                answers, usage = response["answers"], response["usage"]
                if not isinstance(answers, dict) or set(answers) != {item["id"] for item in batch}:
                    raise ClassificationError("invalid-response")
                for name in ("input_tokens", "output_tokens"):
                    if type(usage[name]) is not int or usage[name] < 0:
                        raise ClassificationError("invalid-response")
                records = [classify_response(item, answers[item["id"]], response["model"], policy) for item in batch]
                cost = (usage["input_tokens"] * policy["input_price_usd_per_million"]
                        + usage["output_tokens"] * policy["output_price_usd_per_million"]) / 1_000_000
            except Exception as error:
                entry["failure_signature"] = ('http-' + str(error.code) if isinstance(error, urllib.error.HTTPError)
                    else type(error).__name__)
                if isinstance(response, dict):
                    diagnostic = []
                    for ident, answer in response.get('answers', {}).items():
                        if not isinstance(answer, dict):
                            continue
                        probabilities = answer.get('probabilities', {})
                        if isinstance(probabilities, dict) and all(_number(v) for v in probabilities.values()):
                            diagnostic.append({'id': ident if ident in {x['id'] for x in batch} else '<unknown>',
                                'type': answer.get('type') if answer.get('type') in ('choice','score','noul') else '<unknown>',
                                'choice_known': answer.get('choice') in policy['outcomes'],
                                'confidence': answer.get('confidence') if _number(answer.get('confidence')) else None,
                                'probability_count': len(probabilities), 'probability_sum': sum(probabilities.values()),
                                'chosen_probability': probabilities.get(answer.get('choice')),
                                'max_probability': max(probabilities.values(),default=None)})
                    entry['failure_diagnostics'] = diagnostic
                    failed_usage = response.get('usage', {})
                    if all(type(failed_usage.get(k)) is int and failed_usage[k] >= 0 for k in ('input_tokens','output_tokens')):
                        entry.update(input_tokens=failed_usage['input_tokens'], output_tokens=failed_usage['output_tokens'],
                            usage_derived_usd=(failed_usage['input_tokens'] * policy['input_price_usd_per_million']
                                + failed_usage['output_tokens'] * policy['output_price_usd_per_million']) / 1_000_000)
                entry["status"] = "failed"
                entry["duration_seconds"] = round(time.monotonic() - started, 6)
                _save(usage_path, ledger)
                raise ClassificationError("request-failed") from None
            entry.update(status="succeeded", duration_seconds=round(time.monotonic() - started, 6),
                         input_tokens=usage["input_tokens"], output_tokens=usage["output_tokens"], usage_derived_usd=cost)
            # Persist known usage first: a disk failure must not erase charged attempts.
            _save(usage_path, ledger)
            for item, record in zip(batch, records):
                item["classification"] = record
            _save(catalog_path, catalog)
            result["classified"] += len(batch)
            result["input_tokens"] += usage["input_tokens"]
            result["output_tokens"] += usage["output_tokens"]
            result["usage_derived_usd"] += cost
        result["attempted_judgments"] = ledger["attempted_judgments"]
        result["remaining"] -= result["classified"]
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--catalog", required=True)
    parser.add_argument("--usage-ledger", required=True)
    parser.add_argument("--limit", type=int, required=True)
    parser.add_argument("--max-judgments", type=int, default=2000)
    parser.add_argument("--model", default=None, help="pinned Jev version; aliases are rejected")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)
    try:
        result = run(args.catalog, args.usage_ledger, args.limit, args.max_judgments, args.dry_run, args.model)
    except ClassificationError as error:
        print("classification error: " + str(error), file=sys.stderr)
        return 2
    except Exception:
        print("classification error: local-failure", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
