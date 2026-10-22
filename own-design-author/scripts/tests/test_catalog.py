"""Catalog extraction and stable identity contract; no network/model calls."""
import copy
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import zipfile

import pytest

SCRIPT = Path(__file__).parents[1] / "catalog.py"
spec = importlib.util.spec_from_file_location("design_catalog", SCRIPT)
catalog_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(catalog_module)
generate_catalog = catalog_module.generate_catalog


def archive(tmp_path, body, name="source.zip"):
    path = tmp_path / name
    with zipfile.ZipFile(path, "w") as out:
        out.writestr("foundations/color/index.md", "---\ntitle: Color\nsource_url: https://example.test/color/\n---\n" + body)
        out.writestr("foundations/color/changelog.md", "## Not a catalog item")
        out.writestr("guidance/example/index.md", "## Not a catalog item")
    return path


def upstream_items(catalog):
    return [item for item in catalog["items"] if not item["document"].startswith("supplements/")]


def test_extracts_headings_sections_and_ignores_fences(tmp_path):
    source = archive(tmp_path, "## Colors\nintro\n### Primary\nred\n#### Detail\nmore\n```md\n## False\n```\n~~~\n### Also false\n~~~\n### Primary\nblue\n## Next\nend\n")
    result, ledger, delta = generate_catalog(source, "2026-10-14")
    rows = upstream_items(result)
    assert [row["title"] for row in rows] == ["Colors", "Primary", "Primary", "Next"]
    assert rows[1]["headings"] == ["Colors", "Primary"]
    assert len({row["match_key"] for row in rows}) == 4
    assert "#### Detail" in rows[1]["text"]
    assert "blue" not in rows[1]["text"]
    assert result["source"]["documents"] == {"foundations": 1, "components": 0}
    assert result["source"]["archive_sha256"]
    assert result["edited"] is True and result["attribution"]
    assert all(row["classification"] is None for row in result["items"])
    assert len(result["items"]) == 10
    assert len(delta["new_ids"]) == 10


def test_ids_survive_order_change_and_text_edit(tmp_path):
    first, ledger, _ = generate_catalog(archive(tmp_path, "## A\nold\n## B\nbody\n"), "1")
    saved = copy.deepcopy(ledger)
    second, updated, delta = generate_catalog(archive(tmp_path, "## B\nbody\n## A\nnew\n", "second.zip"), "2", ledger=ledger, previous=first)
    assert {x["match_key"]: x["id"] for x in first["items"]} == {x["match_key"]: x["id"] for x in second["items"]}
    assert ledger == saved  # Caller data is not mutated.
    assert delta["new_ids"] == []
    assert delta["disappeared_keys"] == []


def test_explicit_rename_reuses_id_and_reports_change(tmp_path):
    first, ledger, _ = generate_catalog(archive(tmp_path, "## Old\ntext\n"), "1")
    old = upstream_items(first)[0]
    probe, _, _ = generate_catalog(archive(tmp_path, "## New\ntext\n", "new.zip"), "2")
    new_key = upstream_items(probe)[0]["match_key"]
    result, ledger, delta = generate_catalog(tmp_path / "new.zip", "2", ledger=ledger, previous=first, renames={old["match_key"]: new_key})
    assert upstream_items(result)[0]["id"] == old["id"]
    assert delta["changed_keys"] == [{"id": old["id"], "from": old["match_key"], "to": new_key}]
    assert delta["disappeared_keys"] == []
    assert delta["new_ids"] == []


def test_disappearance_retires_but_never_reuses_id(tmp_path):
    first, ledger, _ = generate_catalog(archive(tmp_path, "## Removed\ntext\n"), "1")
    old = upstream_items(first)[0]
    second, ledger, delta = generate_catalog(archive(tmp_path, "## Added\ntext\n", "next.zip"), "2", ledger=ledger, previous=first)
    assert upstream_items(second)[0]["id"] != old["id"]
    assert old["match_key"] in delta["disappeared_keys"]
    assert next(row for row in ledger["entries"] if row["id"] == old["id"])["retired"] is True
    third, ledger, _ = generate_catalog(tmp_path / "source.zip", "3", ledger=ledger, previous=second)
    assert upstream_items(third)[0]["id"] == old["id"]


def test_rejects_unknown_ambiguous_and_existing_target_renames(tmp_path):
    first, ledger, _ = generate_catalog(archive(tmp_path, "## A\na\n## B\nb\n"), "1")
    a, b = upstream_items(first)
    next_zip = archive(tmp_path, "## C\nc\n", "next.zip")
    probe, _, _ = generate_catalog(next_zip, "2")
    key = upstream_items(probe)[0]["match_key"]
    for renames in ({"unknown": key}, {a["match_key"]: "unknown"}, {a["match_key"]: key, b["match_key"]: key}, {a["match_key"]: b["match_key"]}):
        with pytest.raises(ValueError):
            generate_catalog(next_zip, "2", ledger=ledger, previous=first, renames=renames)


@pytest.mark.parametrize("mutation", ["duplicate_id", "duplicate_key", "next_id_reuse"])
def test_corrupt_ledger_fails_closed(tmp_path, mutation):
    source = archive(tmp_path, "## A\na\n")
    _, ledger, _ = generate_catalog(source, "1")
    if mutation == "duplicate_id":
        ledger["entries"][1]["id"] = ledger["entries"][0]["id"]
    elif mutation == "duplicate_key":
        ledger["entries"][1]["match_key"] = ledger["entries"][0]["match_key"]
    else:
        ledger["next_id"] = 1
    with pytest.raises(ValueError):
        generate_catalog(source, "2", ledger=ledger)


def test_cli_writes_catalog_ledger_and_delta(tmp_path):
    source = archive(tmp_path, "## A\na\n")
    output, ledger = tmp_path / "catalog.json", tmp_path / "ledger.json"
    run = subprocess.run([sys.executable, str(SCRIPT), "generate", "--zip", str(source), "--version", "1", "--output", str(output), "--ledger", str(ledger)], capture_output=True, text=True)
    assert run.returncode == 0, run.stderr
    assert json.loads(output.read_text())["version"] == "1"
    assert json.loads(ledger.read_text())["next_id"] == 8
    assert len(json.loads(run.stdout)["new_ids"]) == 7


def test_real_archive_has_all_560_headings():
    # Opt in to the locally supplied upstream archive without embedding its path.
    source = os.environ.get("DADS_TEST_ARCHIVE")
    if not source:
        pytest.skip("set DADS_TEST_ARCHIVE to exercise the local upstream ZIP")
    result, ledger, _ = generate_catalog(source, "2026-10-14")
    assert result["source"]["documents"] == {"foundations": 8, "components": 49}
    expected = []
    with zipfile.ZipFile(source) as z:
        for name in sorted(z.namelist()):
            if re.fullmatch(r"(foundations|components)/[^/]+/index\.md", name):
                expected.extend((name, line.split(" ", 1)[1].strip()) for line in z.read(name).decode("utf-8-sig").splitlines() if re.match(r"^#{2,3} ", line))
    actual = [(row["document"], row["title"]) for row in upstream_items(result)]
    assert len(expected) == 560
    assert actual == expected
    assert len(result["items"]) == 566
    assert len(ledger["entries"]) == 566


def test_rename_history_is_reserved_and_supports_explicit_reversal(tmp_path):
    old_zip = archive(tmp_path, "## Old\ntext\n")
    new_zip = archive(tmp_path, "## New\ntext\n", "new.zip")
    first, ledger, _ = generate_catalog(old_zip, "1")
    probe, _, _ = generate_catalog(new_zip, "2")
    old, new = upstream_items(first)[0], upstream_items(probe)[0]
    second, ledger, _ = generate_catalog(new_zip, "2", ledger=ledger, previous=first, renames={old["match_key"]: new["match_key"]})
    entry = next(row for row in ledger["entries"] if row["id"] == old["id"])
    assert entry["previous_match_keys"] == [old["match_key"]]
    with pytest.raises(ValueError, match="historical"):
        generate_catalog(old_zip, "3", ledger=ledger, previous=second)
    third, ledger, delta = generate_catalog(old_zip, "3", ledger=ledger, previous=second, renames={new["match_key"]: old["match_key"]})
    assert upstream_items(third)[0]["id"] == old["id"]
    assert delta["new_ids"] == []


def test_unchanged_upstream_items_reuse_recorded_classifications(tmp_path):
    source = archive(tmp_path, '## A\npolicy\n')
    first, ledger, _ = generate_catalog(source, '1')
    first['items'][0]['classification'] = {'recorded':'model-result'}
    second, _, _ = generate_catalog(source, '2', ledger=ledger, previous=first)
    assert second['items'][0]['classification'] == {'recorded':'model-result'}
    changed = archive(tmp_path, '## A\nchanged policy\n', 'changed.zip')
    third, _, _ = generate_catalog(changed, '3', ledger=ledger, previous=first)
    assert third['items'][0]['classification'] is None
