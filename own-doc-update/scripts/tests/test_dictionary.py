"""The docs dictionary: one source for official names, fixed values and development terms.

Each red fixture breaks one property (a malformed dictionary, a variant in docs, a wrong value in HTML,
an expired value in HTML); the faithful copy is green. Run through the real CLI, as a caller would.
"""

import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
SKILLS = SCRIPTS.parents[1]
CLI = SCRIPTS / "dictionary.py"
FIXTURES = Path(__file__).resolve().parent / "fixtures" / "dictionary"
REPO = FIXTURES / "repo"
HTML = FIXTURES / "html"


def run(*args):
    return subprocess.run([sys.executable, str(CLI), *map(str, args)], capture_output=True, text=True)


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class DictionaryShapeTest(unittest.TestCase):
    def test_faithful_dictionary_is_valid_and_counts_what_it_read(self):
        r = run("validate", REPO)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("checked: dictionary 5 entries in 1 files", r.stdout)

    def test_malformed_dictionary_is_red(self):
        r = run("validate", FIXTURES / "broken")
        self.assertNotEqual(r.returncode, 0, r.stdout)
        # A term carries meaning and boundary only: a link to a spec is refused, not ignored.
        self.assertIn("term-layer: unknown key 'spec'", r.stdout)
        self.assertIn("term-layer: kind term needs 'boundary'", r.stdout)
        # A value states its expiry (or that it has none); there is no silent default.
        self.assertIn("value-free: kind value needs exactly one of 'expires' or 'evergreen: true'", r.stdout)

    def test_repository_without_dictionary_is_not_green(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "docs").mkdir()
            r = run("validate", tmp)
        self.assertEqual(r.returncode, 2, r.stdout)
        self.assertIn("docs/dictionary/", r.stdout)

    def test_each_shape_rule_is_reported(self):
        d = load("dictionary_for_shape", CLI)
        base = json.loads((REPO / "docs/dictionary/sample.json").read_text(encoding="utf-8"))
        cases = {
            "id prefix": (lambda e: e[0].update(id="name-layer"), "id prefix 'name' does not match kind 'term'"),
            "duplicate id": (lambda e: e[1].update(id="term-layer", kind="term", boundary="x"), "duplicate id term-layer"),
            "variant is a correct text": (lambda e: e[0]["variants"].append("Sample Score"),
                                          "variant 'Sample Score' is the correct text of name-sample-product"),
            "bad date": (lambda e: e[2].update(expires="2027/10/23"), "expires must be YYYY-MM-DD"),
            "empty text": (lambda e: e[1].update(text=""), "name-sample-product: 'text' must be a non-empty string"),
            "bad kind": (lambda e: e[1].update(kind="brand"), "name-sample-product: kind must be one of"),
        }
        for label, (mutate, message) in cases.items():
            with self.subTest(label):
                data = json.loads(json.dumps(base))
                mutate(data["entries"])
                errors = d.shape_errors({"sample.json": data})
                self.assertTrue(any(message in e for e in errors), errors)

    def test_domain_must_match_file_name(self):
        d = load("dictionary_for_domain", CLI)
        data = json.loads((REPO / "docs/dictionary/sample.json").read_text(encoding="utf-8"))
        data["domain"] = "other"
        self.assertTrue(any("domain 'other' must equal the file name 'sample'" in e
                            for e in d.shape_errors({"sample.json": data})))


class DocsVariantTest(unittest.TestCase):
    def test_variant_in_current_docs_is_red_with_place_and_correct_text(self):
        r = run("check-docs", FIXTURES / "variant-docs")
        self.assertEqual(r.returncode, 1, r.stdout)
        self.assertIn("docs/specs/drift.md:3: 「レイヤー」は term-layer のゆれ（正しい表記は「layer」）", r.stdout)

    def test_history_and_inline_code_are_not_read(self):
        r = run("check-docs", FIXTURES / "variant-docs")
        self.assertNotIn("adrs/", r.stdout)
        self.assertNotIn("archive/", r.stdout)
        clean = run("check-docs", REPO)  # clean.md quotes the variant in an inline code span
        self.assertEqual(clean.returncode, 0, clean.stdout)
        self.assertIn("checked: docs 1 files against 5 entries", clean.stdout)

    def test_docs_check_with_no_documents_is_not_green(self):
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copytree(REPO / "docs/dictionary", Path(tmp) / "docs/dictionary")
            r = run("check-docs", tmp)
        self.assertEqual(r.returncode, 2, r.stdout)
        self.assertIn("not checked: no current docs to read", r.stdout)


class HtmlCheckTest(unittest.TestCase):
    DATE = "2026-11-20"

    def html(self, name, date=None):
        return run("check-html", REPO, HTML / name, "--date", date or self.DATE)

    def test_faithful_copy_is_green(self):
        r = self.html("good.html")
        self.assertEqual(r.returncode, 0, r.stdout)
        self.assertIn("checked: html 1 files, 3 anchors against 5 entries", r.stdout)

    def test_wrong_value_is_red(self):
        r = self.html("wrong-value.html")
        self.assertEqual(r.returncode, 1, r.stdout)
        self.assertIn("DICT_VALUE_MISMATCH", r.stdout)
        self.assertIn("value-answer-time", r.stdout)

    def test_expired_value_is_red_only_after_its_expiry(self):
        self.assertEqual(self.html("expired-value.html").returncode, 0)
        r = self.html("expired-value.html", date="2027-01-05")
        self.assertEqual(r.returncode, 1, r.stdout)
        self.assertIn("DICT_EXPIRED", r.stdout)
        self.assertIn("value-campaign-free", r.stdout)

    def test_expired_value_outside_an_anchor_is_red(self):
        with tempfile.TemporaryDirectory() as tmp:
            page = Path(tmp) / "plain.html"
            page.write_text("<p>今すぐ申し込み 今なら無料</p>", encoding="utf-8")
            r = run("check-html", REPO, page, "--date", "2027-01-05")
        self.assertEqual(r.returncode, 1, r.stdout)
        self.assertIn("DICT_EXPIRED", r.stdout)

    def test_variant_in_html_is_red(self):
        r = self.html("variant.html")
        self.assertEqual(r.returncode, 1, r.stdout)
        self.assertIn("DICT_VARIANT", r.stdout)
        self.assertIn("name-sample-product", r.stdout)

    def test_unknown_anchor_is_red(self):
        with tempfile.TemporaryDirectory() as tmp:
            page = Path(tmp) / "unknown.html"
            page.write_text('<p><span data-dict="value-nope">1</span></p>', encoding="utf-8")
            r = run("check-html", REPO, page, "--date", self.DATE)
        self.assertEqual(r.returncode, 1, r.stdout)
        self.assertIn("DICT_UNKNOWN", r.stdout)

    def test_no_html_given_is_not_green(self):
        r = run("check-html", REPO, "--date", self.DATE)
        self.assertEqual(r.returncode, 2, r.stdout + r.stderr)
        self.assertIn("not checked: no HTML files given", r.stdout)


class ValidatorIntegrationTest(unittest.TestCase):
    def setUp(self):
        self.vrd = load("validate_repo_docs_dict", SCRIPTS / "validate_repo_docs.py")

    def test_broken_dictionary_is_a_validator_error(self):
        errors, warnings = [], []
        self.vrd.validate_dictionary(FIXTURES / "broken", errors, warnings)
        self.assertTrue(any("docs/dictionary/broken.json" in e and "unknown key 'spec'" in e for e in errors), errors)

    def test_variant_in_docs_is_a_validator_error_and_coverage_is_counted(self):
        errors, warnings = [], []
        self.vrd.validate_dictionary(FIXTURES / "variant-docs", errors, warnings)
        self.assertTrue(any("docs/specs/drift.md:3" in e for e in errors), errors)
        cov = self.vrd.LAST_COVERAGE["dictionary"]
        self.assertEqual((cov["entries"], cov["files"], cov["docs"]), (5, 1, 1))

    def test_expired_value_is_a_validator_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copytree(REPO / "docs", Path(tmp) / "docs")
            errors, warnings = [], []
            self.vrd.validate_dictionary(Path(tmp), errors, warnings, date="2027-01-05")
        self.assertEqual(errors, [])
        self.assertTrue(any("value-campaign-free" in w and "2026-12-31" in w for w in warnings), warnings)

    def test_repository_without_dictionary_adds_nothing(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "docs").mkdir()
            errors, warnings = [], []
            self.vrd.validate_dictionary(Path(tmp), errors, warnings)
        self.assertEqual((errors, warnings), ([], []))
        self.assertEqual(self.vrd.LAST_COVERAGE["dictionary"]["files"], 0)


class TemplateTest(unittest.TestCase):
    def test_dictionary_template_is_a_valid_dictionary(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "docs/dictionary/example.json"
            target.parent.mkdir(parents=True)
            shutil.copy(SCRIPTS.parent / "references/dictionary.template.json", target)
            r = run("validate", tmp)
        self.assertEqual(r.returncode, 0, r.stdout)

    def test_spec_templates_have_no_domain_language_table(self):
        # The dictionary owns terms; a spec that grows its own table makes two places to write.
        for path in (SKILLS / "own-spec-grill/references/spec.template.md",
                     SCRIPTS.parent / "references/spec.template.md"):
            with self.subTest(str(path)):
                self.assertNotIn("## Domain Language", path.read_text(encoding="utf-8"))

    def test_spec_grill_reads_the_dictionary_first(self):
        text = (SKILLS / "own-spec-grill/SKILL.md").read_text(encoding="utf-8")
        self.assertIn("docs/dictionary/", text)
        self.assertNotIn("`Domain Language` section", text)


if __name__ == "__main__":
    unittest.main()


class NewSpecTermTableTest(unittest.TestCase):
    """A spec created on or after the dictionary rollout has no term table: terms live in docs/dictionary/."""

    def setUp(self):
        self.vrd = load("validate_repo_docs_terms", SCRIPTS / "validate_repo_docs.py")
        self.tmp = Path(tempfile.mkdtemp())
        (self.tmp / "docs/specs").mkdir(parents=True)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def spec(self, name, created, heading):
        (self.tmp / "docs/specs" / name).write_text(
            f"---\nid: SPEC-{name[:-3]}\nstatus: draft\ncreated_at: {created}\nupdated_at: {created}\n---\n\n"
            f"# T\n\n## Purpose\n\n## {heading}\n\n| Term | Meaning |\n| --- | --- |\n", encoding="utf-8")

    def errors(self):
        errors, warnings = [], []
        self.vrd.validate_spec_placement(self.tmp, errors, warnings)
        return errors

    def test_new_spec_with_a_term_table_is_red_and_names_the_file(self):
        for heading in ("Domain Language", "用語", "用語集", "Glossary"):
            with self.subTest(heading):
                self.spec("new.md", "2026-11-21", heading)
                self.assertTrue(any(e.startswith("docs/specs/new.md:") and "docs/dictionary/" in e for e in self.errors()),
                                self.errors())

    def test_rollout_day_itself_counts_as_new(self):
        self.spec("day.md", self.vrd.DICTIONARY_ROLLOUT_DATE, "Domain Language")
        self.assertTrue(any("docs/specs/day.md:" in e for e in self.errors()))

    def test_older_spec_keeps_its_table(self):
        self.spec("old.md", "2026-11-19", "Domain Language")
        self.assertEqual(self.errors(), [])

    def test_a_heading_that_only_mentions_terms_is_not_a_table(self):
        self.spec("prose.md", "2026-11-21", "用語の方針")
        self.assertEqual(self.errors(), [])
