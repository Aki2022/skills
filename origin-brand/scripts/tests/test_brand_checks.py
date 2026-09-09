"""check_brand_metadata.py / check_tokens_parity.py のテスト。

origin-brand/plugins/*/plugin.yaml の「等級（tier）と取得メタ（source）」の形と、
倉庫 my_company の pptx 媒体プロファイル ↔ origin-pptx/style-guide/tokens.json の
同値を機械で見張るのがこの2本。TDD として、まず陽性対照（わざと壊して赤になることを
実測できるケース）を含めてテストを先に書く。

一時ファイルは必ず TemporaryDirectory を使い、リポジトリ内の正典ファイルは一切
書き換えない（正典を書き換えて戻す方式は、途中で失敗すると壊れたまま残るため禁止）。
"""
from __future__ import annotations

import importlib.util
import io
import json
import re
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = SCRIPTS_DIR.parents[1]  # .../origin-brand/scripts -> parents[1] = リポジトリルート

METADATA_SCRIPT = SCRIPTS_DIR / "check_brand_metadata.py"
PARITY_SCRIPT = SCRIPTS_DIR / "check_tokens_parity.py"

CANON_PLUGINS_DIR = REPO_ROOT / "origin-brand" / "plugins"
CANON_MY_COMPANY_YAML = CANON_PLUGINS_DIR / "my_company" / "plugin.yaml"
CANON_DIGITAL_AGENCY_YAML = CANON_PLUGINS_DIR / "digital-agency" / "plugin.yaml"

CANON_WAREHOUSE_COLOR = CANON_PLUGINS_DIR / "my_company" / "tokens" / "color.json"
CANON_WAREHOUSE_TYPOGRAPHY = CANON_PLUGINS_DIR / "my_company" / "tokens" / "typography.json"
CANON_PPTX_TOKENS = REPO_ROOT / "origin-pptx" / "style-guide" / "tokens.json"

METADATA_SCOPE_LINE = (
    "scope: この検査は等級（tier）と取得メタ（source）の形だけを見る。"
    "上流の内容が実際に新しいかは見ない。"
)
PARITY_SCOPE_LINE = (
    "scope: この検査は上表18項目の値が一致するかだけを見る。"
    "web プロファイルと dataViz は比較しない（媒体が違うので同値である必要がない）。"
)

EXPECTED_PARITY_COUNT = 18


def _load(path: Path):
    spec = importlib.util.spec_from_file_location(path.stem, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _run(module, argv):
    buf = io.StringIO()
    with redirect_stdout(buf):
        code = module.main(argv)
    return code, buf.getvalue()


def _write_single_plugin(tmp: str, box: str, yaml_text: str) -> Path:
    """一時ディレクトリに <box>/plugin.yaml を1件だけ置いた plugins-dir を作る。"""
    plugins_dir = Path(tmp)
    box_dir = plugins_dir / box
    box_dir.mkdir(parents=True)
    (box_dir / "plugin.yaml").write_text(yaml_text, encoding="utf-8")
    return plugins_dir


class IndependentTargetExistenceTest(unittest.TestCase):
    """被検査物（正典ファイル群）が実在し、期待する形を持つことをスクリプトを介さず確かめる。

    これが無いと、スクリプト側にバグがあって常に対象0件を拾っていても
    テストは「一致した／合格した」side の緑を出し続けてしまう。
    """

    def test_four_plugin_yaml_boxes_exist(self):
        plugin_yaml_files = sorted(CANON_PLUGINS_DIR.glob("*/plugin.yaml"))
        self.assertEqual(
            len(plugin_yaml_files), 4, [p.parent.name for p in plugin_yaml_files]
        )
        for p in plugin_yaml_files:
            text = p.read_text(encoding="utf-8")
            self.assertIn("tier:", text, p)
            self.assertIn("source:", text, p)

    def test_warehouse_and_pptx_token_files_have_expected_top_keys(self):
        color = json.loads(CANON_WAREHOUSE_COLOR.read_text(encoding="utf-8"))
        self.assertIn("mediaProfiles", color)
        self.assertIn("pptx", color["mediaProfiles"])
        self.assertIn("shared", color)

        typography = json.loads(CANON_WAREHOUSE_TYPOGRAPHY.read_text(encoding="utf-8"))
        self.assertIn("mediaProfiles", typography)
        self.assertIn("pptx", typography["mediaProfiles"])

        pptx_tokens = json.loads(CANON_PPTX_TOKENS.read_text(encoding="utf-8"))
        self.assertIn("color", pptx_tokens)
        self.assertIn("typography", pptx_tokens)


class CheckBrandMetadataTest(unittest.TestCase):
    def test_a1_canonical_four_boxes_pass(self):
        """陰性対照: 正典の4箱で exit 0。"""
        module = _load(METADATA_SCRIPT)
        code, out = _run(module, ["--plugins-dir", str(CANON_PLUGINS_DIR)])
        self.assertEqual(code, 0, out)

    def test_a2_licensed_tier_is_violation(self):
        """陽性対照: tier: licensed（廃止済み）に書き換えると exit 1。"""
        module = _load(METADATA_SCRIPT)
        original = CANON_MY_COMPANY_YAML.read_text(encoding="utf-8")
        mutated = original.replace("tier: owned", "tier: licensed", 1)
        self.assertNotEqual(original, mutated)
        self.assertIn("tier: licensed", mutated)
        with TemporaryDirectory() as tmp:
            plugins_dir = _write_single_plugin(tmp, "my_company", mutated)
            code, out = _run(module, ["--plugins-dir", str(plugins_dir)])
        self.assertEqual(code, 1, out)

    def test_a3_missing_fetched_at_is_violation(self):
        """陽性対照: source.fetched_at を削ると exit 1（必須3キーの欠落）。"""
        module = _load(METADATA_SCRIPT)
        original = CANON_MY_COMPANY_YAML.read_text(encoding="utf-8")
        lines = [
            line
            for line in original.splitlines(keepends=True)
            if not line.strip().startswith("fetched_at:")
        ]
        mutated = "".join(lines)
        self.assertNotIn("fetched_at:", mutated)
        with TemporaryDirectory() as tmp:
            plugins_dir = _write_single_plugin(tmp, "my_company", mutated)
            code, out = _run(module, ["--plugins-dir", str(plugins_dir)])
        self.assertEqual(code, 1, out)

    def test_a4_owned_with_url_upstream_is_violation(self):
        """陽性対照: tier: owned のまま upstream を URL にすると exit 1。"""
        module = _load(METADATA_SCRIPT)
        original = CANON_MY_COMPANY_YAML.read_text(encoding="utf-8")
        mutated = original.replace(
            "upstream: none", "upstream: https://example.com/not-really-owned", 1
        )
        self.assertNotEqual(original, mutated)
        self.assertIn("tier: owned", mutated)
        with TemporaryDirectory() as tmp:
            plugins_dir = _write_single_plugin(tmp, "my_company", mutated)
            code, out = _run(module, ["--plugins-dir", str(plugins_dir)])
        self.assertEqual(code, 1, out)

    def test_a5_empty_dir_is_exit_2(self):
        """陽性対照: 箱が0件のディレクトリは exit 2（0件だから合格、にしない）。"""
        module = _load(METADATA_SCRIPT)
        with TemporaryDirectory() as tmp:
            code, out = _run(module, ["--plugins-dir", tmp])
        self.assertEqual(code, 2, out)

    def test_a6_non_iso_fetched_at_is_violation(self):
        """陽性対照: fetched_at が YYYY-MM-DD でない（2026/07/04）と exit 1。"""
        module = _load(METADATA_SCRIPT)
        original = CANON_DIGITAL_AGENCY_YAML.read_text(encoding="utf-8")
        mutated = original.replace("fetched_at: 2026-07-04", "fetched_at: 2026/07/04", 1)
        self.assertNotEqual(original, mutated)
        self.assertIn("2026/07/04", mutated)
        with TemporaryDirectory() as tmp:
            plugins_dir = _write_single_plugin(tmp, "digital-agency", mutated)
            code, out = _run(module, ["--plugins-dir", str(plugins_dir)])
        self.assertEqual(code, 1, out)

    def test_scope_line_is_first(self):
        module = _load(METADATA_SCRIPT)
        _, out = _run(module, ["--plugins-dir", str(CANON_PLUGINS_DIR)])
        self.assertEqual(out.splitlines()[0], METADATA_SCOPE_LINE)

    def test_default_plugins_dir_points_at_canonical_dir(self):
        module = _load(METADATA_SCRIPT)
        code, out = _run(module, [])
        self.assertEqual(code, 0, out)


class CheckTokensParityTest(unittest.TestCase):
    def _default_argv(self, color=None, typography=None, pptx_tokens=None):
        return [
            "--warehouse-color", str(color or CANON_WAREHOUSE_COLOR),
            "--warehouse-typography", str(typography or CANON_WAREHOUSE_TYPOGRAPHY),
            "--pptx-tokens", str(pptx_tokens or CANON_PPTX_TOKENS),
        ]

    def test_b1_canonical_files_match_and_compare_18_items(self):
        """陰性対照: 正典の2ファイルで exit 0 かつ比較件数が18。"""
        module = _load(PARITY_SCRIPT)
        code, out = _run(module, self._default_argv())
        self.assertEqual(code, 0, out)
        m = re.search(r"compared:\s*(\d+)\s*/\s*(\d+)", out)
        self.assertIsNotNone(m, out)
        self.assertEqual(int(m.group(1)), EXPECTED_PARITY_COUNT, out)
        self.assertEqual(int(m.group(2)), EXPECTED_PARITY_COUNT, out)

    def test_b2_changed_warehouse_body_color_is_mismatch(self):
        """陽性対照: 倉庫側 mediaProfiles.pptx.text.body を別の色に変えると exit 1。"""
        module = _load(PARITY_SCRIPT)
        color = json.loads(CANON_WAREHOUSE_COLOR.read_text(encoding="utf-8"))
        original_value = color["mediaProfiles"]["pptx"]["text"]["body"]
        color["mediaProfiles"]["pptx"]["text"]["body"] = "#000000"
        self.assertNotEqual(original_value, "#000000")
        with TemporaryDirectory() as tmp:
            color_path = Path(tmp) / "color.json"
            color_path.write_text(json.dumps(color, ensure_ascii=False, indent=2), encoding="utf-8")
            code, out = _run(module, self._default_argv(color=color_path))
        self.assertEqual(code, 1, out)

    def test_b3_missing_warehouse_key_is_exit_2_not_mismatch(self):
        """陽性対照: 倉庫側から shared.tone.positive を削ると exit 2（欠落は不一致と別）。"""
        module = _load(PARITY_SCRIPT)
        color = json.loads(CANON_WAREHOUSE_COLOR.read_text(encoding="utf-8"))
        del color["shared"]["tone"]["positive"]
        with TemporaryDirectory() as tmp:
            color_path = Path(tmp) / "color.json"
            color_path.write_text(json.dumps(color, ensure_ascii=False, indent=2), encoding="utf-8")
            code, out = _run(module, self._default_argv(color=color_path))
        self.assertEqual(code, 2, out)
        m = re.search(r"compared:\s*(\d+)\s*/\s*(\d+)", out)
        self.assertIsNotNone(m, out)
        self.assertLess(int(m.group(1)), EXPECTED_PARITY_COUNT, out)

    def test_b4_both_scripts_print_scope_line_first(self):
        metadata_module = _load(METADATA_SCRIPT)
        _, metadata_out = _run(metadata_module, ["--plugins-dir", str(CANON_PLUGINS_DIR)])
        self.assertEqual(metadata_out.splitlines()[0], METADATA_SCOPE_LINE)

        parity_module = _load(PARITY_SCRIPT)
        _, parity_out = _run(parity_module, self._default_argv())
        self.assertEqual(parity_out.splitlines()[0], PARITY_SCOPE_LINE)

    def test_default_paths_point_at_canonical_files(self):
        module = _load(PARITY_SCRIPT)
        code, out = _run(module, [])
        self.assertEqual(code, 0, out)


if __name__ == "__main__":
    unittest.main()
