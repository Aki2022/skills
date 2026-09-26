from __future__ import annotations

import importlib.util
import json
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory


SCRIPT = Path(__file__).resolve().parents[1] / "plugin_inventory.py"
SPEC = importlib.util.spec_from_file_location("plugin_inventory", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class PluginInventoryTest(unittest.TestCase):
    def test_classifies_skill_and_integration_assets_without_reading_secrets(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "skills" / "sample").mkdir(parents=True)
            (root / "skills" / "sample" / "SKILL.md").write_text("---\nname: sample\n---\n")
            (root / "commands").mkdir()
            (root / "commands" / "run.md").write_text("command")
            (root / "agents").mkdir()
            (root / "agents" / "reviewer.md").write_text("agent")
            (root / "hooks").mkdir()
            (root / "hooks" / "hooks.json").write_text("{}")
            (root / "plugin.json").write_text(
                json.dumps(
                    {
                        "name": "fixture",
                        "license": "Apache-2.0",
                        "mcpServers": {"drive": {"env": {"TOKEN": "must-not-print"}}},
                        "lspServers": {"python": {}},
                    }
                )
            )

            assets = MODULE.scan_plugin_assets(root)

            self.assertEqual(assets.skills, ("sample",))
            self.assertEqual(assets.commands, 1)
            self.assertEqual(assets.agents, 1)
            self.assertTrue(assets.hooks)
            self.assertEqual(assets.mcp_servers, 1)
            self.assertEqual(assets.lsp_servers, 1)
            self.assertEqual(assets.license, "Apache-2.0")
            self.assertEqual(assets.route_kind, "hybrid")
            self.assertNotIn("must-not-print", MODULE.format_plugin("codex", "main", {
                "pluginId": "fixture@local", "version": "1", "source": {"path": str(root)}
            }, assets))

    def test_skill_only_requires_installed_enabled_product_record(self):
        data = {
            "installed": [
                {"pluginId": "enabled@local", "installed": True, "enabled": True},
                {"pluginId": "disabled@local", "installed": True, "enabled": False},
                {"pluginId": "cache-only@local", "installed": False, "enabled": True},
            ]
        }
        self.assertEqual(
            MODULE.active_codex_plugins(data),
            ["enabled@local"],
        )

    def test_claude_inventory_ignores_project_and_disabled_plugins(self):
        rows = [
            {"id": "user-on@market", "scope": "user", "enabled": True},
            {"id": "user-off@market", "scope": "user", "enabled": False},
            {"id": "project-on@market", "scope": "project", "enabled": True},
        ]
        active, ignored_project, disabled = MODULE.active_claude_plugins(rows)
        self.assertEqual([x["id"] for x in active], ["user-on@market"])
        self.assertEqual(ignored_project, 1)
        self.assertEqual(disabled, 1)

    def test_reports_per_skill_licenses_for_mixed_collections(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            for name, license_text in (
                ("permissive", "Apache License Version 2.0"),
                ("restricted", "Proprietary; all rights reserved"),
            ):
                skill = root / "skills" / name
                skill.mkdir(parents=True)
                (skill / "SKILL.md").write_text(f"---\nname: {name}\n---\n")
                (skill / "LICENSE.txt").write_text(license_text)
            assets = MODULE.scan_plugin_assets(root)
            self.assertEqual(
                assets.skill_license_counts,
                (("Apache-2.0", 1), ("Proprietary", 1)),
            )
            rendered = MODULE.format_plugin(
                "Claude", "seat", {"id": "example", "version": "1"}, assets
            )
            self.assertIn("license=per-skill[Apache-2.0:1,Proprietary:1]", rendered)

    def test_installed_plugin_without_local_manifest_is_still_classified(self):
        with TemporaryDirectory() as tmp:
            assets = MODULE.scan_plugin_assets(Path(tmp))
            rendered = MODULE.format_plugin(
                "Codex", "private", {"pluginId": "codex-app-tools@openai-bundled"}, assets
            )
            self.assertIn("route=integration", rendered)
            self.assertIn("manifest=missing", rendered)

    def test_claude_user_setting_overrides_project_effective_enabled_state(self):
        rows = [
            {"id": "shared@market", "scope": "user", "enabled": True},
            {"id": "shared@market", "scope": "project", "enabled": True},
        ]
        active, project, disabled = MODULE.active_claude_plugins(
            rows, {"shared@market": False}
        )
        self.assertEqual(active, [])
        self.assertEqual(project, 1)
        self.assertEqual(disabled, 1)


    def test_codex_config_only_setting_makes_inventory_review_incomplete(self):
        from contextlib import redirect_stdout
        from io import StringIO
        from unittest.mock import patch

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            plugin_dir = root / "active-plugin"
            plugin_dir.mkdir()
            (root / "config.toml").write_text(
                '[plugins."active@local"]\nenabled = true\n'
                '[plugins."config-only@local"]\nenabled = true\n'
            )
            data = {
                "installed": [
                    {
                        "pluginId": "active@local",
                        "installed": True,
                        "enabled": True,
                        "source": {"path": str(plugin_dir)},
                    }
                ]
            }
            output = StringIO()
            with patch.object(MODULE, "_run_json", return_value=data):
                with redirect_stdout(output):
                    status = MODULE.main(["--codex", f"private={root}"])

            self.assertEqual(status, 1)
            self.assertIn("CONFIG-ONLY Codex/private config-only@local", output.getvalue())
            self.assertIn("RESULT: REVIEW (1 unresolved inventory entries)", output.getvalue())

    def test_codex_marketplace_record_resolves_exact_version_cache(self):
        from contextlib import redirect_stdout
        from io import StringIO
        from unittest.mock import patch

        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            cache = root / "plugins" / "cache" / "curated" / "example" / "1.2.3"
            (cache / "skills" / "sample").mkdir(parents=True)
            (cache / "skills" / "sample" / "SKILL.md").write_text("---\nname: sample\n---\n")
            (cache / "plugin.json").write_text(json.dumps({"name": "example", "license": "MIT"}))
            (root / "config.toml").write_text('[plugins."example@curated"]\nenabled = true\n')
            data = {
                "installed": [
                    {
                        "pluginId": "example@curated",
                        "installed": True,
                        "enabled": True,
                        "marketplaceName": "curated",
                        "version": "1.2.3",
                        "source": {"id": "example", "source": "marketplace"},
                    }
                ]
            }
            output = StringIO()
            with patch.object(MODULE, "_run_json", return_value=data):
                with redirect_stdout(output):
                    status = MODULE.main(["--codex", f"main={root}"])

            self.assertEqual(status, 0)
            self.assertIn("example@curated v1.2.3 route=skill-only skills=1[sample]", output.getvalue())
            self.assertIn("RESULT: OK", output.getvalue())
            self.assertNotIn("UNRESOLVED", output.getvalue())


if __name__ == "__main__":
    unittest.main()
