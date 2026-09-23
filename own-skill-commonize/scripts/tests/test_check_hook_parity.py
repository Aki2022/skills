"""hook 参照の実在と席間一致を固定する。

hook の script 実体は共有できるが、「どの設定がそれを呼ぶか」は共通化できない
（Claude の settings.json は permission 等の account 固有 state を含み、丸ごと
symlink にできない）。この非対称が 2 度実害を出した:

  1. 2026-09-21 の origin- → own- 改名で ~/.claude は直したが ~/.claude-seat2 を
     取りこぼし、guard hook と session_start hook が 2 日間死んでいた（rc=127）。
  2. stop_nudge.sh は ~/.claude と Codex にあるのに ~/.claude-seat2 だけ落ちていた。

どちらもどこにも赤が出ないまま動き続けた。固定すべきは:
  1. 設定に書かれた script が実在すること
  2. 同じ製品の席どうしで、呼んでいる script の集合が同じこと
  3. **引数だけの違いは赤にしない**（account 識別子を引数で渡す hook は席ごとに違って正しい）
"""
import json
import os
import subprocess
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "check_hook_parity.py"


def _write(home: Path, seat: str, hooks: dict) -> None:
    d = home / seat
    d.mkdir(parents=True, exist_ok=True)
    (d / "settings.json").write_text(json.dumps({"hooks": hooks}))


def _run(home: Path):
    return subprocess.run([sys.executable, str(SCRIPT)], capture_output=True, text=True,
                          env=dict(os.environ, HOME=str(home)), timeout=30)


def _script(home: Path, name: str) -> Path:
    p = home / "h" / name
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text("#!/bin/bash\n")
    return p


class HookParityTests(unittest.TestCase):
    def test_matching_seats_pass(self):
        with TemporaryDirectory() as t:
            home = Path(t)
            s = _script(home, "own_warn_guards.sh")
            hooks = {"PreToolUse": [{"hooks": [{"command": f"bash {s}"}]}]}
            _write(home, ".claude", hooks)
            _write(home, ".claude-seat2", json.loads(json.dumps(hooks)))
            self.assertEqual(_run(home).returncode, 0)

    def test_dangling_reference_fails(self):
        """改名の取りこぼし。実在しない script を参照したら落ちる。"""
        with TemporaryDirectory() as t:
            home = Path(t)
            s = _script(home, "own_warn_guards.sh")
            _write(home, ".claude", {"PreToolUse": [{"hooks": [{"command": f"bash {s}"}]}]})
            _write(home, ".claude-seat2",
                   {"PreToolUse": [{"hooks": [{"command": f"bash {home}/h/origin_warn_guards.sh"}]}]})
            r = _run(home)
            self.assertEqual(r.returncode, 1)
            self.assertIn("参照先が実在しない", r.stdout)

    def test_hook_present_in_one_seat_only_fails(self):
        """stop_nudge.sh の形。片方の席にしか無い hook を捕まえる。"""
        with TemporaryDirectory() as t:
            home = Path(t)
            g = _script(home, "own_warn_guards.sh")
            n = _script(home, "stop_nudge.sh")
            common = {"PreToolUse": [{"hooks": [{"command": f"bash {g}"}]}]}
            _write(home, ".claude", dict(common, Stop=[{"hooks": [{"command": f"bash {n}"}]}]))
            _write(home, ".claude-seat2", json.loads(json.dumps(common)))
            r = _run(home)
            self.assertEqual(r.returncode, 1)
            self.assertIn("stop_nudge.sh", r.stdout)

    def test_argument_only_difference_passes(self):
        """account 識別子を引数で渡す hook は席ごとに違って正しい。誤検知しない。"""
        with TemporaryDirectory() as t:
            home = Path(t)
            a = _script(home, "aiphetamine_hook.py")
            _write(home, ".claude",
                   {"SessionStart": [{"hooks": [{"command": f"python3 {a} --account-name main"}]}]})
            _write(home, ".claude-seat2",
                   {"SessionStart": [{"hooks": [{"command": f"python3 {a} --account-name alias"}]}]})
            r = _run(home)
            self.assertEqual(r.returncode, 0, "引数の違いを席間のズレと誤認してはいけない")

    def test_no_config_at_all_is_an_error_not_a_pass(self):
        """対象 0 件を合格にしない。"""
        with TemporaryDirectory() as t:
            self.assertEqual(_run(Path(t)).returncode, 2)


if __name__ == "__main__":
    unittest.main()
