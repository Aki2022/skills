"""seat2_clone.sh を偽の app bundle に対して回す。実アプリには触らない。

本物の bundle は 1GB 近くあり、署名も vendor のものなので、テストでは
「Mach-O の実行ファイルを持つ最小の .app」を組み立てて同じ経路を通す。
payload が Mach-O でないと codesign が nested code として蹴るため、
shell script では代用できない。
"""

import json
import os
import plistlib
import shutil
import signal
import subprocess
import sys
import time
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "seat2_clone.sh"

pytestmark = pytest.mark.skipif(
    sys.platform != "darwin"
    or shutil.which("cc") is None
    or shutil.which("codesign") is None,
    reason="macOS と cc / codesign が要る",
)

# 引数と環境変数を読める payload。--sleep で起動中判定のテストに使う。
PAYLOAD_C = r"""
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>
int main(int argc, char **argv) {
    for (int i = 1; i < argc; i++) {
        if (strcmp(argv[i], "--sleep") == 0) { sleep(60); return 0; }
        printf("arg:%s\n", argv[i]);
    }
    const char *home = getenv("CODEX_HOME");
    const char *data = getenv("CODEX_ELECTRON_USER_DATA_PATH");
    printf("CODEX_HOME:%s\n", home ? home : "");
    printf("CODEX_ELECTRON_USER_DATA_PATH:%s\n", data ? data : "");
    return 0;
}
"""


def build_source_app(root: Path, name: str, exe: str, ident: str, version: str, schemes):
    app = root / f"{name}.app"
    macos = app / "Contents" / "MacOS"
    macos.mkdir(parents=True, exist_ok=True)
    src = root / f"{exe}.c"
    src.write_text(PAYLOAD_C)
    subprocess.run(["cc", "-o", str(macos / exe), str(src)], check=True)
    info = {
        "CFBundleExecutable": exe,
        "CFBundleIdentifier": ident,
        "CFBundleName": name,
        "CFBundleDisplayName": name,
        "CFBundleShortVersionString": version,
        "CFBundleURLTypes": [{"CFBundleURLName": name, "CFBundleURLSchemes": schemes}],
    }
    with open(app / "Contents" / "Info.plist", "wb") as fh:
        plistlib.dump(info, fh)
    subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(app)], check=True)
    return app


def run(install_dir: Path, *args):
    env = dict(os.environ, LAUNCHER_INSTALL_DIR=str(install_dir))
    return subprocess.run(
        ["bash", str(SCRIPT), *args], env=env, capture_output=True, text=True
    )


def info_of(app: Path):
    with open(app / "Contents" / "Info.plist", "rb") as fh:
        return plistlib.load(fh)


@pytest.fixture
def claude_src(tmp_path):
    return build_source_app(
        tmp_path, "Claude", "Claude", "com.anthropic.claudefordesktop", "2.2553.13",
        ["claude", "http", "https", "msauth.com.anthropic.claudefordesktop"],
    )


@pytest.fixture
def codex_src(tmp_path):
    return build_source_app(
        tmp_path, "ChatGPT", "ChatGPT", "com.openai.codex", "26.917.51856",
        ["codex", "http", "https"],
    )


def test_install_builds_shim_and_payload(tmp_path, claude_src):
    assert run(tmp_path, "--install", "claude").returncode == 0
    macos = tmp_path / "Claude-Seat2.app" / "Contents" / "MacOS"
    assert (macos / "ClaudeSeat2Payload").exists(), "本体バイナリが改名されていない"
    shim = macos / "Claude"
    assert shim.read_text().startswith("#!/bin/sh"), "shim がシェルスクリプトでない"
    assert os.access(shim, os.X_OK)


def test_identity_is_patched(tmp_path, claude_src):
    run(tmp_path, "--install", "claude")
    info = info_of(tmp_path / "Claude-Seat2.app")
    assert info["CFBundleIdentifier"] == "local.launchers.claude-seat2"
    assert info["CFBundleDisplayName"] == "Claude Seat2"
    schemes = [s for e in info["CFBundleURLTypes"] for s in e["CFBundleURLSchemes"]]
    assert "claude-seat2" in schemes
    # msauth. 系は "." 区切り。identifier 形式の scheme に "-seat2" を足すと別物になる
    assert "msauth.com.anthropic.claudefordesktop.seat2" in schemes
    # 既定ブラウザを争わせない。接尾辞を付けて残す（http-seat2）のも駄目なので前方一致で見る
    assert not [s for s in schemes if s.startswith("http")], schemes
    # 本体の scheme がそのまま残ると二重登録になる
    assert "claude" not in schemes


def test_payload_is_resigned_under_its_new_name(tmp_path, claude_src):
    """--deep が外れると payload は改名前の署名を持ったまま残る。

    実機ではそれが起動時 SIGKILL になる（vendor 署名の本物のバイナリの場合）。
    偽 bundle では kill を再現できないので、原因である「payload を署名し直していない」
    ことを署名識別子で見る。ここが緑のまま実機で落ちる、を防ぐための代理ではなく、
    --deep の有無そのものの検査。
    """
    run(tmp_path, "--install", "claude")
    payload = tmp_path / "Claude-Seat2.app" / "Contents" / "MacOS" / "ClaudeSeat2Payload"
    out = subprocess.run(["codesign", "-dv", str(payload)], capture_output=True, text=True)
    identifier = [l for l in out.stderr.splitlines() if l.startswith("Identifier=")]
    assert identifier, out.stderr
    assert identifier[0].startswith("Identifier=ClaudeSeat2Payload"), identifier[0]


def test_shim_injects_profile_and_passes_args(tmp_path, claude_src):
    run(tmp_path, "--install", "claude")
    out = subprocess.run(
        [str(tmp_path / "Claude-Seat2.app" / "Contents" / "MacOS" / "Claude"), "hello"],
        capture_output=True, text=True,
    )
    assert out.returncode == 0, f"payload が起動しない: {out.stderr}"
    home = os.environ["HOME"]
    assert f"arg:--user-data-dir={home}/.claude-seat2" in out.stdout
    assert "arg:hello" in out.stdout


def test_explicit_user_data_dir_wins(tmp_path, claude_src):
    run(tmp_path, "--install", "claude")
    out = subprocess.run(
        [str(tmp_path / "Claude-Seat2.app" / "Contents" / "MacOS" / "Claude"),
         "--user-data-dir=/somewhere/else"],
        capture_output=True, text=True,
    )
    assert out.stdout.count("--user-data-dir") == 1, "呼び出し側の指定を上書きしている"
    assert "arg:--user-data-dir=/somewhere/else" in out.stdout


def test_codex_shim_exports_codex_home(tmp_path, codex_src):
    assert run(tmp_path, "--install", "codex").returncode == 0
    out = subprocess.run(
        [str(tmp_path / "Codex-Seat2.app" / "Contents" / "MacOS" / "ChatGPT")],
        capture_output=True, text=True,
    )
    home = os.environ["HOME"]
    assert f"CODEX_HOME:{home}/.codex-seat2" in out.stdout
    assert f"CODEX_ELECTRON_USER_DATA_PATH:{home}/.codex-seat2/electron" in out.stdout


def test_check_reports_match_then_drift(tmp_path, claude_src):
    run(tmp_path, "--install", "claude")
    ok = run(tmp_path, "--check", "claude")
    assert ok.returncode == 0 and "OK:" in ok.stdout

    info = info_of(claude_src)
    info["CFBundleShortVersionString"] = "2.9999.0"
    with open(claude_src / "Contents" / "Info.plist", "wb") as fh:
        plistlib.dump(info, fh)

    drifted = run(tmp_path, "--check", "claude")
    assert drifted.returncode == 1, "本体が進んだのに OK を返した"
    assert "DRIFT:" in drifted.stderr


def test_check_skips_when_seat_not_installed(tmp_path, claude_src):
    result = run(tmp_path, "--check", "claude")
    assert result.returncode == 0
    assert "SKIP:" in result.stdout


def test_install_refuses_while_seat_is_running(tmp_path, claude_src):
    run(tmp_path, "--install", "claude")
    shim = tmp_path / "Claude-Seat2.app" / "Contents" / "MacOS" / "Claude"
    proc = subprocess.Popen([str(shim), "--sleep"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        time.sleep(1)
        result = run(tmp_path, "--install", "claude")
        assert result.returncode != 0, "起動中の bundle を差し替えた"
        assert "is running" in result.stderr
    finally:
        proc.send_signal(signal.SIGKILL)
        proc.wait()


def test_rebuild_leaves_no_staging_or_previous(tmp_path, claude_src):
    run(tmp_path, "--install", "claude")
    run(tmp_path, "--install", "claude")
    leftovers = [p.name for p in tmp_path.iterdir() if ".staging" in p.name or ".previous" in p.name]
    assert leftovers == [], f"中間物が残った: {leftovers}"


def test_bad_target_is_rejected(tmp_path, claude_src):
    assert run(tmp_path, "--install", "nosuchseat").returncode == 64
    assert run(tmp_path, "--frobnicate", "claude").returncode == 64


def test_patch_plist_helper_updates_identity_and_schemes(tmp_path):
    plist = tmp_path / "Info.plist"
    original = {
        "CFBundleDisplayName": "Vendor App",
        "CFBundleIdentifier": "com.vendor.app",
        "CFBundleURLTypes": [
            {"CFBundleURLSchemes": ["vendor", "http", "https", "msauth.vendor"]}
        ],
    }
    with open(plist, "wb") as fh:
        plistlib.dump(original, fh)

    helper = SCRIPT.with_name("patch_plist.py")
    result = subprocess.run(
        [sys.executable, str(helper), str(plist), "Vendor Seat2", "local.vendor.seat2"],
        capture_output=True,
        text=True,
        timeout=10,
    )
    assert result.returncode == 0, result.stderr

    with open(plist, "rb") as fh:
        patched = plistlib.load(fh)
    assert patched["CFBundleDisplayName"] == "Vendor Seat2"
    assert patched["CFBundleIdentifier"] == "local.vendor.seat2"
    assert patched["CFBundleURLTypes"][0]["CFBundleURLSchemes"] == [
        "vendor-seat2",
        "msauth.vendor.seat2",
    ]
