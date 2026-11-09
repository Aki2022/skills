#!/usr/bin/env python3
"""refresh_adc.sh から呼ぶ。**heredoc へ戻さないこと**——bash 5.x は heredoc を pipe へ書き、
本文がパイプ容量を超えると誰も読まないパイプへの blocking write になって無言で永久に固まる
（macOS はパイプ KVA 逼迫時に容量が 512 バイトへ縮退する。2026-09-21 実測）。

使い方: mise_insert_adc.py <mise 設定ファイル> <ADC のパス>
  `.mise.toml` でも `mise.toml` でも同じに扱う。挿入先は次の順で決める:
    1. `[env]` 見出しの直下
    2. `[env]` が無ければ CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT の行の直後（同じ表に入れる）
    3. どちらも無ければ末尾に `[env]` ごと足す（末尾の別の表に入れない）
  ADC のパスがホーム配下なら `{{env.HOME}}` で書く（repo に個人のパスを残さない）。
"""
import os
import re
import sys

mise, adc = sys.argv[1], sys.argv[2]
home = os.path.expanduser("~").rstrip("/")
if home and adc.startswith(home + "/"):
    adc = "{{env.HOME}}" + adc[len(home):]

entry = [
    "# SDK(Node/Python)用 ADC。CLIのIMPERSONATE変数はSDKに効かないため必須",
    f'GOOGLE_APPLICATION_CREDENTIALS = "{adc}"',
]
lines = open(mise).read().splitlines()
sa_re = re.compile(r"^\s*CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT\s*=")

idx = next((i for i, ln in enumerate(lines) if ln.strip() == "[env]"), None)
if idx is None:
    idx = next((i for i, ln in enumerate(lines) if sa_re.match(ln)), None)
if idx is not None:
    out = lines[: idx + 1] + entry + lines[idx + 1 :]
else:
    out = lines + ["", "[env]"] + entry
open(mise, "w").write("\n".join(out) + "\n")
