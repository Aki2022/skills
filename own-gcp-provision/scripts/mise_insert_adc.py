#!/usr/bin/env python3
"""refresh_adc.sh から呼ぶ。**heredoc へ戻さないこと**——bash 5.x は heredoc を pipe へ書き、
本文がパイプ容量を超えると誰も読まないパイプへの blocking write になって無言で永久に固まる
（macOS はパイプ KVA 逼迫時に容量が 512 バイトへ縮退する。2026-09-21 実測）。
"""
import sys
mise, adc = sys.argv[1], sys.argv[2]
lines = open(mise).read().splitlines()
out, inserted = [], False
for ln in lines:
    out.append(ln)
    if not inserted and ln.strip() == "[env]":
        out.append(f'# SDK(Node/Python)用 ADC。CLIのIMPERSONATE変数はSDKに効かないため必須')
        out.append(f'GOOGLE_APPLICATION_CREDENTIALS = "{adc}"')
        inserted = True
if not inserted:
    out.append(f'GOOGLE_APPLICATION_CREDENTIALS = "{adc}"')
open(mise, "w").write("\n".join(out) + "\n")
