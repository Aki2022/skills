#!/usr/bin/env python3
"""refresh_adc.sh から呼ぶ。**heredoc へ戻さないこと**——bash 5.x は heredoc を pipe へ書き、
本文がパイプ容量を超えると誰も読まないパイプへの blocking write になって無言で永久に固まる
（macOS はパイプ KVA 逼迫時に容量が 512 バイトへ縮退する。2026-09-21 実測）。
"""
import json, sys
with open(sys.argv[1]) as f:
    d = json.load(f)
t = d.get("type")
if t == "impersonated_service_account":
    src = d.get("source_credentials")
    if not src:
        sys.exit("グローバル ADC に source_credentials が無い")
elif t == "authorized_user":
    src = {k: d[k] for k in d}
else:
    sys.exit(f"未対応の ADC type: {t}")
if src.get("type") != "authorized_user" or "refresh_token" not in src:
    sys.exit("source は authorized_user + refresh_token である必要がある")
print(json.dumps(src))
