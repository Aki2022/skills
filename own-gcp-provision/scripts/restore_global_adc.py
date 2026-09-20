#!/usr/bin/env python3
"""refresh_adc.sh から呼ぶ。**heredoc へ戻さないこと**——bash 5.x は heredoc を pipe へ書き、
本文がパイプ容量を超えると誰も読まないパイプへの blocking write になって無言で永久に固まる
（macOS はパイプ KVA 逼迫時に容量が 512 バイトへ縮退する。2026-09-21 実測）。
"""
import json, os, sys
src = json.loads(os.environ["SOURCE_JSON"])
with open(sys.argv[1], "w") as f:
    json.dump(src, f, indent=2)
