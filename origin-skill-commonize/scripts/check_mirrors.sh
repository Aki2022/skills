#!/usr/bin/env bash
# check_mirrors.sh — 第三者 skill ミラーの同値・鮮度検査（check_mirrors.py の入口）
# 使い方: bash check_mirrors.sh [skills_root] [--max-age-days N]
exec python3 "$(cd "$(dirname "$0")" && pwd)/check_mirrors.py" "$@"
