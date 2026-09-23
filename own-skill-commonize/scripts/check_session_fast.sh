#!/bin/bash
# セッション開始時に走る高速検査の入口。中身は 2 つで、合わせて 0.07 秒。
#
#   check_wiring_fast.sh   配線先に正典から切り離された実体が無いか
#   check_hook_parity.py   hook の参照が実在し、席どうしで揃っているか
#
# 入口を 1 本にまとめる理由: hook のエントリは Claude 2 席・Codex・Gemini の
# 4 箇所に手で書く必要があり（settings は account 固有 state を含むため symlink に
# できない）、エントリを増やすとそれ自体が分岐の種になる。実際 2026-09-21 の改名で
# ~/.claude-seat2 の参照が 2 日間死んでいた。検査を足すときは、この script の中に足す。
#
# 逸脱が無ければ何も出力しない（常に喋る検査は読まれなくなる）。
# 終了コード: 0 = 逸脱なし / 1 = 逸脱あり
set -uo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
rc=0

bash "$HERE/check_wiring_fast.sh" || rc=1

out="$(python3 "$HERE/check_hook_parity.py" 2>&1)" || {
  printf '%s\n' "$out" | grep -v '^coverage:' >&2
  rc=1
}

exit "$rc"
