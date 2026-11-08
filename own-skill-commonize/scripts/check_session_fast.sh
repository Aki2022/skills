#!/bin/bash
# セッション開始時に走る高速検査の入口。中身は 3 つで、合わせて 0.1 秒ほど。
#
#   check_wiring_fast.sh   配線先に正典から切り離された実体が無いか
#   check_mirrors.py       plugin の導入版と正典の第三者 skill がずれていないか（--only-drift）
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

# 第三者 skill の plugin 導入版と正典のずれ（plugin の自動更新後に起きる）。ずれた時だけ 1 行ずつ出す。
plugin_args=()
for f in "$HOME"/.claude*/plugins/installed_plugins.json; do
  [ -f "$f" ] && plugin_args+=(--plugins-file "$f")
done
if [ ${#plugin_args[@]} -gt 0 ]; then
  drift="$(python3 "$HERE/check_mirrors.py" "$(cd "$HERE/../.." && pwd)" "${plugin_args[@]}" --only-drift 2>/dev/null)"
  if [ -n "$drift" ]; then
    printf '%s\n' "$drift" >&2
    rc=1
  fi
fi

out="$(python3 "$HERE/check_hook_parity.py" 2>&1)" || {
  printf '%s\n' "$out" | grep -v '^coverage:' >&2
  rc=1
}

exit "$rc"
