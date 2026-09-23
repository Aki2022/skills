#!/bin/bash
# 配線先に「symlink でない実体」が現れていないかだけを見る。0.02 秒で終わる。
#
# なぜこれだけを分けるか: audit_skill_wiring.py は 1.1 秒、skill_lint は 88 秒かかり、
# どちらも SessionStart hook には置けない。一方この 1 条件は、
# Codex の $skill-installer が $CODEX_HOME/skills/<名前> へ実体を書き込んだ状態と
# 1 対 1 で対応する（2026-09-23 に実測。正典にある名前は
# InstallError: Destination already exists で拒否され、無い名前だけが実体になる）。
# 実体が出来るとその skill は 1 席だけで使えて他のツールから見えず、git にも台帳にも
# 残らない。ADR-20260906「供給源は正典のみ」違反を、ほぼ無料で毎セッション検出する。
#
# 正常時は何も出力しない（常に赤い検査を作らないため）。
# 終了コード: 0 = 逸脱なし / 1 = 逸脱あり
set -uo pipefail

ROOTS=(
  "$HOME/.codex/skills"
  "$HOME/.codex-private/skills"
  "$HOME/.codex-seat2/skills"
  "$HOME/.gemini/config/skills"
)
# 製品が自分で書き込む棚。うちの配線ではないので対象外。
PRODUCT_OWNED_RE='^\.'

found=0
for root in "${ROOTS[@]}"; do
  [ -d "$root" ] || continue
  while IFS= read -r p; do
    n="$(basename "$p")"
    [[ "$n" =~ $PRODUCT_OWNED_RE ]] && continue
    if [ "$found" -eq 0 ]; then
      echo "配線先に正典から切り離された実体があります（Codex の \$skill-installer で入れた skill は" >&2
      echo "1 席だけで使え、Claude / Gemini から見えず、git にも mirrors.yaml にも残りません）:" >&2
      found=1
    fi
    echo "  ${root/#$HOME/~}/$n" >&2
  done < <(find "$root" -maxdepth 1 -mindepth 1 ! -type l 2>/dev/null)
done

if [ "$found" -eq 1 ]; then
  echo "対処: 正典へ移して全 root へ配線し直す →" >&2
  echo "  bash ~/.agents/skills/own-skill-commonize/scripts/mirror_skill.sh --help" >&2
  exit 1
fi
exit 0
