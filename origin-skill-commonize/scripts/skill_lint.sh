#!/usr/bin/env bash
# skill_lint.sh — 決定論的なスキル静的チェック（LLM 不使用・exit code で判定）
#
# 使い方:
#   bash skill_lint.sh [skills_root ...]
# 引数なしなら ~/.agents/skills とカレントリポジトリの .agents/skills を対象にする。
#
# チェック項目:
#   S1: SKILL.md が存在する
#   S2: frontmatter に name: と description: がある
#   S3: frontmatter の name がディレクトリ名と一致する
#   S4: SKILL.md 内で参照している同梱パス (scripts/ references/ assets/) が実在する
#   S5: スキルディレクトリ内に壊れた symlink がない
set -u

FAIL=0
note() { printf '%s\n' "$*"; }
fail() { FAIL=1; printf 'FAIL %s\n' "$*"; }

roots=("$@")
if [ ${#roots[@]} -eq 0 ]; then
  roots=("$HOME/.agents/skills")
  [ -d ".agents/skills" ] && roots+=(".agents/skills")
fi

for root in "${roots[@]}"; do
  [ -d "$root" ] || { note "skip (not a directory): $root"; continue; }
  for dir in "$root"/*/; do
    [ -d "$dir" ] || continue
    name=$(basename "$dir")
    md="$dir/SKILL.md"

    # S1
    if [ ! -f "$md" ]; then
      fail "S1 $name: SKILL.md がない"
      continue
    fi

    # frontmatter を先頭の --- ... --- から抽出
    fm=$(awk 'NR==1 && $0!="---"{exit} NR>1 && $0=="---"{exit} NR>1{print}' "$md")

    # S2
    printf '%s\n' "$fm" | grep -q '^name:' || fail "S2 $name: frontmatter に name: がない"
    printf '%s\n' "$fm" | grep -q '^description:' || fail "S2 $name: frontmatter に description: がない"

    # S3
    fm_name=$(printf '%s\n' "$fm" | sed -n 's/^name:[[:space:]]*//p' | head -1 | sed 's/^["'"'"']//; s/["'"'"']$//')
    if [ -n "$fm_name" ] && [ "$fm_name" != "$name" ]; then
      fail "S3 $name: frontmatter name '$fm_name' がディレクトリ名と不一致"
    fi

    # S4: 同梱リソース参照の実在確認
    while IFS= read -r ref; do
      [ -e "$dir/$ref" ] || fail "S4 $name: 参照 '$ref' が実在しない"
    done < <(grep -oE '(^|[[:space:]`"'"'"'(])(scripts|references|assets)/[A-Za-z0-9._/-]+' "$md" \
      | sed 's/^[[:space:]`"'"'"'(]*//; s/[.,)]*$//' | sort -u)

    # S5: 壊れた symlink
    while IFS= read -r link; do
      fail "S5 $name: 壊れた symlink: $link"
    done < <(find "$dir" -type l ! -exec test -e {} \; -print 2>/dev/null)
  done
done

if [ "$FAIL" -eq 0 ]; then
  note "OK: all skill checks passed"
fi
exit "$FAIL"
