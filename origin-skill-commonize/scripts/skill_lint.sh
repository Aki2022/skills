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
#   S6: 同じ正典にある hooks.json の bash 参照先が実在する
#   S7: skill 間の機械的な重複候補を warn-only で報告する
set -u

FAIL=0
note() { printf '%s\n' "$*"; }
fail() { FAIL=1; printf 'FAIL %s\n' "$*"; }
warn() { printf 'WARN %s\n' "$*"; }

scratch_dir=$(mktemp -d)
trap 'rm -rf "$scratch_dir"' EXIT

roots=("$@")
if [ ${#roots[@]} -eq 0 ]; then
  roots=("$HOME/.agents/skills")
  [ -d ".agents/skills" ] && roots+=(".agents/skills")
fi

root_index=0
for root in "${roots[@]}"; do
  [ -d "$root" ] || { note "skip (not a directory): $root"; continue; }
  names_file="$scratch_dir/names.$root_index"
  : > "$names_file"
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
      if [ "$name" = "design" ]; then
        warn "S3 $name: third-party frontmatter name '$fm_name' がディレクトリ名と不一致"
      else
        fail "S3 $name: frontmatter name '$fm_name' がディレクトリ名と不一致"
      fi
    fi

    if [ -n "$fm_name" ]; then
      previous=$(awk -F '\t' -v key="$fm_name" '$1 == key {print $2; exit}' "$names_file")
      if [ -n "$previous" ]; then
        warn "S7 duplicate frontmatter name '$fm_name': $previous and $name"
      fi
      printf '%s\t%s\n' "$fm_name" "$name" >> "$names_file"
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

  # `source-command-*` wrappers are a deterministic duplicate candidate. Keep the
  # decision warn-only because the actual semantic overlap still needs triage.
  while IFS=$'\t' read -r _fm_name dir_name; do
    case "$dir_name" in
      source-command-*)
        base=${dir_name#source-command-}
        [ -d "$root/$base" ] && warn "S7 source-command duplicate candidate: $dir_name and $base"
        ;;
    esac
  done < "$names_file"

  # Tool-specific config is adjacent to the skills root. Broken first-party hook
  # references are a real failure; a missing config is simply out of scope.
  hook_config="$root/../codex/hooks.json"
  if [ -f "$hook_config" ]; then
    config_dir=$(cd "$(dirname "$hook_config")" && pwd)
    while IFS= read -r command; do
      hook_command=$(printf '%s\n' "$command" | sed -n 's/.*"command"[[:space:]]*:[[:space:]]*"bash[[:space:]]\+\([^"].*\)".*/\1/p')
      hook_path=${hook_command%%[[:space:]]*}
      [ -n "$hook_path" ] || continue
      case "$hook_path" in
        /*) resolved="$hook_path" ;;
        *) resolved="$config_dir/$hook_path" ;;
      esac
      [ -f "$resolved" ] || fail "S6 hooks.json: bash reference does not exist: $hook_path"
    done < <(grep -oE '"command"[[:space:]]*:[[:space:]]*"bash[[:space:]]+[^" ]+"' "$hook_config" || true)
  fi
  root_index=$((root_index + 1))
done

if [ "$FAIL" -eq 0 ]; then
  note "OK: all skill checks passed"
fi
exit "$FAIL"
