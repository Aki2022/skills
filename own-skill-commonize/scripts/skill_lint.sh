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
#   S8: 各 skill の scripts/tests/test_*.py を pytest で実行する
#   S9: 自前 skill の名前が own-<対象>-<動作>（末尾は動詞）で、語数が所属と一致する
#       3語=グローバル正典 $HOME/.agents/skills / 4語=リポジトリ固有 <repo>/.agents/skills
set -u

FAIL=0
note() { printf '%s\n' "$*"; }
fail() { FAIL=1; printf 'FAIL %s\n' "$*"; }
warn() { printf 'WARN %s\n' "$*"; }

scratch_dir=$(mktemp -d)
trap 'rm -rf "$scratch_dir"' EXIT

reference_parser="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)/skill_lint_refs.py"
if [ ! -f "$reference_parser" ] || ! command -v python3 >/dev/null 2>&1; then
  fail "S4: skill_lint_refs.py と python3 が必要"
  exit "$FAIL"
fi

roots=("$@")
if [ ${#roots[@]} -eq 0 ]; then
  roots=("$HOME/.agents/skills")
  [ -d ".agents/skills" ] && roots+=(".agents/skills")
fi

# S8 の事前判定: skip すべき理由があれば一度だけ明示する（黙ってスキップしない）。
pytest_skip_reason=""
if [ "${SKILL_LINT_SKIP_PYTEST:-}" = "1" ]; then
  pytest_skip_reason="環境変数 SKILL_LINT_SKIP_PYTEST=1"
elif ! command -v python3 >/dev/null 2>&1; then
  pytest_skip_reason="python3 が見つからない"
elif ! python3 -m pytest --version >/dev/null 2>&1; then
  pytest_skip_reason="pytest が使えない"
fi
[ -n "$pytest_skip_reason" ] && note "S8: skip（$pytest_skip_reason のため pytest を実行しない）"
pytest_ran=0

# S9 の語彙。allowlist（知らない語は通さない）と、既存逸脱の burn-down リスト。
naming_refs="$(cd "$(dirname "${BASH_SOURCE[0]}")/../references" && pwd)"
naming_verbs_file="$naming_refs/naming-verbs.txt"
naming_exceptions_file="$naming_refs/naming-exceptions.txt"
if [ ! -f "$naming_verbs_file" ]; then
  fail "S9: naming-verbs.txt が無い（動詞 allowlist が読めないと検査が空振りする）"
  exit "$FAIL"
fi
naming_verbs=" $(grep -v '^[[:space:]]*#' "$naming_verbs_file" | tr -s '[:space:]' ' ') "
naming_exceptions=" "
naming_home_root=$(cd "$HOME/.agents/skills" 2>/dev/null && pwd -P || echo "")
[ -f "$naming_exceptions_file" ] && \
  naming_exceptions=" $(grep -v '^[[:space:]]*#' "$naming_exceptions_file" | tr -s '[:space:]' ' ') "

root_index=0
for root in "${roots[@]}"; do
  [ -d "$root" ] || { note "skip (not a directory): $root"; continue; }
  # S9: この root がグローバル正典か、リポジトリ固有かで期待語数が変わる
  naming_root_real=$(cd "$root" 2>/dev/null && pwd -P || echo "$root")
  if [ -n "$naming_home_root" ] && [ "$naming_root_real" = "$naming_home_root" ]; then
    naming_want=3; naming_where="グローバル正典"
  else
    naming_want=4; naming_where="リポジトリ固有"
  fi
  names_file="$scratch_dir/names.$root_index"
  : > "$names_file"
  for dir in "$root"/*/; do
    [ -d "$dir" ] || continue
    name=$(basename "$dir")
    # skill ではない管理ディレクトリを除外（docs/ は origin-doc-update の scaffold）
    # synced/ は Claude Code のスキル同期バケット（UUID ディレクトリと manifest.json）で、
    # skill ではなくアプリ管理のインフラ。正典 root 直下に現れるが S1 の対象ではない。
    case "$name" in docs|node_modules|synced|.git) continue ;; esac
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

    # S9: 自前 skill の命名規則。第三者ミラー（cloudflare-* 等）は対象外。
    case "$name" in
      own-*|origin-*)
        case "$naming_exceptions" in
          *" $name "*) : ;;   # burn-down リスト掲載。改名時に1行消す
          *)
            case "$name" in
              origin-*)
                fail "S9 $name: 接頭辞は own-（origin- は移行対象。references/naming-exceptions.txt 参照）"
                ;;
              *)
                naming_last="${name##*-}"
                naming_segments=$(( $(printf '%s' "$name" | tr -cd '-' | wc -c) + 1 ))
                if [ "$naming_segments" -ne "$naming_want" ]; then
                  fail "S9 $name: ${naming_where}は${naming_want}語（現在 ${naming_segments}語）。語数が所属を表す"
                elif case "$naming_verbs" in *" $naming_last "*) false ;; *) true ;; esac; then
                  fail "S9 $name: 末尾 '$naming_last' が動詞リストに無い（references/naming-verbs.txt）"
                fi
                ;;
            esac
            ;;
        esac
        ;;
    esac

    if [ -n "$fm_name" ]; then
      previous=$(awk -F '\t' -v key="$fm_name" '$1 == key {print $2; exit}' "$names_file")
      if [ -n "$previous" ]; then
        warn "S7 duplicate frontmatter name '$fm_name': $previous and $name"
      fi
      printf '%s\t%s\n' "$fm_name" "$name" >> "$names_file"
    fi

    # S4: 同梱リソース参照の実在確認。相対参照はこのskill配下、
    # /skill-name → references/file 形式はcanonical root配下で解決する。
    refs_output=""
    if ! refs_output=$(python3 "$reference_parser" "$root" "$md"); then
      fail "S4 $name: 参照 parser が失敗"
    fi
    while IFS=$'\t' read -r kind first second; do
      [ -n "$kind" ] || continue
      case "$kind" in
        same)
          [ -e "$dir/$first" ] || fail "S4 $name: 参照 '$first' が実在しない"
          ;;
        cross)
          [ -e "$root/$first/$second" ] || fail "S4 $name: cross-skill 参照 '$first/$second' が実在しない"
          ;;
        unsafe)
          fail "S4 $name: unsafe reference '$second'"
          ;;
        *)
          fail "S4 $name: 参照 parser の出力が不正"
          ;;
      esac
      # process substitution であって here-string ではない。here-string だと bash が
      # 内容をパイプへ書き、その読み手が同じプロセスのこのループになる。参照が多い
      # skill（cloudflare / agents-sdk 等）で内容がパイプバッファ 16KB を超えると、
      # ループが読み始める前に write(2) が埋まって自己デッドロックする。実測では
      # skill_lint が 1 日以上ぶら下がったまま誰も気づかなかった。
    done < <(printf '%s\n' "$refs_output")

    # S5: 壊れた symlink
    while IFS= read -r link; do
      fail "S5 $name: 壊れた symlink: $link"
    done < <(find "$dir" -type l ! -exec test -e {} \; -print 2>/dev/null)

    # S8: scripts/tests/test_*.py があれば pytest を走らせる（赤いまま編集させない）
    if [ -z "$pytest_skip_reason" ] && [ -d "${dir}scripts/tests" ] \
      && ls "${dir}scripts/tests"/test_*.py >/dev/null 2>&1; then
      pytest_ran=$((pytest_ran + 1))
      pytest_out=$(python3 -m pytest -q "${dir}scripts/tests" 2>&1)
      pytest_rc=$?
      summary=$(printf '%s\n' "$pytest_out" | tail -1)
      if [ "$pytest_rc" -ne 0 ]; then
        failed_n=$(printf '%s\n' "$pytest_out" | grep -oE '[0-9]+ failed' | tail -1)
        fail "S8 $name: pytest が失敗（${failed_n:-$summary}）"
      else
        note "S8 $name: pytest 成功（${summary}）"
      fi
    fi
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

if [ -z "$pytest_skip_reason" ] && [ "$pytest_ran" -eq 0 ]; then
  note "S8: pytest を持つ skill が無い"
fi

if [ "$FAIL" -eq 0 ]; then
  note "OK: all skill checks passed"
fi
exit "$FAIL"
