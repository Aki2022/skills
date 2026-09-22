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
[ -f "$naming_exceptions_file" ] && \
  naming_exceptions=" $(grep -v '^[[:space:]]*#' "$naming_exceptions_file" | tr -s '[:space:]' ' ') "

root_index=0
for root in "${roots[@]}"; do
  [ -d "$root" ] || { note "skip (not a directory): $root"; continue; }
  # S9: この root がグローバル正典か、リポジトリ固有かで期待語数が変わる。
  # 判定は**内在的な印**で行う — mirrors.yaml（第三者 skill のミラー台帳）は正典だけが持つ。
  # パス一致（$HOME/.agents/skills）で判定すると、正典を別チェックアウト（worktree・CI・
  # レビュアーの作業コピー）で lint したとき全 skill が誤判定されて赤くなる。実測 25件。
  if [ -f "$root/mirrors.yaml" ]; then
    naming_want=3; naming_where="グローバル正典"
  else
    naming_want=4; naming_where="リポジトリ固有"
  fi
  names_file="$scratch_dir/names.$root_index"
  : > "$names_file"
  for dir in "$root"/*/; do
    [ -d "$dir" ] || continue
    name=$(basename "$dir")
    # skill ではない管理ディレクトリを除外（docs/ は own-doc-update の scaffold）
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

# S10: 別名 root の板が正典と一致しているか。
# 検査そのものは check_global_topology.py が既に持っていたが、実運用で値を埋めて
# 呼ぶ場所が 0 件だった（呼び出しはテストと docs の説明文のみ）。skill を触るたびに
# 走るのはこの lint だけなので、ここから呼ぶ。
# root 一覧を references/alias-roots.txt に置くのは、渡し忘れた root を
# check_global_topology.py が黙って通すため（実測: 渡さなければ RESULT: OK）。
alias_roots_file="$naming_refs/alias-roots.txt"
if [ ! -f "$alias_roots_file" ]; then
  fail "S10: alias-roots.txt が無い（配布先の一覧が読めないと検査が空振りする）"
else
  topo_script="$(cd "$(dirname "$0")" && pwd)/check_global_topology.py"
  s10_canonical=""
  s10_args=()
  s10_checked=0
  s10_out=0
  s10_absent=0
  while IFS=$'\t' read -r s10_role s10_path _; do
    case "$s10_role" in ''|'#'*) continue ;; esac
    case "$s10_role" in "out:") s10_out=$((s10_out + 1)); continue ;; esac
    s10_path="${s10_path/#\~/$HOME}"
    if [ ! -d "$s10_path" ]; then
      # 「その道具を入れていない」と「配線が壊れた」は区別できないので warn。
      # 存在しない root を fail にすると、道具を使っていない環境や、
      # $HOME を差し替えて走るテストで常に赤くなる。
      s10_absent=$((s10_absent + 1))
      warn "S10 $s10_role: 一覧の root が無い（その道具を使っていなければ正常）: $s10_path"
      continue
    fi
    if [ "$s10_role" = "canonical" ]; then
      s10_canonical="$s10_path"
    else
      s10_args+=("--$s10_role" "$s10_path")
      s10_checked=$((s10_checked + 1))
    fi
  done < <(grep -vE '^[[:space:]]*#' "$alias_roots_file")

  # 何件を検査し、何件を対象外としたかを**常に**言う。
  # 緑であることと、何も見ていないことを出力で区別できるようにするため。
  note "S10: 別名 root ${s10_checked} 件を検査（対象外 ${s10_out} 件 / 不在 ${s10_absent} 件）"

  if [ -z "$s10_canonical" ]; then
    # canonical が読めない環境（$HOME を差し替えたテスト、正典を持たないマシン）では
    # 比較の基準が無いので S10 全体を skip する。fail にすると、
    # 別名検査と無関係な検査まで巻き込んで赤くなる。
    if grep -qE '^[[:space:]]*canonical[[:space:]]' "$alias_roots_file"; then
      note "S10: skip（正典 root が読めない。別名の一致は判定できない）"
    else
      fail "S10: alias-roots.txt に canonical の行が無い"
    fi
  elif [ "$s10_checked" -eq 0 ]; then
    if [ "$s10_absent" -gt 0 ]; then
      note "S10: 検査できる別名 root が無い（${s10_absent} 件すべて不在）"
    else
      fail "S10: 検査対象の別名 root が 0 件（一覧が壊れている）"
    fi
  elif [ ! -f "$topo_script" ]; then
    fail "S10: check_global_topology.py が無い: $topo_script"
  else
    if ! s10_report=$(python3 "$topo_script" --canonical "$s10_canonical" "${s10_args[@]}" 2>&1); then
      printf '%s\n' "$s10_report" | grep -E '^(FAIL|ERROR)' | while IFS= read -r s10_line; do
        fail "S10 ${s10_line#FAIL }"
      done
      # サブシェルの fail は親に伝わらないため、ここで明示的に落とす
      FAIL=$((FAIL + 1))
    fi
  fi
fi

if [ "$FAIL" -eq 0 ]; then
  note "OK: all skill checks passed"
fi
exit "$FAIL"
