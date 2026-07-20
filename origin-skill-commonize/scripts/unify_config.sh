#!/usr/bin/env bash
# unify_config.sh — エージェント設定ファイル/スキルを正典への symlink に統一する。
#
# 使い方:
#   unify_config.sh <canonical> <alias> [<alias> ...]
#
# 例:
#   unify_config.sh AGENTS.md CLAUDE.md
#   unify_config.sh ~/.agents/skills ~/.claude/skills ~/.codex/skills/foo
#
# 動作:
#   - canonical が存在しなければエラー終了（先に正典を用意すること）。
#   - 各 alias について:
#       * すでに canonical を指す symlink → 何もしない（冪等）。
#       * 別の対象を指す symlink → 警告して張り替え（バックアップは取らない/元リンクのみ表示）。
#       * 実体ファイル/ディレクトリ → canonical と内容比較。差分があれば警告。
#         必ず <alias>.bak_<YYYYMMDD> に退避してから symlink を作成。
#       * 不在 → そのまま symlink を作成。
#   - 同一ディレクトリ内なら相対 symlink、異なるディレクトリなら絶対パスで張る。
#   - 最後に検証結果を表示する。
#
# 破壊操作（mv による退避）を行うため、重要な環境ではコミット/バックアップ後に実行すること。

set -euo pipefail

if [ "$#" -lt 2 ]; then
  echo "usage: $0 <canonical> <alias> [<alias> ...]" >&2
  exit 2
fi

canonical_in="$1"; shift
# 正典の絶対パス
canonical_abs="$(cd "$(dirname "$canonical_in")" 2>/dev/null && pwd)/$(basename "$canonical_in")"

if [ ! -e "$canonical_abs" ]; then
  echo "ERROR: 正典が存在しない: $canonical_abs" >&2
  echo "  先に正典ファイル/ディレクトリを用意してから実行すること。" >&2
  exit 1
fi

date_tag="$(date +%Y%m%d)"

link_target() {  # 同階層なら相対、違えば絶対を返す
  local alias_dir="$1"
  if [ "$alias_dir" = "$(dirname "$canonical_abs")" ]; then
    basename "$canonical_abs"
  else
    printf '%s' "$canonical_abs"
  fi
}

for alias_in in "$@"; do
  alias_dir="$(cd "$(dirname "$alias_in")" 2>/dev/null && pwd || true)"
  if [ -z "$alias_dir" ]; then
    echo "SKIP: 親ディレクトリが存在しない: $alias_in" >&2
    continue
  fi
  alias_abs="$alias_dir/$(basename "$alias_in")"

  if [ "$alias_abs" = "$canonical_abs" ]; then
    echo "SKIP: alias と canonical が同一: $alias_abs" >&2
    continue
  fi

  target="$(link_target "$alias_dir")"

  if [ -L "$alias_abs" ]; then
    cur="$(readlink "$alias_abs")"
    # 解決先が canonical なら冪等で何もしない
    resolved="$(cd "$alias_dir" && cd "$(dirname "$cur")" 2>/dev/null && pwd)/$(basename "$cur")" || resolved=""
    if [ "$resolved" = "$canonical_abs" ]; then
      echo "OK(冪等): $alias_abs は既に正典を指している"
      continue
    fi
    echo "WARN: $alias_abs は別の対象を指す symlink ($cur) → 張り替える"
    rm "$alias_abs"
  elif [ -e "$alias_abs" ]; then
    # 実体 — 差分チェック
    if diff -rq "$alias_abs" "$canonical_abs" >/dev/null 2>&1; then
      echo "INFO: $alias_abs は正典と同一内容。退避してリンク化する"
    else
      echo "WARN: $alias_abs は正典と差分あり。退避するので必要なら後で確認すること:"
      echo "      バックアップ: ${alias_abs}.bak_${date_tag}"
    fi
    mv "$alias_abs" "${alias_abs}.bak_${date_tag}"
  fi

  ln -s "$target" "$alias_abs"
  echo "LINK: $alias_abs → $target"
done

echo
echo "=== 検証 ==="
for alias_in in "$@"; do
  alias_dir="$(cd "$(dirname "$alias_in")" 2>/dev/null && pwd || true)"
  [ -z "$alias_dir" ] && continue
  alias_abs="$alias_dir/$(basename "$alias_in")"
  [ "$alias_abs" = "$canonical_abs" ] && continue
  if [ -L "$alias_abs" ]; then
    printf '  %-50s → %s\n' "$alias_abs" "$(readlink "$alias_abs")"
  else
    printf '  %-50s (symlink ではない!)\n' "$alias_abs"
  fi
done
