#!/bin/bash
# 正典 ~/.agents/skills のチェックアウトが main で、最後に fetch した origin/main に追いついているかを見る。
#
# なぜ要るか: 正典は全エージェントの読み込み元で、同時に git の作業ツリーでもある。
# 2026-11-12〜13 に、別の作業がそこで feature ブランチに切り替えたまま残り、main に入った
# 13 コミットが実際の作業で読まれていなかった。古い skill が黙って読まれるだけで何も失敗しないので、
# 毎セッションの入口で 1 行だけ知らせる。正典で作業せず、skills の worktree で作業すること。
#
# fetch はしない（SessionStart をネットワークに依存させない）。origin/main は最後の fetch の時点。
# 未コミットの変更だけでは警告しない（別の作業の下書きが残るのは許す）。
# 正常時は何も出力しない。終了コード: 0 = 逸脱なし / 1 = 逸脱あり
set -uo pipefail

canon="$HOME/.agents/skills"
git -C "$canon" rev-parse --is-inside-work-tree >/dev/null 2>&1 || exit 0

branch="$(git -C "$canon" symbolic-ref --quiet --short HEAD 2>/dev/null)" || branch="(detached HEAD)"
if [ "$branch" != "main" ]; then
  echo "skills 正典が main ではない: ${branch}。読み込まれる skill は main と違う（正典で作業せず worktree を使う）" >&2
  exit 1
fi

if git -C "$canon" rev-parse --verify --quiet origin/main >/dev/null; then
  behind="$(git -C "$canon" rev-list --count HEAD..origin/main 2>/dev/null || echo 0)"
  if [ "${behind:-0}" -gt 0 ]; then
    echo "skills 正典の main が origin/main より ${behind} コミット遅れている（git -C ~/.agents/skills pull --ff-only）" >&2
    exit 1
  fi
fi
exit 0
