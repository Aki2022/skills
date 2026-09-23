#!/bin/bash
# Antigravity/Gemini CLI 用の薄いラッパー。中身は check_session_fast.sh と同じ。
#
# 分ける理由は2つ。
#  1. agy の hook は stdin で JSON を受ける。読み捨てないと呼び出し側が詰まりうる。
#  2. SessionStart の戻り値スキーマが公開されていない。セッションを壊さないため
#     **常に exit 0** とし、逸脱は stderr に出すだけにする（agy はそれを cli.log に落とす）。
#     PreToolUse と違い allow/deny の判断は要らないので、decision は返さない。
cat > /dev/null 2>&1 || true
bash "$(cd "$(dirname "$0")" && pwd)/check_session_fast.sh" || true
exit 0
