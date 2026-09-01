#!/bin/bash
# origin-pptx 正典バッチランナー: モックアップ/イラストの codex image_gen 一括生成
#
# 使い方:  bash run_mockups.sh <デッキdir> <NN> [NN...]
#   例:    bash ~/.agents/skills/origin-pptx/scripts/run_mockups.sh "$PWD" 02 03 04
# 前提:    <デッキdir>/process/prompts/mockup_NN_prompt.txt が存在すること
# 並列度:  PARALLEL（既定3）
#
# 設計原則（2026-08-31 の実測失敗より・spec: skills docs/specs/pptx-content-eval-loop.md）:
# 1. 生成前に旧成果物を .stale に退避する——上書き再生成で失敗すると、旧ファイルの存在が
#    「完了」に見えて検分・レビューへ素通りする（stale-file pass-through）
# 2. 回収(collect)の失敗は必ず exit 非0 に反映する（FAIL の echo を握りつぶさない）
# 3. 最初の1枚をスモークとして直列実行し、シートのクレジット切れ等の即死を早期検出する
set -u
DECK="${1:?usage: run_mockups.sh <deck dir> <NN> [NN...]}"; shift
[ $# -ge 1 ] || { echo "no slide numbers given"; exit 2; }
cd "$DECK" || exit 2
PARALLEL="${PARALLEL:-3}"
SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CODEX_BIN=$(command -v codex2 || command -v codex) || { echo "codex not found"; exit 2; }
echo "using: $CODEX_BIN (parallel=$PARALLEL)"

fail=0
run_one() {
  local n=$1
  # 原則1: 旧成果物の退避（失敗時に stale が「完了」を偽装しないため）
  [ -f "process/mockup_${n}.png" ] && mv "process/mockup_${n}.png" "process/mockup_${n}.png.stale"
  "$CODEX_BIN" exec --sandbox workspace-write -c sandbox_workspace_write.network_access=true \
    --cd "$DECK" "$(cat "process/prompts/mockup_${n}_prompt.txt")" \
    < /dev/null > "process/mockup_${n}.log" 2>&1
  if python3 "$SKILL_DIR/scripts/collect_codex_images.py" --take-latest \
       "process/mockup_${n}.log" "process/mockup_${n}.png"; then
    rm -f "process/mockup_${n}.png.stale"
    echo "OK mockup_${n}"
  else
    echo "FAIL mockup_${n} (log: process/mockup_${n}.log / 旧版は .stale に退避済み)"
    return 1
  fi
}

# 原則3: スモーク1枚を直列で
first=$1; shift
run_one "$first" || { echo "SMOKE FAILED — シート/クレジットを確認してから再実行すること"; exit 1; }

# 残りを並列度 PARALLEL で
pids=(); count=0
for n in "$@"; do
  run_one "$n" & pids+=($!); count=$((count+1))
  if [ $((count % PARALLEL)) -eq 0 ]; then
    for p in "${pids[@]}"; do wait "$p" || fail=1; done; pids=()
  fi
done
for p in "${pids[@]}"; do wait "$p" || fail=1; done

# 原則2: 存在チェックは補助であり、collect 失敗が既に fail に反映されている
for n in "$first" "$@"; do
  [ -f "process/mockup_${n}.png" ] || { echo "MISSING mockup_${n}.png"; fail=1; }
done
exit $fail
