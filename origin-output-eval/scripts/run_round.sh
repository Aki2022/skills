#!/usr/bin/env bash
# run_round.sh — 評価ループ1巡ぶんの「判定→record→収束判定」を機械的に通す（穴⑤への対処）。
#
# 目的: 「round 作成→判定→record→収束判定」を毎回手で組む余地を消す。
# 判定は再計算しない。judge_round.py と eval_state.py の exit code と出力を中継するだけ
# （再計算すると2箇所で判定がズレる）。
#
# 使い方:
#   run_round.sh <eval_dir> --thresholds <T.json> [--max-rounds N]
#
# eval_dir の構成（このスクリプトが前提とする）:
#   eval_dir/state.json          無ければ init する（--max-rounds は init のときだけ使う）
#   eval_dir/round_1/ round_2/ … 各 round のディレクトリ（呼び出し側が評価者出力を置く）
#
# exit code:
#   0 pass（judge_round が合格）
#   1 fail（judge_round が不合格・続行可）
#   2 入力の不備（引数不足・round_K/ が無い・judge_round/eval_state のスキーマ違反）
#   3 not-converging（前巡より指摘が減っていない。上限前でも終了）
#   4 max-rounds 到達（終了して最終報告へ）
#   5 前巡で pass 済み（終了して最終報告へ）
#
# judge_round.py の実装は差し替え可能にしてある: 環境変数 JUDGE_ROUND_BIN でコマンドを
# 上書きできる（テストは契約どおりの出力を返すダミーを指す）。本番既定値は同じディレクトリの
# judge_round.py。
set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
EVAL_STATE_PY="$SCRIPT_DIR/eval_state.py"
: "${JUDGE_ROUND_BIN:=python3 $SCRIPT_DIR/judge_round.py}"

if [ "$#" -lt 1 ]; then
  echo "run_round: 使い方: run_round.sh <eval_dir> --thresholds <T.json> [--max-rounds N]" >&2
  exit 2
fi

EVAL_DIR="$1"; shift
THRESHOLDS=""
MAX_ROUNDS=""

while [ "$#" -gt 0 ]; do
  case "$1" in
    --thresholds)
      THRESHOLDS="${2:-}"; shift 2 ;;
    --max-rounds)
      MAX_ROUNDS="${2:-}"; shift 2 ;;
    *)
      echo "run_round: 不明な引数 '$1'" >&2
      exit 2 ;;
  esac
done

if [ -z "$THRESHOLDS" ]; then
  echo "run_round: --thresholds が必要" >&2
  exit 2
fi

STATE_JSON="$EVAL_DIR/state.json"

# 1. state.json が無ければ init（--max-rounds は init のときだけ使う）
if [ ! -f "$STATE_JSON" ]; then
  if [ -n "$MAX_ROUNDS" ]; then
    python3 "$EVAL_STATE_PY" init "$STATE_JSON" --max-rounds "$MAX_ROUNDS"
  else
    python3 "$EVAL_STATE_PY" init "$STATE_JSON"
  fi
  init_rc=$?
  if [ "$init_rc" -ne 0 ]; then
    echo "run_round: eval_state init 失敗 (exit $init_rc)" >&2
    exit 2
  fi
fi

# 2. next を見る。exit 4（上限）/ 5（前巡 pass）ならそのコードで即終了。
next_out="$(python3 "$EVAL_STATE_PY" next "$STATE_JSON")"
next_rc=$?
echo "next: $next_out (exit $next_rc)"
if [ "$next_rc" -eq 4 ] || [ "$next_rc" -eq 5 ]; then
  echo "run_round: 終了して最終報告へ回せ（人間ゲートはこのループの外）"
  exit "$next_rc"
fi
if [ "$next_rc" -ne 0 ]; then
  # 契約上 next は 0/4/5 のみ。それ以外は state が壊れている等の入力の不備であって
  # 「品質の不合格」ではない。そのまま中継すると exit 1（不合格・続行可）に化けて、
  # 壊れた state が品質問題に見える。
  echo "run_round: eval_state next が契約外の exit ($next_rc)。state を直せ" >&2
  echo "$next_out" >&2
  exit 2
fi

# 3. 次巡番号 K を state から決め、round_K/ の存在を確認する
current_round="$(python3 -c "import json,sys; print(json.load(open(sys.argv[1]))['round'])" "$STATE_JSON")"
round_rc=$?
# state から round を読めなければ止める（空文字が算術で 0 と解釈され round_1 を見に行くのを防ぐ）。
# 実際には直前の next が先に落ちるため通常は到達しないが、next の実装が変わっても
# 「巡を取り違えたまま判定する」ことがないよう残してある。
if [ "$round_rc" -ne 0 ] || ! printf '%s' "$current_round" | grep -qE '^[0-9]+$'; then
  echo "run_round: $STATE_JSON から round を読めない（値: '${current_round}'）。state を直せ" >&2
  exit 2
fi
next_round=$((current_round + 1))
ROUND_DIR="$EVAL_DIR/round_${next_round}"
if [ ! -d "$ROUND_DIR" ]; then
  echo "run_round: $ROUND_DIR が無い。評価者出力を置いてから呼べ" >&2
  exit 2
fi

# 4. judge_round.py round_K --thresholds T --json-out round_K/verdict.json
VERDICT_JSON="$ROUND_DIR/verdict.json"
JUDGE_LOG="$ROUND_DIR/judge_round.log"
# shellcheck disable=SC2086
$JUDGE_ROUND_BIN "$ROUND_DIR" --thresholds "$THRESHOLDS" --round "$next_round" --json-out "$VERDICT_JSON" \
  > "$JUDGE_LOG" 2>&1
judge_rc=$?
echo "judge_round: round=$next_round exit=$judge_rc (log: $JUDGE_LOG)"
if [ "$judge_rc" -eq 2 ]; then
  echo "run_round: judge_round が入力の不備 (exit 2)。先に直せ" >&2
  cat "$JUDGE_LOG" >&2
  exit 2
fi
# judge_round の契約は exit 0/1/2 のみ。想定外の exit や verdict.json の欠落・壊れは
# 「不合格」ではなく「判定できていない」なので 2 で止める。混同すると、判定されていない巡を
# 不合格として数えて収束判定が汚れる。
if [ "$judge_rc" -ne 0 ] && [ "$judge_rc" -ne 1 ]; then
  echo "run_round: judge_round が契約外の exit ($judge_rc)。判定は行われていない" >&2
  cat "$JUDGE_LOG" >&2
  exit 2
fi
if ! python3 -c "import json,sys; d=json.load(open(sys.argv[1])); sys.exit(0 if 'findings_count' in d else 1)" "$VERDICT_JSON" 2>/dev/null; then
  echo "run_round: $VERDICT_JSON が無い・壊れている・findings_count が無い。判定は行われていない" >&2
  cat "$JUDGE_LOG" >&2
  exit 2
fi

# 5. eval_state.py record state.json --verdict round_K/verdict.json
record_out="$(python3 "$EVAL_STATE_PY" record "$STATE_JSON" --verdict "$VERDICT_JSON")"
record_rc=$?
echo "record: $record_out (exit $record_rc)"
if [ "$record_rc" -ne 0 ]; then
  # record の失敗は verdict の不備（findings_count 欠落・巡の取り違え等）。判定は記録されていない。
  echo "run_round: eval_state record 失敗 (exit $record_rc)。判定は記録されていない" >&2
  echo "$record_out" >&2
  exit 2
fi

# 6. eval_state.py converged state.json。exit 3 なら 3 を返して終了。
converged_out="$(python3 "$EVAL_STATE_PY" converged "$STATE_JSON")"
converged_rc=$?
echo "converged: $converged_out (exit $converged_rc)"
if [ "$converged_rc" -eq 3 ]; then
  echo "run_round: not-converging。終了して最終報告へ回せ"
  exit 3
fi

# 7. judge の結果が pass なら 0、不合格なら 1 を返す。
if [ "$judge_rc" -eq 0 ]; then
  echo "run_round: round=$next_round pass"
  exit 0
else
  echo "run_round: round=$next_round fail（続行可）"
  exit 1
fi
