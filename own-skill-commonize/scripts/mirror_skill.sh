#!/bin/bash
# 第三者 skill を正典へミラーし、台帳へ登録し、全 root へ配線し、検査するまでを通す。
#
# なぜ script にするか: この4手順は毎回同じで、途中で止めると必ず壊れる。
#   ミラーだけ    → 配線されず Claude からしか見えない
#   台帳を忘れる  → 第三者は git 追跡外なので、消えたら二度と戻せない
#   配線だけ      → 実体が無く宙を指す
# 手で4回やると台帳を忘れる余地が残る。実測で mirrors.yaml は 45 本中 10 本しか
# 登録されていなかった（2026-09-23）。
#
# Codex の $skill-installer は使わないこと。あれは $CODEX_HOME/skills へ実体を書くので、
# その skill は 1 席だけで使え、Claude / Gemini から見えず、git にも台帳にも残らない。
set -euo pipefail

CANON="${AGENTS_SKILLS_ROOT:-$HOME/.agents/skills}"
LEDGER="$CANON/mirrors.yaml"
PER_SKILL_ROOTS=(
  "$HOME/.codex/skills"
  "$HOME/.codex-private/skills"
  "$HOME/.codex-seat2/skills"
  "$HOME/.gemini/config/skills"
)

usage() {
  cat <<'USAGE'
使い方:
  mirror_skill.sh --upstream <owner/repo> --path <repo 内のパス> --name <正典でのディレクトリ名>
                  --license <ライセンス> [--version <sha/tag>] [--note <注記>]

例:
  mirror_skill.sh --upstream openai/skills \
                  --path skills/.curated/gh-fix-ci \
                  --name gh-fix-ci \
                  --license Apache-2.0

やること:
  1. npx degit で上流から正典へ複製
  2. mirrors.yaml へ追記し、書けたことを再解析で確認する（報告だけで済ませない）
  3. .gitignore へ追加（第三者は git で持たない。忘れると PUBLIC repo へ出る）
  4. per-skill 配線先 4 つへ symlink（Claude 3席と Antigravity は棚ごと symlink なので自動）
  5. audit_skill_wiring.py で検査

やらないこと:
  - ライセンスの判断（--license は実物を読んで人が渡す。再配布の可否を決めるため）
  - 上流の版が正しいかの判断
USAGE
}

UPSTREAM= UPPATH= NAME= LICENSE= VERSION= NOTE=
while [ $# -gt 0 ]; do
  case "$1" in
    --upstream) UPSTREAM="$2"; shift 2;;
    --path) UPPATH="$2"; shift 2;;
    --name) NAME="$2"; shift 2;;
    --license) LICENSE="$2"; shift 2;;
    --version) VERSION="$2"; shift 2;;
    --note) NOTE="$2"; shift 2;;
    -h|--help) usage; exit 0;;
    *) echo "不明な引数: $1" >&2; usage >&2; exit 2;;
  esac
done

declare -A REQUIRED=( [UPSTREAM]=--upstream [UPPATH]=--path [NAME]=--name [LICENSE]=--license )
for v in UPSTREAM UPPATH NAME LICENSE; do
  [ -n "${!v}" ] || { echo "${REQUIRED[$v]} は必須です" >&2; usage >&2; exit 2; }
done

case "$NAME" in
  own-*) echo "エラー: own- は自作の接頭辞です。第三者に付けてはいけません" >&2; exit 2;;
esac
[ -e "$CANON/$NAME" ] && { echo "エラー: 正典に既にあります: $CANON/$NAME" >&2; exit 2; }

REINSTALL="npx degit $UPSTREAM/$UPPATH $CANON/$NAME"
REINSTALL_SHOWN="${REINSTALL//$HOME/\~}"

echo "1. 上流から複製: $UPSTREAM/$UPPATH → ${CANON/#$HOME/~}/$NAME"
npx --yes degit "$UPSTREAM/$UPPATH" "$CANON/$NAME"
[ -f "$CANON/$NAME/SKILL.md" ] || { echo "エラー: SKILL.md が無い。パスを確認してください" >&2; exit 1; }

FM_NAME="$(grep -m1 '^name:' "$CANON/$NAME/SKILL.md" | sed 's/^name: *//' | tr -d '"'"'"'')"
if [ -n "$FM_NAME" ] && [ "$FM_NAME" != "$NAME" ]; then
  echo "エラー: frontmatter の name ($FM_NAME) とディレクトリ名 ($NAME) が違います。" >&2
  echo "       第三者は上流名でミラーする規約なので、--name を $FM_NAME にしてください。" >&2
  exit 1
fi

echo "2. 台帳へ登録: ${LEDGER/#$HOME/\~}"
python3 - "$LEDGER" "$NAME" "$UPSTREAM" "$UPPATH" "${VERSION:-unknown（取得時の版を記録していない）}" \
  "$LICENSE" "$REINSTALL_SHOWN" "$NOTE" <<'PY'
import sys, datetime
ledger, name, up, path, ver, lic, cmd, note = sys.argv[1:9]
lines = [f"  - dir: {name}",
         f"    upstream: {up}",
         f"    upstream_path: {path}",
         f"    upstream_version: {ver}",
         f"    fetched_at: {datetime.date.today().isoformat()}",
         f"    license: {lic}",
         f"    reinstall: {cmd}"]
if note:
    lines.append(f"    note: {note}")
raw = open(ledger).read()
block = "\n".join(lines) + "\n"
# mirrors: の並びの末尾＝ retired: の直前へ入れる。retired: の直前の行はコメントで
# あることが多く、raw.index("\nretired:") の位置へそのまま連結するとコメント行の
# 末尾に貼り付いて YAML から消える。2026-09-23 にそれで gh-fix-ci が台帳から消え、
# 直前の wrangler のフィールドが上書きされた（check_mirrors.sh は緑のままだった）。
# 必ず行頭から始まるように改行を挟む。
marker = "\nretired:"
if marker in raw:
    i = raw.index(marker)
    raw = raw[:i].rstrip("\n") + "\n\n" + block + raw[i:]
else:
    raw = raw.rstrip("\n") + "\n" + block
open(ledger, "w").write(raw)

# 書けたことを再解析で確かめる。報告だけで済ませない。
import yaml
d = yaml.safe_load(open(ledger))
dirs = [m["dir"] for m in d.get("mirrors", [])]
if name not in dirs:
    raise SystemExit(f"   ERROR 台帳へ書けていない: {name}（mirrors {len(dirs)} 件）")
if len(dirs) != len(set(dirs)):
    raise SystemExit("   ERROR 台帳に重複した dir がある")
for m in d["mirrors"]:
    if m["dir"] not in m.get("reinstall", ""):
        raise SystemExit(f"   ERROR {m['dir']} の reinstall が別の skill を指している")
print(f"   登録: {name}（mirrors {len(dirs)} 件・整合 OK）")
PY

echo "3. .gitignore へ追加（第三者は git で持たない）"
python3 - "$CANON/.gitignore" "$NAME" <<'GIPY'
import sys
gi, name = sys.argv[1], sys.argv[2]
lines = open(gi).read().splitlines()
if f"/{name}/" in lines:
    print(f"   既にある: /{name}/")
else:
    lines.append(f"/{name}/")
    open(gi, "w").write("\n".join(lines) + "\n")
    print(f"   追加: /{name}/")
GIPY

echo "4. per-skill 配線先へ symlink"
for root in "${PER_SKILL_ROOTS[@]}"; do
  [ -d "$root" ] || { echo "   skip（不在）: ${root/#$HOME/\~}"; continue; }
  ln -s "$CANON/$NAME" "$root/$NAME"
  echo "   ${root/#$HOME/\~}/$NAME"
done

echo "5. 検査"
python3 "$(cd "$(dirname "$0")" && pwd)/audit_skill_wiring.py"
