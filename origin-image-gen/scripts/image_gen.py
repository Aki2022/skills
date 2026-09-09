#!/usr/bin/env python3
"""image_gen.py — 画像生成の完結した窓口（組み立て → 実行 → 回収 → 完了判定）。

使い方:
  image_gen.py --fields <fields.json> --out <path1> [<path2> ...]
               [--anchor <img> ...] [--take-latest|--take-first]
               [--workdir <dir>] [--dry-run]

--workdir が git リポジトリの外なら `--skip-git-repo-check` を窓口が自動で足す
（呼び出し側にフラグを要求しない）。`--skip-git-check` はその強制用。

fields.json は器の10ラベルをキーに持つ object（値は文字列）:
  Use case / Asset type / Style references / Style / House structure /
  Composition / Content / Color palette / Constraints / Avoid

なぜ窓口にするか: 生成の指示だけを書いて実行と回収を呼び出し側に任せると、
(a) プロンプトの「Xとして保存して」を保存の証拠にする、(b) 旧成果物が残ったまま成功に見える、
(c) 並列セッションの生成物を掴む、(d) stdin を閉じずハングする、が繰り返し起きた。
どれも呼び出し側の規律に依存する形だったので、窓口の内側に閉じ込める。

exit: 0 全枚数が揃って回収できた / 1 不足・回収失敗・codex 失敗（成果物を使ってはいけない）
      2 入力不備（fields のキー欠落・JSON 不正・--out が0件・禁止フラグの混入。判定していない）

scope: 本スクリプトは器の組み立てと実行機構・回収・完了判定だけを持つ。
スタイルや配色の値（ブランド軸）と媒体適合の制約（媒体軸）は持たず、呼び出し側が渡す。
生成物の品質評価も行わない（origin-output-eval が持つ）。
"""
from __future__ import annotations

import argparse
import json
import os
import shlex
import shutil
import subprocess
import sys
from pathlib import Path

FIELD_ORDER = [
    "Use case", "Asset type", "Style references", "Style", "House structure",
    "Composition", "Content", "Color palette", "Constraints", "Avoid",
]
# このフラグは窓口が組み立てるコマンドに現れてはならない。安全モードで image_gen は動作し、
# かつ Claude Code の分類器がハードブロックするため AI からは実行できない。
FORBIDDEN = "--dangerously-bypass-approvals-and-sandbox"

SCOPE = (
    "scope: image_gen — 器の組み立て・codex 実行・回収・完了判定までを1つの窓口で通す。"
    "スタイルや配色の値・生成の要否・生成後の配置・品質評価は判定しない。"
)


class InputError(Exception):
    """入力不備。exit 2（判定は行われていない）。"""


def resolve_codex_bin() -> list[str]:
    """実行コマンドを解決する。テストは IMAGE_GEN_CODEX_BIN で差し替える。

    codex2 は CODEX_HOME=~/.codex-seat2 で同じバイナリを起動する別シートのラッパー。
    シートごとに利用枠が分かれるため、あればそちらを優先する。
    """
    injected = os.environ.get("IMAGE_GEN_CODEX_BIN")
    if injected:
        return shlex.split(injected)
    for name in ("codex2", "codex"):
        found = shutil.which(name)
        if found:
            return [found]
    raise InputError("codex コマンドが見つからない（codex2 も codex も PATH に無い）")


def load_fields(path: Path) -> dict[str, str]:
    try:
        data = json.loads(path.read_text())
    except Exception as e:  # noqa: BLE001
        raise InputError(f"{path.name}: JSON として読めない ({e})") from None
    if not isinstance(data, dict):
        raise InputError(f"{path.name}: トップレベルが object でない")
    missing = [k for k in FIELD_ORDER if k not in data]
    if missing:
        raise InputError(f"{path.name}: 器のフィールドが欠落 {missing}")
    extra = [k for k in data if k not in FIELD_ORDER]
    if extra:
        raise InputError(f"{path.name}: 器に無いキー {extra}（器は10フィールド固定）")
    for k, v in data.items():
        if not isinstance(v, str):
            raise InputError(f"{path.name}: {k} の値が文字列でない")
        if FORBIDDEN in v:
            raise InputError(
                f"{path.name}: {k} の値に禁止フラグ {FORBIDDEN} が含まれる"
                "（フィールド値経由でコマンドに混入させることはできない）"
            )
    return {k: data[k] for k in FIELD_ORDER}


def build_prompt(fields: dict[str, str], anchors: list[str]) -> str:
    f = dict(fields)
    if anchors:
        # アンカーは edit-mode で渡す。view_image を先に踏ませないと参照が無視されることがある。
        listing = ", ".join(anchors)
        f["Style references"] = (
            f"First view these style_ref images with view_image, then generate in edit-mode: "
            f"{listing}. Match their visual style exactly — same treatment, same annotation "
            f"devices, same density, same color roles."
        )
    return "\n".join(f"{label}: {f[label]}" for label in FIELD_ORDER)


def inside_git_repo(workdir: Path) -> bool:
    """workdir が git の作業ツリー内か。外なら codex は「Not inside a trusted directory」で即終了する。

    呼び出し側にフラグを要求しない（窓口は呼び出し側の規律に依存しない）。
    """
    r = subprocess.run(["git", "-C", str(workdir), "rev-parse", "--is-inside-work-tree"],
                       capture_output=True, text=True)
    return r.returncode == 0 and r.stdout.strip() == "true"


def build_command(codex_bin: list[str], workdir: Path, prompt: str, skip_git: bool) -> list[str]:
    """安全モード固定。フルバイパスは組み立てない。"""
    cmd = [
        *codex_bin, "exec",
        "--sandbox", "workspace-write",
        "-c", "sandbox_workspace_write.network_access=true",
        "--cd", str(workdir),
    ]
    if skip_git:
        cmd.append("--skip-git-repo-check")
    cmd.append(prompt)
    assert FORBIDDEN not in cmd, "禁止フラグが組み立てに載った"
    return cmd


def stash_existing(outs: list[Path]) -> list[Path]:
    """完了判定3原則(1): 実行前に旧成果物を .stale へ退避する。

    退避しないと生成が失敗しても旧ファイルが残り成功に見える（stale pass-through）。
    退避したファイルは成果物パスとして返さない。
    """
    stashed = []
    for o in outs:
        if o.exists():
            dst = o.with_name(o.name + ".stale")
            if dst.exists():
                dst.unlink()
            o.rename(dst)
            stashed.append(dst)
    return stashed


def main(argv=None) -> int:
    print(SCOPE)
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("--fields", required=True)
    ap.add_argument("--out", nargs="+", required=True)
    ap.add_argument("--anchor", action="append", default=[])
    ap.add_argument("--take-latest", action="store_true")
    ap.add_argument("--take-first", action="store_true")
    ap.add_argument("--workdir", default=".")
    ap.add_argument("--skip-git-check", action="store_true",
                    help="git リポジトリ内でも --skip-git-repo-check を強制する（通常は自動判定）")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args(argv)

    try:
        if a.take_latest and a.take_first:
            raise InputError("--take-latest と --take-first は同時に指定できない")
        outs = [Path(p) for p in a.out]
        fields = load_fields(Path(a.fields))
        for anchor in a.anchor:
            if not Path(anchor).is_file():
                raise InputError(f"アンカー画像が無い: {anchor}")
        workdir = Path(a.workdir).resolve()
        codex_bin = resolve_codex_bin()
        prompt = build_prompt(fields, a.anchor)
        # git リポジトリ外なら窓口が自動で --skip-git-repo-check を足す。
        # 呼び出し側が知っている必要はない（--skip-git-check は強制用の上書き）。
        skip_git = a.skip_git_check or not inside_git_repo(workdir)
        cmd = build_command(codex_bin, workdir, prompt, skip_git)
    except InputError as e:
        print(f"image_gen: 入力不備 — {e}", file=sys.stderr)
        return 2

    outdir = outs[0].parent
    outdir.mkdir(parents=True, exist_ok=True)
    prompt_path = outdir / "prompt.txt"
    prompt_path.write_text(prompt)
    log_path = outdir / "run.log"

    if a.dry_run:
        print(f"prompt: {prompt_path}")
        print("command: " + " ".join(shlex.quote(c) for c in cmd[:-1]) + " <prompt>")
        print(f"outputs({len(outs)}): " + ", ".join(str(o) for o in outs))
        return 0

    stashed = stash_existing(outs)
    if stashed:
        print(f"stashed: {len(stashed)} 件を .stale へ退避した（成果物としては返さない）")

    # < /dev/null は必須。stdin を閉じないと codex がハングする。
    with open(log_path, "w") as lf, open(os.devnull) as devnull:
        proc = subprocess.run(cmd, stdout=lf, stderr=subprocess.STDOUT, stdin=devnull)
    print(f"codex: exit={proc.returncode} (log: {log_path})")
    if proc.returncode != 0:
        print("image_gen: codex が失敗した。成果物は使ってはいけない", file=sys.stderr)
        return 1

    collect = Path(__file__).resolve().parent / "collect_codex_images.py"
    take = ["--take-latest"] if a.take_latest else (["--take-first"] if a.take_first else [])
    r = subprocess.run(
        ["python3", str(collect), *take, str(log_path), *[str(o) for o in outs]],
        capture_output=True, text=True,
    )
    sys.stdout.write(r.stdout)
    sys.stderr.write(r.stderr)
    if r.returncode != 0:
        # 完了判定3原則(2): 不足は救済しない。埋めると「成功に見えて中身が違う」になる。
        print("image_gen: 回収に失敗した（枚数不足は埋めない）。成果物は使ってはいけない",
              file=sys.stderr)
        return 1

    missing = [o for o in outs if not o.exists()]
    if missing:
        print(f"image_gen: 回収は成功を返したが出力が無い: {missing}", file=sys.stderr)
        return 1
    print("artifacts:")
    for o in outs:
        print(f"  - {o}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
