#!/usr/bin/env python3
"""
collect_codex_images.py — codex exec が生成した画像を、ログのsession idから決定的に回収する。

なぜ必要か（2026-07-14 実証）: 並列セッションに「generated_images直下の最新を拾え」と
指示すると他セッションの生成物を掴む競合が起きる（11枚中6枚がシャッフルされた実失敗）。
正しい回収 = ログの `session id:` → `$CODEX_HOME/generated_images/<session-id>/` 配下のみを
mtime昇順で並べる（タスクファイルの生成順と一致する）。

使い方:
  python3 collect_codex_images.py [--take-latest|--take-first] <codexログ> <出力名1> <出力名2> ...
  例: python3 collect_codex_images.py process/icons_1.log \
        assets/icon_factory.png assets/icon_rocket.png assets/icon_launchpad.png
生成枚数と出力名の数が一致しない場合はエラー終了する（黙って間引かない）。

edit-mode（参照画像を渡す生成）では、入力画像のコピーも session dir に保存されるため
枚数が出力名より多くなる（2026-08-27 実測。`count mismatch: ... has 2 images, but 1
output names given`）。この場合だけ `--take-latest`（生成物は後にできる）で明示的に
救済する。`--take-first` は逆側。**足りない側（生成 < 出力名）は救済しない** —
無音で欠落を埋めるのが最悪の失敗だから。
"""
import glob
import os
import shutil
import sys


def select_files(files, want, take=None):
    """mtime昇順の files から want 枚を選ぶ。take は None / "latest" / "first"。

    files は mtime 昇順で渡すこと。戻り値も mtime 昇順（= 生成順）を保つ。
    """
    if len(files) == want:
        return list(files)
    if take is not None and len(files) > want:
        return list(files[-want:]) if take == "latest" else list(files[:want])
    raise SystemExit(
        f"count mismatch: session dir has {len(files)} images, but {want} output names given"
        + (
            ""
            if len(files) < want
            else "（edit-mode では入力画像のコピーも保存される。"
            "生成物だけを取るなら --take-latest を付ける）"
        )
    )


def main():
    argv = sys.argv[1:]
    take = None
    rest = []
    for arg in argv:
        if arg == "--take-latest":
            take = "latest"
        elif arg == "--take-first":
            take = "first"
        else:
            rest.append(arg)
    if len(rest) < 2:
        print(__doc__)
        sys.exit(2)
    log, outs = rest[0], rest[1:]
    sid = None
    for line in open(log, errors="ignore"):
        if "session id:" in line:
            sid = line.split()[-1].strip()
            break
    if not sid:
        raise SystemExit(f"session id not found in {log}（codexがstdin待ちでハングした可能性。"
                         "起動時に < /dev/null を付けたか確認）")
    # session id からシート横断で generated_images を探索する。
    # codex2 等のマルチシートラッパーは CODEX_HOME を子プロセス内だけで切り替えるため、
    # 呼び出し側シェルの CODEX_HOME に頼ると別シートの生成物を見失う（2026-08-31 実測）。
    # sid は UUID で一意なので、候補ルート全部から <root>/generated_images/<sid> を探せば決定的。
    roots = []
    env_home = os.environ.get("CODEX_HOME")
    if env_home:
        roots.append(env_home)
    roots += sorted(glob.glob(os.path.expanduser("~/.codex*")))
    session_dirs = []
    for root in roots:
        d = os.path.join(root, "generated_images", sid)
        if os.path.isdir(d) and d not in session_dirs:
            session_dirs.append(d)
    if not session_dirs:
        raise SystemExit(
            f"generated_images/{sid} が見つからない（探索ルート: {', '.join(roots) or '~/.codex*'}）"
        )
    if len(session_dirs) > 1:
        raise SystemExit(f"session {sid} が複数シートに存在して曖昧: {session_dirs}")
    files = sorted(glob.glob(f"{session_dirs[0]}/*.png"), key=os.path.getmtime)
    try:
        picked = select_files(files, len(outs), take=take)
    except SystemExit as exc:
        raise SystemExit(f"session {sid}: {exc}") from None
    if len(picked) != len(files):
        skipped = [os.path.basename(f) for f in files if f not in picked]
        print(f"note: {take} で {len(skipped)} 枚を除外した: {', '.join(skipped)}")
    for src, dst in zip(picked, outs):
        os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
        shutil.copy(src, dst)
        print(f"{os.path.basename(src)[:16]}... -> {dst}")


if __name__ == "__main__":
    main()
