---
name: own-docs-maintain
description: >-
  docs/ を AI のコンテキストとして自律的に維持する。close-session ごとに、機械修正（docs_hygiene --fix）の後、
  判断が要る整理 — 閉じる候補の判定と archive、参照されない guide/spec の降格、32KB 超や履歴混在の
  guide/spec の分割・圧縮 — を subagent が一次証拠で判定して適用し、判断の記録を repo に残す。人間の tick を
  待たない。予算（1 回の close で判定 1 本＋分割/圧縮 2 本まで）で token 消費を有界にする。
  必ず使うこと: own-session-close の Step 2（own-doc-update）の直後、docs/00_index.md を持つ repo で。
  「docs を整理して」「docs を圧縮して」「index が膨れている」「guide が大きすぎる」「docs を自動で維持」
  と言われた時。使わない場面: 規約・テンプレート・validator・hygiene の機械修正そのもの（own-doc-update
  が持つ）、コード変更に伴う guide の更新（各 issue の slice が持つ）、fixture repo（ws-loop-fixture 等）。
---

# own-docs-maintain

docs/ は次のセッションが最初に読むコンテキストで、放っておくと三つの経路で膨れる:
閉じない（完了した issue が active に残る）、書き足す（index や guide に経緯が積まれる）、
大きい guide を丸ごと参照する。own-doc-update の `docs_hygiene.py` はこのうち機械で決まる
部分を直し、決まらない部分を候補として数える。この skill はその候補を **subagent が一次証拠で
判定して適用する**。人間は判断を求められず、判断の記録（sweep / review ファイルへの刻印）を
後から読める。

決定の背景: 2026-09-21、yorisoi_kaigo の index は digest が 140 行を落とし、hot な guide/spec
の合計は 1.2 MB、32 KB 超が 12 本。sweep 149 件を人間が tick する設計は成立しない、と判断された
（ADR-20260921-autonomous-docs-curation、biz_ops）。

## 契約

- 入力: `docs_hygiene.py <repo> --fix --report --json` の結果と、それが書いた
  `docs/log/review-YYYYMMDD.md`・`docs/log/sweep-YYYYMMDD.md`。
- 出力: archive・分割・圧縮が commit 済み（docs/ のみ）で、判断の根拠が sweep / review ファイルに
  1 行ずつ残り、validator の error が実行前より増えていない。
- 予算（1 回の close-session あたり）: 判定 subagent 1 本（候補は全件）、分割または圧縮の subagent
  **2 本まで**。残りは次の close に持ち越す（review が毎回候補を出し直すので取りこぼさない）。
  1 本あたりの実測は判定 8〜28 万 token、分割 15〜28 万 token（2026-09-20〜21、sonnet）。
- 安全側の規則: 証拠が無い候補は keep（`— keep: undecided: <欠けている証拠>`）。archive は
  可逆（git と archive/ に残る）、削除は行わない。移動した本文は `docs/log/<slug>-history.md` に
  逐語で残し、事実は消さない。会社ドメイン・org ラベル・絶対パスは新規ファイルに書かない
  （vibe-guard が commit を止める）。
- 実行場所: 対象 repo が main で docs/ が clean ならその場、作業中 branch なら
  `docs/hygiene-<date>` ブランチの worktree。他セッションの未コミット docs 変更がある repo は触らず、
  その事実を報告する。

## 手順

1. **機械修正**: `python3 ~/.agents/skills/own-doc-update/scripts/docs_hygiene.py <repo> --fix --report --json`
   を実行し、JSON を保存する（R1〜R9 の件数、sweep 候補、review の split/demote 候補）。
2. **計画**: `python3 ~/.agents/skills/own-docs-maintain/scripts/plan.py <repo> <hygiene.json>` が、
   予算内で今回やる作業を JSON で出す: `judge`（sweep 候補があれば 1）、`docs`（分割/圧縮の対象、
   hot かつ大きい順に最大 2 本、各 `mode: split|condense`）。`docs` の選び方は
   「active な作業が参照している（hot）」＞「index が参照（warm）」＞ cold、同順位ならサイズ降順。
   cold の降格候補は `judge` に渡す。
3. **判定 subagent（sonnet）**: `references/judge-sweep.prompt.md` を repo パスとファイル名で埋めて
   起こす。sweep と review の archive 行を一次証拠（Completion・成果物の実在・git log・参照の有無）で
   `[x]` / `— keep: 理由` に書き分けさせる。終わったら
   `docs_hygiene.py <repo> --apply-sweep <sweep>` と `--apply-sweep <review>` を実行する。
4. **分割/圧縮 subagent（sonnet、`docs` の本数だけ並列）**: `mode: split` は
   `references/split-doc.prompt.md`、`mode: condense` は `references/condense-doc.prompt.md`。
   split は「現在の真実を task 単位の guide（32 KB 以下）に分け、経緯は history、決定は ADR、元パスは
   frontmatter を保った pointer」。condense は「一つの guide のまま、重複・経緯・死んだ参照を除いて
   現在の真実だけに書き直し、除いた本文は history に逐語で残す」。どちらも全行の配置を
   プログラムで検証させる。subagent は commit しない。
5. **統合**: index の Guides/Specs 行を追加（link ＋ 200 字以内、status 語なし）、
   `validate_repo_docs.py <repo>` で error が実行前以下であることを確認、`docs_hygiene.py <repo> --json`
   で R9（digest の落とし行・32 KB 超・hot 集合）の前後を記録、docs/ だけを **1 commit**。
   commit は単独で実行して成否を確認し、その後に別コマンドで worktree を片付ける
   （`git worktree remove` に `-f` を付けない — 2026-09-21 に未コミット成果を消した）。
6. **報告**: before/after の R9、archive 件数、分割/圧縮した文書、持ち越した候補数を 1 表で返す。
7. **push まで**: own-session-close の中で呼ばれた場合は Step 3（own-git-clean）が main を同期する。
   **単独で回した場合は自分で push する**（main が push 可能なら main を、guard のある repo は branch を push して
   PR）。2026-09-21 に 14 repo で「commit したが push しない」まま残り、人間の指摘で気づいた。
   commit だけで終わる報告は完了ではない。

## いつ止まるか

- validator の error が実行前より増えた（統合の誤り）。差分を残して報告し、commit しない。
- vibe-guard が commit を止めた。原因を直す（マスク漏れ・絶対パス）。allowlist は足さない。
- 対象 repo の docs/ に他セッションの未コミット変更がある。触らず報告。

## Resources

- `references/judge-sweep.prompt.md` — 判定 subagent への指示（sweep と review の archive 行）
- `references/split-doc.prompt.md` — 分割 subagent への指示
- `references/condense-doc.prompt.md` — 圧縮 subagent への指示
- `scripts/plan.py <repo> <hygiene.json> [--budget-docs 2]` — 予算内の作業計画を JSON で出す
- `scripts/tests/test_plan.py`
