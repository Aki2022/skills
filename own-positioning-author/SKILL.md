---
name: own-positioning-author
description: "repo 直下 design/info_arch/info_arch.yaml の positioning（なぜ・誰に・何を。セグメントごとのジョブ・痛み・得たいもの・代替・カテゴリー・強み・提供価値・証拠と、それらを結ぶ矢印）を読み書きし、検査し、オーナーに見せる承認依頼のページを機械で作る。承認は提供価値ごとの経路の単位。「positioning を作って」「ポジショニングを整理して」「承認依頼を作って」「承認のページを出して」「提供価値の経路を見せて」「承認を読み戻して」で使う。成果物ごとの outline（誰に・何を・どの順で）は own-infoarch-author、見た目は own-design-author と own-web-build。"
---

# own-positioning-author — positioning と承認依頼

positioning は 3 つの layer（positioning・infoarch・design）の最初の 1 つで、repo に 1 つ。サービスを「誰の・どのジョブと痛みに・どの代替と比べて・どの強みから生まれる・どの提供価値で・どの証拠が支えるか」で言う。下の layer（infoarch の outline）は、その id を引く。引けなければ失格。
`design/info_arch/info_arch.yaml`（スキーマ版 5）が正典で、オーナーが**提供価値ごとの経路**（主張→差別化の根拠→前提→証拠）を承認する。この skill は、それを読み書きし（作る側）、形・参照・id・経路の失格を検査し（保守する側）、承認依頼の表示を作る。
成果物ごとの outline、描画後の確認、出荷ゲートは own-infoarch-author が持つ（その検査が、この skill の positioning を読む）。

## 重要ルール

- **承認の状態を書かない。** `state` は `empty` / `proposed` だけ。承認は `design/approvals.yaml` の `ia:path:<提供価値 id>` の行（経路全体のハッシュ）が持ち、オーナーの承認の言葉を受けた AI が own-design-author の `approve` で記録する（ADR-20261118-ai-records-owner-approvals）。この skill は台帳に書かない。
- **正式名称と固定値は持たない。** `docs/dictionary/` の辞書が持つ（SPEC-deliverable-production-layers 要件 9）。声は廃止した（要件 8）。スキーマ版 4 の木は警告つきで読み、`infoarch-author migrate-dictionary` で辞書へ移す。
- **白紙から発案しない。** 資料から読み取れないものは `state: empty` にして、足りないものを依頼主に伝える。値の出所は `provenance` に書く。
- **提供価値 1 つは経路 1 つ。** 経路の中身が 1 つでも変わると、その経路の承認が外れる。共有する原子（セグメント・ジョブ・所属カテゴリーなど）を直すと、それを含む経路がまとめて外れる。
- **対象 0 件は失格。** 提供価値が 0 件なら終了コード 2（IA_NO_VALUES）。

## 使い方

```sh
positioning-author check --root . [--strict]      # positioning の決定論の検査。終了コード 0 / 1（警告）/ 2（失格）。--strict は全部の経路を出荷の outline が指すものとして失格にする
positioning-author positioning --root .           # セグメントごとの positioning と、提供価値ごとの経路（比べる集合・失格・承認済みか）
positioning-author targets --root .               # 承認の対象（ia:path:<提供価値 id> と経路のハッシュ）。出荷が要るキーは infoarch-author targets
positioning-author request --root . --out DIR     # オーナーに見せる承認依頼のページ（○ / ×）と、approve --prepare の行を書く。既存のファイルは上書きしない
positioning-author request --root . --read-decisions rows.json  # オーナーが保存した ○ / × を、今の経路と突き合わせて読み戻す（ok / ng / stale / undecided）
positioning-author review --root . --request design/approval-request-ia-positioning.json  # 承認の束の中身を読める形で。古い束は失格
positioning-author migrate-approvals --root . [--write]  # 原子と矢印ごとの旧い承認の行を、経路の行へ一度だけ移す
```

`scripts/positioning-author` を PATH に置き、PyYAML と jsonschema のある Python を `DESIGN_AUTHOR_PYTHON` で選ぶ（自動インストールしない。無ければ終了コード 2）。repo の設定にローカルパスを書かない。
新しい repo では `templates/design/info_arch/info_arch.yaml` を写す（outline のチャネルと散文は own-infoarch-author の `templates/`）。

## 承認依頼を出す手順

1. `positioning-author check --root .` を緑にする（失格が残る経路は、ページの「決めてほしいこと」に並ぶ）。
2. 出荷が要る経路を示すなら `infoarch-author targets --root . > targets.json` を作り、`--must-file targets.json` を付ける。
3. `positioning-author request --root . --out <一時フォルダ>`。出力は `approval-request.html`（自己完結の 1 枚）と `approval-request.json`（束の行）。
4. HTML を**非公開の Artifact** として出す（`capabilities: {db: {}}`。artifact-capabilities を読んでから）。repo やホームにファイルを残さない。ページには経路ごとのカード（主張→差別化の根拠→前提→証拠、ハッシュ先頭 12 桁、決めてほしいこと）と、○（承認する）・×（差し戻す）・メモがあり、選ぶと Artifact の保存先 `decisions/<提供価値 id>` に残る。
5. オーナーが選んだら、`ArtifactData` の list で `decisions` を読み戻し、`--read-decisions` に渡す。stale（経路の中身が変わった後の判断）は数えない。**○ は承認の言葉の代わりにならない。** 台帳に記録するのは、オーナーの承認の言葉（と日付）があるものだけ。

## 参照

- `references/contract.md` — スキーマ版 5 の原子・矢印・経路・承認の契約、承認依頼の保存先、検査の code。
- `references/schema/info_arch.v5.schema.json`（読む形）・`info_arch.v4.schema.json`（警告つきで読む旧い形）。
- `fixtures/positioning-red-cases.json` — 要件 5・要件 2・要件 15 の失格条件と赤ケースの対応表。outline を持つ `infoarch-author check` と、outline 無しの `check --strict` の両方で走らせる（要件 15 は outline が要るので infoarch の側だけ）。自己テストは `uv run --no-project --with-requirements <skill>/scripts/tests/requirements.txt -- pytest -q <skill>/scripts/tests`。
