# 立ち上げ時の抽出

新しい repo で、既存の資料（企画書・公開中の LP・配布済みの資料・製品の仕様）から、positioning（`info_arch.yaml`）の下書き（原子と矢印）と、辞書（`docs/dictionary/`）の固定値・正式名称の下書きを起こす手順。positioning は own-positioning-author の持ち物で、辞書は own-doc-update の手順で足す。outline は成果物ごとに AI が作るので、ここでは作らない。

## やること

1. `own-positioning-author/templates/design/info_arch/info_arch.yaml` と `own-infoarch-author/templates/design/info_arch/`（channels と INFO_ARCH.md）を repo の `design/info_arch/` に写す。辞書の置き場所と形は own-doc-update（`docs/dictionary/`）。
2. 依頼主に、読んでよい既存の資料の場所を聞く（repo の中のパス、または URL）。資料を勝手に探し回らない。
3. 資料を読み、次の順に下書きする。どの原子にも `provenance: extracted:<repo 相対パス>#<見出し>` か `extracted:<url>` を書く。
   1. 正式名称（製品名・会社名・機能名）→ 辞書の `name`
   2. 固定値（数値・連絡先・免責）→ 辞書の `value`。期限を `expires` で決めるか `evergreen: true` にする（その値の中身の判断なので、決まらなければオーナーに聞く）
   3. セグメント（資料が誰に向けて書かれているか。払う人と使う人が違えば別のセグメント）
   4. positioning の原子と矢印（セグメントごとに、ジョブ・痛み・得たいもの・代替と満たす度合い・カテゴリーと所属・強み・提供価値。資料から読み取れる分だけ。承認はオーナーが提供価値ごとの経路で）
   5. 証拠（事例・数値・受賞。出典と時点を付け、segment と、支える原子・矢印への supports を書く）
4. **白紙から発案しない。** 資料から読み取れないものは `state: empty` にして、何が足りないかを依頼主に伝える。
5. 資料の間で矛盾する値（古い数字と新しい数字）は、新しい日付の方を `proposed` にし、古い方を説明に残す。
6. `infoarch-author check` で SSOT の検査を通す（アウトラインが 0 件なので、IA_NO_OUTLINES の失格だけが残るのが正常）。
7. SSOT の変更を PR にし、検査が緑なら自律マージする（人のレビューは置かない）。`empty` のまま残した原子と、資料の間で矛盾していた値は、PR の説明に一覧で書いて、オーナーが後から直せるようにする。

## 出典

- 内容を一度構造化し、表示は媒体側が決める（COPE / adaptive content）：Karen McGrane（https://karenmcgrane.com/talks/adapting-ourselves-to-adaptive-content/）。
- 抽出の手順そのもの：出典なし・目安（この skill の運用ルール）。
