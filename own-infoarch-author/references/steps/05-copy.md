# 段 5 文案を書く

目的：各セクションの文案（`copy`）を、SSOT の声とトーン、証拠、固定値を使って書く。

## やること

1. SSOT の `voices`（声）の各特性の「〜だが」と「〜ではない」を読んでから書く。声は媒体が変わっても変えない。
2. 対象者の `tone` に合わせて調子を決める。
3. 数値・連絡先・免責など、固定値にあるものは**必ず `{{fct-NNNN}}` で参照**し、値を直接書かない（検査では IA_OUTLINE_FACT_LITERAL）。
4. 製品名・会社名は、SSOT の正式名称（`entities`）の表記で書く。
5. 平易な語で、結論を文の前に置く。1 文に言いたいことは 1 つ。
6. SSOT と証拠に無い事実を書かない。新しい事実が要るなら、SSOT の変更として別の PR で提案する。

## よくある失敗

- 媒体の慣習に引っ張られて、声が変わる（チラシだけ煽る、資料だけ硬い）。
- 数字を直接書いて、SSOT の更新に追従しなくなる。
- 証拠の無い最上級（「業界初」「No.1」）を書く。景品表示法の観点でも根拠が要る。

## アウトラインへの書き方

`pages[].sections[].copy`。

## 出典

- 平易な語、結論を前に：Nielsen Norman Group「How Users Read on the Web」（https://www.nngroup.com/articles/how-users-read-on-the-web/）。
- 文の長さの数値（1 文何語など）：出典なし・目安。この skill では数値の基準を持たない。
- No.1 表示の根拠：消費者庁「No.1表示に関する実態調査報告書」（https://www.caa.go.jp/policies/policy/representation/fair_labeling/survey）。
