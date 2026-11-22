# 人間確認とハッシュに基づく承認

正典はbiz_opsのSPEC-design-system-standardizationと
ADR-20261111-defer-approval-signatures・ADR-20261118-ai-records-owner-approvals。署名必須のDeferredは2026-10-22から実装済みで、所有者が2026-11-11に承認して記録した。
「記録はオーナー自身」は2026-11-18にオーナーの指示で「承認の判断はオーナー、記録はAI」に置き換えた。

## 実装と境界

L0はkey・value_hash・approver・ISO dateを持つ行と媒体別JCS hashを照合する。
値/理由/散文/alias先/profileが変われば承認が失効する。
署名・鍵・agent・allowed_signersの設定は不要。承認者名と日付は記録であり、本人性の証明ではない。
同じ権限のAIによる台帳の偽造を技術的に防ぐ保証は持たない。
承認の判断はオーナーがする。記録はAIがする。AIが記録してよいのは、会話やPRにオーナー自身の承認の言葉がある時だけ。
言葉の無い記録は道具が断る（`--owner-words` と `--approved-on` が必須）。言葉は記録の正しさを機械で証明しない。引用と日付を commit と PR に残し、後から人が辿れるようにする。
承認方式の簡素化への同意は、個別のデザイン値への承認ではない。

## 確認と記録

```sh
design-author approve --prepare --must --repo . > request.json
```

prepareはschema_version 2 / approval_scheme hash-jcs-v1の確認用データだけを返す。
値・理由・散文と対象hashを人間が確認する。部分承認は--target ds:web:item:<id>で絞る。
オーナーが会話やPRで対象を承認したら、AIが次を実行する。`--approved-on` は確認用データの各行の date と同じ日でなければ断る（違う日の承認は `prepare --date` で作り直す）。

```sh
design-author approve --request request.json --approver owner \
  --owner-words '<オーナーの承認の言葉をそのまま>' --approved-on <YYYY-MM-DD> --repo .
```

結果の `commit_trailer`（`Owner-Approval: 「…」 (日付)`）を commit メッセージと PR 本文に書く。
positioning は outline より先に承認するので、`IA_NO_OUTLINES` と `zero IA check targets` は prepare を止めない。
positioning を分けて承認する途中の「まだ足りない」（代替が 0 件・強みの比べ先が無い・満たす度合いの欠けなど、承認済みの部分から導く失格）も prepare を止めない。残りを承認すれば直るから。矢印の形・重複・セグメントをまたぐ矢印・参照の切れなど、下書きそのものの壊れは止める。
design/design.yaml の無い repo（IA だけ）では、IA の解決だけで対象を作り、記録する。

確認用データはorigin URLのdigestでrepositoryを識別し、URLそのものを出力しない。
記録前・記録中に現在のhashとrepositoryを照合し、変更されていれば拒否する。
台帳は既存の他の行を保持しatomic replaceする。同時記録はlockで拒否し、失敗時に部分記録を残さない。
記録後もL0に残件があればexit 2となり、recorded_rowsを別途報告する。
承認行やmanifestを第三者からコピーしただけでは、人間承認の根拠にならない。

## JCSと検証範囲

既存NodeでECMAScriptのprimitive表現とUTF-16順を使う。DESIGN_AUTHOR_NODEで選ぶ。
[RFC 8785](https://www.rfc-editor.org/rfc/rfc8785.html)に従い、Unicodeを正規化しない。
不正surrogate・非有限数・IEEE 754で精度を失う整数を拒否する。新規依存/ネットワークは不要。

自己テストは実CLI・実consumer・既存scannerを通るhookとmise入口で承認前の赤と承認後の緑を検証する。
これは人間が内容を判断した証明ではない。実パイロットの台帳は対象承認まで空のままにする。

判定台帳の読取・記録/移行、5区分報告、IA接続口はjudgments-and-modules.md。
実プロジェクトの回答取得と判定台帳への記録は別工程。人間ラベルが無い時はschema2で精度未測定と明示し、ラベルを発案しない。
固定L2 2回と5区分報告、出荷SHA/成果物hashとship_command/doctor、維持モードは実装済み。
出荷契約はship-gates.md、終了時の点検はmaintenance.md、A/B/Cの相互参照はdocs validatorが持つ。
実パイロットへの正式なチャネル導入・実台帳記入・判定台帳投入と全画面の合格は別工程。
署名は、AIから独立した承認者の認証を必要とする運用になった時に再検討する。
