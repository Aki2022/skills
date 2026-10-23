# 人間確認とハッシュに基づく承認

正典はbiz_opsのSPEC-design-system-standardizationと
ADR-20261022-defer-approval-signatures。所有者が2026-10-22に署名必須のDeferredを承認した。

## 実装と境界

L0はkey・value_hash・approver・ISO dateを持つ行と媒体別JCS hashを照合する。
値/理由/散文/alias先/profileが変われば承認が失効する。
署名・鍵・agent・allowed_signersの設定は不要。承認者名と日付は記録であり、本人性の証明ではない。
同じ権限のAIによる台帳の偽造を技術的に防ぐ保証は持たない。
AIは実台帳を書かない。オーナー自身が対象を確認して記録する。
承認方式の簡素化への同意は、個別のデザイン値への承認ではない。

## 確認と記録

```sh
design-author approve --prepare --must --repo . > request.json
```

prepareはschema_version 2 / approval_scheme hash-jcs-v1の確認用データだけを返す。
値・理由・散文と対象hashを人間が確認する。部分承認は--target ds:web:item:<id>で絞る。
対象への明示承認後、オーナー自身が次を実行する。AIが記録操作を代行しない。

```sh
design-author approve --request request.json --approver owner --repo .
```

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
実パイロットのモデル回答と人間ラベル付きholdoutはまだ未投入。
固定L2 2回と5区分報告、出荷SHA/成果物hashとship_command/doctor、維持モードは実装済み。
出荷契約はship-gates.md、終了時の点検はmaintenance.md、A/B/Cの相互参照はdocs validatorが持つ。
実パイロットへの正式なチャネル導入・実台帳記入・判定台帳投入と全画面の合格は別工程。
署名は、AIから独立した承認者の認証を必要とする運用になった時に再検討する。
