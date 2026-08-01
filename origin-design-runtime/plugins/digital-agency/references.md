---
title: Digital Agency Plugin — References & License Gate
license_gate_status: APPROVED
updated_at: 2026-07-04
---

# Digital Agency Plugin — References & License Gate

デジタル庁の公開アセットを利用するためのライセンス・出典・利用条件の整理。
**このゲートが承認されるまで、資産値（tokens実値・フォント・ロゴ・テンプレート・境界データ）を
成果物に反映してはいけない**（`plugin.yaml` の `usage_policy.license_gate_approved` は `false` のまま）。
承認されたら本ファイル末尾に承認記録を追記し、`license_gate_approved: true` に更新する。

## アセット別ライセンス一覧（調査結果 2026-07-04）

| アセット                                   | ライセンス                                   | 帰属表示                                               | 再配布/利用条件                                                                                                 | 出典                                                            |
| ------------------------------------------ | -------------------------------------------- | ------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------- |
| デザインデータ（Figma）                    | CC BY 4.0                                    | 原則必要（大幅改変・UI部品化時は不要）                 | 改変物を「デジタル庁作成」と誤認させる形での公表は禁止                                                          | design.digital.go.jp                                            |
| コードスニペット HTML版                    | MIT                                          | 未改変で公開再配布する場合のみ必要                     | PR受付なし。改変可                                                                                              | github.com/digital-go-jp/design-system-example-components-html  |
| コードスニペット React版                   | MIT                                          | 同上                                                   | 同上                                                                                                            | github.com/digital-go-jp/design-system-example-components-react |
| アイコン（Material Symbols 由来分）        | Apache License 2.0                           | Apache 2.0 に従う                                      | 改変・再配布可                                                                                                  | Figmaデザインデータ内                                           |
| フォント Noto Sans JP / Noto Sans Mono     | SIL Open Font License 1.1                    | OFL に従う                                             | Webフォント化・埋め込み可。**フォント単体販売は禁止**。本プロジェクト方針では**同梱・再配布しない**（参照のみ） | design.digital.go.jp/dads/foundations/typography                |
| policy-dashboard-assets（テンプレ/テーマ） | Public Data License v1.0 (PDL1.0)            | 未改変配布時は出典記載。ダッシュボード用に改変時は不要 | LICENSEファイルなし・READMEに記載                                                                               | github.com/digital-go-jp/policy-dashboard-assets                |
| 行政区域ポリゴンデータ                     | 国土数値情報DLサイト規約に従う（別途同意要） | サイト規約に従う                                       | **別ライセンス。利用前に規約確認必須**                                                                          | nlftp.mlit.go.jp/ksj/other/agreement.html                       |
| デジタル庁ロゴ・ロゴタイプ                 | 個別申請・制限あり                           | —                                                      | **本プロジェクトでは使用しない**（公式性誤認防止）                                                              | digital.go.jp/applications/logotype                             |

## 本プロジェクトでの取り扱い方針（要承認）

1. **コードスニペット（MIT）**: 参照・改変利用可。改変して UI 部品化する前提のため帰属表示は原則不要だが、
   references として出典を残す。
2. **デザインデータ（CC BY 4.0）**: tokens（色・タイポ・余白）の**思想・役割**は参照するが、値を写す場合は
   出典を明記。「デジタル庁が作成した」と誤認させない。
3. **フォント（OFL）**: フォントファイルは**同梱・再配布しない**。DESIGN.md では「Noto Sans JP を前提」と
   記述するに留め、実ファイルはユーザー環境/Webフォント参照に委ねる。
4. **ロゴ・紋章**: 使用しない。公的機関の公式資料と誤認される表現を作らない。
5. **境界ポリゴンデータ**: 国土数値情報の規約を個別確認するまで同梱・反映しない。
6. **policy-dashboard-assets（PDL1.0）**: ダッシュボード用に改変利用。未改変再配布時のみ出典記載。

## 出典URL

- デジタル庁デザインシステム: https://www.digital.go.jp/policies/servicedesign/designsystem
- デザインシステムβ版サイト: https://design.digital.go.jp/
- 利用上の注意事項: https://design.digital.go.jp/introduction/notices/
- タイポグラフィ（フォント/OFL）: https://design.digital.go.jp/dads/foundations/typography/
- HTML版コードスニペット: https://github.com/digital-go-jp/design-system-example-components-html
- React版コードスニペット: https://github.com/digital-go-jp/design-system-example-components-react
- policy-dashboard-assets: https://github.com/digital-go-jp/policy-dashboard-assets
- ダッシュボード実践ガイドブック: https://www.digital.go.jp/resources/dashboard-guidebook
- ロゴタイプ利用申請: https://www.digital.go.jp/en/applications/logotype
- 国土数値情報DLサイト規約: https://nlftp.mlit.go.jp/ksj/other/agreement.html

## 承認記録（ゲート）

- 状態: **APPROVED**（承認済み）
- 承認日: 2026-07-04
- 承認者: リポジトリオーナー（Aki）
- 承認した利用範囲: 上記「本プロジェクトでの取り扱い方針」の全項目をそのまま承認。
  - コードスニペット（MIT）: 改変利用可
  - デザインデータ/tokens思想（CC BY 4.0）: 値を写す場合は出典明記・誤認防止
  - フォント（OFL）: ファイル同梱・再配布せず、前提記述のみ
  - Material Symbols アイコン（Apache 2.0）: 利用可
  - policy-dashboard-assets（PDL1.0）: 改変利用可
  - 行政区域ポリゴンデータ: 国土数値情報規約の個別確認まで反映しない（今回は使わない）
  - ロゴ・紋章: 使用しない
- 方針からの変更点: なし
- 対応: `plugin.yaml` の `usage_policy.license_gate_approved` を `true` に更新済み。
