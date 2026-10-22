# 署名付き承認への改訂（実装途中）

正典はbiz_opsのSPEC-design-system-standardizationと
ADR-20261022-approval-provenance-and-ship-gate-are-structural（PR #70、8f10578）。

## 現在の境界

旧L0のtoken_path/item_id行は署名を持たず、新specのオーナー承認ではない。
旧hash案へのチャット確認だけで台帳を書かない。実プロジェクトの台帳は空のまま。
旧shipはL0だけなので、新specの出荷ゲートとして合格を主張しない。
旧契約は既存fixtureを検証する実装の説明として残し、新版へ段階的に差し替える。

## 追加した基礎

scripts/jcs.pyのcanonicalize_batchは既存NodeでJCSをバッチ実行し、media_hashesは媒体・profile・解決値を束縛する。
NodeはDESIGN_AUTHOR_NODEで選ぶ。依存の自動導入・ネットワーク利用はしない。
この基礎はまだL0の承認照合へ接続していない。旧catalogの分類入力hashを無断で変更しない。

[RFC 8785](https://www.rfc-editor.org/rfc/rfc8785.html)に従い、ECMAScriptのprimitive表現と
UTF-16順を使う。オブジェクトの数値風keyをJSON.stringify任せにしない。
Unicodeを正規化せず、不正surrogate・非有限数・IEEE 754で精度を失う整数を拒否する。
テストは数値表現、数値風keyと補助平面文字の順序、媒体/profile/解決先の変更、不正入力を含む。

## 次の実装と人間境界

- keyをds:/ia:で名前空間化し、媒体別JCS hashを用いる。署名payloadはrepo id/key/hash/date、署名namespaceはds-approval/ia-approval。
- 無署名・別repo/namespace・未登録鍵・alias先/profile変更をL0で拒否する赤fixtureを先に作る。
- オーナー自身がapproveを起動し、タッチか都度確認の鍵で署名する。AIが実鍵を起動・新設・変更しない。
- allowed_signersはAIに変更できない信頼境界へ置く。repo内やAIが差し替えられる通常ファイルを、保護された台帳の証拠と扱わない。
- 実鍵/ssh-agent/allowed_signersの導入・変更は認証/secret設定の確認ゲート。人工鍵のfixtureで本番の人間介在を証明しない。
- 判定台帳、5区分、IAモジュール境界は未実装。分類610/2,000を勝手に再実行しない。

固定環境2回のL2、出荷SHA/成果物hash、ship_command/doctor、維持モードとA/B/Cも別の追加受入として残る。
