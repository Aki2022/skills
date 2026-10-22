# 署名付き承認（オーナー環境の導入待ち）

正典はbiz_opsのSPEC-design-system-standardizationと
ADR-20261022-approval-provenance-and-ship-gate-are-structural（PR #70、8f10578）。

## 実装と制約

L0はSSH署名・媒体別JCS hashを照合する。無署名、未登録鍵、別repo/namespace、
値/alias先/profileの変更は失格。結果は評価日時とPython版・JCS実装digestを持つ。
これはL2の固定環境指紋を実装済みという意味ではない。

signersをDESIGN_ALLOWED_SIGNERSで指定する。外部の通常ファイルで、agent accountが
ファイルまたは親を変更できる場合は失格。所有者によるchmod変更や親からの差替えも防ぐ。
パスは実行環境で渡し、repoに端末固有のパスを書かない。出荷時の起動環境にも保護された設定が必要。
署名鍵には物理タッチのFIDO2鍵、または都度確認のssh-agentを使う。agentの確認制約は
公開鍵一覧からは証明できないため、オーナー設定と実際の確認動作の検証が必要。

## オーナーの操作

AIはprepareまで実行できる。台帳はまだ書かない。

```sh
design-author approve --prepare --must --repo . > request.json
```

公開鍵と、hashが束縛する値/理由/散文を確認する。オーナー自身の端末で次を実行する。

```sh
design-author approve --request request.json --key <owner-public-key.pub> --approver owner --repo .
```

実鍵を扱うのはオーナーだけ。公開鍵は既存agentが署名できるEd25519/FIDO2鍵を指定する。
既存環境を接続するか、新設するかは人間ゲート。skillは鍵・agent・signersを自動設定しない。
署名は行ごとなのでmust 471件なら471署名となり、各行でタッチ/確認が必要な場合がある。
部分承認はprepareの--target ds:web:item:<id>を使える。

署名前・署名後に現在のhashとrepo idを再計算し、途中変更なら台帳へ書かない。
全署名検証後に他の既存行を保持してatomic replaceする。失敗時に部分承認を書かない。
同時承認コマンドはlockで拒否する。記録後もL0に残件があればexit 2となり、記録数を別途報告する。
CLIのprepareはリクエスト生成だけなので起動可能でも、署名/出荷の成功とは扱わない。

## JCSと検証範囲

既存NodeでECMAScriptのprimitive表現とUTF-16順を使う。DESIGN_AUTHOR_NODEで選ぶ。
[RFC 8785](https://www.rfc-editor.org/rfc/rfc8785.html)に従い、Unicodeを正規化しない。
不正surrogate・非有限数・IEEE 754で精度を失う整数を拒否する。新規依存/ネットワークは不要。

in-process正例は人工鍵の実SSH署名と検証を使い、OS trust境界だけを代替する。
実CLI/実hookはagent所有の人工trustを拒否する。実hookのgreen対照はnice-onlyの状態であり、
署名済みmustの実出荷を証明しない。実オーナーの信頼基盤と確認動作は導入後の必須検証。

判定台帳・5区分・IAモジュール境界、固定L2 2回、出荷SHA/成果物hash、ship_command/doctor、
維持モードとA/B/Cはまだ追加受入として残る。現在のL0-only shipを正式な出荷ゲートと呼ばない。
