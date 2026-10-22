# 判定台帳・5区分報告・IA接続

## 判定対象と入力

L0は、カタログ項目に対するdoc証拠の散文とn-a理由だけを決定論で抽出する。
判定の対象はdoc-relevance（項目を扱っていない可能性）とna-rationale（適用外理由が不十分な可能性）。
JevをL0/L1/L2から呼ばない。カタログのmust分類は既存classify.pyが所有し、この台帳と混ぜない。

```sh
design-author check --repo . --include-candidates
```

通常のcheckは候補の本文・文脈を出力しない。include-candidatesを明示した時だけinputを含む。
inputはkind・catalog_id・空白を正規化した本文・元カタログ項目とpreset等の文脈。
公開できないプロジェクト内容をチャットやログへ転記しない。機密の分類や外部送信の認可はこの出力では変わらない。

キーはspan_hash＋context_hash＋model_digest＋prompt_version＋prompt_digest＋thetaの辞書のSHA256。
同じ文章でもカタログ項目・証拠参照・presetが違えば再利用しない。
model_digestはproviderと版付きmodel idの識別情報のdigestであり、提供元の重みファイルを検証したものではない。
現在は既存分類で観測済みのjev-1.13.0をpinし、thetaは中央のjudgment-policy.jsonに置く。

## 保存形式と読み取り

任意のdesign/judgments.jsonに次の構造を持つ。rowのscoreは「問題がある確率」。
policyは中央policyの全フィールドと一致必須で、judgment_hashは現在の候補から得る。

```json
{
  "schema_version": 1,
  "policy": {"schema_version": 1, "model": "jev-<pinned-version>", "theta": 0.5,
             "prompt_version": "<version>", "prompts": {"doc-relevance": "<criteria>", "na-rationale": "<criteria>"}},
  "rows": [{"judgment_hash": "<candidate-hash>", "score": 0.1}],
  "holdout": {"source": "owner", "cases": [
    {"id": "<positive-id>", "positive": true, "score": 0.9},
    {"id": "<negative-id>", "positive": false, "score": 0.1}
  ]}
}
```

これは形式例で実データではない。scoreをAIが発案してJevの結果として登録しない。
本物の回答の取得は既存Jevの認証・呼出経路で行い、実行数と使用量をWSの予算へ計上する。
このsliceのCLIは回答ファイルの検証・保存を担当し、モデル実行アダプタを追加していない。

台帳なし・現在のjudgment_hashなしは未カバー。mustの未カバーはship失格。
scoreがtheta以上なら証拠の問題として失格。theta未満はscreenedとし、holdoutのrecall・陽性件数・偽陰性数を併記する。
空のholdout、陽性だけ/陰性だけ、重複id、非有限確率、重複keyは拒否する。
人間の正解ラベルの確認が必要で、AIのラベル案をsource: ownerとして自動登録しない。
小さな人工例のrecallを実運用全体のrecallへ一般化しない。

```sh
python3 <skill>/scripts/judgment_ledger.py record --repo . --replacement answers.json
python3 <skill>/scripts/judgment_ledger.py record --repo . --replacement answers.json --write
```

初回recordは台帳が既にあれば拒否する。通常はpreviewだけで、writeが明示された時にatomic replaceする。
判定台帳を記録してもapprovals.yamlは変えない。

## 版・モデル・thetaの変更

中央policyを変えたら古い台帳とmanifestは失効する。新しい回答と同じid・同じ正解ラベルのholdoutを用意する。

```sh
python3 <skill>/scripts/judgment_ledger.py migrate --repo . --replacement new-answers.json
python3 <skill>/scripts/judgment_ledger.py migrate --repo . --replacement new-answers.json --write
```

前後のrecall・偽陰性数と閾値判定の変わったcase idを記録する。ラベル集合を都合よく変えた比較は拒否する。
保存するmigrationの数値は前後のholdoutから再計算して照合する。モデル/閾値を変えたまま古いkeyを使わない。
書込みはlockを取り、policyが記録中に変われば拒否する。

## 報告

L0とdesign-checkはreportにexecuted / skipped / missing / uncovered / screenedの5配列を持つ。
L0単独ならL1/L2は理由付きskipped。design-checkは実行した段階をexecutedへ移す。
実装がないIAはmissing、記録がない判定はuncovered、閾値未満はscreened。
L0と後段consumerは中央policyのバイトdigestを照合し、policy変更後に古い生成物を使わない。

## IA resolverの接続口

本体の実装は兄弟WSが所有する。ここでは接続契約だけを実装する。
design.yamlのmodulesにiaの版を宣言する。コードのパスや任意コマンドは宣言できない。

```json
{"modules": {"ia": "<module-version>"}}
```

固定の隣接skill own-infoarch-author/scripts/ia_module.pyだけを読み、VERSIONと宣言版を照合する。
resolverはresolve(root=..., approvals=..., profiles=...)を1回だけ呼ばれる。
戻り値はmedia（媒体別解決結果）、approval_hashes（ia:キーと現在hash）、must_approval_keys、
errors、warnings、5区分report。正常なのに承認対象0件なら拒否する。
runnerは共用台帳を全targetへ照合し、approveのprepareにもDS/IA両方の対象hashを渡す。

解決結果はdesign/.resolved/<media>.jsonのmodules.iaへ入り、後段は再解決しない。
resolver実装のdigestが変わればconsumerも拒否する。解決中の実装変更も失格。
IA入力があるのに版がない、宣言したのに実装がない、版や契約が違う場合は失格で古いmanifestを削除する。
IAそのものの原子・抽出・検査はこの接続口だけでは実装済みにならない。
