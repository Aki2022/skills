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

design/judgments.jsonのschema2は、中央policyと同じpolicy、rows、holdoutを持つ。
rowsは`{"candidate": <check --include-candidatesの候補全体>, "score": <実回答>}`の配列。
candidateの入力・文脈・全hashを再計算し、policyとの一致を確認する。scoreはNoulのYes確率で、説明不足を確定する値ではない。
ラベルが無い場合は`"holdout": {"source": "unmeasured"}`を明示し、recall・誤検出率を作らない。
実際の人間ラベルがある場合はsource: ownerとcases（id・positive・実score）を持ち、recall等を再計算する。
旧schema1のhash/score行と正規のowner holdoutも読める。旧形式をラベルなしへ黙って変更しない。

上記は形式の説明で実データではない。scoreをAIが発案してJevの結果として登録しない。
本物の回答はjudgment_run.pyで取得し、既存classify.pyの認証・固定endpoint・リダイレクト拒否・原子的保存を再利用する。
Noulのyes確率を「問題がある確率」として保存する。版付きモデル、完全な回答集合、確率、使用量を検証する。
公式契約: [Noul](https://docs.typesafe.ai/primitives/noul)、[HTTP API](https://docs.typesafe.ai/api/system-one)、[料金](https://docs.typesafe.ai/models)。

```sh
design-author check --repo . --include-candidates > candidates.json
python3 <skill>/scripts/judgment_run.py --candidates candidates.json --answers design/judgment-answers.json --limit 50
python3 <skill>/scripts/judgment_run.py --candidates candidates.json --answers design/judgment-answers.json --limit 50 --execute
```

checkが未承認mustでexit 2でも候補は取得できる。候補以外のcheck結果を合格扱いにしない。
既定はdry-runで認証読取・通信・書込みなし。executeは別途許可された外部送信と従量予算内だけで使う。
新しいSDK・認証設定を作らず、classify.pyと同じcatalog/classification-usage.jsonをlockして累計2,000判定を守る。
既存使用量ファイルが無ければ停止する。CLIから別ファイルや上限増加を指定できない。
10問/16,000文字以内に分割し、通信前に問数を予約する。失敗や未検証回答も予約数を消費し、自動再試行しない。
小規模試行後の実使用量に基づいて残りを見積もる。usage_derived_usdは使用量と設定単価からの計算で、請求額・金額上限ではない。
成功バッチごとに保存し、同じcandidate/policyの回答は再取得しない。入力・policy変更、破損した回答、symlink、対象0件は拒否する。

出力名はjudgment-answers.jsonに限定する。全candidate・Noulのraw回答・scoreをstatus: model-answers-unlabeledとして保存し、
公開可能な入力かを送信前に確認する。使用量台帳には本文・認証情報を残さず、request hashと件数・token数・時間を残す。
この出力はjudgments.jsonとは別で、holdout・人間ラベル・承認を生成しない。L0は直接読まない。
現在のcandidateに対応する実回答のcandidate/scoreだけをschema2へ写し、評価用ラベルが無ければunmeasuredとして下記recordへ渡す。実ラベルが無いことは失格ではない。

台帳なし・現在のjudgment_hashなしは未カバー。全current candidateの未カバーはship失格。入力・版不一致や壊れた台帳も停止する。
保存済み判定は全件screened。scoreがtheta以上ならflagged_advisoryとして対象id・hash・scoreを報告し、原文照合用の警告にする。scoreだけで出荷を止めない。
意味の十分さは自動保証しない。実ラベルがあればholdout指標を併記し、なければevaluation.statusはunmeasured。通常の出力に原文は出さず、明示した候補出力で照合する。
空のholdout、陽性だけ/陰性だけ、重複id、非有限確率、重複keyは拒否する。
source: ownerを使う時だけ人間の正解ラベルの確認が必要。専門ラベルを所有者へ再度求めず、AIのラベル案をsource: ownerとして自動登録しない。
小さな人工例のrecallを実運用全体のrecallへ一般化しない。

```sh
python3 <skill>/scripts/judgment_ledger.py record --repo . --replacement answers.json
python3 <skill>/scripts/judgment_ledger.py record --repo . --replacement answers.json --write
```

初回recordは台帳が既にあれば拒否する。通常はpreviewだけで、writeが明示された時にatomic replaceする。
判定台帳を記録してもapprovals.yamlは変えない。

## 同じ判定条件で文章を更新する

変更した入力だけの新しい実回答を取得し、現在の候補集合のschema2 replacementを用意する。

```sh
python3 <skill>/scripts/judgment_ledger.py refresh --repo . --replacement current-answers.json
python3 <skill>/scripts/judgment_ledger.py refresh --repo . --replacement current-answers.json --write
```

refreshは同一policyのschema2だけで、旧行と追加・削除・保持したhashを証拠として保存する。
保持した入力のscoreと既存holdoutは変更できない。追加/削除と証拠の改ざんも照合する。
同一policyのmigrateでもscoreやholdoutを変更できず、通る回答を選ぶ抜け道にしない。
旧schema1の元の移行記録は、その当時のholdout比較の保証範囲で読み込み、改ざんを拒否する。

## 版・モデル・thetaの変更

中央policyを変えたら古い台帳とmanifestは失効する。同じ入力集合への新しい実回答を用意し、新旧scoreと閾値判定差分を記録する。実ラベルがあれば同じid・同じ正解ラベルも維持する。旧schema1は同一policyで完全candidateを持つschema2へ更新してからpolicyを変える。

```sh
python3 <skill>/scripts/judgment_ledger.py migrate --repo . --replacement new-answers.json
python3 <skill>/scripts/judgment_ledger.py migrate --repo . --replacement new-answers.json --write
```

同一入力の新旧hash・score・指摘差分と評価用ラベル有無を記録する。実ラベルがある場合はrecall・偽陰性数等も比較し、ラベル集合を都合よく変えた比較やラベルの削除は拒否する。
保存するmigrationの証拠は旧台帳と新台帳から再計算して照合する。モデル/閾値を変えたまま古いkeyを使わない。
書込みはlockを取り、policyが記録中に変われば拒否する。

## 報告

L0とdesign-checkはreportにexecuted / skipped / missing / uncovered / screenedの5配列を持つ。
L0単独ならL1/L2は理由付きskipped。design-checkは実行した段階をexecutedへ移す。
実装がないIAはmissing、記録がない判定はuncovered、保存済みDS意味判定は高scoreもscreened（補助指摘を明示）。
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

## IA v4との接続

上記の承認対象を持つ旧resolver契約と、IA v4の承認状態を持たないSSOTは区別する。
modules.ia: 0.2.0は、固定のia_module.pyから既存IAのcheckを同じPythonで呼ぶ。ia-git-history-v2契約に限定してIAのapproval_hashesを空にでき、DS承認の必要条件は変えない。
入力hashと実際のチャネル・アウトライン・セクション件数を持ち、対象0件・壊れた参照・期限違反は失格。IA入力や実装のsymlink/逸脱は拒否する。
マニフェストのIA実装digestは橋渡しファイルだけでなく、固定のscripts/references/schema全体を束縛し、Python/JS両consumerが古い結果を拒否する。
このL0接続はIAのDOM検査や未実装の意味判定を実行済みにはしない。後段のDOMは別に実行し、意味判定の欠落も報告する。
