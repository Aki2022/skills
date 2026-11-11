---
name: own-marketing-route
description: >-
  marketing-automation-core（@aki2022/* の leaf package）の機能を、いま開いている repo に
  標準 CI・隔離 runner・認証・docs の決まりも含めて「必要なものだけ選んで取り込みたい」ときの入口。
  core をまだ何も入れていない repo でも使う。core の採用メニュー全体を見せて人と 1 問ずつ話し、
  npm の機能を選んだ時だけ読み取りの診断（marketing-capabilities start）を流して、決まったものを
  それぞれの手順（own-marketing-capability-route・標準 CI の guide など）へ渡す。
  使うこと: 「core の機能を入れたい」「MA の機能を選んで入れたい」「マーケの機能で何が使える？」
  「@aki2022 の package を入れたい」「marketing-automation-core を使いたい」「この repo に計測や
  レポートの機能を足したい」「core の標準 CI を使いたい」「隔離 runner（Apple のコンテナ）に載せたい」と言われた時。
  使わない場面: 「Excel にしたい」「PDF を作りたい」のように core に触れない一般的な依頼
  （core の機能が候補かどうかを人に聞いてから使う）、core repo 自体の開発、
  既に機能 ID が決まっていて install するだけの時（own-marketing-capability-route を直接使う）。
---

# marketing-automation-core の機能を選んで入れる入口

この skill は入口だけを持つ。何を取り込むかを決めるのはこの repo の人で、AI は材料を
見せて 1 問ずつ聞く。core が出しているのは npm の機能だけではない。全体の一覧は core の
`docs/guides/consumer-adoption-menu.md`（採用メニュー）が正典で、npm の catalog はその一部。
catalog に無いことを「core に無い」と言わない。

## 1. core の checkout を見つける

core の CLI と catalog は、core の正典の checkout から読む。どこかから CLI を
ダウンロードして実行しない。

1. 環境変数 `MARKETING_CORE_DIR` があれば、それを core の checkout とみなす。
2. 無ければ人に 1 回だけ場所を聞く。推測で探し回らない。
3. `$MARKETING_CORE_DIR/packages/marketing-core/bin/capabilities.cjs` と
   `$MARKETING_CORE_DIR/packages/marketing-core/capabilities/catalog.json` が両方あることを確かめる。
   無ければ止まって人に伝える。
4. catalog の `sharedVersion` を人に見せる。入れるのはこの版になる。

## 2. 採用メニューを見せて、目的を聞く

core の checkout の `docs/guides/consumer-adoption-menu.md` を読み、目的ごとの一覧
（開発・検査、外部サービス接続、サイト計測、データ、改善・承認、レポート・通知、運用）を
平易な言葉で見せる。最初の問いは「何に困っているか・何をしたいか」。答えから、次のどれに進むかを決める。

- **標準 CI**: core の `docs/guides/ci-standard.md` の「Calling it」に従う。npm の準備は要らない。
- **隔離 runner**: ADR-20261111 の 4 段（owner が token の対象に repo を足す → agent が並走の PR で一致を確かめる →
  owner が必須チェックを差し替える → 見張り番の一覧に足す）。owner の操作が要る段は、その旨を伝えて止まる。
- **GCP の認証**: user 単位の own-gcp-provision。
- **docs の決まり**: own-doc-update。
- **npm の機能（catalog の capability）**: 下の 3 と 4 へ進む。
- **core が出していないもの**（deploy の manifest、Cloud Run の job、汎用の GCP API など）: メニューの
  「Not offered by core」を見せ、この repo で作るものだと伝える。

## 3. npm の機能を選ぶ前の診断

いま開いている repo の root で実行する。何も書き込まない。

```sh
node "$MARKETING_CORE_DIR/packages/marketing-core/bin/capabilities.cjs" start --repo .
```

- `errors` は進む前に直すもの、`warnings` は知っておくもの。どちらも `next` に次の操作がある。
  これを平易な言葉で一覧にして見せる。exit 2 は「まだ準備が足りない」という意味で、失敗ではない。
- `ready` は「`plan` を流せる」という意味だけ。host peer・auth scope・lockfile・版は、選んだ
  capability ごとに `plan` が見る。`ready` を導入済みの証拠にしない。
- `package.json` や `.npmrc` を作るのは、この repo への書き込み。人の了承を得てから行う。
  repo の `AGENTS.md` や docs の決まりがあれば先に読み、それに従う。
- `setup.*` を選びそうなら、先に `package.json` を手で作らない。`consumer-init` が scaffold として
  作り、既にあると全体を断るので、provision skill の順番に従う。
- token（`NODE_AUTH_TOKEN`）の値は表示しない・書かない。用意の仕方はこの repo の秘密の
  扱い（例: own-secret-resolve）に任せる。診断は token が設定されているかだけを見る。
- `@aki2022/*` の pin が古い・catalog 外だと warning に出る。新しい機能を足す前に、その repo が
  使っている capability を全部挙げて `upgrade`（core の checkout の CLI）で上げる。`upgrade` の後は
  `npm ci` と `skills` をやり直し、配布 skill の差分は人と確かめる。

## 4. 人と話して capability を決める

1 回に 1 問だけ聞く。

1. 最初に、何に困っているか・何をしたいかを聞く。機能名からではなく、困りごとから始める。
2. `start` が返した `capabilities` の `userNeed` を読み、合いそうな候補を 1〜3 個に絞る。
   **言葉が似ているだけで選ばない。** 例えば `reporting.gcs` は「作ったレポートを GCS に置く」機能で、
   任意のファイルを GCS で読み書きする機能ではない。合うか迷うものは、その違いを見せて人に聞く。
3. 候補ごとに次を見せる: `userNeed`、`leafCount` と `requiredPackages`、`sdkPackages`、
   `prerequisites`（先に要るもの）、`effects`（使うと外で起きること・費用）、
   `hostRequirements`（React など repo 側に要るもの）、`requiredAuthScopes`。
4. 人に選んでもらう。複数でもよい。決めるのは人で、AI は推しを 1 つ言ってよいが選ばない。
5. catalog にもメニューにも合うものが無ければ、3 つに分けて人に見せる。
   (a) core の設定や plugin で表せる → その方法を示す。
   (b) ほかの repo にも役立つ汎用の不足 → core の repo で `own-marketing-capability-escalate` の流れ（要望の issue）。
   (c) この repo だけの事情 → この repo で作る。どれに当たるかは人と決め、勝手に作り始めない。

## 5. 決めたことを残す

選んだ capability ID、理由、見せた `prerequisites` と `effects`、決めた日を、この repo の
決定の置き場所に残す（`docs/00_index.md` があれば own-doc-update の流儀で WS か ADR）。
置き場所が無ければ、どこに残すかを人に聞く。会話だけで終わらせない。

## 6. install へ渡す

core の checkout にある `.agents/skills/own-marketing-capability-route/SKILL.md` を読み、
決まった ID でその手順に従う（`plan` → 人の承認 → `install --apply` → `npm ci` → `skills`）。
`setup.*` を選んだ場合は、同じ場所の `own-marketing-consumer-provision` も読む。

- `install --apply` と `npm ci` は private registry から package を取ってくる。
  人の承認を得てから実行する。
- `plan` や `doctor` の exit 0 を「入った」の証拠にしない。`addedLeafPackages` と
  `lockfileNeedsUpdate` を読む。

## しないこと

- 人の代わりに capability を決めない。
- token や秘密の値を表示・保存しない。
- core の checkout を書き換えない。core に直すべき点が見つかったら、人に伝える。
- catalog に無い package、`latest`、URL、Git ref を入れない。
