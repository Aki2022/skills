---
name: own-marketing-route
description: >-
  marketing-automation-core（@aki2022/* の leaf package）の機能を、いま開いている repo に
  「必要なものだけ選んで入れたい」ときの入口。core をまだ何も入れていない repo でも使う。
  読み取りだけの診断（marketing-capabilities start）で足りないものを示し、人と 1 問ずつ話して
  capability を決め、決まった ID を core の own-marketing-capability-route へ渡す。
  使うこと: 「core の機能を入れたい」「MA の機能を選んで入れたい」「マーケの機能で何が使える？」
  「@aki2022 の package を入れたい」「marketing-automation-core を使いたい」「この repo に計測や
  レポートの機能を足したい」と言われた時。
  使わない場面: 「Excel にしたい」「PDF を作りたい」のように core に触れない一般的な依頼
  （core の機能が候補かどうかを人に聞いてから使う）、core repo 自体の開発、
  既に機能 ID が決まっていて install するだけの時（own-marketing-capability-route を直接使う）。
---

# marketing-automation-core の機能を選んで入れる入口

この skill は入口だけを持つ。何を入れるかを決めるのはこの repo の人で、AI は材料を
見せて 1 問ずつ聞く。install の手順そのものは core の `own-marketing-capability-route` が持つ。

## 1. core の checkout を見つける

core の CLI と catalog は、core の正典の checkout から読む。どこかから CLI を
ダウンロードして実行しない。

1. 環境変数 `MARKETING_CORE_DIR` があれば、それを core の checkout とみなす。
2. 無ければ人に 1 回だけ場所を聞く。推測で探し回らない。
3. `$MARKETING_CORE_DIR/packages/marketing-core/bin/capabilities.cjs` と
   `$MARKETING_CORE_DIR/packages/marketing-core/capabilities/catalog.json` が両方あることを確かめる。
   無ければ止まって人に伝える。
4. catalog の `sharedVersion` を人に見せる。入れるのはこの版になる。

## 2. 読み取りだけの診断を流す

いま開いている repo の root で実行する。何も書き込まない。

```sh
node "$MARKETING_CORE_DIR/packages/marketing-core/bin/capabilities.cjs" start --repo .
```

- `errors` は進む前に直すもの、`warnings` は知っておくもの。どちらも `next` に次の操作がある。
  これを平易な言葉で一覧にして見せる。exit 2 は「まだ準備が足りない」という意味で、失敗ではない。
- `package.json` や `.npmrc` を作るのは、この repo への書き込み。人の了承を得てから行う。
  repo の `AGENTS.md` や docs の決まりがあれば先に読み、それに従う。
- token（`NODE_AUTH_TOKEN`）の値は表示しない・書かない。用意の仕方はこの repo の秘密の
  扱い（例: own-secret-resolve）に任せる。診断は token が設定されているかだけを見る。
- `existingPins` に `@aki2022/*` が既にあれば、新規導入ではなく追加か版上げの話になる。
  catalog より古い版なら、版上げ（`upgrade`）が要ることを伝える。

## 3. 人と話して capability を決める

1 回に 1 問だけ聞く。

1. 最初に、何に困っているか・何をしたいかを聞く。機能名からではなく、困りごとから始める。
2. `start` が返した `capabilities` の `userNeed` を読み、合いそうな候補を 1〜3 個に絞る。
   **言葉が似ているだけで選ばない。** 例えば `reporting.gcs` は「作ったレポートを GCS に置く」機能で、
   任意のファイルを GCS で読み書きする機能ではない。合うか迷うものは、その違いを見せて人に聞く。
3. 候補ごとに次を見せる: `userNeed`、`requiredPackages`（入る leaf の数）、SDK、
   `prerequisites`（先に要るもの）、`effects`（使うと外で起きること・費用）、
   `hostRequirements`（React など repo 側に要るもの）、`requiredAuthScopes`。
4. 人に選んでもらう。複数でもよい。決めるのは人で、AI は推しを 1 つ言ってよいが選ばない。
5. catalog に合うものが無ければ、そう伝えて止まる。core に足してほしい機能なら、
   core の repo で `own-marketing-capability-escalate` の流れ（要望の issue）になることを伝える。
   この repo で代わりのものを作り始めない。

## 4. 決めたことを残す

選んだ capability ID、理由、見せた `prerequisites` と `effects`、決めた日を、この repo の
決定の置き場所に残す（`docs/00_index.md` があれば own-doc-update の流儀で WS か ADR）。
置き場所が無ければ、どこに残すかを人に聞く。会話だけで終わらせない。

## 5. install へ渡す

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
