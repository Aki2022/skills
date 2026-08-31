---
name: origin-trouble-log
description: >-
  エージェント作業中に起きたトラブルを、後の一括分析に足る形で1件記録する。また
  溜まった記録をトリアージして hook / skill への形式化候補を提示し、過去対策の status と
  有効性評価を更新する。必ず使うこと:
  ユーザーから作業のやり方を指摘された時（「それは違う」「なぜ聞くのか」「それは既に
  自動化されている」「なぜ確認しないのか」等）、無音の no-op・誤った完了報告・不要な
  質問・空振りの検査に気づいた時、「トラブルを記録」「今のを記録して」「摩擦を記録」
  「溜まりを分析」「トリアージして」「記録を棚卸し」と言われた時。skill を使っていない
  場面のトラブルも対象で、skill 名は任意。使用しない場面: リポジトリ固有の技術的バグの
  起票（各リポジトリの issue が持つ）、skill 本文の改訂そのもの（origin-skill-commonize
  が所有）、hook や settings.json の作成（settings を扱う skill が所有）。
---

# Origin Trouble Log

エージェント（Claude / Codex 等）の作業トラブルを集積する。**集積と、集積物の
トリアージ、対策 status の管理まで**をこの skill が所有する。形式化の実行は所有しない。
entry 本体は証拠として append-only に保ち、可変な status は保管ルート内の
`triage/status.tsv` と `triage/responses.tsv` で管理する。

なぜ必要か: 前身の `~/.agents/skills/FRICTION.md` は 1 行形式・skill 名必須で、
skill を使っていない場面のトラブルを記録できなかった。2026-07-31 に規約が作られ
2026-08-01 に 3 件入った後、2026-08-08 まで追記ゼロだった。**発火条件を持たない
規約は使われない**ため、この skill は description に発火語を持つ。

## 保管ルート

パスをこの skill に書かない（端末ごとに異なり、クラウド同期ディレクトリ名は
記録に残せないため）。次の順で解決する。

1. 環境変数 `ORIGIN_TROUBLE_LOG_ROOT`
2. 無ければ `~/.config/origin-trouble-log/root`（ルートのパスを 1 行書いたファイル）

```bash
ROOT="${ORIGIN_TROUBLE_LOG_ROOT:-$(cat ~/.config/origin-trouble-log/root 2>/dev/null)}"
[ -n "$ROOT" ] && [ -d "$ROOT" ] || { echo "保管ルート未設定"; exit 1; }
```

**2 を主経路とみなす。** 環境変数はシェル設定に書く必要があり、この環境ではシェル設定が
nix（home-manager）管理でリビルドを要する。加えてエージェントのシェルはプロファイルを
読み込めないことがある（実際にこのセッションで PATH が欠けた）。ポインタファイルなら
どちらの制約も受けない。

どちらからも解決できないときは**記録を諦めず、ユーザーに設定を求めて停止する**。
黙って別の場所へ書かない — 記録が散ると集積の目的が消える。

制約:

- ルートは **git working tree の外**に置く。`~/.agents/` 配下にも置かない
  （そこは remote を持たないローカル git repo で、グローバル hook が検査を掛ける）。
- git を使わない。commit / push / 同期 cron を持たない。バックアップは
  クラウド同期クライアントに委ねる。
- ルートには `README.md` を置き、後から開いた人間が何のディレクトリか分かるようにする。

## 1 件を記録する

配置は解決した `$ROOT` 配下の `entries/YYYY-MM/YYYY-MM-DD-<slug>.md`。
**entry の新規 Write のみで完結する**（既存 entry の証拠本文を読まないため、並行セッションでも
競合しない）。後から変わる triage / response status は別の status ledger に書くため、
entry 本体へ結論を書き戻さない。

テンプレートは `references/entry.template.md`。frontmatter は絞り込みに使える
機械可読な値だけを持つ。

| Field     | 必須             | 用途                                                    |
| --------- | ---------------- | ------------------------------------------------------- |
| `date`    | 必須             | 時系列の絞り込み                                        |
| `summary` | 必須             | 本文を開かずに一覧・トリアージする                      |
| `skills`  | 必須（空配列可） | skill 単位の需要駆動トリガの絞り込み                    |
| `repo`    | 任意             | 集計軸                                                  |
| `canon`   | 必須             | **そのとき読んでいた横断正典の版**。下記の取り方で得る  |
| `paths`   | 必須（空配列可） | **トラブルの対象ファイル**。hook 化の可否を判断する材料 |

`canon` の取り方:

```bash
readlink ~/.agents/AGENTS.md | sed 's|.*/nix/store/\([a-z0-9]\{8\}\).*|\1|'
```

**なぜ必要か（2026-08-09 の初回トリアージで判明）**: そのとき最も知りたかったのは
「**新しい規律を読んだ上での違反か**」だったが、記録にその情報が無く**判定できなかった**。
正典はセッション開始時に読まれるため、活性化より前に始まったセッションは古い版を持つ。
`canon` があれば1値で判定でき、セッション開始時刻との突き合わせが要らない。

**`paths` が必要な理由**: AC1 検証で、文脈を持たない読み手が
**「機械抽出できるか分からないので hook 案を書けない」**と報告した。
対象パスが本文に散らばるだけでは hook 化の可否を判断できない。

本文は固定 7 節。**分類ではなく、事実と仮説を分けて書く。**

テンプレートに加えて `scripts/check_entry_format.sh` を使える。これは必須 frontmatter、
固定 7 見出し、実文コードブロック、本文を含むユーザー名付き絶対パスを決定論的に検査する。
過去の entry を一括修正する用途には使わず、今後の新規記録とトリアージ時の形式確認に使う。
status ledger の同期・更新・検証には `scripts/status_ledger.py` を使う（詳細は
`references/status.md`）。

### 書き方の必須要件

1. **行動は実文で引用する。** 実行したコマンド、書いたコード、出した報告文を
   そのまま貼る。hook の検知条件は「観測可能な行動」に対してしか書けず、
   要約からは導けない。これが実文主義の理由であり、「厚く書く」ためではない。
2. **既に存在していた正解の在り処を具体パスで書く。** 記録は修正後に書かれるため
   在り処は既に手元にある。**「無かった」と「不明」を書き分ける** — 前者は規律ではなく
   機能欠落を意味し、形式化先が変わる。
3. **「本来どうすべきだったか」「なぜそうしなかったか」は仮説として書く。**
   「わからない」と書いてよい。失敗直後の自己診断は外れることがあり、断定形で書くと
   分析を誤った形式化へ誘導する。

### secret の扱い

**secret は書く時点で値をマスクし、構造を残す**（`Bearer <redacted>`、
`--api-key <redacted>` 等）。git の有無と無関係に必須 — 生のトークンをディスクへ
書かないこと自体が目的であり、保管ルートは第三者のクラウドへ同期される。

マスクは実文主義と衝突しない。検知条件に必要なのは構造（コマンド名・引数の形・
エラーの型）であって、秘密の値そのものに材料価値は無い。

会社ドメイン・顧客名・リポジトリ名はマスクしない。マスクすると「どのリポジトリ・どの案件で
再発しているか」が追えず、集積の目的が消える。一方、`repo` はリポジトリ名だけ、`paths` は
リポジトリ相対パスだけで記録する。グローバル設定は `~/.agents/...`、アプリケーション内の
対象は `<app>/...` のようなユーザー名を含まない論理パスにする。`/Users/<user>/...`、
`/home/<user>/...`、保管ルートを解決した絶対パスは frontmatter と本文のどちらにも残さない。

## いつ記録するか

**自主的な追記に任せない。** 7 節形式の重さでは書かれなくなる。二段構えとする。

1. **主力: ユーザーからの指摘を即時 trigger にする。** 作業のやり方への指摘を受けたら
   その場で 1 件記録する。指摘は「本来どうすべきだったか」を既に含むため 7 節が埋まりやすい。
2. **補助: セッション終了時の棚卸し。** `origin-close-session` の手順に組み込み、
   指摘されなかったトラブルを拾う。

Stop hook による全セッション強制は**採らない**。毎セッションのトークン消費が記録の
価値に見合わず、毎回問われることで形式的に「無し」と答える形骸化も招く。

## status を管理する

status は「トリアージ済みか」と「対策が効いたか」を分けて記録する。対象欄に列挙された
だけでは対策済みとはみなさない。

- `triage_status`: `untriaged` / `triaged`
- `response_status`: `unknown` / `legacy_unrecorded` / `none` / `proposed` /
  `implemented_unverified` / `effective_provisional` / `recurred` / `ineffective` /
  `manual_only` / `repo_specific_pending`
- `response_ids`: `triage/responses.tsv` に登録した、再利用可能な対策 ID。日付や一時的な
  セッション番号を ID にしない。

意味と遷移の詳細は `references/status.md` に従う。特に次を守る。

1. 新規 entry は `untriaged / unknown` で始める。
2. 過去レポートに列挙済みだが outcome の証拠が無い entry は、backfill で
   `triaged / legacy_unrecorded` とする。実装済み・有効とは推測しない。
3. 実装と局所テストを確認しただけなら `implemented_unverified`。運用上の再発・見逃し・
   誤警告を定義した期間で測っていない限り、`effective_provisional` に進めない。
4. 対策が存在する状態で同型が出たら `recurred` とし、同じ修正をもう一度適用するのではなく、
   coverage gap・運用逸脱・別原因を切り分ける。
5. 各 status 変更は、次回レポートの `## 過去対策の評価` に根拠・再発数・見逃し/誤警告数・
   次の測定を書く。過去 entry と過去レポートは書き換えない。

起動時・トリアージ開始時には次を実行し、過去分の取りこぼしを確認する。

```bash
python3 ~/.agents/skills/origin-trouble-log/scripts/status_ledger.py --root "$ROOT" sync
python3 ~/.agents/skills/origin-trouble-log/scripts/status_ledger.py --root "$ROOT" summary
python3 ~/.agents/skills/origin-trouble-log/scripts/status_ledger.py --root "$ROOT" validate
```

`sync` は不足行だけを追加して既存 status を上書きしない。status ledger が未作成なら、
全 entry を列挙し、全過去レポートの対象欄の和集合に含まれるものを
`triaged / legacy_unrecorded`、含まれないものを `untriaged / unknown` として初期化する。

## トリアージする

**人間の明示指示でのみ起動する。** 自動起動（cron / scheduled agent / 遅延実行）は
持たない。トリアージの出力は「形の一致が N 件ある。hook 化を提案する」であり、
その場で承認を取れないと次の工程へ進めないためである。想定周期は週次だが、
その担保はこの skill の外（ユーザーの todo 管理）にある。

手順:

1. `$ROOT/triage/` にある**全ての過去レポート**（`YYYY-MM-DD.md`）を読み、各レポートの
   `## 対象 entry 一覧` 見出しから次の `##` 見出しまでだけを対象欄として抽出する。対象欄に
   列挙された entry ファイル名の**和集合**を既トリアージとみなす。レポート本文全体を grep
   して参照名を拾ってはならない。レポートが無ければ全件を対象にする。
   **日付で切らない** — `last-triage.txt` の日付単位判定では、同じ日にトリアージ後へ
   追加された記録が次回拾われない穴が実測されている（2026-08-09 に8件、うち `canon`
   保有4件が対象から外れかけた）。`last-triage.txt` は実施日の表示だけを担い、対象の
   確定には使わない。

   対象欄だけを抽出する最小形は次のとおり。`grep` の対象は各レポートの対象欄に限定する。

   ```bash
   TARGETS() {
     awk '/^## 対象 entry 一覧/{on=1; next} on && /^## /{exit} on' "$1" \
       | grep -oE '[0-9]{4}-[0-9]{2}-[0-9]{2}-[a-z0-9-]+\.md'
   }
   for report in "$ROOT"/triage/[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9].md; do
     [ -f "$report" ] && TARGETS "$report"
   done | sort -u
   ```

2. **対象集合をこの時点の `ls` で確定する（snapshot）。** 和集合に含まれないファイル名一覧を
   固定してから読み始める。トリアージ中に並行セッションが追記した entry は今回の対象に
   **入れない**（次回対象として自然に残る）。2026-08-16 の実走で、読了 114 件の裏で
   4 件が追加された — snapshot が無いと「読んだ集合」と「列挙する集合」がズレる。
3. `status.tsv` と `responses.tsv` を読み、過去の response status を先に確認する。既に
   `implemented_unverified` / `effective_provisional` の response は再実装せず、過去文を
   参照して有効性を測る。`recurred` / `ineffective` は次の修正対象を決める材料にする。
4. snapshot の `entries/` を読み、`summary` と 7 節から**形の一致**を探す。
   件数閾値で判定しない — 「気づけない失敗」はリポジトリを跨いで分散するため、
   件数駆動では永久に発火しない。**形の一致は件数閾値より早く見える。**
   **対象が数十件を超える場合は並列サブエージェントで digest してからクラスタリングする**
   （2026-08-16 に 114 件で実証した手順）。digest は 1 件につき:
   frontmatter（date/summary/skills/repo/canon/paths）／失敗の形 1 文（観測可能な行動）／
   本人分類の仮説（知らなかった・知っていて飛ばした・unclear — 仮説として扱う）／
   機械検知できるコマンド・パスのパターン（無ければ「none」）。
   フォーマット逸脱の entry（frontmatter 欠落・7 節不備）も digest で拾い、レポートに列挙する。
5. **新規対象が 0 件で、評価期限の response も無ければ何も出力せずに終了する。** 空振りのノイズでトリアージ自体が
   読まれなくなるのを防ぐ。
   新規対象が無くても、`implemented_unverified` / `effective_provisional` の評価期限が来ていれば、
   対象欄を空にした評価レポートを作り、`過去対策の評価` だけを更新する。
6. 形式化候補を提示する。振り分けは 2 軸で行う
   （`ADR-20260816-adopt-warn-only-hooks-for-syntactic-shapes`・biz_ops）:
   - 「知らなかった」型・機能欠落型 → skill / docs の不足、またはスクリプトのバグ修正。
   - 「知っていて飛ばした」型 → hook 候補。**ただし hook 化するのは、検知条件が
     ツール呼び出しの構文・実行時状態だけで書ける型のみ**。意図・網羅性・断定の真偽など
     自然文判定が要る型は hook 化しない（誤検知の常時警告化が hook 全体の信頼を毀損する）。
     強度は **warn-only を既定**とし、的中率を見てから block 昇格を個別に判断する。
   - 前方一致の許可・拒否で足りるものは hook ではなく **permission リスト**（allow/deny）へ。
   - リポジトリ固有の構成知識に依存するものは横断化せず、**当該リポジトリの issue / guide** へ。
7. **トリアージ1回につき `$ROOT/triage/YYYY-MM-DD.md` を新規 Write する。**
   「今回読んだ記録」（**次回の除外基準になるため、手順2の snapshot をそのまま実名列挙する**）
   「見つけた形の一致」「フォーマット逸脱の entry」「過去対策の評価」「起票した issue」を書く。
   `過去対策の評価` には少なくとも `response_id`、前回 status、今回の一次証拠、再発数、
   見逃し/誤警告数、次 status、次の測定を記録する。
   記録本体に結論欄を持たせない理由は、後から書き戻すと「entry は新規 Write のみで完結する」
   （並行セッションで競合しない）設計を崩すため。結論はレポート側が持つ。
8. 実施日を `triage/last-triage.txt` に記録する。

**この skill の起動時（記録時・トリアージ時とも）に未トリアージ件数と、未評価の
entry / response 件数を必ず報告する。** 起動が人間に依存するため、忘れている状態と、同じ対策を
再実行せず評価へ回すべき状態が見えるようにする。
算出は「全過去レポートの対象欄の和集合」と `entries/` の現在一覧の差分で行う。対象欄以外の
本文は絶対に入力に含めない:

```bash
triaged=$(mktemp)
for report in "$ROOT"/triage/[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9].md; do
  [ -f "$report" ] || continue
  awk '/^## 対象 entry 一覧/{on=1; next} on && /^## /{exit} on' "$report" \
    | grep -oE '[0-9]{4}-[0-9]{2}-[0-9]{2}-[a-z0-9-]+\.md' || true
done | sort -u > "$triaged"
comm -13 "$triaged" \
  <(cd "$ROOT/entries" && find . -mindepth 2 -maxdepth 2 -type f -name '*.md' \
      -exec basename {} \; | sort -u)
rm -f "$triaged"
```

件数の status 集計は上記の `status_ledger.py summary` を使う。`untriaged` は新規対象、
`legacy_unrecorded` と `implemented_unverified` は過去文と対策の一次証拠を評価する対象であり、
実装を繰り返す件数ではない。

## 形式化は所有しない

| 操作                                          | 所有                                        |
| --------------------------------------------- | ------------------------------------------- |
| 1 件の追記（7 節テンプレートの提供を含む）    | この skill                                  |
| トリアージ（人間が起動）                      | この skill                                  |
| 形式化の実行: skill 新設・改訂                | `origin-skill-commonize` へ委譲             |
| 形式化の実行: hook / settings.json            | settings.json を扱う skill へ委譲           |
| 形式化の実行: permission リスト（allow/deny） | settings.json を扱う skill へ委譲           |
| 形式化の実行: リポジトリ固有の guide / issue  | 当該リポジトリの `origin-doc-update` へ委譲 |

**受け渡しの境界を明示する。** 「skill の記述が現実とズレていた」型のトラブルは、
**集めるのがこの skill・直すのが `origin-skill-commonize`**。境界を書かないと
どちらも動かないケースが生じる。

**人間ゲートは形式化の直前に置く。** `~/.agents/` は全リポジトリ・全セッションに効く
横断正典であり、エージェントが独断で触らない領域として既に規定されている。
記録を読んで仮説を立てるところまでは承認不要、形式化は承認必須。

分析の自動実行はしない。仮説欄には誤った自己診断が混ざる前提であり、それを素材に
自動で形式化すると誤ったルールが全リポジトリに効く。

## 既存 FRICTION.md との関係

`~/.agents/skills/FRICTION.md` は**廃止**。既存 3 行は移送しない（7 節形式を満たさず
分析素材にならず、情報自体は既に別の場所で消化済み）。ファイルは廃止予告付きで残し、
古い規約で動くエージェントがここへ着地できるようにする。

## 採らなかったもの

- **分量の下限（各節 N 文以上）**: 中身のない水増しで通り、何も担保しない。
- **記録専用の対話的な埋め込み**: 品質は上がるが、記録が重くなって書かれなくなる損失が大きい。
- **「失敗の型」の分類フィールド**: 分類は分析時に決める。収集時に振り分けると、
  曖昧な記録が失われる。
- **単一ファイルへの追記 / JSONL**: 前者は並行追記が競合し書き込みコストが増え続ける。
  後者は厚い記述の escaping が壊れ差分が読めない。

## Resource

- `references/entry.template.md` — 1 件分の frontmatter と 7 節テンプレート。
- `references/status.md` — status ledger の schema、状態遷移、backfill、対策評価の契約。
- `scripts/check_entry_format.sh` — 新規 entry の形式・実文・パス表記を決定論的に検査する。
- `scripts/status_ledger.py` — entry と過去レポートを照合し、status ledger を同期・検証・更新する。
