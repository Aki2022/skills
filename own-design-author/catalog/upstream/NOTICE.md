# 同梱の上流 ZIP（デジタル庁デザインシステム Markdown 版）

| 項目 | 内容 |
| --- | --- |
| ファイル | `dads-markdown-20261021.zip` |
| 版 | 2026年10月21日版（公式サイトの現行の取得ページが案内する版） |
| 取得元 | https://design.digital.go.jp/dads/dads-markdown-20261021.zip |
| 取得日 | 2026-11-07 |
| sha256 | `ccab493b68878592d6f4537c24014cbe7e07f9335f6789be3ec6b5819bc9c994` |
| 編集 | なし。公式の ZIP をそのまま同梱している |
| 著作権表示 | © Digital Agency, Government of Japan（公式サイトの表示） |
| ライセンス | 上流はライセンス名を示していない。公式サイトのライセンス表記は図・アイコン等の別資源のもので、この ZIP には付いていない |

## この ZIP と catalog の関係

- `../2026-10-14.json` のカタログは、2026-10-14 版から作られ、`archive_sha256` は `562f1e19…` を pin している。この ZIP は別の版で、sha256 は一致しない。
- `scripts/tests/test_catalog.py` の `test_real_archive_has_all_560_headings` は、この ZIP から生成した項目が、ZIP の見出し（foundations 8 文書・components 49 文書の h2/h3、560 件）と一致することを確かめる。カタログ pin との一致は確かめない。
- 環境変数 `DADS_TEST_ARCHIVE` を設定すると、この ZIP の代わりに別の ZIP を使う。
- カタログの pin を 2026-10-21 版に上げるかは未決。上げる場合は `catalog.py generate --previous` で差分を出して判断する。

## 扱い

- 上流の文書を含む。デジタル庁が作成したものとして示さない。
- カタログの出典表示は `../2026-10-14.json` の `attribution` に従う。
- 版を替えるときは、この表の全項目を更新し、古い ZIP を削除する。
