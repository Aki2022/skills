# 上流 ZIP（デジタル庁デザインシステム Markdown 版）について

このリポジトリには、上流の ZIP を含めていない。照合テストを走らせるときは、取得元から手元に取得し、`DADS_TEST_ARCHIVE` でそのパスを指定する。

| 項目 | 内容 |
| --- | --- |
| 取得元 | https://design.digital.go.jp/dads/ の取得ページ（ファイル名 `dads-markdown-20261021.zip`、2026年10月21日版） |
| sha256（2026年10月21日版） | `ccab493b68878592d6f4537c24014cbe7e07f9335f6789be3ec6b5819bc9c994` |
| 著作権表示 | © Digital Agency, Government of Japan（公式サイトの表示） |
| ライセンスの状況 | 未確認。DADS のページと ZIP には、利用条件の記載が無い。サイト全体の著作権ポリシー（https://www.digital.go.jp/copyright-policy）は、権利表記が無い限り公共データ利用規約（第1.0版）（PDL1.0）を適用すると定めている。ただし、ZIP がその対象かは明記されていない |

## カタログとの関係

- `../2026-10-14.json` のカタログは、2026-10-14 版から作られ、`archive_sha256` は `562f1e19…` を pin している。2026年10月21日版の ZIP とは sha256 が一致しない。
- `scripts/tests/test_catalog.py` の `test_real_archive_has_all_560_headings` は、与えられた ZIP から生成した項目が、ZIP の見出し（foundations 8 文書・components 49 文書の h2/h3、560 件）と一致することを確かめる。カタログ pin との一致は確かめない。
- カタログの pin を 2026年10月21日版に上げるかは未決。上げる場合は `catalog.py generate --previous` で差分を出して判断する。

## 扱い

- 上流の ZIP は、この repo に置かない。ライセンスが確認できるまで、再配布しない。
- 出典は `../2026-10-14.json` の `attribution` に従う。デジタル庁が作成したものとして示さない。
- ZIP を一度含めたコミット（`3b7b49f`、main では `8a26bad`）は、リポジトリの履歴に残っている。履歴から消すかどうかは、別に判断する。
- 利用条件を確認するには、デジタル庁の問い合わせフォーム（https://www.digital.go.jp/contact）を使う。
