---
id: ADR-20261104-skill-test-deps-and-wiring-shelf-rule
status: accepted
scope: development
created_at: 2026-11-04
updated_at: 2026-11-04
source_workstreams: []
source_issues: []
related_specs: []
related_guides: []
supersedes: ""
superseded_by: ""
---

# Decision: skill のテスト依存の宣言方式と、配線の棚判定・CLI 報告を分離する

## Context

skill_lint の S8 は、各 skill の `scripts/tests/` を pytest で走らせる。pytest が使えない端末では
S8 が skip され、16 件の skill テストが一度も走らなかった。pytest を入れると、テストが import する
ライブラリ（yaml・pptx）が無いものと、テストが本当に失敗するものが見えた。依存を宣言する場所が無いこと
が根因だった。

S10 の配線監査は、正典を指す symlink の親ディレクトリを「配線先（棚）」として数えていた。mise の
状態ディレクトリが正典の `mise.toml` を指す symlink を持つと、その状態ディレクトリを 1 冊ずつの棚と
誤認し、2146 件の誤検出を出した。`~/.local/bin` の個別スクリプトへの symlink も同じ規則で棚扱いされていた。
同じ規則が、宙ぶらりんの CLI リンク（本物の問題だが skill の配線の問題ではない）も RESULT に混ぜていた。

## Decision

1. **テスト依存は skill ごとに宣言する。** 依存を持つ skill は `scripts/tests/requirements.txt` に
   pin 済みの依存を置き（`uv pip compile` で解決）、S8 は `uv run --no-project --with-requirements <file>
   -- pytest` で走らせる。依存を持たない skill は mise の global の pytest（pypi:pytest 9.2.0、
   `~/.config/mise/mise.lock` で固定）で走らせる。
2. **棚は正典直下の skill ディレクトリそのものを指す symlink だけで判定する。** 個別スクリプトへの symlink、
   正典直下のファイル、mise の状態ファイルは棚として数えない。
3. **宙ぶらりんの CLI リンクは判定から分離し、CLI 行で報告する。** `--cli-bin`（既定は `~/.local/bin`）の
   宙ぶらりん symlink は RESULT と S10 の失敗に入れず、`CLI 宙を指している:` 行として出す。lint は
   `CLI ` 行を fail にせず note として出す。

## Alternatives

- **global の Python に pip で入れる**: 依存が混ざり、skill 間で分離が崩れるので却下。
- **全 skill を uv に一本化する**: pytest の固定先が一つになるが、mise を中心にするという既存方針を
  崩すので却下。mise の global は CLI の固定に、uv は依存付きテストの実行に分ける。
- **skill ごとの永続 venv**: リポが増えると容量を圧迫するので却下。uv のキャッシュを共有する
  （clone 方式で同じ実体を二重に持たない）ほうを採る。
- **`~/.local/bin` を棚として数え続ける**: 誤判定が常時赤になるので却下。
- **宙ぶらりんの CLI リンクも S10 の失敗にする**: CLI リンクは skill の配線ではなく分類が違うので却下。
  代わりに、lint は CLI リンク切れでは赤くならない。見落とすリスクは残るので、Consequences に明記する。

## Consequences

- S8 は依存を宣言した skill で実テストを走らせる。依存を宣言した 4 skill のうち通ったのは 1 件
  （`own-project-update`、387 件）。残る 3 件（`own-design-author`・`own-infoarch-author`・`own-pptx-build`）は
  依存が揃った上でのテスト本体の失敗で、依存不足ではない。依存を持たない `own-doc-update` は、未追跡の作業中
  テストが未実装の関数を呼んで失敗する。
- S10 の配線先は 8 件（棚ごと 4 / 1冊ずつ 4）。見なくなったもの: `~/.local/bin` 内の「正典の外を指す」
  「正典に無い名前」「symlink でない実体」。これらは skill の配線監査としてはノイズだった。
- 宙ぶらりんの CLI リンク（現在 10 件）は報告のみ。lint は赤くならないので、見落としやすくなる。
- `requirements.txt` は pin 済みだが、macOS（arm64・Python 3.12）で解決したもの。他のプラットフォームでは
  再解決が要る。

## Verification

- 先に赤のテストを書き、今の実装で赤を確認してから直した（監査テスト、S8 の uv 経路テスト）。
- 監査テスト 17 件、skill-commonize 全体 109 件が緑。
- 実物の S10: 2251 件の誤検出から、FAIL 0 件へ。S8: 16 件中 12 件が成功（残る 4 件は既存の失敗）。

## Open

- 宙ぶらりんの CLI リンク 10 件を直すか消すかは未決（人が決める）。
- S8 の残る 4 件の失敗（`own-design-author`・`own-infoarch-author`・`own-doc-update`・`own-pptx-build`）は
  未決。`own-design-author` と `own-infoarch-author` は未追跡の作業中ディレクトリで、その `requirements.txt`
  も未コミット。
