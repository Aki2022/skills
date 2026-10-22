# 出荷ゲートの接続

実装元: biz_ops `WS-20261021-design-system-standardization` ISSUE-04。

`design-ship` は同じ `design-author check --ship` を実行する。L0失格はexit 2、警告だけなら内容を表示してexit 0。
起動不能・JSON破損・終了コード不一致も停止する。提案状態のnice項目だけで出荷を止めない。
任意の文字列をshell評価せず、`DESIGN_AUTHOR_COMMAND` は単一の実行ファイルを指定する。

## 導入前

対象repoのL0を確認し、既存のvibe-guard pre-pushが `pre-push.d` 拡張を実行することを読む。
既存hook・scanner・他の拡張は保持する。対応しないhookを自動で書き換えない。
本番repoへの導入はユーザー確認後に行う。skillは自動インストーラーを提供しない。

## 接続するもの

- PATHへskillの `scripts/` を追加し、既存依存を持つPythonを `DESIGN_AUTHOR_PYTHON` で選ぶ。
  repoの設定には端末固有のパスを保存しない。
- `templates/mise.design.toml` のtaskを既存mise設定へ統合する。`mise run design:ship` が同じ入口を呼ぶ。
- `hooks/pre-push.d/60-design-ship` を既存vibe-guardの拡張ディレクトリへ追加する。
  既存hookを置換せず、core.hooksPathを変更しない。
- 有効化するrepoだけ `.vibe-guard/design-ship.enabled` に `templates/design-ship.enabled` の内容を置く。
  markerは文字列1のみ。壊れたmarker・欠落コマンドは失格。

global拡張はmarkerを持たないrepoでは何もしない。markerを持つrepoではdesign/不在も出荷を止める。
`DESIGN_SHIP_COMMAND` でPATHコマンドの代わりに単一実行ファイルを指定できる。repoに個別端末パスを書かない。

## 実経路の検証

自己テストは一時Git設定に元設定をincludeし、実vibe-guard scannerと既存拡張を保持したhookコピーへ
新拡張を足す。ユーザーのGit設定・global hook・repo hook設定は編集しない。
人工repoの作業branchからローカルbare remoteへ実際にpushし、未承認mustは停止、人工承認後は到達を検証する。
scannerが両pushで実行されたことも確認する。mise taskも同じ未承認・承認状態で実行する。

実環境の試験は `DESIGN_TEST_VIBE_ROOT` と、miseも検証する場合は `DESIGN_TEST_MISE` を指定する。
未指定時のskipは実hookの合格とは扱わない。実プロジェクトの承認台帳はこの試験で作成しない。
