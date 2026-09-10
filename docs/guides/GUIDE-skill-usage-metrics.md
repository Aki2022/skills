---
id: GUIDE-skill-usage-metrics
updated_at: 2026-09-09
source_issues: [ISSUE-20260909-measure-skill-usage, ISSUE-20260909-skill-usage-schedule, ISSUE-20260910-improve-loop-review-handoff]
source_workstreams: []
related_specs: []
---

# Codex / Claude skill 使用量メトリクス

## What It Does

`origin-skill-commonize/scripts/measure_skill_usage.py` は、Claude Code と Codex の
アカウント別 JSONL 履歴を同じ形式に正規化し、指定期間の明示的な skill 呼び出しを数える。
対象は正典 root にある `SKILL.md` の名前だけで、実行実績がない skill は削除候補にせず
`unknown` として残す。正典 skill 間の `/skill/references/...` 参照は依存専用の判定に使う。

## How To Use

正典 root と各 account の履歴 root を毎回明示する。ラベルは `ACCOUNT=PATH` 形式で、同じ
agent の複数 account は `--claude` / `--codex` を繰り返して指定する。

```bash
python3 ~/.agents/skills/origin-skill-commonize/scripts/measure_skill_usage.py \
  --canonical ~/.agents/skills \
  --claude seat1=~/.claude/projects \
  --claude seat2=~/.claude-seat2/projects \
  --codex seat1=~/.codex/sessions \
  --codex seat2=~/.codex-seat2/sessions \
  --days 30
```

`--since` と `--until` は ISO-8601（`--until` は排他的）で固定期間を指定でき、
`--days` は `--until` から遡る期間になる。出力はデフォルトで stdout に出る。
保存が必要なときだけ `--output PATH`（1 スナップショットを原子的に置換）または
`--append PATH`（1 行 1 JSON）を明示する。スケジューラや保存先の retention はこの
スクリプト単体からはインストールしない。

### macOS の定期収集

`install_skill_usage_launchd.py` を使うと、現在のユーザーだけに
`com.origin.skill-usage-metrics` を登録できる。既定値は毎日03:30（Macのローカル時刻）、
30日窓、180日保持で、`~/Library/LaunchAgents` と
`~/.local/state/origin-skill-usage/` 以外へ書き込まない。state directory は `0700`、
スナップショット・lock・launchd ログは `0600` で作成し、集計形式でない既存行や privacy
契約に反する行があれば追記・prune を停止する。インストーラは既存 plist が
異なる場合に上書きせず停止する。`--check` は plist と launchd の登録状態を読み取り専用で
確認し、`--install` と `--bootstrap` は明示した場合だけ永続状態を変更する。
既存 plist の差し替えは `--replace` を使うが、ロード中の job は自動で置き換えず、先に
停止手順を実行してから人間が内容を確認する。

```bash
python3 ~/.agents/skills/origin-skill-commonize/scripts/install_skill_usage_launchd.py --check
python3 ~/.agents/skills/origin-skill-commonize/scripts/install_skill_usage_launchd.py --install --bootstrap
```

登録直後に一度だけ実行結果を確認する場合は、`--kickstart` を追加する。これは四つの
履歴 root を走査して一つの集計行を保存するため、履歴量に応じて時間がかかる。

```bash
python3 ~/.agents/skills/origin-skill-commonize/scripts/install_skill_usage_launchd.py \
  --install --bootstrap --kickstart
```

停止するときは次を実行してから plist を削除する。スナップショットは別途保持・削除を判断する。

```bash
launchctl bootout "gui/$(id -u)/com.origin.skill-usage-metrics"
```

## 承認前レビュー

commit または push の前に、次の4項目を人間ゲートとして確認する。これは自動承認ではなく、
`origin-close-session` が完了時に案内するための現在の checklist である。

1. **スケジュール** — 毎日03:30、30日窓、180日保持、`RunAtLoad=false` の設定を許可する。
2. **保存データ** — `~/.local/state/origin-skill-usage/` への aggregate JSONL 保存を許可し、
   raw prompt・履歴 path・session ID・credential を保存しない契約を確認する。
3. **変更範囲** — collector・runner・LaunchAgent・docs だけが対象で、alias topology、
   Codex `.system`、plugin cache、認証・履歴は変更しないことを確認する。
4. **Git 反映** — canonical skills repository の変更を commit / 公開 push してよいか決める。

実物の登録状態は次で確認できる。`--check` が `template_valid=true plist_installed=true
plist_exact=true loaded=true`、`launchctl print` の直近終了コードが `0` であることを確認する。

```bash
python3 ~/.agents/skills/origin-skill-commonize/scripts/install_skill_usage_launchd.py --check
launchctl print "gui/$(id -u)/com.origin.skill-usage-metrics"
```

この checklist 自体は `origin-trouble-log` の entry に複製しない。実際に確認を漏らした、
誤った完了を報告した、または同型の摩擦が発生した場合だけ、同 skill の証拠 entry として記録する。

出力の主な項目は次の通り。

- `skills[]`: `used`、`dependency-only`、`unknown` の状態、総呼び出し数、account 別集計
- `sources[]`: account ごとの走査件数、期間内 prompt 数、曖昧な Codex message 数、履歴 root の可用性
- `summary`: 状態別の skill 数。`candidate` は常に 0 で、退役判断をしない
- `privacy`: prompt 本文、履歴 path、session ID、credential を保存していないこと

## Maintenance Notes

Claude/Codex の JSONL schema が変わったら、実物を本文を表示しない形で調査し、
`origin-skill-commonize/scripts/tests/test_measure_skill_usage.py` に正常系と誤検出防止の
fixture を先に追加する。その後、次を実行する。

```bash
python3 -m pytest -q origin-skill-commonize/scripts/tests/test_measure_skill_usage.py
bash origin-skill-commonize/scripts/skill_lint.sh ~/.agents/skills
```

集計結果は Git 管理の台帳へ自動転記しない。削除・無効化・mirror の変更は、完全な履歴と
依存関係を人間が確認した後に、別の承認付き変更として行う。

## Known Limitations

- 明示トークンが履歴に残らない暗黙 routing や native telemetry は数えない。
- token / context cost は skill 呼び出しに安全に帰属できないため、この集計には含めない。
- Claude の `skill-doctor` は account-local の補助証拠であり、この collector の代替ではない。
- Codex の user message は `user.text` metadata がない場合に fail-closed で `ambiguous` とし、
  使用とは数えない。古い履歴ほど `unknown` の幅が広くなる。
- 第三者 skill の本文品質や upstream parity はこの collector の責務ではない。
