# Skill Friction Log（廃止）

> **このログは廃止しました。ここに追記しないでください。**
>
> **記録先は `origin-trouble-log` skill です。**
> 保管ルートは環境変数 `ORIGIN_TROUBLE_LOG_ROOT` から解決し、
> `entries/YYYY-MM/YYYY-MM-DD-<slug>.md` に 1 件 1 ファイルで記録します。
> テンプレートと手順は `~/.agents/skills/origin-trouble-log/SKILL.md` にあります。
>
> **なぜ移したか**: この形式は `<skill-name>` を必須にしており、skill を使っていない
> 場面で起きた作業規律の欠落（無音の no-op・誤った完了報告・不要な質問）を記録できません。
> 実際、2026-07-31 の規約作成から 2026-08-01 に 3 件入った後、2026-08-08 まで追記ゼロでした。
> また 1 行形式では、後の分析で hook の検知条件や skill の改訂案を導けません
> （行動の実文と「既に存在していた正解の在り処」が残らないため）。
>
> **このファイルの削除予定: 2026-11-09**（新 skill 稼働から 3 ヶ月）。
> 古い規約で動くエージェントの道標として、それまで残します。
> 既存 3 行は新形式を満たさないため移送していません。

<!-- 以下は履歴。追記しないこと -->

- 2026-08-01 | origin-git-cleanup | vibe-guard の pretooluse guard が同一コマンド内の無関係な `grep -n` を `git commit -n` と誤検知し commit がブロックされた（commit を単独実行して回避）
- 2026-08-01 | origin-git-cleanup | 同 guard が読み取りだけの hooks-path 設定値の確認もブロックする。survey 目的の読み取りと書き換えが区別されていない
- 2026-08-01 | origin-skill-commonize | 棚卸し手順が symlink 状態しか見ておらず、正典 ~/.agents/skills 自身で origin-* skill 4 件が長期 untracked だったことに気付けなかった（不変条件 7 の検証手順が無い）
