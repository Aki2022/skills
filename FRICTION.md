# Skill Friction Log

スキルの実使用で気づいた摩擦（不発・曖昧な指示・手戻り・現実とのズレ）を 1 件 1 行で追記する。
形式: `- YYYY-MM-DD | <skill-name> | <症状を一文で>`
運用規約は `origin-skill-commonize/SKILL.md` の「スキル品質の保守」を参照。
追記は自由、スキル本文の改変はこのログを根拠に人間が判断してから行う。

<!-- entries below -->

- 2026-08-01 | origin-git-cleanup | vibe-guard の pretooluse guard が同一コマンド内の無関係な `grep -n` を `git commit -n` と誤検知し commit がブロックされた（commit を単独実行して回避）
- 2026-08-01 | origin-git-cleanup | 同 guard が読み取りだけの hooks-path 設定値の確認もブロックする。survey 目的の読み取りと書き換えが区別されていない
- 2026-08-01 | origin-skill-commonize | 棚卸し手順が symlink 状態しか見ておらず、正典 ~/.agents/skills 自身で origin-* skill 4 件が長期 untracked だったことに気付けなかった（不変条件 7 の検証手順が無い）
