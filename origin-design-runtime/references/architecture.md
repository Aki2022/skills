---
title: Architecture — Runtime / Media / Plugin / Validator の境界
---

# Architecture

design-runtime は 4 つの層に責務を分離する。この境界を崩さないことが、Plugin追加時に
Skill本体（SKILL.md）を変更しなくて済む（拡張容易性）ための前提である。

## 層と責務

| 層                      | 実体                                             | 責務                                                                     | 持たないもの                       |
| ----------------------- | ------------------------------------------------ | ------------------------------------------------------------------------ | ---------------------------------- |
| **Runtime**             | `SKILL.md` + `references/`                       | タスク理解・媒体判定・Plugin選択・合成・生成指示・検証ループ制御         | 具体的なデザイン値、媒体固有ルール |
| **Media Plugin**        | `media/<media>/`                                 | 媒体固有の制約と推奨（HTML/PPTX/Dashboard）                              | ブランド・配色・ロゴ等の意匠       |
| **Design/Asset Plugin** | `plugins/<plugin>/`                              | ブランド・デザインシステム固有の DESIGN.md / tokens / assets / templates | 媒体固有の実装制約                 |
| **Validator**           | `validators/` + 各 plugin/media の `validators/` | 生成物の点検（[自動]/[自己点検]）                                        | デザイン生成そのもの               |

## データフロー

```text
User Task
  → Task Analyzer（媒体・目的・文脈を抽出）
  → Media Resolver（media/ をスキャンし媒体を選ぶ）
  → Design Plugin Resolver（plugins/ をスキャンしPluginを選ぶ。extendsは1階層）
  → Context Composer（DESIGN.md + Media Rule + tokens を優先度順に合成: context-composer.md）
  → Generator（成果物を生成）
  → Validator（[自動]→[自己点検]、不合格なら最大2回まで修正）
  → Final Artifact / Spec + 使用Plugin・制約・未解決項目の記録
```

## Media Plugin と Asset Plugin を分離する理由

同じ Design Plugin（例: digital-agency）を HTML にも PPTX にも Dashboard にも適用できるようにするため。
媒体制約（レスポンシブ、スライド比率など）とブランド意匠（配色、余白の思想）を別軸で組み合わせる。

## 継承（extends）

- MVP では **1階層まで**（子 → 親）。親がさらに親を持つ多段継承は不可。
- 継承の解決は Runtime 内部で行い、ユーザーからは「1個の Plugin を選んだ」ように見える。
- 合成順序と競合解決は `context-composer.md` に従う。

## Runtime デフォルト方針（Pluginなし時のフォールバック）

候補 Plugin が無い場合に適用する最小限の中立方針。

- 可読性優先（十分なコントラスト、本文14px相当以上）。
- 余白を確保し情報を詰め込みすぎない。
- 日本語（CJK）の行間・字間に配慮する。
- 媒体の `media.yaml` の `type: constraint` は常に守る。
- Pluginなしで生成した旨を成果物に明記する。

## 発見機構

Runtime は起動時に `plugins/*/plugin.yaml` と `media/*/media.yaml` をスキャンして一覧を作る。
専用レジストリファイルは持たない。追加はディレクトリと yaml を置くだけ。複数の検索パス
（プロジェクトローカル / global `~/.agents/skills/design-runtime/`）がある場合はローカルを優先。
