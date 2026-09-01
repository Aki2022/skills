---
id: SPEC-pptx-content-eval-loop
status: active # draft | active | superseded
created_at: 2026-08-31
updated_at: 2026-08-31
related_guides: []
affected_workstreams: []
---

# origin-pptx コンテンツ評価ループの公式化

## Purpose

20260831_nissay デッキ制作（kanpu_kaigo リポ）でアドホックに運用して有効だった「subagent 評価を基準点まで反復するループ」を、origin-pptx スキルの正式工程として組み込む。あわせて同セッションの実測トラブル3件（検分の誤ok・バッチ完了判定の穴・審査合格と可読性の乖離）への対策を skill 本体へ反映する。

## Problem

- origin-pptx の評価系は ②.5 persona-eval（任意・合成デッキ対象）と ④ VLM 比較（構図一致）のみで、**①テキスト段階（base.md / outline.md）の内容品質を測る工程が無い**。今回そこをアドホックに補い、v8→v14 の品質を駆動した
- 審査員型ペルソナだけを回すと防御的追記が積み上がり可読性が劣化する（2026-08-31 実測: AI審査6回合格の直後に人間初見レビューで本編全面差し戻し・モックアップ20枚廃棄）
- 評価はトークン大量消費（今回の Luna 審査は入力 0.8〜1.4M/人/回）。**何を読ませ・どのモデルで回すか**の規律が無いと運用できない

## Goals

- base.md / outline.md / 画像（mockup・preview）の各段階に、評価 subagent のループを正式工程として定義する
- 「採点して合否を出す審査」と「合否に関与しない助言」を役割分離する
- 段階×役割ごとのモデル割当・入力範囲・回数上限を skill-config で規定し、トークン消費を予算内に収める
- 今回実測済みの対策（seen 書き出しの義務化・バッチ完了判定・codex2）を正典化する

## Non-Goals

- 評価ループによる人間ゲートの代替（①共同確定・②承認・④.5 は従来どおり人間）
- persona-eval（②.5 合成デッキ評価）の廃止

## Domain Language

| Term                     | Meaning                                                                    | Boundary or contrast                           |
| ------------------------ | -------------------------------------------------------------------------- | ---------------------------------------------- |
| 審査員（judge）          | rubric で採点し verdict / must_fix を出す評価者。ループの合否ゲートに使う  | アドバイザーと違い合否に効く                   |
| 初見読者（fresh reader） | 予備知識ゼロで対象だけを読み、理解可否と素朴な疑問を返す評価者             | 採点しないが「分からない」判定は must-fix 相当（2026-08-31 決定） |
| アドバイザー（advisor）  | 修正フェーズで不合格箇所に構成案3つ＋推奨を出す設計者。評価ループには参加しない（2026-08-31 決定）                                 | advice 欄=審査員の軽い助言／アドバイザー=fail 箇所の再設計 |
| フレッシュ評価           | 過去版・過去評価の文脈を与えない新規 subagent による評価                   | 同一エージェントの再評価（文脈汚染あり）と区別 |
| seen 書き出し            | 画像検分で、判定前に「実際に写っている主図解・要素数」を記述させる必須出力 | 2026-08-31 実測で誤 ok を根絶                  |

## Actors and Scenarios

- orchestrator（メインループ）: 評価計画の提示・subagent 起動・結果集約・修正指示
- 評価 subagent: 審査員 / 初見読者 / アドバイザー / 画像検分（Haiku）
- 人間: 評価計画の承認、ループ結果を材料に各ゲート（①②④.5）を判定

## Requirements

### Functional

- **発火条件＝計画承認制（2026-08-31 決定・案A）**: 各段階（base/outline/画像）で orchestrator が評価計画（対象・評価者の構成と人数・モデル・回数上限・概算トークン）を必ず提案し、人間の承認後にのみ実行する。承認なしの評価ループは回さない。自律ラン（origin-goal 等）で使う場合は、WS の Authorization Envelope に評価計画の事前承認を明記することで往復を省略できる

- **評価者の構成＝2役評価＋修正時アドバイザー（2026-08-31 決定・案A）**: 評価ループは**審査員**（rubric採点・verdict・must_fix・advice欄）と**初見読者**（採点なし・予備知識ゼロで対象だけを読み、スライド別の 分かる/引っかかる/分からない と素朴な疑問を返す）の2役で回す。**初見読者の「分からない」が本編スライドに1枚以上あれば must-fix 相当としてループ継続**（可読性を合否に組み込む）。**アドバイザー（ゼロベース設計者）は評価者ではなく修正フェーズの道具**——不合格箇所に対して「構成案を3つ＋推奨1つ＋根拠」を出させる形式で投入する（2026-08-31 の S2/S3/S9 再設計で実証済み・fail 箇所だけに使うためトークン効率が良い）
- **モデル割当の既定（2026-08-31 決定・skill-config `contentEval` 節に記載・デッキごとに変更可）**:
  - 読者評価: 審査員・初見読者 = **opus**。画像段階は全スライドのレンダ画像を1体が通しで見る（読書体験の再現。画像16枚≒30〜40k入力）
  - アドバイザー（fail箇所の再設計・修正フェーズのみ）= **opus**
  - 一致検査: VLM構図比較・数値監査・ビルダー = **sonnet**／画像検分・④.5送付前検分 = **haiku**（seen 書き出し義務化）
  - 計画承認時の見積り式: 概算トークン ≒ 評価者数 × 回数 ×（対象サイズ×1.5＋出力10k）
- **合格条件・ループ上限の既定（2026-08-31 決定・案A・skill-config に置き計画承認時に上書き可）**:
  - 合格 = 各審査員75以上 **かつ** 平均80以上 **かつ** 重大must-fixゼロ **かつ** 初見読者の「分からない」が本編0枚
  - 回数 = 最大3回。超過継続は人間承認で+N（承認時にコスト再見積りを提示）
  - **再評価は毎回フレッシュ**（過去版・過去評価の文脈を与えない新規 subagent）。ただし「前回 must-fix の反映検証」だけは事実として審査員プロンプトに渡す
  - 収束ガード = 前巡より指摘が減らなければ上限前でも人間へエスカレーション（既存 verifyLoop と同型）
- 確定済み（2026-08-31 ユーザー指示）:
  - 各段階（base.md / outline.md / 画像）の評価ループを skill の正式工程として組み込む
  - 審査（採点）と別に、アドバイスをもらう機能の付加を検討する（形は grill で確定）
  - subagent のモデル割当はトークン消費の観点で明示的に規定する
  - codex は `codex2` エイリアス優先で呼び出す（実装・push 済み: skills bc4e776 / skill-config imageGen.codexBin）

### Quality

- **読者評価と一致検査の分離（2026-08-31 決定・人間指摘）**: 評価活動を2種に分け、混ぜない
  - **読者評価**（審査員・初見読者）: 評価対象＝**その段階で読み手が実際に受け取る成果物・それ単独**。base.md 段階は base.md のみ／outline 段階は outline.md のみ／**画像段階（mockup・preview）は画像のみ**——.md をサイド情報として併給しない（読み手が持たない情報で評価が現実から歪むため）
  - **一致検査**（VLM構図比較・数値監査・検分）: 仕様との照合が本務であり、outline・mockup・base.md 該当節を参照する。読者評価の代替にはならない
- 評価ループの追加トークンは、上記の対象単独規約（＝自足化）が主レバー（実測: base全文+research併給 0.8〜1.4M/体 → outline単独 67〜86k/体）。outline は本文・構図メモ・出典で自足するよう書く

## Design Direction

- **実装先（2026-08-31 決定・案A・改訂統合）**: `references/persona-eval.md` を `references/content-eval.md` へ改名・全面改訂して1文書に一般化する（①テキスト段階／画像段階の読者評価・一致検査との分離・計画承認制・2役＋修正時アドバイザー・閾値/モデル既定）。旧②.5 の合成デッキ手順は同文書の1節として吸収し、SKILL.md 等の参照を更新する。閾値・モデル既定は `style-guide/skill-config.json` の `contentEval` 節が実行時の正
- 併せて反映する実測済み対策: ④検分の seen 義務化（pipeline.md）／バッチ生成の完了判定・スモーク生成（image_gen.md＋ランナー雛形）／①ゲートのダイジェスト提示（SKILL.md）

## Constraints and Tradeoffs

- skills リポは public——push 前の人間確認は不要になった前例を作らない（本 spec の実装コミットも人間確認後に push）
- 評価回数×入力サイズ×モデル単価がコストの3因子。回数を固定しても入力が野放しなら破綻する（今回実測: 1審査員1回で入力最大1.4M）

## Acceptance Criteria

1. `origin-pptx/references/content-eval.md` が存在し、本 spec の決定（計画承認制・2役＋修正時アドバイザー・読者評価の対象単独規約・段階マップ・閾値/回数既定・フレッシュ評価・収束ガード）を全て含む。`persona-eval.md` は残存しない
2. `git grep -l "persona-eval" -- origin-pptx/` が 0 件（参照の付け替え漏れなし）
3. `style-guide/skill-config.json` に `contentEval` 節（モデル割当・passThreshold・maxRounds）があり JSON として妥当
4. `references/pipeline.md` の④検分・VLM比較に「seen（実写の書き出し）を判定前に必須出力とし、書き出しなしの ok は無効」が明記されている
5. `references/image_gen.md` にバッチ生成規律（旧成果物の事前退避・回収失敗の exit 反映・開始前スモーク1枚）が明記され、`scripts/run_mockups.sh`（正典ランナー雛形）が存在する
6. `SKILL.md` の①に「ダイジェスト提示（キーメッセージ一覧＋各枚1行要約）で人間確定を取る」と content-eval への参照があり、パイプライン表が更新されている
7. `origin-skill-commonize/scripts/skill_lint.sh` が FAIL ゼロで通る
8. 実装コミットは人間確認後に push（public リポ）

## Impact on Existing System

- `origin-pptx/SKILL.md`（①手順・パイプライン表）、`references/persona-eval.md`、`references/pipeline.md`（④検分）、`references/image_gen.md`（バッチ）、`style-guide/skill-config.json`（モデル割当）
- 実装は origin-skill-commonize の規約（in-place 編集・lint）に従う

## Deferred Decisions


