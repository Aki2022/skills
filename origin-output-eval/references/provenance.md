---
title: provenance.json — 評価者の申告
---

# provenance.json

各 `round_K/` に**必ず**置く。無ければ `judge_round.py` は判定せず `exit 2` で止まる。

なぜ機械検査にするか: 「作った本人に採点させない」「毎巡フレッシュ」「読者にサイド情報を渡さない」
「キャッシュを外す」は、文書に書くだけでは破られた実績がある（`origin-image-gen` で禁止形が破られ、
窓口化した）。申告を必須にして食い違いを拒否すれば、規律が構造になる。
そして申告の内容はそのまま最終報告の「評価者の起動条件」欄になる。

```json
{
  "round": 2,
  "evaluators": [
    { "role": "judge",  "id": "j1", "fresh": true, "inputs": ["artifact", "rubric", "carry"] },
    { "role": "judge",  "id": "j2", "fresh": true, "inputs": ["artifact", "rubric", "carry"] },
    { "role": "reader", "id": "r1", "fresh": true, "inputs": ["artifact"] }
  ],
  "gates_passed": ["audit_html.py exit 0", "measure_lp.mjs 失格ゼロ"],
  "cache_cleared": true
}
```

## 検査される内容（違反はすべて exit 2）

| 項目 | 内容 | 塞いでいる失敗 |
| --- | --- | --- |
| `round` | `round_K` の K と一致する | 前巡の残骸が混ざったまま集計される |
| `evaluators[].id` | `judges/<id>.json` / `readers/<id>.json` が実在する | **評価者の起動が失敗したのに合格に見える**（計画3人で1人しか届かない） |
| 逆向き | 実ファイルに対応する申告がある | 未申告の評価者が紛れ込む |
| `fresh` | すべて `true` | 同一エージェントの再評価は文脈汚染で甘くなる |
| `inputs`（reader） | `artifact` のみ | 読者にサイド情報を渡すと「そう書いてあるから読める」に流れる。**carry も渡さない** |
| `inputs`（judge） | `artifact` / `rubric` / `carry` のみ | 同上（設計意図・仕様書・過去版を渡さない） |
| `cache_cleared` | `true` | 採点者が修正前の版を見て報告した実測がある |

`gates_passed` は前段の決定論ゲートの申告（`protocol.md` §1）。内容は検査せず、最終報告へ中継する。

## 書くタイミング

**評価者を立てる前**に書く。後から書くと「実際に何を渡したか」ではなく「渡したつもり」を書くことになる。
`id` は評価者出力のファイル名（拡張子なし）に一致させる。
