# DESIGN.md — awesome-design-md-jp Import Policy

日本語UI向けの第三者 DESIGN.md を**参照デザインライブラリ**として安全に利用するためのポリシー。
参照元は `awesome-design-md-jp`（https://github.com/kzhrknt/awesome-design-md-jp）に限定する。

## 目的

日本語UIの雰囲気・レイアウト密度・CJKタイポグラフィ・コンポーネント表現を**抽象化して**参考にする。
自社/指定 Plugin と合成し、直接コピーはしない。

## 使い方

1. **ユーザーが参考にする個別サイトを指示する**（Skillが自律的に探索・全件スキャンしない）。
2. 指示されたサイトの DESIGN.md を取得する。
3. 内部形式（Visual Theme / Color Roles / Typography / Spacing / Layout / Components /
   Iconography / Motion / Accessibility / Content Tone / Do-Don't / Examples）に正規化する（mapping.md）。
4. 抽象化されたデザイン特性として扱い、自社 Plugin または指定 Plugin と合成する。
5. Context Composer 上では**優先度6（参照 DESIGN.md）**。具体値はコピーせず、方針・密度・トーンのみ反映。

## 禁止事項（spec §15.3）

- 既存サービスの UI コピー
- ロゴ・商標・固有ビジュアルの再利用
- 競合サービスと誤認される表現
- 著作権・商標権を侵害するアセット利用

## ライセンス/権利

- 各サイトの DESIGN.md の権利は各提供元に帰属する。参照利用に限定し、`references.md` に出典を記録する。
- 具体資産は取り込まないため `allow_asset_copy: false`。
