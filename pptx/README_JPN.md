# PPTX スタイルシステム

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE.txt)

PowerPointプレゼンテーションのスタイリングを自動化するテンプレートベースのシステム。Python、R、Mermaidで一貫したデザインを実現します。

[English README](README.md)

## 特徴

- **Single Source of Truth**: すべてのスタイルを`style.yaml`で一元管理、テンプレートから自動抽出
- **テンプレートファースト**: `template.pptx`と`Chart.crtx`をビジュアル編集し、プログラムで抽出
- **多言語対応**: Python、R、Mermaidで統一されたスタイリング
- **ネイティブ編集可能オブジェクト**: 表、グラフ、Mermaid図表をすべてネイティブPowerPointシェイプとして描画
- **自動検証**: データ検証機能と詳細なエラーログ
- **自動セットアップ**: プロジェクトディレクトリとスタイル設定を初回実行時に自動作成

## クイックスタート

### インストール

```bash
# Python依存関係
pip install python-pptx lxml pyyaml pillow

# R依存関係（オプション）
R -e "install.packages(c('ggplot2', 'yaml', 'dplyr', 'tidyr'))"

# Mermaid CLI（オプション、図表生成用）
npm install -g @mermaid-js/mermaid-cli
```

### 基本的な使い方

```python
import sys
sys.path.insert(0, '~/.claude/skills/pptx')

from pptx import Presentation
from scripts.native_objects import create_styled_table, create_styled_chart

# テンプレート読み込み
prs = Presentation('templates/template.pptx')
slide = prs.slides.add_slide(prs.slide_layouts[10])

# コンテンツプレースホルダーを探す
for shape in slide.shapes:
    if shape.is_placeholder and shape.placeholder_format.idx == 1:
        # 表を作成
        table_spec = {
            'data': [
                ['項目', '値A', '値B'],
                ['データ1', '100', '200'],
                ['データ2', '150', '250']
            ],
            'header_row': True
        }
        create_styled_table(slide, shape, table_spec)

prs.save('output.pptx')
```

## プロジェクト構成

```
├── README.md              # 英語版README
├── README_JPN.md          # このファイル
├── LICENSE.txt            # MITライセンス
├── SKILL.md               # スキル定義（Claude Code用）
├── style.yaml             # 自動生成されたスタイル定義
├── templates/
│   ├── template.pptx      # スライドテンプレート（手動編集）
│   └── template.crtx      # チャートテンプレート（手動編集）
├── scripts/
│   ├── extract_style.py         # テンプレートからスタイル抽出
│   ├── style_config.py          # Pythonスタイルローダー
│   ├── style_config.R           # Rスタイルローダー
│   ├── native_objects.py        # 表/グラフ/図表作成
│   ├── crtx_utils.py            # チャートテンプレートユーティリティ
│   ├── mermaid_to_shapes.py     # Mermaid → ネイティブシェイプ
│   ├── logging_utils.py         # 自動設定ログ
│   ├── layout_registry.py       # レイアウト管理
│   └── generate_template.py     # TEMPLATE.md自動生成
└── examples/                    # 使用例
```

## ワークフロー

### 1. スタイル抽出

テンプレートからスタイルを抽出して`style.yaml`を生成:

```bash
python scripts/extract_style.py
```

抽出元:
- `templates/template.crtx` - チャートスタイル（系列色、軸、凡例、データラベル）
- `templates/template.pptx` スライド33 - 表スタイル
- `templates/template.pptx` スライド34 - フローチャート/図表スタイル

### 2. プレゼンテーション作成

表とグラフは必ず`native_objects.py`を使用 - 完全なスタイル適用とデータ検証を実行:

```python
from scripts.native_objects import create_styled_table, create_styled_chart

# スタイル付きグラフを作成
chart_spec = {
    'chart_kind': 'column',  # 'line', 'bar', 'pie'
    'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
    'series': [
        {'name': '売上', 'values': [100, 120, 110, 130]},
        {'name': 'コスト', 'values': [80, 90, 85, 95]}
    ]
}
create_styled_chart(slide, placeholder, chart_spec)
```

### 3. Rでの使用

```r
source("scripts/style_config.R")
style <- load_style("style.yaml")

# ggplot2用の色を取得
colors <- get_series_colors(style, 3)

# 一貫したスタイルでプロットを作成
p <- ggplot(data, aes(x, y)) +
  geom_bar(fill = get_primary_color(style)) +
  theme_minimal()
```

### 4. Mermaid図表

```python
from scripts.mermaid_to_shapes import create_flowchart_shapes

mermaid_code = """flowchart LR
    A[開始] --> B{判断}
    B -->|Yes| C[処理]
    B -->|No| D[終了]"""

create_flowchart_shapes(slide, placeholder, mermaid_code)
```

## 高度な機能

### Rグラフ

複雑なggplot2グラフには、style.yamlを使用して一貫したスタイリングを適用:

```r
source("scripts/style_config.R")
style <- load_style("style.yaml")

p <- ggplot(data, aes(x, y)) +
  geom_bar(fill = get_primary_color(style)) +
  theme_minimal()

# PNGとして保存し、PowerPointに挿入
ggsave("chart.png", p, width = 10, height = 6, dpi = 300)
```

### ログ機能

すべての操作は自動的に`processing/pptx_generation.log`に記録:

```bash
# エラーや警告を確認
cat processing/pptx_generation.log
```

### テンプレートレイアウト

`template.pptx`の主要レイアウト:

| インデックス | 名前 | 用途 |
|-------|------|-----|
| 0 | Title Slide | タイトルページ |
| 8 | Section Header | セクション区切り |
| 10 | Title and Content_withKeyMessage | メインコンテンツ |
| 19 | Content with Caption_withKeyMessage | グラフ+説明 |

## トラブルシューティング

### グラフ作成が失敗する

ログを確認:
```bash
cat processing/pptx_generation.log
```

よくある問題:
- **"Series 'X' contains non-numeric value"** → すべてのグラフ値は数値である必要があります
- **"Chart.crtx not found"** → テンプレートパスの問題（templates/ディレクトリを確認）
- **"Unknown theme color"** → style.yamlのテーマカラー定義を確認

### 表作成が失敗する

- **"Row X has Y columns, expected Z"** → データ配列の列数の一貫性を確認
- **"Table spec.data is empty"** → データが空でないことを確認

### スタイルが適用されない

- **"Failed to apply category axis styling"** → template.crtxの互換性を確認
- コンソールに警告が表示される → 詳細は`processing/pptx_generation.log`を確認

## セキュリティ

このプロジェクトはセキュリティレビュー済み:

- ✅ ハードコードされた認証情報や秘密情報なし
- ✅ 安全なYAMLパース（`yaml.safe_load()`使用）
- ✅ サブプロセス呼び出しは引数リストを使用（シェルインジェクションリスクなし）
- ✅ 任意コード実行なし
- ✅ lxmlによる安全なXML/OOXMLパース
- ✅ 検証済みファイル操作

## コントリビューション

貢献を歓迎します！以下の手順でお願いします:

1. リポジトリをフォーク
2. フィーチャーブランチを作成
3. 変更を加える
4. 該当する場合はテストを追加
5. プルリクエストを提出

## ライセンス

MITライセンス - 詳細は[LICENSE.txt](LICENSE.txt)を参照

## 謝辞

- [python-pptx](https://python-pptx.readthedocs.io/)で構築
- [Mermaid](https://mermaid.js.org/)を図表生成に使用
- [Claude Code](https://claude.com/claude-code)での使用を想定して設計
