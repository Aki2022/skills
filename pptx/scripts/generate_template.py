#!/usr/bin/env python3
"""
TEMPLATE.mdを自動生成

template.pptxから動的に情報を抽出し、Markdown形式でドキュメントを生成する。
スライド数やレイアウトが変更されても自動的に対応する。
"""

import sys
import os
from datetime import datetime

# pptxスキルのスクリプトをインポート
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from layout_registry import LayoutRegistry

from pptx import Presentation


def generate_template_md(template_path: str, output_path: str = None):
    """
    TEMPLATE.mdを自動生成

    Args:
        template_path: template.pptxへのパス
        output_path: 出力先（省略時は標準出力）
    """
    registry = LayoutRegistry(template_path)
    prs = registry.prs

    lines = []

    # ヘッダー
    lines.append("# Template Inventory (Auto-generated)")
    lines.append("")
    lines.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"**Template:** {os.path.basename(template_path)}")
    lines.append(f"**Total Slides:** {registry.get_slide_count()}")
    lines.append(f"**Total Layouts:** {registry.get_layout_count()}")
    lines.append(f"**Used Layouts:** {len(registry.get_used_layouts())}")
    lines.append(f"**Unused Layouts:** {len(registry.get_unused_layouts())}")
    lines.append("")
    lines.append("> NOTE: スライド番号は0-indexed（Slide 0が最初のスライド）")
    lines.append("> レイアウト番号はoutline.mdで指定する際に使用")
    lines.append("")
    lines.append("---")
    lines.append("")

    # レイアウト一覧
    lines.append("## Layouts Overview")
    lines.append("")
    lines.append("| Index | Layout Name | Used | Examples |")
    lines.append("|-------|-------------|------|----------|")

    for idx in sorted(registry._layouts.keys()):
        info = registry._layouts[idx]
        used = "✓" if info.example_slides else "✗"
        examples = ", ".join([f"Slide {s}" for s in info.example_slides[:3]])
        if len(info.example_slides) > 3:
            examples += f" ... (+{len(info.example_slides)-3})"
        if not examples:
            examples = "なし"

        lines.append(f"| {idx} | {info.name} | {used} | {examples} |")

    lines.append("")
    lines.append("---")
    lines.append("")

    # スライド詳細
    lines.append("## Slides Detail")
    lines.append("")

    for slide_idx, slide in enumerate(prs.slides):
        # このスライドのレイアウトを特定
        layout_idx = None
        layout_name = ""
        for idx, layout in enumerate(prs.slide_layouts):
            if slide.slide_layout == layout:
                layout_idx = idx
                layout_name = layout.name
                break

        lines.append(f"### Slide {slide_idx}: {layout_name} [Layout {layout_idx}]")
        lines.append("")

        # タイトルを抽出（あれば）
        if slide.shapes.title:
            try:
                title_text = slide.shapes.title.text
                if title_text:
                    lines.append(f"**Title:** {title_text}")
                    lines.append("")
            except:
                pass

        # プレースホルダー情報
        placeholders = []
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_idx = shape.placeholder_format.idx
                ph_type = str(shape.placeholder_format.type)
                placeholders.append(f"idx={ph_idx} ({ph_type})")

        if placeholders:
            lines.append(f"**Placeholders:** {', '.join(placeholders)}")
            lines.append("")

        # レイアウト情報
        lines.append(f"**Layout:** `Layout {layout_idx}: {layout_name}`")
        lines.append("")

        # outline.mdでの使用例
        lines.append(f"**outline.mdでの指定例:**")
        lines.append("```markdown")
        lines.append(f"## スライドX: タイトル")
        lines.append(f"**レイアウト**: {layout_name} (layout {layout_idx})")
        lines.append("```")
        lines.append("")
        lines.append("---")
        lines.append("")

    # 未使用レイアウト
    unused = registry.get_unused_layouts()
    if unused:
        lines.append("## Unused Layouts")
        lines.append("")
        lines.append("以下のレイアウトは実例スライドがありません:")
        lines.append("")

        for idx in unused:
            info = registry._layouts[idx]
            lines.append(f"- **Layout {idx}: {info.name}**")

        lines.append("")
        lines.append("> これらのレイアウトも使用可能ですが、実例を参照できません。")
        lines.append("")

    # フッター
    lines.append("---")
    lines.append("")
    lines.append("## Usage Notes")
    lines.append("")
    lines.append("### outline.mdでレイアウトを指定する方法")
    lines.append("")
    lines.append("```markdown")
    lines.append("## スライド1: タイトルスライド")
    lines.append("**レイアウト**: Title Slide (layout 0)")
    lines.append("")
    lines.append("## スライド2: コンテンツ")
    lines.append("**レイアウト**: Title and Content_withKeyMessage (layout 10)")
    lines.append("**コンテンツタイプ**: TABLE")
    lines.append("```")
    lines.append("")
    lines.append("### AIによる自動選択")
    lines.append("")
    lines.append("generate_presentation.pyは、コンテンツタイプに応じて")
    lines.append("適切なレイアウトを自動選択できます:")
    lines.append("")
    lines.append("- 1カラムテキスト → Layout 10")
    lines.append("- 1カラムグラフ/表 → Layout 11")
    lines.append("- 2カラム比較 → Layout 12/13")
    lines.append("- 3カラム → Layout 16/17")
    lines.append("")

    # 出力
    content = "\n".join(lines)

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✅ Generated: {output_path}")
        print(f"   Slides: {registry.get_slide_count()}")
        print(f"   Layouts: {registry.get_layout_count()}")
    else:
        print(content)


def main():
    """メイン関数"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(script_dir)

    template_path = os.path.join(skill_root, 'templates', 'template.pptx')
    output_path = os.path.join(skill_root, 'templates', 'TEMPLATE.md')

    generate_template_md(template_path, output_path)


if __name__ == '__main__':
    main()
