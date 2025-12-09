#!/usr/bin/env python3
"""
Layout Registry - 動的にtemplate.pptxからレイアウト情報を取得・管理

ハードコードを避け、template.pptxの変更に自動対応する。
"""

from pptx import Presentation
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from collections import defaultdict


@dataclass
class LayoutInfo:
    """レイアウト情報"""
    index: int
    name: str
    placeholders: List[Tuple[int, str]]  # [(idx, type), ...]
    example_slides: List[int]  # このレイアウトを使用しているスライド番号

    def __repr__(self):
        return f"Layout({self.index}, {self.name}, {len(self.example_slides)} examples)"


class LayoutRegistry:
    """
    template.pptxのレイアウト情報を動的に管理

    使い方:
        registry = LayoutRegistry('templates/template.pptx')
        layout_idx = registry.find_layout('Title Slide')
        layout_info = registry.get_layout_info(10)
    """

    def __init__(self, template_source):
        """
        Args:
            template_source: Either a file path (str) or a Presentation instance
        """
        if isinstance(template_source, str):
            self.template_path = template_source
            self.prs = Presentation(template_source)
            self._owns_presentation = True
        else:
            # Assume it's a Presentation instance
            self.prs = template_source
            self.template_path = None
            self._owns_presentation = False
        
        self._layouts: Dict[int, LayoutInfo] = {}
        self._name_to_index: Dict[str, int] = {}
        self._analyze()

    def _analyze(self):
        """template.pptxを解析してレイアウト情報を構築"""
        # レイアウト情報を収集
        for idx, layout in enumerate(self.prs.slide_layouts):
            placeholders = []
            for ph in layout.placeholders:
                placeholders.append((
                    ph.placeholder_format.idx,
                    str(ph.placeholder_format.type)
                ))

            self._layouts[idx] = LayoutInfo(
                index=idx,
                name=layout.name,
                placeholders=placeholders,
                example_slides=[]
            )
            self._name_to_index[layout.name] = idx

        # 各スライドがどのレイアウトを使っているか記録
        for slide_idx, slide in enumerate(self.prs.slides):
            for layout_idx, layout in enumerate(self.prs.slide_layouts):
                if slide.slide_layout == layout:
                    self._layouts[layout_idx].example_slides.append(slide_idx)
                    break

    def get_layout_count(self) -> int:
        """レイアウト総数を取得"""
        return len(self._layouts)

    def get_slide_count(self) -> int:
        """スライド総数を取得"""
        return len(self.prs.slides)

    def get_layout_info(self, index: int) -> Optional[LayoutInfo]:
        """インデックスからレイアウト情報を取得"""
        return self._layouts.get(index)

    def get_layout_by_name(self, name: str):
        """
        レイアウト名からレイアウトオブジェクトを取得（完全一致）

        Args:
            name: レイアウト名（例: "Title Slide", "Title and Content_withKeyMessage"）

        Returns:
            SlideLayout object or None

        例:
            layout = registry.get_layout_by_name('Title Slide')
            if layout is None:
                raise ValueError(f"レイアウト '{name}' が見つかりません")
        """
        return self.prs.slide_layouts.get_by_name(name)

    def find_layout(self, name: str) -> Optional[int]:
        """
        レイアウト名からインデックスを検索（部分一致）

        非推奨: get_layout_by_name() を使用してください
        後方互換性のため残しています

        例:
            find_layout('Title Slide') -> 0
            find_layout('KeyMessage') -> 10 (最初にマッチしたもの)
        """
        # 完全一致を優先
        if name in self._name_to_index:
            return self._name_to_index[name]

        # 部分一致
        for layout_name, idx in self._name_to_index.items():
            if name in layout_name:
                return idx

        return None

    def find_layouts_by_pattern(self, pattern: str) -> List[int]:
        """パターンに一致するすべてのレイアウトを検索"""
        results = []
        for layout_name, idx in self._name_to_index.items():
            if pattern.lower() in layout_name.lower():
                results.append(idx)
        return results

    def get_used_layouts(self) -> List[int]:
        """実際に使用されているレイアウトのリスト"""
        return [idx for idx, info in self._layouts.items() if info.example_slides]

    def get_unused_layouts(self) -> List[int]:
        """未使用のレイアウトのリスト"""
        return [idx for idx, info in self._layouts.items() if not info.example_slides]

    def suggest_layout(self, content_type: str, columns: int = 1,
                      has_keymessage: bool = True) -> Optional[str]:
        """
        コンテンツタイプから適切なレイアウトを推奨

        Args:
            content_type: 'text', 'table', 'chart', 'diagram', 'image'
            columns: カラム数（1, 2, 3）
            has_keymessage: KeyMessageプレースホルダーが必要か

        Returns:
            推奨レイアウトの名前（例: "Title Slide", "Title and Content_withKeyMessage"）
        """
        # KeyMessageありの場合
        if has_keymessage:
            if columns == 1:
                if content_type in ['table', 'chart', 'diagram', 'image']:
                    return '1_Title and Object_withKeyMessage'
                else:
                    return 'Title and Content_withKeyMessage'

            elif columns == 2:
                if content_type in ['table', 'chart', 'diagram', 'image']:
                    return '1_Two Object_withKeyMessage'
                else:
                    return 'Two Content_withKeyMessage'

            elif columns == 3:
                if content_type in ['table', 'chart', 'diagram', 'image']:
                    return '2_Three Object_withKeyMessage'
                else:
                    return '1_Three Content_withKeyMessage'

        # KeyMessageなしの場合
        else:
            if columns == 1:
                return 'Title and Content'
            elif columns == 2:
                return 'Two Content'
            elif columns == 3:
                return 'Comparison'

        # デフォルト
        return 'Title and Content'

    def print_summary(self):
        """レイアウト情報のサマリーを出力"""
        print(f"=== Layout Registry Summary ===")
        print(f"Template: {self.template_path}")
        print(f"Layouts: {self.get_layout_count()}")
        print(f"Slides: {self.get_slide_count()}")
        print(f"Used layouts: {len(self.get_used_layouts())}")
        print(f"Unused layouts: {len(self.get_unused_layouts())}")

        print(f"\n=== All Layouts ===")
        for idx in sorted(self._layouts.keys()):
            info = self._layouts[idx]
            status = "✓" if info.example_slides else "✗"
            examples = f"({len(info.example_slides)} examples)" if info.example_slides else ""
            print(f"[{idx:2d}] {status} {info.name} {examples}")

        if self.get_unused_layouts():
            print(f"\n=== Unused Layouts ===")
            for idx in self.get_unused_layouts():
                info = self._layouts[idx]
                print(f"[{idx:2d}] {info.name}")


def main():
    """テスト用メイン関数"""
    import sys
    import os

    # pptxスキルのルートディレクトリを取得
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(script_dir)
    template_path = os.path.join(skill_root, 'templates', 'template.pptx')

    registry = LayoutRegistry(template_path)
    registry.print_summary()

    print(f"\n=== Layout Suggestions ===")
    print(f"1カラムテキスト: Layout {registry.suggest_layout('text', 1, True)}")
    print(f"1カラムグラフ: Layout {registry.suggest_layout('chart', 1, True)}")
    print(f"2カラム比較: Layout {registry.suggest_layout('text', 2, True)}")
    print(f"3カラム図表: Layout {registry.suggest_layout('diagram', 3, True)}")

    print(f"\n=== Search Examples ===")
    print(f"'Title Slide': {registry.find_layout('Title Slide')}")
    print(f"'KeyMessage': {registry.find_layouts_by_pattern('KeyMessage')}")
    print(f"'Three': {registry.find_layouts_by_pattern('Three')}")


if __name__ == '__main__':
    main()
