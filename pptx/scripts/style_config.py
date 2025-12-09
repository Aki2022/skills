#!/usr/bin/env python3
"""Python helper for loading and using style.yaml.

This module provides a convenient interface for accessing style.yaml
values in Python, with automatic conversion to python-pptx types.

Usage:
    from style_config import StyleConfig

    config = StyleConfig.load()

    # Get primary color as RGBColor
    primary = config.colors.primary  # '#4F4F70'

    # Get series colors
    rgb = config.get_series_rgb(0)  # RGBColor
    theme_info = config.get_series_theme(1)  # {'theme_color': MSO_THEME_COLOR.BACKGROUND_1, 'brightness': -0.25}

    # Get axis properties
    width = config.category_axis.line_width  # Pt(0.75)
    size = config.category_axis.font_size  # Pt(11)

    # Get legend position
    position = config.legend.position  # XL_LEGEND_POSITION.BOTTOM
"""

import os
import yaml
from typing import Dict, Any, Optional

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.util import Pt


class AttrDict(dict):
    """Dictionary that allows attribute access."""

    def __getattr__(self, key):
        try:
            value = self[key]
            if isinstance(value, dict):
                return AttrDict(value)
            return value
        except KeyError:
            raise AttributeError(f"'{type(self).__name__}' object has no attribute '{key}'")


class StyleConfig:
    """Configuration class for style.yaml."""

    _instance = None
    _style_data = None

    def __init__(self, style_data: Dict[str, Any]):
        self._style_data = style_data

    @classmethod
    def load(cls, yaml_path: Optional[str] = None) -> 'StyleConfig':
        """Load style.yaml and return StyleConfig instance.

        Search order:
        1. Provided yaml_path (if specified)
        2. ./powerpoint/processing/style.yaml (project-specific, recommended)
        3. ./processing/style.yaml (legacy, for backward compatibility)
        4. ~/.claude/skills/pptx/templates/style.yaml (master template)

        Args:
            yaml_path: Optional path to style.yaml. If None, auto-detects.

        Returns:
            StyleConfig instance
        """
        if yaml_path is None:
            # Search order: powerpoint/processing -> processing (legacy) -> skill master
            candidates = [
                os.path.join(os.getcwd(), 'powerpoint', 'processing', 'style.yaml'),
                os.path.join(os.getcwd(), 'processing', 'style.yaml'),
                os.path.join(os.path.expanduser('~/.claude/skills/pptx/templates'), 'style.yaml'),
            ]

            for candidate in candidates:
                if os.path.exists(candidate):
                    yaml_path = candidate
                    break

            if yaml_path is None:
                raise FileNotFoundError(
                    "style.yaml not found. Please run:\n"
                    "  mkdir -p powerpoint/processing\n"
                    "  cp ~/.claude/skills/pptx/templates/style.yaml powerpoint/processing/\n"
                    "Or generate it with:\n"
                    "  cd ~/.claude/skills/pptx && python scripts/extract_style.py\n"
                    "  cp templates/style.yaml /path/to/your/project/powerpoint/processing/"
                )

        with open(yaml_path, 'r') as f:
            style_data = yaml.safe_load(f)

        return cls(style_data)

    @property
    def colors(self) -> AttrDict:
        """Get colors section."""
        return AttrDict(self._style_data.get('colors', {}))

    @property
    def category_axis(self) -> 'AxisConfig':
        """Get category axis configuration."""
        return AxisConfig(self._style_data.get('category_axis', {}))

    @property
    def value_axis(self) -> 'AxisConfig':
        """Get value axis configuration."""
        return AxisConfig(self._style_data.get('value_axis', {}))

    @property
    def legend(self) -> 'LegendConfig':
        """Get legend configuration."""
        return LegendConfig(self._style_data.get('legend', {}))

    @property
    def gridlines(self) -> AttrDict:
        """Get gridlines configuration."""
        return AttrDict(self._style_data.get('gridlines', {}))

    @property
    def table(self) -> 'TableConfig':
        """Get table configuration."""
        return TableConfig(self._style_data.get('table', {}))

    @property
    def diagram(self) -> AttrDict:
        """Get diagram configuration."""
        return AttrDict(self._style_data.get('diagram', {}))

    @property
    def mermaid(self) -> AttrDict:
        """Get mermaid configuration."""
        return AttrDict(self._style_data.get('mermaid', {}))

    @property
    def shape(self) -> AttrDict:
        """Get shape configuration."""
        return AttrDict(self._style_data.get('shape', {}))

    @property
    def flowchart(self) -> AttrDict:
        """Get flowchart configuration."""
        return AttrDict(self._style_data.get('flowchart', {}))

    def get_series_rgb(self, index: int) -> RGBColor:
        """Get series color as RGBColor.

        Args:
            index: Series index

        Returns:
            RGBColor for the series
        """
        series_list = self._style_data.get('colors', {}).get('series', [])

        if index < len(series_list):
            series = series_list[index]
            if series.get('type') == 'rgb':
                return self.hex_to_rgb(series.get('value', '#000000'))

        # Default: primary color
        return self.hex_to_rgb(self._style_data.get('colors', {}).get('primary', '#4F4F70'))

    def get_series_theme(self, index: int) -> Dict[str, Any]:
        """Get series theme color info.

        Args:
            index: Series index

        Returns:
            Dict with 'theme_color' (MSO_THEME_COLOR) and 'brightness' (float)
        """
        series_list = self._style_data.get('colors', {}).get('series', [])

        if index < len(series_list):
            series = series_list[index]
            if series.get('type') == 'theme':
                theme_color = self.get_theme_color(series.get('value', 'bg1'))
                brightness = series.get('brightness', 0)
                return {
                    'theme_color': theme_color,
                    'brightness': brightness
                }

        # Default: BACKGROUND_1 with no brightness
        return {
            'theme_color': MSO_THEME_COLOR.BACKGROUND_1,
            'brightness': 0
        }

    def get_data_label_style(self, index: int) -> Dict[str, Any]:
        """Get data label style for series.

        Args:
            index: Series/data label index

        Returns:
            Dict with data label style properties
        """
        data_labels = self._style_data.get('data_labels', [])

        if index < len(data_labels):
            return data_labels[index]

        # Default: first style or empty
        if data_labels:
            return data_labels[0]
        return {'show_value': False}

    @staticmethod
    def get_theme_color(theme_name: str) -> MSO_THEME_COLOR:
        """Convert theme color name to MSO_THEME_COLOR.

        Args:
            theme_name: Theme name from OOXML (tx1, bg1, accent1, etc.)

        Returns:
            MSO_THEME_COLOR enum value
        """
        theme_map = {
            'tx1': MSO_THEME_COLOR.TEXT_1,
            'tx2': MSO_THEME_COLOR.TEXT_2,
            'bg1': MSO_THEME_COLOR.BACKGROUND_1,
            'bg2': MSO_THEME_COLOR.BACKGROUND_2,
            'accent1': MSO_THEME_COLOR.ACCENT_1,
            'accent2': MSO_THEME_COLOR.ACCENT_2,
            'accent3': MSO_THEME_COLOR.ACCENT_3,
            'accent4': MSO_THEME_COLOR.ACCENT_4,
            'accent5': MSO_THEME_COLOR.ACCENT_5,
            'accent6': MSO_THEME_COLOR.ACCENT_6,
            'dk1': MSO_THEME_COLOR.DARK_1,
            'dk2': MSO_THEME_COLOR.DARK_2,
            'lt1': MSO_THEME_COLOR.LIGHT_1,
            'lt2': MSO_THEME_COLOR.LIGHT_2,
        }
        return theme_map.get(theme_name, MSO_THEME_COLOR.BACKGROUND_1)

    @staticmethod
    def hex_to_rgb(hex_color: str) -> RGBColor:
        """Convert hex color string to RGBColor.

        Args:
            hex_color: Hex color string (e.g., '#4F4F70' or '4F4F70')

        Returns:
            RGBColor object
        """
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return RGBColor(r, g, b)


class AxisConfig:
    """Configuration for chart axis."""

    def __init__(self, axis_data: Dict[str, Any]):
        self._data = axis_data

    @property
    def visible(self) -> bool:
        """Whether axis is visible."""
        return self._data.get('visible', True)

    @property
    def tick_marks(self) -> str:
        """Tick marks setting."""
        return self._data.get('tick_marks', 'none')

    @property
    def line_width(self) -> Pt:
        """Line width as Pt."""
        line = self._data.get('line', {})
        width_pt = line.get('width_pt', 0.75)
        return Pt(width_pt)

    @property
    def line_color_type(self) -> str:
        """Line color type (rgb or theme)."""
        line = self._data.get('line', {})
        return line.get('color_type', 'theme')

    @property
    def line_color_value(self) -> str:
        """Line color value."""
        line = self._data.get('line', {})
        return line.get('color_value', 'tx1')

    @property
    def line_brightness(self) -> float:
        """Line brightness."""
        line = self._data.get('line', {})
        return line.get('brightness', 0)

    @property
    def font_size(self) -> Pt:
        """Font size as Pt."""
        font = self._data.get('font', {})
        size_pt = font.get('size_pt', 11)
        return Pt(size_pt)

    @property
    def font_color_type(self) -> str:
        """Font color type."""
        font = self._data.get('font', {})
        return font.get('color_type', 'theme')

    @property
    def font_color_value(self) -> str:
        """Font color value."""
        font = self._data.get('font', {})
        return font.get('color_value', 'tx1')

    @property
    def font_brightness(self) -> float:
        """Font brightness."""
        font = self._data.get('font', {})
        return font.get('brightness', 0)


class TableHeaderConfig:
    """Configuration for table header."""
    def __init__(self, header_data: Dict[str, Any]):
        self._data = header_data
    
    @property
    def vertical_align(self) -> str:
        return self._data.get('vertical_align', 'middle')
    
    @property
    def horizontal_align(self) -> str:
        return self._data.get('horizontal_align', 'right')
    
    @property
    def font_family(self) -> str:
        return self._data.get('font_family', 'Arial')
    
    @property
    def font_size_pt(self) -> float:
        return self._data.get('font_size_pt', 12)
    
    @property
    def font_bold(self) -> bool:
        return self._data.get('font_bold', False)
    
    @property
    def font_italic(self) -> bool:
        return self._data.get('font_italic', False)
    
    @property
    def font_underline(self) -> bool:
        return self._data.get('font_underline', False)
    
    @property
    def margin_left_emu(self) -> int:
        return self._data.get('margin_left_emu', 91440)
    
    @property
    def margin_right_emu(self) -> int:
        return self._data.get('margin_right_emu', 91440)
    
    @property
    def margin_top_emu(self) -> int:
        return self._data.get('margin_top_emu', 45720)
    
    @property
    def margin_bottom_emu(self) -> int:
        return self._data.get('margin_bottom_emu', 45720)


class TableBodyConfig:
    """Configuration for table body."""
    def __init__(self, body_data: Dict[str, Any]):
        self._data = body_data
    
    @property
    def vertical_align(self) -> str:
        return self._data.get('vertical_align', 'middle')
    
    @property
    def horizontal_align(self) -> str:
        return self._data.get('horizontal_align', 'right')
    
    @property
    def font_family(self) -> str:
        return self._data.get('font_family', 'Arial')
    
    @property
    def font_size_pt(self) -> float:
        return self._data.get('font_size_pt', 12)
    
    @property
    def font_bold(self) -> bool:
        return self._data.get('font_bold', False)
    
    @property
    def font_italic(self) -> bool:
        return self._data.get('font_italic', False)
    
    @property
    def font_underline(self) -> bool:
        return self._data.get('font_underline', False)
    
    @property
    def margin_left_emu(self) -> int:
        return self._data.get('margin_left_emu', 91440)
    
    @property
    def margin_right_emu(self) -> int:
        return self._data.get('margin_right_emu', 91440)
    
    @property
    def margin_top_emu(self) -> int:
        return self._data.get('margin_top_emu', 45720)
    
    @property
    def margin_bottom_emu(self) -> int:
        return self._data.get('margin_bottom_emu', 45720)


class TableConfig:
    """Configuration for table styling."""

    def __init__(self, table_data: Dict[str, Any]):
        self._data = table_data

    @property
    def header(self) -> TableHeaderConfig:
        """Header configuration."""
        return TableHeaderConfig(self._data.get('header', {}))
    
    @property
    def body(self) -> TableBodyConfig:
        """Body configuration."""
        return TableBodyConfig(self._data.get('body', {}))

    @property
    def border_color(self) -> RGBColor:
        """Border color as RGBColor."""
        border = self._data.get('border', {})
        hex_color = border.get('color', '#4F4F70')
        return StyleConfig.hex_to_rgb(hex_color)

    @property
    def border_width_outer(self) -> Pt:
        """Outer border width."""
        border = self._data.get('border', {})
        return Pt(border.get('width_outer_pt', 1.5))

    @property
    def border_width_inner(self) -> Pt:
        """Inner border width."""
        border = self._data.get('border', {})
        return Pt(border.get('width_inner_pt', 1.0))

    @property
    def header_fill_theme(self) -> MSO_THEME_COLOR:
        """Header fill theme color."""
        header = self._data.get('header', {})
        return StyleConfig.get_theme_color(header.get('fill_theme', 'bg1'))

    @property
    def header_fill_brightness(self) -> float:
        """Header fill brightness."""
        header = self._data.get('header', {})
        return header.get('fill_brightness', -0.5)

    @property
    def header_text_theme(self) -> MSO_THEME_COLOR:
        """Header text theme color."""
        header = self._data.get('header', {})
        return StyleConfig.get_theme_color(header.get('text_color_theme', 'lt1'))

    @property
    def header_text_brightness(self) -> float:
        """Header text brightness."""
        header = self._data.get('header', {})
        return header.get('text_color_brightness', 0)

    @property
    def header_font_bold(self) -> bool:
        """Header font bold."""
        header = self._data.get('header', {})
        return header.get('font_bold', True)

    @property
    def header_font_size(self) -> Pt:
        """Header font size."""
        header = self._data.get('header', {})
        return Pt(header.get('font_size_pt', 12))

    @property
    def body_fill_theme(self) -> MSO_THEME_COLOR:
        """Body fill theme color."""
        body = self._data.get('body', {})
        return StyleConfig.get_theme_color(body.get('fill_theme', 'bg1'))

    def get_body_brightness(self, col_idx: int) -> float:
        """Get body cell brightness for column."""
        body = self._data.get('body', {})
        brightnesses = body.get('column_brightness', [-0.15, -0.05, -0.05, -0.05])
        if col_idx < len(brightnesses):
            return brightnesses[col_idx]
        return brightnesses[-1] if brightnesses else -0.05

    @property
    def body_text_theme(self) -> MSO_THEME_COLOR:
        """Body text theme color."""
        body = self._data.get('body', {})
        return StyleConfig.get_theme_color(body.get('text_color_theme', 'dk1'))

    @property
    def body_text_brightness(self) -> float:
        """Body text brightness."""
        body = self._data.get('body', {})
        return body.get('text_color_brightness', 0)

    @property
    def body_font_size(self) -> Pt:
        """Body font size."""
        body = self._data.get('body', {})
        return Pt(body.get('font_size_pt', 12))

    @property
    def alignment(self) -> str:
        """Default alignment (deprecated, use header/body.horizontal_align)."""
        return self._data.get('alignment', 'right')


class LegendConfig:
    """Configuration for chart legend."""

    def __init__(self, legend_data: Dict[str, Any]):
        self._data = legend_data

    @property
    def position(self) -> XL_LEGEND_POSITION:
        """Legend position as XL_LEGEND_POSITION."""
        pos = self._data.get('position', 'bottom')
        pos_map = {
            'bottom': XL_LEGEND_POSITION.BOTTOM,
            'top': XL_LEGEND_POSITION.TOP,
            'left': XL_LEGEND_POSITION.LEFT,
            'right': XL_LEGEND_POSITION.RIGHT,
        }
        return pos_map.get(pos, XL_LEGEND_POSITION.BOTTOM)

    @property
    def font_size(self) -> Pt:
        """Font size as Pt."""
        font = self._data.get('font', {})
        size_pt = font.get('size_pt', 11)
        return Pt(size_pt)

    @property
    def font_color_type(self) -> str:
        """Font color type."""
        font = self._data.get('font', {})
        return font.get('color_type', 'theme')

    @property
    def font_color_value(self) -> str:
        """Font color value."""
        font = self._data.get('font', {})
        return font.get('color_value', 'tx1')

    @property
    def font_brightness(self) -> float:
        """Font brightness."""
        font = self._data.get('font', {})
        return font.get('brightness', 0)


if __name__ == '__main__':
    # Demo usage
    config = StyleConfig.load()

    print("Style Configuration Demo")
    print("=" * 40)
    print(f"Primary color: {config.colors.primary}")
    print(f"Series 0 RGB: {config.get_series_rgb(0)}")
    print(f"Series 1 theme: {config.get_series_theme(1)}")
    print(f"Category axis line width: {config.category_axis.line_width}")
    print(f"Category axis font size: {config.category_axis.font_size}")
    print(f"Legend position: {config.legend.position}")
    print(f"Gridlines enabled: {config.gridlines.enabled}")
