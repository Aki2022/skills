#!/usr/bin/env python3
"""Utility functions for applying .crtx chart template styling."""

import zipfile
from lxml import etree
from typing import Dict, List, Any, Optional
from pptx.util import Pt
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

# Import logging
from scripts.logging_utils import get_logger


def lummod_to_brightness(lummod_val: int, lumoff_val: int = 0) -> float:
    """Convert lumMod/lumOff values to brightness (-1.0 to 1.0).

    Args:
        lummod_val: lumMod value from OOXML (e.g., 75000 = 75%)
        lumoff_val: lumOff value from OOXML (e.g., 85000 = 85%)

    Returns:
        brightness: Value for python-pptx brightness property

    For dark base colors (like tx1/black):
        lumMod 15000 + lumOff 85000 → brightness +0.85 (light gray)

    For light base colors (like bg1/white):
        lumMod 75000 → brightness -0.25 (25% darker)
    """
    # lumMod is multiplication, lumOff is offset
    # For tx1 (black): result = 0 * lumMod + lumOff = lumOff
    # For bg1 (white): result = 1 * lumMod = lumMod

    if lumoff_val > 0:
        # Has lumOff - used to lighten dark colors
        # brightness = lumOff / 100000
        return lumoff_val / 100000.0
    else:
        # No lumOff - lumMod only
        # brightness = lumMod - 1.0
        percentage = lummod_val / 100000.0
        return percentage - 1.0


def extract_crtx_styling(crtx_path: str) -> Dict[str, Any]:
    """Extract styling information from .crtx chart template file.

    Args:
        crtx_path: Path to .crtx file

    Returns:
        Dictionary with styling info:
        {
            'series': [
                {
                    'fill_type': 'rgb' or 'theme',
                    'fill_value': '#RRGGBB' or 'bg1',
                    'fill_lummod': int (optional),
                    'line_lummod': int (optional),
                },
                ...
            ],
            'category_axis': {
                'visible': bool,
                'line_width_emu': int,
            },
            'value_axis': {
                'visible': bool,
            }
        }
    """
    with zipfile.ZipFile(crtx_path, 'r') as z:
        chart_xml = z.read('chart/chart.xml')

    root = etree.fromstring(chart_xml)

    # Define namespaces
    ns = {
        'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }

    result = {
        'series': [],
        'category_axis': {},
        'value_axis': {},
        'data_labels': [],
        'legend': {},
    }

    # Extract series styling
    for ser in root.findall('.//c:ser', ns):
        series_style = {}

        spPr = ser.find('.//c:spPr', ns)
        if spPr is not None:
            # Fill styling
            solidFill = spPr.find('.//a:solidFill', ns)
            if solidFill is not None:
                # Check for RGB color
                srgbClr = solidFill.find('.//a:srgbClr', ns)
                if srgbClr is not None:
                    series_style['fill_type'] = 'rgb'
                    series_style['fill_value'] = f"#{srgbClr.get('val')}"

                # Check for theme color
                schemeClr = solidFill.find('.//a:schemeClr', ns)
                if schemeClr is not None:
                    series_style['fill_type'] = 'theme'
                    series_style['fill_value'] = schemeClr.get('val')

                    # Get lumMod (brightness modifier)
                    lumMod = schemeClr.find('.//a:lumMod', ns)
                    if lumMod is not None:
                        series_style['fill_lummod'] = int(lumMod.get('val'))

            # Line styling
            ln = spPr.find('.//a:ln', ns)
            if ln is not None:
                solidFill_ln = ln.find('.//a:solidFill', ns)
                if solidFill_ln is not None:
                    schemeClr = solidFill_ln.find('.//a:schemeClr', ns)
                    if schemeClr is not None:
                        lumMod = schemeClr.find('.//a:lumMod', ns)
                        if lumMod is not None:
                            series_style['line_lummod'] = int(lumMod.get('val'))

        result['series'].append(series_style)

    # Extract category axis styling
    catAx = root.find('.//c:catAx', ns)
    if catAx is not None:
        delete_elem = catAx.find('.//c:delete', ns)
        result['category_axis']['visible'] = (delete_elem is None or delete_elem.get('val') == '0')

        # Tick marks
        majorTickMark = catAx.find('.//c:majorTickMark', ns)
        if majorTickMark is not None:
            result['category_axis']['major_tick_mark'] = majorTickMark.get('val')

        minorTickMark = catAx.find('.//c:minorTickMark', ns)
        if minorTickMark is not None:
            result['category_axis']['minor_tick_mark'] = minorTickMark.get('val')

        # Text properties (font)
        txPr = catAx.find('.//c:txPr', ns)
        if txPr is not None:
            defRPr = txPr.find('.//a:defRPr', ns)
            if defRPr is not None:
                sz = defRPr.get('sz')
                if sz:
                    result['category_axis']['font_size_pt'] = int(sz) / 100

                # Font color
                solidFill = defRPr.find('.//a:solidFill', ns)
                if solidFill is not None:
                    schemeClr = solidFill.find('.//a:schemeClr', ns)
                    if schemeClr is not None:
                        result['category_axis']['font_color_theme'] = schemeClr.get('val')

                        lumMod = schemeClr.find('.//a:lumMod', ns)
                        if lumMod is not None:
                            result['category_axis']['font_lummod'] = int(lumMod.get('val'))

                        lumOff = schemeClr.find('.//a:lumOff', ns)
                        if lumOff is not None:
                            result['category_axis']['font_lumoff'] = int(lumOff.get('val'))

        # Shape properties (line)
        spPr = catAx.find('.//c:spPr', ns)
        if spPr is not None:
            ln = spPr.find('.//a:ln', ns)
            if ln is not None:
                width = ln.get('w')
                if width:
                    result['category_axis']['line_width_emu'] = int(width)

                # Extract line color
                solidFill = ln.find('.//a:solidFill', ns)
                if solidFill is not None:
                    # RGB color
                    srgbClr = solidFill.find('.//a:srgbClr', ns)
                    if srgbClr is not None:
                        result['category_axis']['line_color_type'] = 'rgb'
                        result['category_axis']['line_color_value'] = f"#{srgbClr.get('val')}"

                    # Theme color
                    schemeClr = solidFill.find('.//a:schemeClr', ns)
                    if schemeClr is not None:
                        result['category_axis']['line_color_type'] = 'theme'
                        result['category_axis']['line_color_value'] = schemeClr.get('val')

                        # lumMod and lumOff
                        lumMod = schemeClr.find('.//a:lumMod', ns)
                        if lumMod is not None:
                            result['category_axis']['line_lummod'] = int(lumMod.get('val'))

                        lumOff = schemeClr.find('.//a:lumOff', ns)
                        if lumOff is not None:
                            result['category_axis']['line_lumoff'] = int(lumOff.get('val'))

    # Extract value axis styling
    valAx = root.find('.//c:valAx', ns)
    if valAx is not None:
        delete_elem = valAx.find('.//c:delete', ns)
        result['value_axis']['visible'] = (delete_elem is None or delete_elem.get('val') == '0')

        # Tick marks
        majorTickMark = valAx.find('.//c:majorTickMark', ns)
        if majorTickMark is not None:
            result['value_axis']['major_tick_mark'] = majorTickMark.get('val')

        minorTickMark = valAx.find('.//c:minorTickMark', ns)
        if minorTickMark is not None:
            result['value_axis']['minor_tick_mark'] = minorTickMark.get('val')

    # Extract data label styling
    for dLbls in root.findall('.//c:dLbls', ns):
        dl_style = {}

        # Show options
        showVal = dLbls.find('.//c:showVal', ns)
        if showVal is not None:
            dl_style['show_value'] = showVal.get('val') == '1'

        # Position
        dLblPos = dLbls.find('.//c:dLblPos', ns)
        if dLblPos is not None:
            dl_style['position'] = dLblPos.get('val')

        # Number format
        numFmt = dLbls.find('.//c:numFmt', ns)
        if numFmt is not None:
            dl_style['number_format'] = numFmt.get('formatCode')
            # sourceLinked="0" means custom format
            source_linked = numFmt.get('sourceLinked')
            if source_linked is not None:
                dl_style['number_format_linked'] = source_linked == '1'

        # Text properties
        txPr = dLbls.find('.//c:txPr', ns)
        if txPr is not None:
            defRPr = txPr.find('.//a:defRPr', ns)
            if defRPr is not None:
                sz = defRPr.get('sz')
                if sz:
                    dl_style['font_size_pt'] = int(sz) / 100

                # Font color
                solidFill = defRPr.find('.//a:solidFill', ns)
                if solidFill is not None:
                    # Check for RGB color first
                    srgbClr = solidFill.find('.//a:srgbClr', ns)
                    if srgbClr is not None:
                        dl_style['font_color_rgb'] = srgbClr.get('val')
                    else:
                        # Check for scheme color
                        schemeClr = solidFill.find('.//a:schemeClr', ns)
                        if schemeClr is not None:
                            dl_style['font_color_theme'] = schemeClr.get('val')

                            lumMod = schemeClr.find('.//a:lumMod', ns)
                            if lumMod is not None:
                                dl_style['font_lummod'] = int(lumMod.get('val'))

                            lumOff = schemeClr.find('.//a:lumOff', ns)
                            if lumOff is not None:
                                dl_style['font_lumoff'] = int(lumOff.get('val'))

        result['data_labels'].append(dl_style)

    # Extract legend styling
    legend = root.find('.//c:legend', ns)
    if legend is not None:
        legendPos = legend.find('.//c:legendPos', ns)
        if legendPos is not None:
            result['legend']['position'] = legendPos.get('val')

        # Overlay
        overlay = legend.find('.//c:overlay', ns)
        if overlay is not None:
            result['legend']['overlay'] = overlay.get('val') == '1'

        txPr = legend.find('.//c:txPr', ns)
        if txPr is not None:
            defRPr = txPr.find('.//a:defRPr', ns)
            if defRPr is not None:
                sz = defRPr.get('sz')
                if sz:
                    result['legend']['font_size_pt'] = int(sz) / 100

                solidFill = defRPr.find('.//a:solidFill', ns)
                if solidFill is not None:
                    schemeClr = solidFill.find('.//a:schemeClr', ns)
                    if schemeClr is not None:
                        result['legend']['font_color_theme'] = schemeClr.get('val')

                        lumMod = schemeClr.find('.//a:lumMod', ns)
                        if lumMod is not None:
                            result['legend']['font_lummod'] = int(lumMod.get('val'))

                        lumOff = schemeClr.find('.//a:lumOff', ns)
                        if lumOff is not None:
                            result['legend']['font_lumoff'] = int(lumOff.get('val'))

    return result


def _get_theme_color_map() -> Dict[str, MSO_THEME_COLOR]:
    """Get comprehensive theme color mapping."""
    return {
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
        'hlink': MSO_THEME_COLOR.HYPERLINK,
        'folHlink': MSO_THEME_COLOR.FOLLOWED_HYPERLINK,
        'dk1': MSO_THEME_COLOR.DARK_1,
        'lt1': MSO_THEME_COLOR.LIGHT_1,
        'dk2': MSO_THEME_COLOR.DARK_2,
        'lt2': MSO_THEME_COLOR.LIGHT_2,
    }


def apply_crtx_styling_to_chart(chart, crtx_styling: Dict[str, Any], limited_mode=False):
    """Apply .crtx styling to an existing python-pptx chart.

    Args:
        chart: python-pptx Chart object
        crtx_styling: Styling dict from extract_crtx_styling()
        limited_mode: If True, only apply series colors (for compatibility)
    """
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    logger = get_logger()
    chart_type = chart.chart_type
    theme_map = _get_theme_color_map()

    # Apply series styling
    for idx, series in enumerate(chart.series):
        if idx >= len(crtx_styling['series']):
            break

        style = crtx_styling['series'][idx]

        # Apply fill for column/bar/area charts
        if chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED,
                          XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.BAR_STACKED,
                          XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
            series.format.fill.solid()

            if style.get('fill_type') == 'rgb':
                # RGB color
                rgb_hex = style['fill_value'].lstrip('#')
                rgb_color = RGBColor.from_string(rgb_hex)
                series.format.fill.fore_color.rgb = rgb_color

            elif style.get('fill_type') == 'theme':
                # Theme color
                theme_val = style['fill_value']
                if theme_val in theme_map:
                    series.format.fill.fore_color.theme_color = theme_map[theme_val]
                else:
                    logger.warning(f"Unknown theme color '{theme_val}' in series {idx} fill")

                # Apply lumMod as brightness
                if 'fill_lummod' in style:
                    brightness = lummod_to_brightness(style['fill_lummod'])
                    series.format.fill.fore_color.brightness = brightness

            # Apply line styling for borders
            if 'line_lummod' in style:
                try:
                    series.format.line.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
                    brightness = lummod_to_brightness(style['line_lummod'])
                    series.format.line.color.brightness = brightness
                except Exception as e:
                    get_logger().warning(f"Failed to apply line styling to series {idx}: {e}")

        # Apply line for line charts
        elif chart_type == XL_CHART_TYPE.LINE:
            if style.get('fill_type') == 'rgb':
                # RGB color for line
                rgb_hex = style['fill_value'].lstrip('#')
                rgb_color = RGBColor.from_string(rgb_hex)
                series.format.line.color.rgb = rgb_color

            elif style.get('fill_type') == 'theme':
                # Theme color for line
                theme_val = style['fill_value']
                if theme_val in theme_map:
                    series.format.line.color.theme_color = theme_map[theme_val]
                else:
                    logger.warning(f"Unknown theme color '{theme_val}' in series {idx} line")

                # Apply lumMod as brightness
                if 'fill_lummod' in style:
                    brightness = lummod_to_brightness(style['fill_lummod'])
                    series.format.line.color.brightness = brightness

        # Apply to pie charts - use series colors for points
        elif chart_type == XL_CHART_TYPE.PIE:
            # For pie chart, apply series styling to individual points
            # Point 0 -> Series 0 colors, Point 1 -> Series 1 colors, etc.
            # Points beyond series count get BACKGROUND_1 with continuing brightness
            for point_idx, point in enumerate(series.points):
                point.format.fill.solid()

                if point_idx < len(crtx_styling['series']):
                    # Use series styling
                    point_style = crtx_styling['series'][point_idx]

                    if point_style.get('fill_type') == 'rgb':
                        # RGB color
                        rgb_hex = point_style['fill_value'].lstrip('#')
                        rgb_color = RGBColor.from_string(rgb_hex)
                        point.format.fill.fore_color.rgb = rgb_color

                    elif point_style.get('fill_type') == 'theme':
                        # Theme color
                        theme_val = point_style['fill_value']
                        if theme_val in theme_map:
                            point.format.fill.fore_color.theme_color = theme_map[theme_val]
                        else:
                            logger.warning(f"Unknown theme color '{theme_val}' in pie point {point_idx}")

                        # Apply lumMod as brightness
                        if 'fill_lummod' in point_style:
                            brightness = lummod_to_brightness(point_style['fill_lummod'])
                            point.format.fill.fore_color.brightness = brightness
                else:
                    # Points beyond series count: continue BACKGROUND_1 with brightness gradient
                    # Point 3: -0.10, Point 4: -0.05, Point 5: 0.0, etc.
                    point.format.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
                    extra_idx = point_idx - len(crtx_styling['series'])
                    brightness = -0.10 + (extra_idx * 0.05)
                    brightness = min(0.0, brightness)  # Cap at 0.0
                    point.format.fill.fore_color.brightness = brightness

    # Apply category axis styling (skip in limited mode for compatibility)
    if not limited_mode and chart_type != XL_CHART_TYPE.PIE:
        cat_axis = crtx_styling.get('category_axis', {})
        if cat_axis:
            try:
                # Tick marks
                from pptx.enum.chart import XL_TICK_MARK
                tick_map = {
                    'none': XL_TICK_MARK.NONE,
                    'inside': XL_TICK_MARK.INSIDE,
                    'outside': XL_TICK_MARK.OUTSIDE,
                    'cross': XL_TICK_MARK.CROSS,
                }
                if 'major_tick_mark' in cat_axis:
                    tick_val = cat_axis['major_tick_mark']
                    if tick_val in tick_map:
                        chart.category_axis.major_tick_mark = tick_map[tick_val]

                if 'minor_tick_mark' in cat_axis:
                    tick_val = cat_axis['minor_tick_mark']
                    if tick_val in tick_map:
                        chart.category_axis.minor_tick_mark = tick_map[tick_val]

                # Line width
                if 'line_width_emu' in cat_axis:
                    chart.category_axis.format.line.width = cat_axis['line_width_emu']

                # Line color
                if 'line_color_type' in cat_axis:
                    if cat_axis['line_color_type'] == 'rgb':
                        rgb_hex = cat_axis['line_color_value'].lstrip('#')
                        rgb_color = RGBColor.from_string(rgb_hex)
                        chart.category_axis.format.line.color.rgb = rgb_color
                    elif cat_axis['line_color_type'] == 'theme':
                        theme_val = cat_axis['line_color_value']
                        if theme_val in theme_map:
                            chart.category_axis.format.line.color.theme_color = theme_map[theme_val]

                            # Apply lumMod/lumOff as brightness
                            if 'line_lummod' in cat_axis:
                                lummod = cat_axis['line_lummod']
                                lumoff = cat_axis.get('line_lumoff', 0)
                                brightness = lummod_to_brightness(lummod, lumoff)
                                chart.category_axis.format.line.color.brightness = brightness
                        else:
                            logger.warning(f"Unknown theme color '{theme_val}' in category axis line")

                # Font properties
                if 'font_size_pt' in cat_axis:
                    from pptx.util import Pt
                    chart.category_axis.tick_labels.font.size = Pt(cat_axis['font_size_pt'])

                if 'font_color_theme' in cat_axis:
                    theme_val = cat_axis['font_color_theme']
                    if theme_val in theme_map:
                        chart.category_axis.tick_labels.font.color.theme_color = theme_map[theme_val]

                        if 'font_lummod' in cat_axis:
                            lummod = cat_axis['font_lummod']
                            lumoff = cat_axis.get('font_lumoff', 0)
                            brightness = lummod_to_brightness(lummod, lumoff)
                            chart.category_axis.tick_labels.font.color.brightness = brightness
                    else:
                        logger.warning(f"Unknown theme color '{theme_val}' in category axis font")

            except Exception as e:
                get_logger().warning(f"Failed to apply category axis styling: {e}")

        # Apply value axis styling
        val_axis = crtx_styling.get('value_axis', {})
        if val_axis:
            try:
                if 'visible' in val_axis:
                    chart.value_axis.visible = val_axis['visible']

                # Tick marks
                from pptx.enum.chart import XL_TICK_MARK
                tick_map = {
                    'none': XL_TICK_MARK.NONE,
                    'inside': XL_TICK_MARK.INSIDE,
                    'outside': XL_TICK_MARK.OUTSIDE,
                    'cross': XL_TICK_MARK.CROSS,
                }
                if 'major_tick_mark' in val_axis:
                    tick_val = val_axis['major_tick_mark']
                    if tick_val in tick_map:
                        chart.value_axis.major_tick_mark = tick_map[tick_val]

                if 'minor_tick_mark' in val_axis:
                    tick_val = val_axis['minor_tick_mark']
                    if tick_val in tick_map:
                        chart.value_axis.minor_tick_mark = tick_map[tick_val]

            except Exception as e:
                get_logger().warning(f"Failed to apply value axis styling: {e}")

        # Disable gridlines (template has no gridlines)
        # NOTE: Skip for area charts - has_major_gridlines causes XML corruption
        if chart_type not in [XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
            try:
                chart.value_axis.has_major_gridlines = False
                chart.value_axis.has_minor_gridlines = False
            except Exception as e:
                get_logger().warning(f"Failed to disable value axis gridlines: {e}")

            try:
                chart.category_axis.has_major_gridlines = False
                chart.category_axis.has_minor_gridlines = False
            except Exception as e:
                get_logger().warning(f"Failed to disable category axis gridlines: {e}")
        else:
            get_logger().info("Skipping gridline settings for area chart (python-pptx compatibility)")

    # Apply data label styling (skip in limited mode for compatibility)
    if limited_mode:
        get_logger().info("Limited mode: skipping data label styling")
        return

    data_labels = crtx_styling.get('data_labels', [])
    if data_labels:
        # Use first data label style as default for all series
        default_dl = data_labels[0] if data_labels else {}

        for series_idx, series in enumerate(chart.series):
            # Get series-specific or default style
            dl_style = data_labels[series_idx] if series_idx < len(data_labels) else default_dl

            if dl_style.get('show_value', False):
                series.has_data_labels = True
                dl = series.data_labels
                dl.show_value = True

                # Position - default based on chart type if not in template
                # NOTE: Skip position for area charts - causes XML corruption in python-pptx
                position_val = dl_style.get('position')
                if not position_val:
                    # Set default based on chart type
                    from pptx.enum.chart import XL_CHART_TYPE
                    if chart_type == XL_CHART_TYPE.LINE:
                        position_val = 't'  # top for line charts
                    elif chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED]:
                        position_val = 'outEnd'  # outside end for bar/column
                    elif chart_type in [XL_CHART_TYPE.COLUMN_STACKED, XL_CHART_TYPE.BAR_STACKED]:
                        position_val = 'ctr'  # center for stacked
                    elif chart_type == XL_CHART_TYPE.PIE:
                        position_val = 'outEnd'  # outside for pie
                    # Skip area charts - position setting causes XML corruption
                    # elif chart_type in [XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
                    #     position_val = 't'

                # Skip position setting for area charts (python-pptx compatibility)
                if position_val and chart_type not in [XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
                    from pptx.enum.chart import XL_DATA_LABEL_POSITION
                    pos_map = {
                        't': XL_DATA_LABEL_POSITION.ABOVE,
                        'b': XL_DATA_LABEL_POSITION.BELOW,
                        'l': XL_DATA_LABEL_POSITION.LEFT,
                        'r': XL_DATA_LABEL_POSITION.RIGHT,
                        'ctr': XL_DATA_LABEL_POSITION.CENTER,
                        'inBase': XL_DATA_LABEL_POSITION.INSIDE_BASE,
                        'inEnd': XL_DATA_LABEL_POSITION.INSIDE_END,
                        'outEnd': XL_DATA_LABEL_POSITION.OUTSIDE_END,
                    }
                    if position_val in pos_map:
                        try:
                            dl.position = pos_map[position_val]
                        except Exception as e:
                            logger.warning(f"Failed to set data label position for series {series_idx}: {e}")

                # Number format
                if 'number_format' in dl_style:
                    try:
                        dl.number_format = dl_style['number_format']
                    except Exception as e:
                        logger.warning(f"Failed to set data label number format for series {series_idx}: {e}")

                # Font size
                if 'font_size_pt' in dl_style:
                    from pptx.util import Pt
                    dl.font.size = Pt(dl_style['font_size_pt'])

                # Font color - RGB or theme
                if 'font_color_rgb' in dl_style:
                    # Apply RGB color
                    try:
                        from pptx.dml.color import RGBColor
                        rgb_val = dl_style['font_color_rgb']
                        r = int(rgb_val[0:2], 16)
                        g = int(rgb_val[2:4], 16)
                        b = int(rgb_val[4:6], 16)
                        dl.font.color.rgb = RGBColor(r, g, b)
                    except Exception as e:
                        logger.warning(f"Failed to apply RGB color to data label for series {series_idx}: {e}")
                elif 'font_color_theme' in dl_style:
                    theme_val = dl_style['font_color_theme']
                    if theme_val in theme_map:
                        try:
                            dl.font.color.theme_color = theme_map[theme_val]

                            if 'font_lummod' in dl_style:
                                lummod = dl_style['font_lummod']
                                lumoff = dl_style.get('font_lumoff', 0)
                                brightness = lummod_to_brightness(lummod, lumoff)
                                dl.font.color.brightness = brightness
                        except Exception as e:
                            logger.warning(f"Failed to apply data label font color for series {series_idx}: {e}")
                    else:
                        logger.warning(f"Unknown theme color '{theme_val}' in data label font for series {series_idx}")

    # Apply legend styling
    legend_style = crtx_styling.get('legend', {})
    if legend_style:
        chart.has_legend = True

        # Position
        if 'position' in legend_style:
            from pptx.enum.chart import XL_LEGEND_POSITION
            pos_map = {
                'b': XL_LEGEND_POSITION.BOTTOM,
                't': XL_LEGEND_POSITION.TOP,
                'r': XL_LEGEND_POSITION.RIGHT,
                'l': XL_LEGEND_POSITION.LEFT,
            }
            pos_val = legend_style['position']
            if pos_val in pos_map:
                chart.legend.position = pos_map[pos_val]

        # Overlay
        if 'overlay' in legend_style:
            chart.legend.include_in_layout = legend_style['overlay']

        # Font size
        if 'font_size_pt' in legend_style:
            from pptx.util import Pt
            chart.legend.font.size = Pt(legend_style['font_size_pt'])

        # Font color
        if 'font_color_theme' in legend_style:
            theme_val = legend_style['font_color_theme']
            if theme_val in theme_map:
                try:
                    chart.legend.font.color.theme_color = theme_map[theme_val]

                    if 'font_lummod' in legend_style:
                        lummod = legend_style['font_lummod']
                        lumoff = legend_style.get('font_lumoff', 0)
                        brightness = lummod_to_brightness(lummod, lumoff)
                        chart.legend.font.color.brightness = brightness
                except Exception as e:
                    logger.warning(f"Failed to apply legend font color: {e}")
            else:
                logger.warning(f"Unknown theme color '{theme_val}' in legend font")


if __name__ == '__main__':
    # Test extraction
    crtx_path = 'templates/template.crtx'
    styling = extract_crtx_styling(crtx_path)

    print("Extracted styling:")
    import json
    print(json.dumps(styling, indent=2))

    # Convert lumMod to brightness
    print("\nBrightness conversions:")
    for idx, ser in enumerate(styling['series']):
        if 'fill_lummod' in ser:
            brightness = lummod_to_brightness(ser['fill_lummod'])
            print(f"  Series {idx} fill: lumMod {ser['fill_lummod']} → brightness {brightness}")
        if 'line_lummod' in ser:
            brightness = lummod_to_brightness(ser['line_lummod'])
            print(f"  Series {idx} line: lumMod {ser['line_lummod']} → brightness {brightness}")
