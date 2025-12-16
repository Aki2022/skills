#!/usr/bin/env python3
"""NATIVE mode object creators: Table, Chart, Diagram.

These functions create PowerPoint-native objects that can be edited after generation.
All styling follows style.md specifications using style_spec.py constants.

Usage:
    from native_objects import create_styled_table, create_styled_chart

    # In process_object() or replacement workflow
    table_shape = create_styled_table(slide, placeholder, spec)
    chart_shape = create_styled_chart(slide, placeholder, spec)
"""

from typing import Any, Dict, List
import os
from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

# Use style_config for all styling (Single Source of Truth)
from scripts.style_config import StyleConfig

# Import logging
from scripts.logging_utils import get_logger

# Import crtx utilities if available
try:
    from scripts.crtx_utils import extract_crtx_styling, apply_crtx_styling_to_chart
    CRTX_AVAILABLE = True
except ImportError:
    CRTX_AVAILABLE = False
    get_logger().error("crtx_utils not available - chart styling will fail")

# Import snapshot utilities
from scripts.snapshot_utils import create_generation_snapshot

# Path to Chart.crtx template (absolute path from skill directory)
SKILL_DIR = Path(__file__).parent.parent
CHART_CRTX_PATH = SKILL_DIR / 'templates' / 'template.crtx'

# Track whether snapshot has been created in this session
_SNAPSHOT_CREATED = False


def create_styled_table(slide, placeholder_shape, spec: Dict[str, Any]):
    """Create native table with style.yaml styling.

    Args:
        slide: Slide object where table will be inserted
        placeholder_shape: Shape to replace (uses its position/size)
        spec: Table specification dict with keys:
            - data: List[List[str]] - 2D array of table data
            - header_row: bool - whether first row is header (default: True)
            - column_types: List[str] - optional, 'text' or 'number' for each column

    Returns:
        Table shape object

    Styling applied (from style.yaml table section):
        - Border: PRIMARY color (outer thick 1.5pt, inner 1pt)
        - Header: bg1 + brightness -0.5, LIGHT_1 text, bold, 12pt
        - Body: bg1 + column brightness, DARK_1 text, 12pt
        - Alignment: right (all cells)
    """
    logger = get_logger()
    logger.info("Creating styled table")

    # Create generation snapshot (once per session)
    _ensure_snapshot_created()

    # Load style configuration
    style = StyleConfig.load()
    table_style = style.table

    data = spec.get('data', [])
    header_row = spec.get('header_row', True)
    column_types = spec.get('column_types', [])

    # Validate data
    if not data or not data[0]:
        logger.error("Table spec.data is empty or invalid")
        raise ValueError("Table spec.data must be non-empty 2D array")

    # Validate column count consistency
    cols = len(data[0])
    for row_idx, row_data in enumerate(data):
        if len(row_data) != cols:
            logger.error(f"Row {row_idx} has {len(row_data)} columns, expected {cols}")
            raise ValueError(f"All rows must have same number of columns (expected {cols}, row {row_idx} has {len(row_data)})")

    logger.debug(f"Table validated: {len(data)} rows x {cols} columns")

    rows = len(data)
    cols = len(data[0])

    # Auto-detect column types if not provided
    if not column_types:
        column_types = []
        # Use first data row (row 1 if header exists, row 0 otherwise)
        data_row_idx = 1 if header_row and len(data) > 1 else 0
        for c_idx in range(cols):
            cell_value = str(data[data_row_idx][c_idx]).strip()
            # Simple number detection: if parseable as float and contains digits
            is_number = False
            try:
                # Remove common number formatting (commas, yen sign, units, etc.)
                cleaned = cell_value.replace(',', '').replace('¥', '').replace('%', '').replace('万', '').replace('円', '')
                float(cleaned)
                is_number = True
            except ValueError:
                is_number = False
            column_types.append('number' if is_number else 'text')

    # Try to use placeholder.insert_table() if available
    from pptx.enum.shapes import PP_PLACEHOLDER_TYPE
    
    table_shape = None
    placeholder_removed = False
    
    if placeholder_shape.is_placeholder:
        try:
            # OBJECT placeholders support insert_table()
            ph_type = placeholder_shape.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER_TYPE.TABLE, PP_PLACEHOLDER_TYPE.OBJECT):
                table_shape = placeholder_shape.insert_table(rows, cols)
                placeholder_removed = True  # insert_table() replaces the placeholder
        except (AttributeError, NotImplementedError, Exception) as e:
            # Placeholder doesn't support insert_table(), fall back to manual insertion
            pass
    
    # Fall back to manual insertion if placeholder method not available
    if table_shape is None:
        left = placeholder_shape.left
        top = placeholder_shape.top
        width = placeholder_shape.width
        height = placeholder_shape.height
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    
    table = table_shape.table

    # Fill data and apply styling
    for r_idx, row_data in enumerate(data):
        is_header = header_row and r_idx == 0

        for c_idx, cell_value in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)

            # Format cell value with thousand separator if it's a number
            cell_text = str(cell_value)
            if not is_header and c_idx < len(column_types) and column_types[c_idx] == 'number':
                # Try to format numbers with thousand separator
                import re
                # Match patterns like "3500万", "1234567円", "98765"
                match = re.match(r'([\d.]+)(.*)', cell_text.replace(',', ''))
                if match:
                    number_part = match.group(1)
                    suffix = match.group(2)  # unit like "万", "円", "%"
                    try:
                        # Parse as number and format with comma
                        num = float(number_part)
                        if num == int(num):
                            formatted = f"{int(num):,}"
                        else:
                            formatted = f"{num:,}"
                        cell_text = formatted + suffix
                    except ValueError:
                        pass  # Keep original if parsing fails

            cell.text = cell_text

            # Apply fill
            cell.fill.solid()
            if is_header:
                # Header: theme color + brightness from style.yaml
                cell.fill.fore_color.theme_color = table_style.header_fill_theme
                cell.fill.fore_color.brightness = table_style.header_fill_brightness
            else:
                # Body: theme color + column-specific brightness from style.yaml
                cell.fill.fore_color.theme_color = table_style.body_fill_theme
                cell.fill.fore_color.brightness = table_style.get_body_brightness(c_idx)

            # Apply text formatting
            text_frame = cell.text_frame
            
            # Get style config based on row type
            if is_header:
                style_config = table_style.header
            else:
                style_config = table_style.body
            
            # Text margins
            from pptx.util import Emu
            text_frame.margin_left = Emu(style_config.margin_left_emu)
            text_frame.margin_right = Emu(style_config.margin_right_emu)
            text_frame.margin_top = Emu(style_config.margin_top_emu)
            text_frame.margin_bottom = Emu(style_config.margin_bottom_emu)
            
            # Vertical alignment from style.yaml
            v_align_str = style_config.vertical_align
            
            v_align_map = {
                'top': MSO_ANCHOR.TOP,
                'middle': MSO_ANCHOR.MIDDLE,
                'bottom': MSO_ANCHOR.BOTTOM
            }
            text_frame.vertical_anchor = v_align_map.get(v_align_str, MSO_ANCHOR.MIDDLE)
            
            # Ensure vertical centering via OOXML (python-pptx API sometimes doesn't work)
            from pptx.oxml.xmlchemy import OxmlElement
            txBody = cell._tc.txBody
            if txBody is not None:
                bodyPr = txBody.bodyPr
                if bodyPr is not None:
                    ooxml_anchor = {'top': 't', 'middle': 'ctr', 'bottom': 'b'}.get(v_align_str, 'ctr')
                    bodyPr.set('anchor', ooxml_anchor)

            for para in text_frame.paragraphs:
                # Horizontal alignment from style.yaml
                h_align_str = style_config.horizontal_align
                
                h_align_map = {
                    'left': PP_ALIGN.LEFT,
                    'center': PP_ALIGN.CENTER,
                    'right': PP_ALIGN.RIGHT,
                    'justify': PP_ALIGN.JUSTIFY
                }
                para.alignment = h_align_map.get(h_align_str, PP_ALIGN.RIGHT)

                for run in para.runs:
                    if is_header:
                        # Header: all attributes from style.yaml
                        run.font.name = table_style.header.font_family
                        run.font.size = Pt(table_style.header.font_size_pt)
                        run.font.bold = table_style.header.font_bold
                        run.font.italic = table_style.header.font_italic
                        run.font.underline = table_style.header.font_underline
                        run.font.color.theme_color = table_style.header_text_theme
                        run.font.color.brightness = table_style.header_text_brightness
                    else:
                        # Body: all attributes from style.yaml
                        run.font.name = table_style.body.font_family
                        run.font.size = Pt(table_style.body.font_size_pt)
                        run.font.bold = table_style.body.font_bold
                        run.font.italic = table_style.body.font_italic
                        run.font.underline = table_style.body.font_underline
                        run.font.color.theme_color = table_style.body_text_theme
                        run.font.color.brightness = table_style.body_text_brightness

            # Apply borders
            _apply_table_borders(cell, r_idx, c_idx, rows, cols)

    # Remove placeholder shape (only if not already removed by insert_table)
    if not placeholder_removed:
        try:
            sp = placeholder_shape._element
            sp.getparent().remove(sp)
            logger.debug("Placeholder shape removed successfully")
        except Exception as e:
            logger.warning(f"Could not remove placeholder shape: {e}")

    return table_shape


def _apply_table_borders(cell, row_idx, col_idx, total_rows, total_cols):
    """Apply borders to table cell according to style.yaml.

    Borders:
    - Outer edges: thick (1.5pt), PRIMARY color
    - Inner edges: standard (1pt), PRIMARY color

    Note: Requires OOXML manipulation as python-pptx has limited table border API.
    """
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.util import Pt

    # Load style configuration
    style = StyleConfig.load()
    table_style = style.table

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Remove existing borders
    for child in list(tcPr):
        if child.tag.endswith(('lnL', 'lnR', 'lnT', 'lnB')):
            tcPr.remove(child)

    # Determine if this is an outer edge
    is_left_edge = col_idx == 0
    is_right_edge = col_idx == total_cols - 1
    is_top_edge = row_idx == 0
    is_bottom_edge = row_idx == total_rows - 1

    # Get border color from style.yaml (theme color reference)
    border_config = table_style.get('border', {}) if hasattr(table_style, 'get') else getattr(table_style, 'border', {})
    border_theme = border_config.get('color_theme', 'ACCENT_1') if isinstance(border_config, dict) else getattr(border_config, 'color_theme', 'ACCENT_1')
    border_brightness = border_config.get('color_brightness', 0.0) if isinstance(border_config, dict) else getattr(border_config, 'color_brightness', 0.0)

    # Helper to create border line element with theme color
    def create_border(width_pt):
        ln = OxmlElement('a:ln')
        ln.set('w', str(int(width_pt * 12700)))  # Convert pt to EMU
        
        solidFill = OxmlElement('a:solidFill')
        
        # Use theme color instead of RGB
        schemeClr = OxmlElement('a:schemeClr')
        theme_color_map = {
            'ACCENT_1': 'accent1',
            'LIGHT_1': 'lt1',
            'DARK_1': 'dk1',
            'bg1': 'bg1',
            'tx1': 'tx1'
        }
        schemeClr.set('val', theme_color_map.get(border_theme, 'accent1'))
        
        # Apply brightness if needed
        if border_brightness != 0.0:
            lumMod = OxmlElement('a:lumMod')
            lumMod.set('val', str(int((1.0 + border_brightness) * 100000)))
            schemeClr.append(lumMod)
        
        solidFill.append(schemeClr)
        ln.append(solidFill)
        
        return ln

    # Apply borders
    # Left border
    lnL = create_border(1.5 if is_left_edge else 1.0)
    lnL.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}lnL'
    tcPr.append(lnL)

    # Right border
    lnR = create_border(1.5 if is_right_edge else 1.0)
    lnR.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}lnR'
    tcPr.append(lnR)

    # Top border
    lnT = create_border(1.5 if is_top_edge else 1.0)
    lnT.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}lnT'
    tcPr.append(lnT)

    # Bottom border
    lnB = create_border(1.5 if is_bottom_edge else 1.0)
    lnB.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}lnB'
    tcPr.append(lnB)


def create_styled_chart(slide, placeholder_shape, spec: Dict[str, Any]):
    """Create native chart with style.yaml styling.

    Args:
        slide: Slide object where chart will be inserted
        placeholder_shape: Shape to replace (uses its position/size)
        spec: Chart specification dict with keys:
            - chart_kind: str - 'line', 'column', 'bar', 'pie' (default: 'line')
            - categories: List[str] - X-axis category labels
            - series: List[Dict] - Series data, each with:
                - name: str - Series name
                - values: List[float] - Y-axis values

    Returns:
        Chart shape object

    Styling applied (from style.yaml):
        - Series colors from style.yaml colors.series
        - Category axis from style.yaml category_axis
        - Value axis from style.yaml value_axis
        - Legend from style.yaml legend
        - Data labels from style.yaml data_labels
    """
    logger = get_logger()
    logger.info(f"Creating styled chart (type: {spec.get('chart_kind', 'line')})")

    # Create generation snapshot (once per session)
    _ensure_snapshot_created()

    # Load style configuration
    style = StyleConfig.load()

    chart_kind = spec.get('chart_kind', 'line')
    categories = spec.get('categories', [])
    series_specs = spec.get('series', [])

    # Validate data
    if not categories:
        logger.error("Chart spec.categories is empty")
        raise ValueError("Chart spec must have categories")

    if not series_specs:
        logger.error("Chart spec.series is empty")
        raise ValueError("Chart spec must have series")

    # Validate series data
    for idx, series_spec in enumerate(series_specs):
        series_name = series_spec.get('name', f'Series{idx}')
        series_values = series_spec.get('values', [])

        if not series_values:
            logger.error(f"Series '{series_name}' has no values")
            raise ValueError(f"Series '{series_name}' must have values")

        if len(series_values) != len(categories):
            logger.error(f"Series '{series_name}' has {len(series_values)} values, expected {len(categories)}")
            raise ValueError(f"Series '{series_name}' must have same length as categories ({len(categories)})")

        # Validate numeric values
        for val_idx, val in enumerate(series_values):
            try:
                float(val)
            except (ValueError, TypeError):
                logger.error(f"Series '{series_name}' value at index {val_idx} is not numeric: {val}")
                raise ValueError(f"Series '{series_name}' contains non-numeric value: {val}")

    logger.debug(f"Chart validated: {len(categories)} categories, {len(series_specs)} series")

    # Map chart_kind to XL_CHART_TYPE
    chart_type_map = {
        'line': XL_CHART_TYPE.LINE,
        'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'column_stacked': XL_CHART_TYPE.COLUMN_STACKED,
        'bar': XL_CHART_TYPE.BAR_CLUSTERED,
        'bar_stacked': XL_CHART_TYPE.BAR_STACKED,
        'pie': XL_CHART_TYPE.PIE,
        'area': XL_CHART_TYPE.AREA,
        'area_stacked': XL_CHART_TYPE.AREA_STACKED,
    }

    chart_type = chart_type_map.get(chart_kind.lower(), XL_CHART_TYPE.LINE)

    # Prepare chart data
    chart_data = CategoryChartData()
    chart_data.categories = categories

    for series_spec in series_specs:
        series_name = series_spec.get('name', 'Series')
        series_values = series_spec.get('values', [])
        chart_data.add_series(series_name, series_values)

    # Try to use placeholder.insert_chart() if available
    from pptx.enum.shapes import PP_PLACEHOLDER_TYPE
    
    chart_shape = None
    placeholder_removed = False
    
    if placeholder_shape.is_placeholder:
        try:
            # OBJECT placeholders support insert_chart()
            ph_type = placeholder_shape.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER_TYPE.CHART, PP_PLACEHOLDER_TYPE.OBJECT):
                chart_shape = placeholder_shape.insert_chart(chart_type, chart_data)
                placeholder_removed = True  # insert_chart() replaces the placeholder
        except (AttributeError, NotImplementedError, Exception) as e:
            # Placeholder doesn't support insert_chart(), fall back to manual insertion
            pass
    
    # Fall back to manual insertion if placeholder method not available
    if chart_shape is None:
        left = placeholder_shape.left
        top = placeholder_shape.top
        width = placeholder_shape.width
        height = placeholder_shape.height
        chart_shape = slide.shapes.add_chart(
            chart_type, left, top, width, height, chart_data
        )
    
    chart = chart_shape.chart

    # Apply .crtx template styling
    # Chart.crtx is the Single Source of Truth for chart styling
    if not CRTX_AVAILABLE:
        logger.error("crtx_utils not available - required for chart styling")
        raise RuntimeError("crtx_utils not available - required for chart styling")

    # Load chart template styling
    if not CHART_CRTX_PATH.exists():
        logger.error(f"Chart template not found at {CHART_CRTX_PATH}")
        raise FileNotFoundError(f"Chart template not found at {CHART_CRTX_PATH}")

    logger.debug(f"Loading chart template from {CHART_CRTX_PATH}")
    crtx_styling = extract_crtx_styling(str(CHART_CRTX_PATH))

    # Apply styling (area charts have special handling in apply_crtx_styling_to_chart)
    # NOTE: Area charts skip position and gridlines due to python-pptx bugs
    apply_crtx_styling_to_chart(chart, crtx_styling, limited_mode=False)
    logger.info("Chart styling applied successfully")

    # Remove placeholder shape (only if not already removed by insert_chart)
    if not placeholder_removed:
        try:
            sp = placeholder_shape._element
            sp.getparent().remove(sp)
            logger.debug("Placeholder shape removed successfully")
        except Exception as e:
            logger.warning(f"Could not remove placeholder shape: {e}")

    return chart_shape


def create_styled_diagram(slide, placeholder_shape, spec: Dict[str, Any]):
    """Create native diagram with basic nodes (simplified implementation).

    Note: Full diagram with connectors is complex. This is a simplified version
    that creates only nodes. For complex diagrams, use IMAGE_MERMAID mode.

    Args:
        slide: Slide object where diagram will be inserted
        placeholder_shape: Shape to replace
        spec: Diagram specification dict with keys:
            - nodes: List[Dict] - Node definitions, each with:
                - text: str - Node label
                - position: tuple - (x, y) relative position (0-1 range)
                - width: float - Node width in inches (optional)
                - height: float - Node height in inches (optional)

    Returns:
        List of shape objects (nodes)

    Future: Add connector support with bentConnector3 style
    """
    # Load style configuration
    style = StyleConfig.load()
    diagram_style = style.diagram

    nodes = spec.get('nodes', [])

    if not nodes:
        raise ValueError("Diagram spec must have nodes")

    # Get bounds from placeholder
    base_left = placeholder_shape.left
    base_top = placeholder_shape.top
    base_width = placeholder_shape.width
    base_height = placeholder_shape.height

    created_shapes = []

    for node_spec in nodes:
        text = node_spec.get('text', '')
        position = node_spec.get('position', (0.5, 0.5))  # Default center
        node_width = Inches(node_spec.get('width', 2.0))
        node_height = Inches(node_spec.get('height', 1.0))

        # Calculate absolute position
        left = base_left + int(position[0] * base_width) - node_width // 2
        top = base_top + int(position[1] * base_height) - node_height // 2

        # Create rounded rectangle (node)
        from pptx.enum.shapes import MSO_SHAPE
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            left, top, node_width, node_height
        )

        # Apply styling from style.yaml diagram section
        # Node fill: Primary color (RGB)
        shape.fill.solid()
        node_style = diagram_style.node
        shape.fill.fore_color.rgb = style.hex_to_rgb(node_style.fill)

        # Node border: Primary color (RGB)
        shape.line.color.rgb = style.hex_to_rgb(node_style.border_color)
        shape.line.width = Pt(node_style.border_width_pt)

        # Add text
        text_frame = shape.text_frame
        text_frame.text = text

        # Text styling: Light color (theme)
        for para in text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
            for run in para.runs:
                run.font.color.theme_color = style.get_theme_color(node_style.text_color_theme)
                run.font.size = Pt(node_style.font_size_pt)
                run.font.bold = node_style.font_bold

        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        created_shapes.append(shape)

    # Remove placeholder if it's a placeholder
    # Note: Diagrams don't have a native placeholder.insert_diagram() method,
    # so we always need to remove the placeholder after creating the shapes
    if placeholder_shape.is_placeholder:
        try:
            sp = placeholder_shape._element
            sp.getparent().remove(sp)
            logger.debug("Placeholder shape removed successfully")
        except Exception as e:
            logger.warning(f"Could not remove placeholder shape: {e}")

    return created_shapes


def _ensure_snapshot_created():
    """Ensure generation snapshot is created once per session.

    Creates a snapshot of templates/styles from ~/.claude/skills/pptx/templates/
    to powerpoint/processing/snapshot/ for audit purposes.

    This function is called automatically by create_styled_table() and
    create_styled_chart() on first use.
    """
    global _SNAPSHOT_CREATED

    if _SNAPSHOT_CREATED:
        return

    logger = get_logger()
    try:
        create_generation_snapshot()
        _SNAPSHOT_CREATED = True
    except Exception as e:
        logger.warning(f"Failed to create generation snapshot: {e}")
        # Don't fail the entire generation if snapshot fails
        _SNAPSHOT_CREATED = True  # Don't retry
