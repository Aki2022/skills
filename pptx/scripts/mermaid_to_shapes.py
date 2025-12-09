#!/usr/bin/env python3
"""Convert Mermaid flowchart to native PowerPoint shapes.

This module parses Mermaid flowchart code and creates editable
PowerPoint shapes (rectangles, diamonds, connectors) instead of
embedded images.

Usage:
    from mermaid_to_shapes import create_flowchart_shapes

    mermaid_code = '''
    flowchart LR
        A[Start] --> B{Decision}
        B -->|Yes| C[Action]
        B -->|No| D[End]
    '''

    shapes = create_flowchart_shapes(slide, placeholder, mermaid_code)
"""

import re
from typing import Dict, List, Tuple, Any

from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR, MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

# Try to import style_config
try:
    from scripts.style_config import StyleConfig
    STYLE_CONFIG_AVAILABLE = True
except ImportError:
    try:
        from style_config import StyleConfig
        STYLE_CONFIG_AVAILABLE = True
    except ImportError:
        STYLE_CONFIG_AVAILABLE = False


def get_style():
    """Get styling from style.yaml."""
    if STYLE_CONFIG_AVAILABLE:
        try:
            return StyleConfig.load()
        except:
            pass
    return None


def get_flowchart_config(style):
    """Get flowchart configuration from style.yaml.

    Returns default values if style.yaml not available.
    """
    if style and hasattr(style, '_style_data') and 'flowchart' in style._style_data:
        return style._style_data['flowchart']

    # Default configuration
    return {
        'direction': 'LR',
        'node': {
            'shape': 'rounded_rectangle',
            'fill': '#4F4F70',
            'fill_theme': None,
            'border_width_pt': 0,
            'corner_radius_pt': 8,
            'shadow': {
                'enabled': True,
                'blur_pt': 4,
                'distance_pt': 3,
                'direction_deg': 45,
                'opacity': 0.4
            },
            'text': {
                'color': '#FFFFFF',
                'font_size_pt': 11,
                'bold': False,
                'vertical_align': 'middle',
                'horizontal_align': 'center'
            }
        },
        'connector': {
            'type': 'elbow',
            'color_theme': 'bg1',
            'color_brightness': -0.25,
            'width_pt': 1.0,
            'dash_style': 'dash',
            'arrow': {
                'type': 'triangle',
                'size': 'medium'
            }
        },
        'label': {
            'font_size_pt': 9,
            'color_theme': 'bg1',
            'color_brightness': -0.5
        }
    }


def parse_mermaid_flowchart(mermaid_code: str) -> Tuple[Dict, List]:
    """Parse Mermaid flowchart code into nodes and edges.

    Args:
        mermaid_code: Mermaid flowchart code

    Returns:
        Tuple of (nodes dict, edges list)
        nodes: {id: {'text': str, 'shape': 'rect'|'diamond'|'rounded'}}
        edges: [{'from': id, 'to': id, 'label': str}]
    """
    nodes = {}
    edges = []

    # Remove comments and normalize
    lines = mermaid_code.strip().split('\n')

    for line in lines:
        line = line.strip()

        # Skip flowchart declaration and empty lines
        if not line or line.startswith('flowchart') or line.startswith('graph'):
            continue

        # Parse node definitions and connections
        # Pattern: A[Text] or A{Text} or A(Text) or A((Text))
        # Connection: A --> B or A -->|label| B

        # Find connections first
        conn_match = re.search(r'(\w+)\s*--+>?\s*(?:\|([^|]*)\|)?\s*(\w+)', line)
        if conn_match:
            from_id = conn_match.group(1)
            label = conn_match.group(2) or ''
            to_id = conn_match.group(3)
            edges.append({'from': from_id, 'to': to_id, 'label': label.strip()})

        # Find node definitions
        # [text] = rectangle
        # {text} = diamond
        # (text) = rounded
        # ((text)) = circle
        node_patterns = [
            (r'(\w+)\[([^\]]+)\]', 'rect'),      # A[Text]
            (r'(\w+)\{([^}]+)\}', 'diamond'),    # A{Text}
            (r'(\w+)\(([^)]+)\)', 'rounded'),    # A(Text)
        ]

        for pattern, shape_type in node_patterns:
            for match in re.finditer(pattern, line):
                node_id = match.group(1)
                text = match.group(2).strip()
                if node_id not in nodes:
                    nodes[node_id] = {'text': text, 'shape': shape_type}

    return nodes, edges


def calculate_layout(nodes: Dict, edges: List, bounds: Tuple, direction: str = 'LR') -> Dict:
    """Calculate positions for nodes.

    Args:
        nodes: Node definitions
        edges: Edge definitions
        bounds: (left, top, width, height) in EMU
        direction: 'LR' for left-to-right, 'TD' for top-down

    Returns:
        Dict mapping node_id to (left, top, width, height)
    """
    left, top, width, height = bounds

    # Build dependency graph to determine levels
    node_ids = list(nodes.keys())
    levels = {}
    incoming = {n: 0 for n in node_ids}

    for edge in edges:
        if edge['to'] in incoming:
            incoming[edge['to']] += 1

    # Find root nodes (no incoming edges)
    queue = [n for n in node_ids if incoming[n] == 0]
    for n in queue:
        levels[n] = 0

    # BFS to assign levels
    processed = set(queue)
    while queue:
        current = queue.pop(0)
        current_level = levels[current]

        for edge in edges:
            if edge['from'] == current:
                next_node = edge['to']
                if next_node not in levels:
                    levels[next_node] = current_level + 1
                else:
                    levels[next_node] = max(levels[next_node], current_level + 1)

                if next_node not in processed:
                    queue.append(next_node)
                    processed.add(next_node)

    # Assign levels to any remaining nodes
    for n in node_ids:
        if n not in levels:
            levels[n] = 0

    # Group nodes by level
    level_groups = {}
    for node_id, level in levels.items():
        if level not in level_groups:
            level_groups[level] = []
        level_groups[level].append(node_id)

    # Calculate positions
    positions = {}
    num_levels = max(levels.values()) + 1 if levels else 1

    # Node dimensions
    node_width = Inches(1.5)
    node_height = Inches(0.8)

    if direction == 'LR':
        # Horizontal layout
        level_width = width // num_levels
        for level, group in level_groups.items():
            x = left + level * level_width + (level_width - node_width) // 2
            num_nodes = len(group)
            node_spacing = height // (num_nodes + 1)

            for i, node_id in enumerate(group):
                y = top + (i + 1) * node_spacing - node_height // 2
                positions[node_id] = (int(x), int(y), int(node_width), int(node_height))
    else:
        # Vertical layout (TD)
        level_height = height // num_levels
        for level, group in level_groups.items():
            y = top + level * level_height + (level_height - node_height) // 2
            num_nodes = len(group)
            node_spacing = width // (num_nodes + 1)

            for i, node_id in enumerate(group):
                x = left + (i + 1) * node_spacing - node_width // 2
                positions[node_id] = (int(x), int(y), int(node_width), int(node_height))

    return positions


def create_flowchart_shapes(slide, placeholder, mermaid_code: str, direction: str = None):
    """Create native PowerPoint shapes from Mermaid flowchart.

    Args:
        slide: PowerPoint slide object
        placeholder: Placeholder shape (for position/size)
        mermaid_code: Mermaid flowchart code
        direction: 'LR' or 'TD' (default from style.yaml)

    Returns:
        List of created shape objects
    """
    # Get styling from style.yaml
    style = get_style()
    config = get_flowchart_config(style)

    # Get direction from config if not specified
    if direction is None:
        direction = config.get('direction', 'LR')

    # Get node config
    node_config = config.get('node', {})
    connector_config = config.get('connector', {})
    label_config = config.get('label', {})

    # Get primary color - use theme color if specified, otherwise RGB
    fill_theme = node_config.get('fill_theme')
    fill_rgb = node_config.get('fill')

    # Prefer theme color, but fall back to RGB if theme is null
    if fill_theme:
        primary_theme = fill_theme
        primary_rgb = None
    elif fill_rgb:
        primary_theme = None
        primary_rgb = fill_rgb
    else:
        # Default to ACCENT_1 theme color
        primary_theme = 'ACCENT_1'
        primary_rgb = None

    # Parse Mermaid code
    nodes, edges = parse_mermaid_flowchart(mermaid_code)

    if not nodes:
        raise ValueError("No nodes found in Mermaid code")

    # Get bounds from placeholder
    bounds = (placeholder.left, placeholder.top, placeholder.width, placeholder.height)

    # Calculate layout
    positions = calculate_layout(nodes, edges, bounds, direction)

    # Remove placeholder
    sp = placeholder._element
    sp.getparent().remove(sp)

    created_shapes = []
    shape_refs = {}  # node_id -> shape

    # Create node shapes
    for node_id, node_info in nodes.items():
        if node_id not in positions:
            continue

        x, y, w, h = positions[node_id]
        text = node_info['text']
        shape_type = node_info['shape']

        # Determine PowerPoint shape type
        # Template uses ROUNDED_RECTANGLE for most shapes
        if shape_type == 'diamond':
            mso_shape = MSO_SHAPE.DIAMOND
        else:  # rect and rounded both use rounded rectangle
            mso_shape = MSO_SHAPE.ROUNDED_RECTANGLE

        # Create shape
        shape = slide.shapes.add_shape(mso_shape, x, y, w, h)

        # Apply fill (from style.yaml) - use theme color or RGB
        shape.fill.solid()
        if primary_theme:
            # Use theme color
            shape.fill.fore_color.theme_color = style.get_theme_color(primary_theme) if style else MSO_THEME_COLOR.ACCENT_1
            # Apply fill brightness
            fill_brightness = node_config.get('fill_brightness', 0.0)
            if fill_brightness != 0.0:
                shape.fill.fore_color.brightness = fill_brightness
        elif primary_rgb:
            # Use RGB color
            rgb_str = primary_rgb.strip('#')
            r = int(rgb_str[0:2], 16)
            g = int(rgb_str[2:4], 16)
            b = int(rgb_str[4:6], 16)
            shape.fill.fore_color.rgb = RGBColor(r, g, b)

        # Border (from style.yaml)
        border_width = node_config.get('border_width_pt', 0)
        if border_width == 0:
            shape.line.fill.background()
        else:
            shape.line.width = Pt(border_width)
            # Apply border color
            border_theme = node_config.get('border_theme')
            if border_theme and style:
                shape.line.color.theme_color = style.get_theme_color(border_theme)
                border_brightness = node_config.get('border_brightness', 0.0)
                if border_brightness != 0.0:
                    shape.line.color.brightness = border_brightness

        # Shadow - explicitly disable (override theme defaults)
        # Always disable shadow as style.yaml has shadow.enabled: false
        spPr = shape._element.spPr
        # Remove any existing effectLst
        existing_effectLst = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst')
        if existing_effectLst is not None:
            spPr.remove(existing_effectLst)
        # Add empty effectLst to disable shadow
        etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst')

        # Add text (from style.yaml)
        text_config = node_config.get('text', {})
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = text

        # Horizontal alignment
        h_align = text_config.get('horizontal_align', 'center')
        align_map = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
        p.alignment = align_map.get(h_align, PP_ALIGN.CENTER)

        # Text styling - use theme colors or RGB from style.yaml
        font_family = text_config.get('font_family', 'Arial')
        font_size = text_config.get('font_size_pt', 12)
        font_bold = text_config.get('bold', False)
        font_italic = text_config.get('italic', False)

        # Get color - prefer theme, fallback to RGB
        color_theme = text_config.get('color_theme')
        color_rgb = text_config.get('color')
        color_brightness = text_config.get('color_brightness', 0.0)

        # Apply font formatting to runs
        for run in p.runs:
            run.font.name = font_family
            run.font.size = Pt(font_size)
            run.font.bold = font_bold
            run.font.italic = font_italic

            # Apply color
            if color_theme and style:
                # Use theme color
                run.font.color.theme_color = style.get_theme_color(color_theme)
                if color_brightness != 0.0:
                    run.font.color.brightness = color_brightness
            elif color_rgb:
                # Use RGB color
                rgb_str = color_rgb.strip('#')
                r = int(rgb_str[0:2], 16)
                g = int(rgb_str[2:4], 16)
                b = int(rgb_str[4:6], 16)
                run.font.color.rgb = RGBColor(r, g, b)

        # Vertical alignment
        v_align = text_config.get('vertical_align', 'middle')
        v_align_map = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE, 'bottom': MSO_ANCHOR.BOTTOM}
        tf.vertical_anchor = v_align_map.get(v_align, MSO_ANCHOR.MIDDLE)
        tf.word_wrap = True
        tf.auto_size = None

        shape_refs[node_id] = shape
        created_shapes.append(shape)

    # Create connectors
    for edge in edges:
        from_id = edge['from']
        to_id = edge['to']

        if from_id not in shape_refs or to_id not in shape_refs:
            continue

        from_shape = shape_refs[from_id]
        to_shape = shape_refs[to_id]

        # Get connector points
        # From: right center, To: left center (for LR)
        if direction == 'LR':
            begin_x = from_shape.left + from_shape.width
            begin_y = from_shape.top + from_shape.height // 2
            end_x = to_shape.left
            end_y = to_shape.top + to_shape.height // 2
        else:  # TD
            begin_x = from_shape.left + from_shape.width // 2
            begin_y = from_shape.top + from_shape.height
            end_x = to_shape.left + to_shape.width // 2
            end_y = to_shape.top

        # Create connector (from style.yaml)
        conn_type = connector_config.get('type', 'elbow')
        conn_type_map = {'straight': MSO_CONNECTOR.STRAIGHT, 'elbow': MSO_CONNECTOR.ELBOW}
        mso_conn = conn_type_map.get(conn_type, MSO_CONNECTOR.ELBOW)

        connector = slide.shapes.add_connector(
            mso_conn,
            begin_x, begin_y,
            end_x, end_y
        )

        # Remove shadow from connector
        conn_spPr = connector._element.spPr
        existing_effectLst = conn_spPr.find(qn('a:effectLst'))
        if existing_effectLst is not None:
            conn_spPr.remove(existing_effectLst)
        etree.SubElement(conn_spPr, qn('a:effectLst'))

        # Style connector (from style.yaml)
        color_theme = connector_config.get('color_theme', 'bg1')
        color_brightness = connector_config.get('color_brightness', -0.25)
        conn_width = connector_config.get('width_pt', 1.0)
        dash_style = connector_config.get('dash_style', 'dash')

        connector.line.color.theme_color = style.get_theme_color(color_theme) if style else MSO_THEME_COLOR.BACKGROUND_1
        connector.line.color.brightness = color_brightness
        connector.line.width = Pt(conn_width)

        # Dash style
        dash_map = {
            'solid': MSO_LINE_DASH_STYLE.SOLID,
            'dash': MSO_LINE_DASH_STYLE.DASH,
            'dot': MSO_LINE_DASH_STYLE.ROUND_DOT
        }
        connector.line.dash_style = dash_map.get(dash_style, MSO_LINE_DASH_STYLE.DASH)

        # Add arrow at end (from style.yaml)
        arrow_config = connector_config.get('arrow', {})
        arrow_type = arrow_config.get('type', 'triangle')
        arrow_size = arrow_config.get('size', 'medium')
        size_map = {'small': 'sm', 'medium': 'med', 'large': 'lg'}

        ln = connector.line._ln
        tailEnd = ln.find(qn('a:tailEnd'))
        if tailEnd is None:
            tailEnd = connector.line._ln.makeelement(qn('a:tailEnd'))
            ln.append(tailEnd)
        tailEnd.set('type', arrow_type)
        tailEnd.set('w', size_map.get(arrow_size, 'med'))
        tailEnd.set('len', size_map.get(arrow_size, 'med'))

        created_shapes.append(connector)

        # Add label if present
        if edge['label']:
            # Create text box for label
            label_x = (begin_x + end_x) // 2 - Inches(0.5)
            label_y = (begin_y + end_y) // 2 - Inches(0.2)

            label_box = slide.shapes.add_textbox(
                label_x, label_y,
                Inches(1), Inches(0.4)
            )

            # Remove shadow from label
            label_spPr = label_box._element.spPr
            existing_effectLst = label_spPr.find(qn('a:effectLst'))
            if existing_effectLst is not None:
                label_spPr.remove(existing_effectLst)
            etree.SubElement(label_spPr, qn('a:effectLst'))

            tf = label_box.text_frame
            p = tf.paragraphs[0]
            p.text = edge['label']
            p.alignment = PP_ALIGN.CENTER

            # Label styling (from style.yaml)
            label_font_size = label_config.get('font_size_pt', 9)
            label_color_theme = label_config.get('color_theme', 'bg1')
            label_brightness = label_config.get('color_brightness', -0.5)

            for run in p.runs:
                run.font.size = Pt(label_font_size)
                run.font.color.theme_color = style.get_theme_color(label_color_theme) if style else MSO_THEME_COLOR.BACKGROUND_1
                run.font.color.brightness = label_brightness

            created_shapes.append(label_box)

    return created_shapes


if __name__ == '__main__':
    # Test parsing
    test_code = '''flowchart LR
        A[開始] --> B{条件1}
        B -->|Yes| C[処理A]
        B -->|No| D{条件2}
        D -->|Yes| E[処理B]
        D -->|No| F[処理C]
        C --> G[終了]
        E --> G
        F --> G
    '''

    nodes, edges = parse_mermaid_flowchart(test_code)

    print("Parsed nodes:")
    for node_id, info in nodes.items():
        print(f"  {node_id}: {info}")

    print("\nParsed edges:")
    for edge in edges:
        print(f"  {edge['from']} --> {edge['to']} [{edge['label']}]")
