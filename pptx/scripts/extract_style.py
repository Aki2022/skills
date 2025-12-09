#!/usr/bin/env python3
"""Extract styling from Chart.crtx and template.pptx to generate style.yaml.

This module creates a Single Source of Truth (style.yaml) from:
- Chart.crtx: Chart styling (series, axes, legend, data labels)
- template.pptx Slide 1: Table styling
- template.pptx Slide 2: Flowchart/Diagram styling

The generated YAML can be used by Python, R, and JavaScript/Mermaid.

Usage:
    python scripts/extract_style.py

Example:
    python scripts/extract_style.py
    python scripts/extract_style.py --template path/to/template.pptx
"""

import os
import sys
import yaml
from typing import Dict, Any, Optional

# Add scripts directory to path for imports
sys.path.insert(0, os.path.dirname(__file__))

from crtx_utils import extract_crtx_styling, lummod_to_brightness
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR, MSO_COLOR_TYPE
from pptx.util import Emu


def rgb_to_hex(rgb) -> str:
    """Convert RGBColor to hex string."""
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def theme_to_str(theme_color) -> str:
    """Convert MSO_THEME_COLOR to string."""
    theme_map = {
        MSO_THEME_COLOR.TEXT_1: 'tx1',
        MSO_THEME_COLOR.TEXT_2: 'tx2',
        MSO_THEME_COLOR.BACKGROUND_1: 'bg1',
        MSO_THEME_COLOR.BACKGROUND_2: 'bg2',
        MSO_THEME_COLOR.DARK_1: 'dk1',
        MSO_THEME_COLOR.DARK_2: 'dk2',
        MSO_THEME_COLOR.LIGHT_1: 'lt1',
        MSO_THEME_COLOR.LIGHT_2: 'lt2',
        MSO_THEME_COLOR.ACCENT_1: 'accent1',
        MSO_THEME_COLOR.ACCENT_2: 'accent2',
    }
    return theme_map.get(theme_color, 'bg1')


def extract_table_style(prs: Presentation, slide_index: int = 1) -> Dict[str, Any]:
    """Extract table styling from template.pptx.

    Args:
        prs: Presentation object
        slide_index: Slide number (1-based, default 1)

    Returns:
        Table style dictionary
    """
    slide = prs.slides[slide_index - 1]

    table_style = {
        'border': {
            'color': '#4F4F70',
            'width_outer_pt': 1.5,
            'width_inner_pt': 1.0,
        },
        'header': {
            'fill_theme': 'bg1',
            'fill_brightness': -0.5,
            'text_color_theme': 'lt1',
            'font_bold': True,
            'font_size_pt': 12,
        },
        'body': {
            'fill_theme': 'bg1',
            'column_brightness': [-0.15, -0.05, -0.05, -0.05],
            'text_color_theme': 'dk1',
            'font_size_pt': 12,
        },
        'alignment': 'right',
    }

    # Find table shape
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table

            # Extract header styling from first row
            if len(table.rows) > 0:
                header_cell = table.cell(0, 0)
                fill = header_cell.fill

                if fill.type is not None:
                    try:
                        fc = fill.fore_color
                        if fc.type == MSO_COLOR_TYPE.SCHEME:
                            table_style['header']['fill_theme'] = theme_to_str(fc.theme_color)
                            table_style['header']['fill_brightness'] = round(fc.brightness, 2)
                        elif fc.type == MSO_COLOR_TYPE.RGB:
                            table_style['header']['fill_rgb'] = rgb_to_hex(fc.rgb)
                    except:
                        pass

                # Text styling
                if header_cell.text_frame.paragraphs:
                    para = header_cell.text_frame.paragraphs[0]
                    if para.runs:
                        run = para.runs[0]
                        if run.font.size:
                            table_style['header']['font_size_pt'] = int(run.font.size.pt)
                        table_style['header']['font_bold'] = run.font.bold or False

                        try:
                            if run.font.color.type == MSO_COLOR_TYPE.SCHEME:
                                table_style['header']['text_color_theme'] = theme_to_str(run.font.color.theme_color)
                                if run.font.color.brightness is not None:
                                    table_style['header']['text_color_brightness'] = round(run.font.color.brightness, 2)
                        except:
                            pass

            # Extract body styling from second row
            if len(table.rows) > 1:
                brightnesses = []
                for col_idx in range(len(table.columns)):
                    try:
                        body_cell = table.cell(1, col_idx)
                        fill = body_cell.fill
                        if fill.type is not None:
                            fc = fill.fore_color
                            if fc.type == MSO_COLOR_TYPE.SCHEME:
                                brightnesses.append(round(fc.brightness, 2))
                                if col_idx == 0:
                                    table_style['body']['fill_theme'] = theme_to_str(fc.theme_color)
                    except:
                        brightnesses.append(-0.05)

                if brightnesses:
                    table_style['body']['column_brightness'] = brightnesses

                # Body text styling
                body_cell = table.cell(1, 0)
                if body_cell.text_frame.paragraphs:
                    para = body_cell.text_frame.paragraphs[0]
                    if para.runs:
                        run = para.runs[0]
                        if run.font.size:
                            table_style['body']['font_size_pt'] = int(run.font.size.pt)
                        try:
                            if run.font.color.type == MSO_COLOR_TYPE.SCHEME:
                                table_style['body']['text_color_theme'] = theme_to_str(run.font.color.theme_color)
                                if run.font.color.brightness is not None:
                                    table_style['body']['text_color_brightness'] = round(run.font.color.brightness, 2)
                        except:
                            pass

            break

    return table_style


def extract_flowchart_style(prs: Presentation, slide_index: int = 2) -> Dict[str, Any]:
    """Extract flowchart/diagram styling from template.pptx.

    Args:
        prs: Presentation object
        slide_index: Slide number (1-based, default 2)

    Returns:
        Flowchart style dictionary
    """
    slide = prs.slides[slide_index - 1]

    flowchart_style = {
        'direction': 'LR',
        'node': {
            'shape': 'rounded_rectangle',
            'fill': '#4F4F70',
            'fill_theme': None,
            'border_width_pt': 0,
            'shadow': {
                'enabled': True,
                'blur_pt': 4,
                'distance_pt': 3,
                'direction_deg': 45,
                'opacity': 0.4,
            },
            'text': {
                'color': '#FFFFFF',
                'font_size_pt': 11,
                'bold': False,
            },
        },
        'connector': {
            'type': 'elbow',
            'color_theme': 'bg1',
            'color_brightness': -0.25,
            'width_pt': 1.0,
            'dash_style': 'dash',
            'arrow': {
                'type': 'triangle',
                'size': 'medium',
            },
        },
        'label': {
            'font_size_pt': 9,
            'color_theme': 'bg1',
            'color_brightness': -0.5,
        },
    }

    # Find shapes to analyze - prioritize RGB-filled shapes (like "Good" box)
    primary_shape = None
    connector_found = False

    for shape in slide.shapes:
        shape_name = shape.name.lower()

        # Extract connector styling (first one found)
        if shape.shape_type == 9 and not connector_found:  # LINE (connector)
            try:
                line = shape.line
                if line.color.type == MSO_COLOR_TYPE.SCHEME:
                    flowchart_style['connector']['color_theme'] = theme_to_str(line.color.theme_color)
                    flowchart_style['connector']['color_brightness'] = round(line.color.brightness, 2)

                if line.width:
                    flowchart_style['connector']['width_pt'] = round(line.width.pt, 1)

                if hasattr(line, 'dash_style') and line.dash_style:
                    dash_map = {1: 'solid', 4: 'dash', 2: 'dot'}
                    flowchart_style['connector']['dash_style'] = dash_map.get(int(line.dash_style), 'dash')
                connector_found = True
            except:
                pass
            continue

        # Skip non-auto shapes
        if shape.shape_type != 1:
            continue

        # Check if it's a node shape (rounded rectangle with fill)
        if 'rounded' in shape_name or 'rectangle' in shape_name or '角丸' in shape_name:
            try:
                fill = shape.fill
                if fill.type == 1:  # SOLID
                    fc = fill.fore_color
                    # Prioritize RGB-filled shapes (primary color like "Good")
                    if fc.type == MSO_COLOR_TYPE.RGB:
                        primary_shape = shape
                        break  # Found primary color shape
                    elif primary_shape is None:
                        # Keep as fallback if no RGB shape found yet
                        primary_shape = shape
            except:
                pass

    # Extract styling from the selected shape
    if primary_shape:
        try:
            fill = primary_shape.fill
            if fill.type == 1:
                fc = fill.fore_color
                if fc.type == MSO_COLOR_TYPE.RGB:
                    flowchart_style['node']['fill'] = rgb_to_hex(fc.rgb)
                    flowchart_style['node']['fill_theme'] = None  # Explicitly set to None
                elif fc.type == MSO_COLOR_TYPE.SCHEME:
                    flowchart_style['node']['fill_theme'] = theme_to_str(fc.theme_color)
                    if fc.brightness:
                        flowchart_style['node']['fill_brightness'] = round(fc.brightness, 2)

            # Border
            line = primary_shape.line
            if line.width:
                flowchart_style['node']['border_width_pt'] = round(line.width.pt, 1)

            # Text styling
            if primary_shape.has_text_frame and primary_shape.text_frame.paragraphs:
                para = primary_shape.text_frame.paragraphs[0]
                if para.runs:
                    run = para.runs[0]
                    if run.font.size:
                        flowchart_style['node']['text']['font_size_pt'] = int(run.font.size.pt)
                    flowchart_style['node']['text']['bold'] = run.font.bold or False

                    try:
                        if run.font.color.type == MSO_COLOR_TYPE.RGB:
                            flowchart_style['node']['text']['color'] = rgb_to_hex(run.font.color.rgb)
                        elif run.font.color.type == MSO_COLOR_TYPE.SCHEME:
                            flowchart_style['node']['text']['color_theme'] = theme_to_str(run.font.color.theme_color)
                    except:
                        pass
        except:
            pass

    return flowchart_style


def convert_to_style_yaml(crtx_styling: Dict[str, Any]) -> Dict[str, Any]:
    """Convert crtx_styling dict to style.yaml format.

    Args:
        crtx_styling: Raw styling from extract_crtx_styling()

    Returns:
        Dictionary in style.yaml format
    """
    style = {
        'colors': {
            'primary': '#4F4F70',  # Default, will be updated from series[0]
            'series': []
        },
        'category_axis': {},
        'value_axis': {},
        'legend': {},
        'data_labels': [],
        'gridlines': {'enabled': False},
    }

    # Convert series colors
    for idx, ser in enumerate(crtx_styling.get('series', [])):
        series_entry = {}

        if ser.get('fill_type') == 'rgb':
            series_entry['type'] = 'rgb'
            series_entry['value'] = ser.get('fill_value', '#000000')
            if idx == 0:
                style['colors']['primary'] = ser.get('fill_value', '#4F4F70')

        elif ser.get('fill_type') == 'theme':
            series_entry['type'] = 'theme'
            series_entry['value'] = ser.get('fill_value', 'bg1')

            if 'fill_lummod' in ser:
                brightness = lummod_to_brightness(ser['fill_lummod'])
                series_entry['brightness'] = round(brightness, 2)

        style['colors']['series'].append(series_entry)

    # Convert category axis
    cat_axis = crtx_styling.get('category_axis', {})
    if cat_axis:
        style['category_axis'] = {
            'visible': cat_axis.get('visible', True),
            'tick_marks': cat_axis.get('major_tick_mark', 'none'),
            'line': {},
            'font': {}
        }

        # Line styling
        if 'line_width_emu' in cat_axis:
            width_pt = cat_axis['line_width_emu'] / 12700.0
            style['category_axis']['line']['width_pt'] = round(width_pt, 2)

        if 'line_color_type' in cat_axis:
            style['category_axis']['line']['color_type'] = cat_axis['line_color_type']
            style['category_axis']['line']['color_value'] = cat_axis.get('line_color_value', 'tx1')

            if 'line_lummod' in cat_axis or 'line_lumoff' in cat_axis:
                lummod = cat_axis.get('line_lummod', 100000)
                lumoff = cat_axis.get('line_lumoff', 0)
                brightness = lummod_to_brightness(lummod, lumoff)
                style['category_axis']['line']['brightness'] = round(brightness, 2)

        # Font styling
        if 'font_size_pt' in cat_axis:
            style['category_axis']['font']['size_pt'] = int(cat_axis['font_size_pt'])

        if 'font_color_theme' in cat_axis:
            style['category_axis']['font']['color_type'] = 'theme'
            style['category_axis']['font']['color_value'] = cat_axis['font_color_theme']

            if 'font_lummod' in cat_axis or 'font_lumoff' in cat_axis:
                lummod = cat_axis.get('font_lummod', 100000)
                lumoff = cat_axis.get('font_lumoff', 0)
                brightness = lummod_to_brightness(lummod, lumoff)
                style['category_axis']['font']['brightness'] = round(brightness, 2)

    # Convert value axis
    val_axis = crtx_styling.get('value_axis', {})
    if val_axis:
        style['value_axis'] = {
            'visible': val_axis.get('visible', True),
            'tick_marks': val_axis.get('major_tick_mark', 'none'),
        }

    # Convert legend
    legend = crtx_styling.get('legend', {})
    if legend:
        # Map position codes to full names
        pos_map = {'b': 'bottom', 't': 'top', 'r': 'right', 'l': 'left'}
        pos_code = legend.get('position', 'b')
        position = pos_map.get(pos_code, pos_code)

        style['legend'] = {
            'position': position,
            'font': {}
        }

        if 'font_size_pt' in legend:
            style['legend']['font']['size_pt'] = int(legend['font_size_pt'])

        if 'font_color_theme' in legend:
            style['legend']['font']['color_type'] = 'theme'
            style['legend']['font']['color_value'] = legend['font_color_theme']

            if 'font_lummod' in legend or 'font_lumoff' in legend:
                lummod = legend.get('font_lummod', 100000)
                lumoff = legend.get('font_lumoff', 0)
                brightness = lummod_to_brightness(lummod, lumoff)
                style['legend']['font']['brightness'] = round(brightness, 2)

    # Convert data labels
    for dl in crtx_styling.get('data_labels', []):
        dl_entry = {
            'show_value': dl.get('show_value', False),
        }

        if 'font_size_pt' in dl:
            dl_entry['font_size_pt'] = int(dl['font_size_pt'])

        if 'font_color_theme' in dl:
            dl_entry['font_color_type'] = 'theme'
            dl_entry['font_color_value'] = dl['font_color_theme']

            if 'font_lummod' in dl or 'font_lumoff' in dl:
                lummod = dl.get('font_lummod', 100000)
                lumoff = dl.get('font_lumoff', 0)
                brightness = lummod_to_brightness(lummod, lumoff)
                dl_entry['brightness'] = round(brightness, 2)

        style['data_labels'].append(dl_entry)

    return style


def extract_and_save_style(crtx_path: str, template_path: str, output_path: str) -> Dict[str, Any]:
    """Extract styling from .crtx and template.pptx, save as YAML.

    Args:
        crtx_path: Path to .crtx file
        template_path: Path to template.pptx file
        output_path: Path to output .yaml file

    Returns:
        The generated style dictionary
    """
    # Extract chart styling from crtx
    crtx_styling = extract_crtx_styling(crtx_path)
    style = convert_to_style_yaml(crtx_styling)

    # Extract table and flowchart styling from template.pptx
    if os.path.exists(template_path):
        prs = Presentation(template_path)

        # Table style from Slide 1
        try:
            table_style = extract_table_style(prs, slide_index=1)  # Slide 1 has table
            style['table'] = table_style
        except Exception as e:
            print(f"Warning: Could not extract table style: {e}")

        # Flowchart style from Slide 2
        try:
            flowchart_style = extract_flowchart_style(prs, slide_index=2)  # Slide 2 has shapes
            style['flowchart'] = flowchart_style
        except Exception as e:
            print(f"Warning: Could not extract flowchart style: {e}")

        # Mermaid settings (defaults, not extracted)
        style['mermaid'] = {
            'direction': 'LR',
            'format': 'svg',
            'scale': 3,
            'background': 'transparent',
            'theme': {
                'primary': style['colors']['primary'],
                'primary_text': '#FFFFFF',
                'secondary': '#BFBFBF',
                'secondary_text': '#000000',
                'line_color': style['colors']['primary'],
                'text_color': '#000000',
                'font_family': 'Arial, sans-serif',
                'font_size': '14px',
            },
        }

        # Diagram settings (for native_objects.py)
        style['diagram'] = {
            'node': {
                'fill': style['colors']['primary'],
                'text_color_theme': 'lt1',
                'border_color': style['colors']['primary'],
                'border_width_pt': 1.0,
                'font_size_pt': 12,
                'font_bold': False,
            },
        }

    # Write YAML with header comment
    header = """# Generated from Chart.crtx and template.pptx
# DO NOT EDIT MANUALLY - regenerate with: python scripts/extract_style.py
#
# This file is the Single Source of Truth for styling.
# Sources:
#   - Chart.crtx: Chart styling (series, axes, legend, data labels)
#   - template.pptx Slide 1: Table styling
#   - template.pptx Slide 2: Flowchart/Diagram styling
#
# Used by: Python, R, Mermaid/JavaScript

"""

    with open(output_path, 'w') as f:
        f.write(header)
        yaml.dump(style, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

    return style


def main():
    """Main entry point for CLI usage."""
    # Default paths
    script_dir = os.path.dirname(__file__)
    base_dir = os.path.dirname(script_dir)

    default_crtx = os.path.join(base_dir, 'templates', 'template.crtx')
    default_template = os.path.join(base_dir, 'templates', 'template.pptx')
    default_output = os.path.join(base_dir, 'templates', 'style.yaml')

    # Parse arguments
    crtx_path = default_crtx
    template_path = default_template
    output_path = default_output

    # Simple argument parsing
    for i, arg in enumerate(sys.argv[1:], 1):
        if arg == '--template' and i < len(sys.argv):
            template_path = sys.argv[i + 1]
        elif arg == '--crtx' and i < len(sys.argv):
            crtx_path = sys.argv[i + 1]
        elif arg == '--output' and i < len(sys.argv):
            output_path = sys.argv[i + 1]

    # Check files exist
    if not os.path.exists(crtx_path):
        print(f"Error: Chart.crtx not found at {crtx_path}")
        sys.exit(1)

    if not os.path.exists(template_path):
        print(f"Error: template.pptx not found at {template_path}")
        sys.exit(1)

    # Extract and save
    style = extract_and_save_style(crtx_path, template_path, output_path)

    print(f"Generated {output_path}")
    print(f"  Sources:")
    print(f"    - {crtx_path}")
    print(f"    - {template_path}")
    print(f"  Contents:")
    print(f"    - {len(style['colors']['series'])} series colors")
    print(f"    - {len(style['data_labels'])} data label styles")
    if 'table' in style:
        print(f"    - Table styling (from Slide 1)")
    if 'flowchart' in style:
        print(f"    - Flowchart styling (from Slide 2)")


if __name__ == '__main__':
    main()
