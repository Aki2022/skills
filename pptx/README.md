# PPTX Style System

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE.txt)

A template-based PowerPoint presentation styling system that automates consistent design across Python, R, and Mermaid diagrams.

[日本語版 README はこちら](README_JPN.md)

## Features

- **Single Source of Truth**: All styling defined in `style.yaml`, automatically extracted from templates
- **Template-First Approach**: Edit `template.pptx` and `Chart.crtx` visually, extract styles programmatically
- **Multi-Language Support**: Consistent styling across Python, R, and Mermaid
- **Native Editable Objects**: Tables, charts, and Mermaid diagrams rendered as native PowerPoint shapes
- **Automatic Validation**: Built-in data validation and detailed error logging
- **Auto-Setup**: Project directories and style configurations created automatically on first run

## Quick Start

### Installation

```bash
# Python dependencies
pip install python-pptx lxml pyyaml pillow

# R dependencies (optional)
R -e "install.packages(c('ggplot2', 'yaml', 'dplyr', 'tidyr'))"

# Mermaid CLI (optional, for diagram generation)
npm install -g @mermaid-js/mermaid-cli
```

### Basic Usage

```python
import sys
sys.path.insert(0, '~/.claude/skills/pptx')

from pptx import Presentation
from scripts.native_objects import create_styled_table, create_styled_chart

# Load template
prs = Presentation('templates/template.pptx')
slide = prs.slides.add_slide(prs.slide_layouts[10])

# Find content placeholder
for shape in slide.shapes:
    if shape.is_placeholder and shape.placeholder_format.idx == 1:
        # Create table
        table_spec = {
            'data': [
                ['Item', 'Value A', 'Value B'],
                ['Data 1', '100', '200'],
                ['Data 2', '150', '250']
            ],
            'header_row': True
        }
        create_styled_table(slide, shape, table_spec)

prs.save('output.pptx')
```

## Project Structure

```
├── README.md              # This file
├── README_JPN.md          # Japanese documentation
├── LICENSE.txt            # MIT License
├── SKILL.md               # Skill definition (for Claude Code)
├── style.yaml             # Auto-generated style definitions
├── templates/
│   ├── template.pptx      # Slide template (human-edited)
│   └── template.crtx      # Chart template (human-edited)
├── scripts/
│   ├── extract_style.py         # Extract styles from templates
│   ├── style_config.py          # Python style loader
│   ├── style_config.R           # R style loader
│   ├── native_objects.py        # Create tables/charts/diagrams
│   ├── crtx_utils.py            # Chart template utilities
│   ├── mermaid_to_shapes.py     # Mermaid → native shapes
│   ├── logging_utils.py         # Auto-configured logging
│   ├── layout_registry.py       # Layout management
│   └── generate_template.py     # TEMPLATE.md auto-generation
└── examples/                    # Usage examples
```

## Workflow

### 1. Extract Styles

Extract styling from templates to create `style.yaml`:

```bash
python scripts/extract_style.py
```

This reads:
- `templates/template.crtx` - Chart styling (series colors, axes, legend, data labels)
- `templates/template.pptx` Slide 33 - Table styling
- `templates/template.pptx` Slide 34 - Flowchart/diagram styling

### 2. Create Presentations

Always use `native_objects.py` for tables and charts - it applies complete styling and validates data:

```python
from scripts.native_objects import create_styled_table, create_styled_chart

# Create styled chart
chart_spec = {
    'chart_kind': 'column',  # 'line', 'bar', 'pie'
    'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
    'series': [
        {'name': 'Sales', 'values': [100, 120, 110, 130]},
        {'name': 'Cost', 'values': [80, 90, 85, 95]}
    ]
}
create_styled_chart(slide, placeholder, chart_spec)
```

### 3. Use in R

```r
source("scripts/style_config.R")
style <- load_style("style.yaml")

# Get colors for ggplot2
colors <- get_series_colors(style, 3)

# Create plot with consistent styling
p <- ggplot(data, aes(x, y)) +
  geom_bar(fill = get_primary_color(style)) +
  theme_minimal()
```

### 4. Mermaid Diagrams

```python
from scripts.mermaid_to_shapes import create_flowchart_shapes

mermaid_code = """flowchart LR
    A[Start] --> B{Decision}
    B -->|Yes| C[Action]
    B -->|No| D[End]"""

create_flowchart_shapes(slide, placeholder, mermaid_code)
```

## Advanced Features

### R Charts

For complex ggplot2 charts, use R with style.yaml for consistent styling:

```r
source("scripts/style_config.R")
style <- load_style("style.yaml")

p <- ggplot(data, aes(x, y)) +
  geom_bar(fill = get_primary_color(style)) +
  theme_minimal()

# Save as PNG and insert into PowerPoint
ggsave("chart.png", p, width = 10, height = 6, dpi = 300)
```

### Logging

All operations are automatically logged to `processing/pptx_generation.log`:

```bash
# Check logs for errors or warnings
cat processing/pptx_generation.log
```

### Template Layouts

Key layouts in `template.pptx`:

| Index | Name | Use |
|-------|------|-----|
| 0 | Title Slide | Title page |
| 8 | Section Header | Section dividers |
| 10 | Title and Content_withKeyMessage | Main content |
| 19 | Content with Caption_withKeyMessage | Chart + description |

## Troubleshooting

### Chart Creation Fails

Check logs:
```bash
cat processing/pptx_generation.log
```

Common issues:
- **"Series 'X' contains non-numeric value"** → All chart values must be numbers
- **"Chart.crtx not found"** → Template path issue (check templates/ directory)
- **"Unknown theme color"** → Check style.yaml theme color definitions

### Table Creation Fails

- **"Row X has Y columns, expected Z"** → Check data array consistency
- **"Table spec.data is empty"** → Verify data is not empty

### Styling Not Applied

- **"Failed to apply category axis styling"** → Check template.crtx compatibility
- Console shows warnings → Check `processing/pptx_generation.log` for details

## Security

This project has been reviewed for security:

- ✅ No hardcoded credentials or secrets
- ✅ Safe YAML parsing (`yaml.safe_load()`)
- ✅ Subprocess calls use argument lists (no shell injection risk)
- ✅ No arbitrary code execution
- ✅ Safe XML/OOXML parsing with lxml
- ✅ Validated file operations

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - see [LICENSE.txt](LICENSE.txt) for details.

## Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/)
- Uses [Mermaid](https://mermaid.js.org/) for diagram generation
- Designed for use with [Claude Code](https://claude.com/claude-code)
