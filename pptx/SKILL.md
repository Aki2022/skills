---
name: pptx
description: "Generates PowerPoint presentations from templates with consistent styling across tables, charts, and Mermaid diagrams. Use when creating PPTX files, working with template-based presentations, or applying unified styling from style.yaml. Supports Python, R, and native PowerPoint shapes."
---

# PPTX Skill (Template-based)

## 0. Scope & Prerequisites

This skill is **template-first** and uses **style.yaml as Single Source of Truth**.

- ✅ **ALWAYS reference TEMPLATE.md for layout selection** - different layouts have different placeholder indices
- ✅ Always begin with `template.pptx`
- ✅ Extract styles from `Chart.crtx` and `template.pptx` into `style.yaml`
- ✅ Use consistent styling across Python, R, and Mermaid

---

## 1. Working Directory Structure

To keep the skill directory clean, all working files should be placed in a separate project directory:

```
{project}/
├── outline.md           # Content definition (input)
└── powerpoint/          # All PowerPoint outputs (auto-created)
    ├── output.pptx      # Final output
    └── processing/      # Intermediate files (auto-created)
        ├── style.yaml   # Project-specific styles (auto-copied on first run)
        ├── pptx_generation.log  # Debug logs (auto-generated)
        ├── charts/      # R-generated SVG/PNG (optional)
        ├── diagrams/    # Mermaid-generated SVG (optional)
        └── temp/        # Other temporary files (optional)
```

### Directory Roles

- **outline.md** - Markdown file defining slide content and structure (in project root)
- **powerpoint/** - All PowerPoint-related outputs (auto-created)
  - **output.pptx** - Final generated PowerPoint presentation
  - **processing/** - All intermediate/temporary files (can be deleted after completion)
    - **style.yaml** - Project-specific style configuration (auto-copied from skill templates on first run)
    - **pptx_generation.log** - Detailed debug and error logs (auto-generated)

### Setup

No manual setup required! The `powerpoint/processing/` directory and `style.yaml` are automatically created on first run.

Optional: Create subdirectories for R charts or Mermaid diagrams if needed:
```bash
mkdir -p powerpoint/processing/{charts,diagrams,temp}
```

### Logging (Automatic)

All PPTX generation activities are automatically logged to `powerpoint/processing/pptx_generation.log`:

- **Auto-detection**: Finds `powerpoint/processing/` directory automatically
- **Console**: Shows warnings/errors only
- **Log file**: Records all debug information, validation errors, and styling issues

No manual setup required - logging initializes on first use of table/chart creation functions.

---

## 2. Files and Roles

### Core Files

- **templates/template.pptx** - Slide layouts, theme colors/fonts (human-edited, shared across projects)
- **templates/template.crtx** - Chart template with styling (human-edited, shared across projects)
- **templates/style.yaml** - Master style definitions (auto-generated from templates, shared across projects)
- **{project}/powerpoint/processing/style.yaml** - Project-specific style (auto-copied from templates/style.yaml)

### Scripts

- **scripts/extract_style.py** - Generate style.yaml from templates
- **scripts/style_config.py** - Python style loader
- **scripts/style_config.R** - R style loader
- **scripts/mermaid_to_shapes.py** - Mermaid → native PowerPoint shapes
- **scripts/native_objects.py** - Native table/chart/diagram creation (with validation & logging)
- **scripts/crtx_utils.py** - Chart.crtx utilities (with detailed error logging)
- **scripts/logging_utils.py** - Auto-configured logging to processing/pptx_generation.log
- **scripts/layout_registry.py** - Layout management
- **scripts/generate_template.py** - TEMPLATE.md auto-generation (maintenance tool)

---

## 3. Style System

### Generate Master Style (templates/style.yaml)

When you update `template.pptx` or `template.crtx`, regenerate the master style:

```bash
cd ~/.claude/skills/pptx
python scripts/extract_style.py
```

This extracts styling from:

- `templates/template.crtx` - Series colors, axes, legend, data labels
- `templates/template.pptx` Slide 1 - Table styling
- `templates/template.pptx` Slide 2 - Flowchart/diagram styling

Output: `templates/style.yaml` (master template)

### Project-Specific Style

Each project gets its own copy of `style.yaml` in `powerpoint/processing/`:

- **Auto-setup**: On first run, `generate_presentation.py` copies `templates/style.yaml` to `powerpoint/processing/style.yaml`
- **Customization**: You can edit `powerpoint/processing/style.yaml` for project-specific styling
- **Fallback**: If `powerpoint/processing/style.yaml` doesn't exist, the system uses `templates/style.yaml`

### style.yaml Structure

```yaml
colors:
  primary: "#4F4F70"
  series:
    - type: rgb
      value: "#4F4F70"
    - type: theme
      value: bg1
      brightness: -0.25

category_axis:
  visible: true
  font:
    size_pt: 11
    color_type: theme
    color_value: tx1
    brightness: 0.35

value_axis:
  visible: false

legend:
  position: bottom
  font:
    size_pt: 11

table:
  header:
    fill_theme: bg1
    fill_brightness: -0.5
  body:
    column_brightness: [-0.15, -0.05, -0.05, -0.05]

flowchart:
  node:
    fill: "#4F4F70"
    shadow:
      enabled: false
  connector:
    type: elbow
    dash_style: solid
```

---

## 4. Usage

### Creating Presentations (REQUIRED)

**IMPORTANT**: Always use `native_objects.py` for creating tables and charts. This ensures:

- Complete styling from `.crtx` template is applied
- Automatic data validation
- Detailed error logging to `powerpoint/processing/pptx_generation.log`

**CRITICAL**: Different layouts have different placeholder indices. Always check TEMPLATE.md or use the debug script to find the correct idx for your layout.

```python
import sys
import os
sys.path.insert(0, os.path.expanduser('~/.claude/skills/pptx'))

from pptx import Presentation
from scripts.native_objects import create_styled_table, create_styled_chart

# Load template
skill_dir = os.path.expanduser('~/.claude/skills/pptx')
template_path = os.path.join(skill_dir, 'templates', 'template.pptx')
prs = Presentation(template_path)

# Example 1: Create a chart slide (use Layout 5)
slide = prs.slides.add_slide(prs.slide_layouts[5])  # Handout_Single_Chart_Pos
slide.shapes.title.text = "Sales Report"
slide.placeholders[13].text = "Q1-Q4 performance analysis"
# Chart placeholder is idx=15 for this layout
chart_spec = {
    'chart_kind': 'column',  # 'line', 'bar', 'pie'
    'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
    'series': [
        {'name': 'Sales', 'values': [100, 120, 110, 130]},
        {'name': 'Cost', 'values': [80, 90, 85, 95]}
    ]
}
create_styled_chart(slide, slide.placeholders[15], chart_spec)

# Example 2: Create a table slide (use Layout 7)
slide = prs.slides.add_slide(prs.slide_layouts[7])  # Handout_Single_Table_Pos
slide.shapes.title.text = "Summary Data"
slide.placeholders[13].text = "Key metrics overview"
# Table placeholder is idx=16 for this layout
table_spec = {
    'data': [
        ['項目', '値A', '値B'],
        ['データ1', '100', '200'],
        ['データ2', '150', '250']
    ],
    'header_row': True
}
create_styled_table(slide, slide.placeholders[16], table_spec)

prs.save('powerpoint/output.pptx')
```

### Reading Styles (Advanced)

For custom styling beyond native objects, use `StyleConfig`:

```python
from scripts.style_config import StyleConfig

# Auto-detects: powerpoint/processing/style.yaml -> processing/style.yaml (legacy) -> templates/style.yaml
style = StyleConfig.load()
primary = style.colors['primary']  # '#4F4F70'
table_config = style.table

# Or specify path explicitly:
style = StyleConfig.load('powerpoint/processing/style.yaml')
```

**StyleConfig.load() search order:**
1. `powerpoint/processing/style.yaml` (project-specific, recommended)
2. `processing/style.yaml` (legacy, for backward compatibility)
3. `~/.claude/skills/pptx/templates/style.yaml` (master)

**WARNING**: Using `StyleConfig` directly requires manual application of all styles. Prefer `native_objects.py` instead.

### R

```r
source("scripts/style_config.R")
style <- load_style("style.yaml")
colors <- get_series_colors(style, 3)
```

### Mermaid → Native Shapes

```python
from scripts.mermaid_to_shapes import create_flowchart_shapes

code = """flowchart LR
    A[Start] --> B{Decision}
    B -->|Yes| C[Action]
    B -->|No| D[End]"""

create_flowchart_shapes(slide, placeholder, code)
```

---

## 5. Render Modes

| Type    | Mode   | Description                 | Editable |
| ------- | ------ | --------------------------- | -------- |
| TABLE   | NATIVE | python-pptx table           | ✅       |
| CHART   | NATIVE | python-pptx with Chart.crtx | ✅       |
| DIAGRAM | NATIVE | Mermaid → native shapes     | ✅       |

**Note:** All rendering uses NATIVE mode for maximum editability in PowerPoint.

---

## 6. Template Layouts

**IMPORTANT**: For complete layout reference, see **[TEMPLATE.md](TEMPLATE.md)**.

TEMPLATE.md provides:
- All 124 available layouts with detailed descriptions
- Naming convention: `{Usage}_{Layout}_{Content}_{Variant}`
- Selection guidelines for Handout vs Preso layouts
- AI guidelines for outline.md creation

### Quick Reference

**Foundation Layouts**:
- `0`: `00_Title` - Opening slide
- `1`: `01_Contents` - Table of contents
- `2`: `02_Section` - Section divider

**Common Layouts**:
- `0`: `00_Title` - Title slide
- `5`: `Handout_Single_Chart_Pos` - Full-width chart with key message
- `7`: `Handout_Single_Table_Pos` - Full-width table with key message
- `11`: `Handout_Single_Object_Pos` - Full-width object (for Mermaid diagrams)
- `66`: `Preso_Single_Chart_Pos` - Presentation mode chart

**Key Placeholder Indices** (vary by layout - check TEMPLATE.md):
- `idx=0`: TITLE (most layouts)
- `idx=13`: KeyMessage (most content layouts)
- `idx=15`: CHART (chart layouts like 5, 6)
- `idx=16`: TABLE (table layouts like 7, 8)
- `idx=1`: OBJECT (object layouts like 11, 12)

**IMPORTANT**: Always reference TEMPLATE.md for exact placeholder indices for each layout.

---

## 7. Dependencies

### Python

```bash
pip install python-pptx lxml pyyaml pillow
```

### R

```r
install.packages(c("ggplot2", "yaml", "dplyr", "tidyr"))
```

### Mermaid (optional)

```bash
npm install -g @mermaid-js/mermaid-cli
```

---

## 8. Workflow Example

### Complete Example

```python
#!/usr/bin/env python3
import sys
import os
sys.path.insert(0, os.path.expanduser('~/.claude/skills/pptx'))

from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from scripts.native_objects import create_styled_table, create_styled_chart

# Setup (run once in project directory)
# mkdir -p powerpoint/processing/{charts,diagrams,temp}
# cp ~/.claude/skills/pptx/templates/template.pptx powerpoint/

# Load template
prs = Presentation('template.pptx')

# Delete existing slides
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[0]

# Slide 1: Title slide
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Presentation Title"
slide.placeholders[1].text = "Subtitle\nDate"

# Slide 2: Chart slide (use Layout 5 for charts)
slide = prs.slides.add_slide(prs.slide_layouts[5])  # Handout_Single_Chart_Pos
slide.shapes.title.text = "Chart Example"
slide.placeholders[13].text = "Key message about this chart"
# Chart placeholder is idx=15
chart_spec = {
    'chart_kind': 'column',
    'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
    'series': [
        {'name': 'Sales', 'values': [100, 120, 110, 130]},
        {'name': 'Cost', 'values': [80, 90, 85, 95]}
    ]
}
create_styled_chart(slide, slide.placeholders[15], chart_spec)

# Slide 3: Table slide (use Layout 7 for tables)
slide = prs.slides.add_slide(prs.slide_layouts[7])  # Handout_Single_Table_Pos
slide.shapes.title.text = "Table Example"
slide.placeholders[13].text = "Summary statistics"
# Table placeholder is idx=16
table_spec = {
    'data': [
        ['Item', 'Value A', 'Value B', 'Total'],
        ['Product 1', '100', '200', '300'],
        ['Product 2', '150', '250', '400']
    ],
    'header_row': True
}
create_styled_table(slide, slide.placeholders[16], table_spec)

# Save
prs.save('powerpoint/output.pptx')
print("✅ Presentation created: powerpoint/output.pptx")
print("📋 Check logs: cat powerpoint/processing/pptx_generation.log")
```

### R Charts (Advanced)

For complex ggplot2 charts, use R with style.yaml:

```r
source("~/.claude/skills/pptx/scripts/style_config.R")
style <- load_style("powerpoint/processing/style.yaml")

p <- ggplot(data, aes(x, y)) +
  geom_bar(fill = get_primary_color(style)) +
  theme_style(style)

# Save as PNG and insert manually into PowerPoint
ggsave("powerpoint/processing/charts/chart.png", p, width = 10, height = 6, dpi = 300)
```

---

## 9. Troubleshooting

### Check Logs

If tables or charts fail to generate correctly:

```bash
cat powerpoint/processing/pptx_generation.log
```

### Common Issues

**Table creation fails**

- Log shows: `Row X has Y columns, expected Z` → Check data array consistency
- Log shows: `Table spec.data is empty` → Verify data is not empty

**Chart creation fails**

- Log shows: `Series 'X' contains non-numeric value` → All chart values must be numbers
- Log shows: `Chart.crtx not found` → Template path issue (auto-fixed in latest version)
- Log shows: `Unknown theme color 'accentX'` → Check style.yaml theme color definitions

**Styling not applied**

- Log shows: `Failed to apply category axis styling` → Check template.crtx compatibility
- Console shows warnings → Check `powerpoint/processing/pptx_generation.log` for details

### Error Prevention

All input data is now validated:

- Table: Column count consistency, non-empty data
- Chart: Numeric values, matching series/category lengths, non-empty series
- Template paths use absolute paths (no longer dependent on working directory)
