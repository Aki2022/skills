---
name: origin-quarto
description: Create, modify, or render any Quarto document or project, including `.qmd` files, `_quarto.yml`, RevealJS slides, and PowerPoint output. Always use the bundled organization template and extensions when creating Quarto or QMD content.
---

# Origin Quarto

Create Quarto deliverables from the bundled template. Keep the template files together so its RevealJS plugins, SCSS, citation style, and PowerPoint reference document remain available.

## Template

Use `assets/template/` as the source of truth. It contains:

- `qmd_template.qmd` — default document and slide template.
- `_extensions/` — required Quarto extensions, styling, CSL, and PowerPoint reference document.
- `marp_template.md` — related Marp template; use only for an explicitly requested Marp deliverable.

## Workflow

1. Inspect the requested deliverable, audience, output format, and destination. Ask only if the missing choice materially changes the output.
2. Before creating a new Quarto document, copy the contents of `assets/template/` into the destination project directory. Do not overwrite existing user files; merge or ask for direction when a filename conflicts.
3. Start the document from `qmd_template.qmd`, rename it for the deliverable, and replace all placeholder metadata and content. Preserve the relative `_extensions/` paths unless deliberately changing the design.
4. Keep YAML valid and retain the relevant output configuration. For RevealJS and PowerPoint output, preserve the included extensions and PowerPoint reference document unless the user requests a different visual system.
5. Render the requested format with Quarto when the environment and dependencies are available. Inspect render errors and correct the source rather than suppressing failures.
6. Report the created source and rendered outputs, plus any dependency that prevented rendering.

## Editing existing work

When editing an existing Quarto project, inspect its current configuration first. Reuse the bundled template only when creating new material or when the user asks to apply this design; do not replace an established project style without authorization.

## Validation

- Confirm that every local extension, theme, citation, and reference-document path resolves from the destination project.
- Render at least the requested target format when feasible.
- For slide output, visually inspect the rendered result when layout matters.
