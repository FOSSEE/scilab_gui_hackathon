# Results workbook

Edit `results.xlsx` to update `/results/`.

- Each visible worksheet becomes a tab, using its worksheet name.
- The first used row supplies the column headings.
- Horizontal and vertical merged cells are preserved.
- Cell values use the website's styling; Excel colors, images, and fonts are not imported.
- Formulas use the results saved by Excel; save/recalculate the workbook before building.

Run `npm run build` and redeploy to publish changes. During development, refresh
the results page after saving the workbook (restart the dev server if necessary).
