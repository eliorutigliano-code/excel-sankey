# Claude Code Prompt for Sankey Diagram Excel Add-in

Copy and paste this into Claude Code to continue development and iterate on the add-in.

---

## Initial Setup Prompt

```
I have an Excel Office Web Add-in project in the current directory that creates interactive Sankey diagrams from spreadsheet data. Here's the architecture:

**Stack**: D3.js v7 + d3-sankey, Office.js API, Webpack 5, Babel
**Key files**:
- manifest.xml — Office Add-in manifest (sideloaded via Insert tab)
- src/taskpane/sankey-engine.js — Core rendering engine (SankeyEngine class)
  - parseData(): Auto-detects Source/Target/Value vs multi-level format
  - render(): D3 Sankey layout with tooltips, hover effects
  - exportSVG() / exportPNG(): Image export
- src/taskpane/taskpane.js — Controller that reads Excel selection via Office.js
- src/taskpane/taskpane.html/css — Task pane UI with settings panel
- webpack.config.js — Bundles with HTTPS dev server on port 3000

**To run**: `npm install && npm run dev-server`, then sideload manifest.xml in Excel.

Please help me get this set up and running. Run `npm install` first, then let me know if there are any issues. After that, here are some enhancements I'd like to work on:

1. Test it end-to-end: build it, verify no errors, and walk me through sideloading
2. [ADD YOUR SPECIFIC REQUESTS HERE]
```

## Useful Follow-up Prompts

### Add features
```
Add a "Save Diagram" feature that saves the current Sankey configuration (data range reference, settings) to a named range in the workbook so users can reload it later.
```

```
Add node drag-and-drop so users can manually reposition nodes vertically in the diagram.
```

```
Support importing data from a CSV file in addition to the Excel selection, using a file picker button in the task pane.
```

### Improve the UI
```
Add a dark mode toggle that switches the task pane and diagram to a dark color scheme. Use CSS custom properties for easy theming.
```

```
Add a "fullscreen" mode that opens the Sankey diagram in a dialog window for better viewing and presentation.
```

### Data handling
```
Add data validation that highlights problematic rows in the Excel sheet (circular references, negative values, missing data) before rendering.
```

```
Add support for percentage-based flows where values represent percentages and the diagram shows proportional widths with % labels.
```

### Production readiness
```
Help me set up the project for production deployment. I want to host the add-in files on [Azure/Vercel/Netlify] and create a proper production manifest that points to the hosted URL instead of localhost.
```

```
Add unit tests for the SankeyEngine class using Jest, covering all three data format parsers and edge cases like empty cells, single-row data, and circular references.
```
