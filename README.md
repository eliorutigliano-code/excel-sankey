# Excel Sankey Diagram Add-in

An Office Web Add-in that creates interactive Sankey diagrams from your Excel data. Supports simple Source→Target→Value tables and multi-level flow data with automatic format detection.

## Features

- **Auto-detect data format**: Works with 3-column (Source/Target/Value) and multi-level (Stage1/Stage2/.../Value) layouts
- **Interactive**: Hover tooltips showing flow values and node totals
- **Customizable**: Alignment, colors, opacity, spacing, and label options
- **Export**: SVG, PNG, or insert directly as an image in your worksheet
- **5 color schemes**: Default, Tableau 10, D3 Category, Warm, Cool

## Prerequisites

- **Node.js** >= 18 (https://nodejs.org)
- **Excel** desktop (Windows or Mac) or Excel on the web

## Quick Start

```bash
# 1. Install dependencies
npm install

# 2. Install dev certificates (required for HTTPS)
npx office-addin-dev-certs install

# 3. Start the dev server
npm run dev-server

# 4. Sideload the add-in in Excel
#    Option A: Use the Office sideloading tool
npm run start
#    Option B: Manual sideload (see below)
```

### Manual Sideload (Excel Desktop - Windows)

1. Open Excel
2. Go to **Insert** → **My Add-ins** → **Upload My Add-in**
3. Browse to `manifest.xml` in this project folder
4. Click **Upload**

### Manual Sideload (Excel on the Web)

1. Open Excel Online
2. Go to **Insert** → **Office Add-ins** → **Upload My Add-in**
3. Browse to `manifest.xml`
4. Click **Upload**

## How to Use

### Data Format A: Source → Target → Value

| Source     | Target     | Value |
|------------|------------|-------|
| Revenue    | Product    | 500   |
| Revenue    | Services   | 300   |
| Product    | Online     | 350   |
| Product    | Retail     | 150   |
| Services   | Consulting | 200   |
| Services   | Support    | 100   |

### Data Format B: Multi-Level Stages → Value

| Region    | Department | Category  | Amount |
|-----------|------------|-----------|--------|
| North     | Sales      | Hardware  | 120    |
| North     | Sales      | Software  | 80     |
| South     | Marketing  | Digital   | 95     |
| South     | Sales      | Hardware  | 60     |

### Steps

1. Enter your flow data in a spreadsheet (including headers)
2. Select the entire data range (headers + data)
3. Click the **Sankey Diagram** button in the Insert tab ribbon
4. Click **Read Selection** in the task pane
5. Adjust settings as needed
6. Export or insert into your sheet

## Development

```bash
# Build for production
npm run build

# Validate manifest
npm run validate
```

## Project Structure

```
excel-sankey-diagram/
├── manifest.xml              # Office Add-in manifest
├── package.json
├── webpack.config.js
├── .babelrc
├── assets/                   # Icons (generate your own or use placeholders)
└── src/
    ├── taskpane/
    │   ├── taskpane.html     # Task pane UI
    │   ├── taskpane.css      # Styles
    │   ├── taskpane.js       # Controller (reads Excel, manages settings)
    │   └── sankey-engine.js  # D3 Sankey rendering engine
    └── commands/
        └── commands.html     # Ribbon command handler
```

## Claude Code Iteration Prompt

Use this prompt in Claude Code to iterate on the add-in:

> I have an Excel Office Web Add-in project for Sankey diagrams in the current directory. The project uses D3.js + d3-sankey for rendering, webpack for bundling, and the Office.js API for Excel integration. The main files are: manifest.xml, src/taskpane/taskpane.js (controller), src/taskpane/sankey-engine.js (rendering engine), and webpack.config.js. Please help me [YOUR REQUEST HERE — e.g., "add drag-to-reorder nodes", "support CSV import alongside Excel selection", "add a dark mode theme", etc.].

## License

MIT
