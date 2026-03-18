/**
 * Excel Sankey Diagram Add-in — Task Pane Controller
 */
import { SankeyEngine } from "./sankey-engine";
import "./taskpane.css";

/* global Office, Excel */

// Color scheme presets
const COLOR_SCHEMES = {
  default: [
    "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f",
    "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac",
  ],
  tableau: [
    "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f",
    "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac",
  ],
  category: [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
  ],
  warm: [
    "#e6550d", "#fd8d3c", "#fdae6b", "#fdd0a2", "#d94701",
    "#a63603", "#feedde", "#f16913", "#e6550d", "#8c2d04",
  ],
  cool: [
    "#3182bd", "#6baed6", "#9ecae1", "#c6dbef", "#2171b5",
    "#08519c", "#deebf7", "#4292c6", "#2171b5", "#084594",
  ],
};

let sankeyEngine = null;
let currentData = null;

// ==========================================
// Office Initialization
// ==========================================
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeUI();
  }
});

function initializeUI() {
  // Read data button
  document.getElementById("btn-read-data").addEventListener("click", readSelectedData);

  // Settings toggle
  document.getElementById("toggle-settings").addEventListener("click", () => {
    const panel = document.getElementById("settings-panel");
    const chevron = document.querySelector(".chevron");
    panel.classList.toggle("collapsed");
    chevron.classList.toggle("open");
  });

  // Settings change handlers
  document.getElementById("opt-title").addEventListener("input", onSettingsChange);
  document.getElementById("opt-alignment").addEventListener("change", onSettingsChange);
  document.getElementById("opt-color-scheme").addEventListener("change", onSettingsChange);
  document.getElementById("opt-number-format").addEventListener("change", onSettingsChange);
  document.getElementById("opt-show-values").addEventListener("change", onSettingsChange);

  const rangeInputs = ["opt-node-width", "opt-node-padding", "opt-link-opacity"];
  rangeInputs.forEach((id) => {
    const input = document.getElementById(id);
    input.addEventListener("input", () => {
      const valSpan = document.getElementById(id + "-val");
      if (id === "opt-link-opacity") {
        valSpan.textContent = (input.value / 100).toFixed(2);
      } else {
        valSpan.textContent = input.value;
      }
    });
    input.addEventListener("change", onSettingsChange);
  });

  // Export buttons
  document.getElementById("btn-export-svg").addEventListener("click", exportSVG);
  document.getElementById("btn-export-png").addEventListener("click", exportPNG);
  document.getElementById("btn-insert-image").addEventListener("click", insertImageInSheet);

  // Window resize
  window.addEventListener("resize", () => {
    if (sankeyEngine) sankeyEngine.resize();
  });
}

// ==========================================
// Read data from Excel selection
// ==========================================
async function readSelectedData() {
  try {
    showStatus("Reading selected data...", "info");

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values, rowCount, columnCount, address");
      await context.sync();

      const values = range.values;
      const rowCount = range.rowCount;
      const colCount = range.columnCount;

      if (rowCount < 2 || colCount < 2) {
        showStatus("Please select at least 2 columns and 2 rows (header + data).", "error");
        return;
      }

      currentData = values;

      // Initialize or update engine
      const container = document.getElementById("sankey-container");
      container.innerHTML = ""; // clear placeholder

      const containerRect = container.getBoundingClientRect();

      const options = getCurrentOptions();
      options.width = Math.max(containerRect.width, 400);
      options.height = Math.max(350, Math.min(600, rowCount * 20));

      sankeyEngine = new SankeyEngine(container, options);

      const parsedData = sankeyEngine.parseData(values);

      if (parsedData.nodes.length === 0 || parsedData.links.length === 0) {
        showStatus("No valid flow data found. Check your data format.", "error");
        return;
      }

      sankeyEngine.render(parsedData);

      // Update UI
      const dataSummary = document.getElementById("data-summary");
      dataSummary.textContent = `${parsedData.nodes.length} nodes, ${parsedData.links.length} flows from ${range.address}`;
      document.getElementById("data-preview").classList.remove("hidden");
      document.getElementById("export-section").style.display = "";

      showStatus(`Diagram created from ${range.address}`, "success");
    });
  } catch (error) {
    console.error("Error reading data:", error);
    showStatus(`Error: ${error.message}`, "error");
  }
}

// ==========================================
// Settings
// ==========================================
function getCurrentOptions() {
  return {
    title: document.getElementById("opt-title").value.trim(),
    alignment: document.getElementById("opt-alignment").value,
    nodeWidth: parseInt(document.getElementById("opt-node-width").value),
    nodePadding: parseInt(document.getElementById("opt-node-padding").value),
    linkOpacity: parseInt(document.getElementById("opt-link-opacity").value) / 100,
    showValues: document.getElementById("opt-show-values").checked,
    numberFormat: document.getElementById("opt-number-format").value,
    colorScheme: COLOR_SCHEMES[document.getElementById("opt-color-scheme").value] || COLOR_SCHEMES.default,
  };
}

function onSettingsChange() {
  if (!sankeyEngine || !currentData) return;

  const options = getCurrentOptions();
  sankeyEngine.update(options);
  const parsedData = sankeyEngine.parseData(currentData);
  sankeyEngine.render(parsedData);
}

// ==========================================
// Export functions
// ==========================================
function exportSVG() {
  if (!sankeyEngine) return;

  const svgString = sankeyEngine.exportSVG();
  if (!svgString) return;

  downloadFile(svgString, "sankey-diagram.svg", "image/svg+xml");
  showStatus("SVG exported", "success");
}

async function exportPNG() {
  if (!sankeyEngine) return;

  try {
    showStatus("Generating PNG...", "info");
    const dataUrl = await sankeyEngine.exportPNG(2);
    const link = document.createElement("a");
    link.download = "sankey-diagram.png";
    link.href = dataUrl;
    link.click();
    showStatus("PNG exported", "success");
  } catch (err) {
    showStatus(`PNG export failed: ${err.message}`, "error");
  }
}

async function insertImageInSheet() {
  if (!sankeyEngine) return;

  try {
    showStatus("Inserting image into sheet...", "info");

    const dataUrl = await sankeyEngine.exportPNG(2);
    const base64 = dataUrl.split(",")[1];

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();

      const image = sheet.shapes.addImage(base64);
      image.name = "SankeyDiagram";
      image.left = 0;
      image.top = 0;

      await context.sync();
      showStatus("Image inserted into the worksheet", "success");
    });
  } catch (err) {
    showStatus(`Insert failed: ${err.message}`, "error");
  }
}

// ==========================================
// Utilities
// ==========================================
function showStatus(message, type = "info") {
  const area = document.getElementById("status-area");
  const msg = document.getElementById("status-message");
  const icon = document.getElementById("status-icon");

  area.className = `status-area ${type}`;
  msg.textContent = message;

  const icons = { success: "\u2713", error: "\u2717", info: "\u2139" };
  icon.textContent = icons[type] || "";

  area.classList.remove("hidden");

  if (type === "success" || type === "info") {
    setTimeout(() => {
      area.classList.add("hidden");
    }, 4000);
  }
}

function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
