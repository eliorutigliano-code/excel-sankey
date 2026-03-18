/**
 * Sankey Diagram Rendering Engine
 * Uses D3.js + d3-sankey to render interactive Sankey diagrams
 */
import * as d3 from "d3";
import { sankey as d3Sankey, sankeyLinkHorizontal, sankeyCenter, sankeyLeft, sankeyRight, sankeyJustify } from "d3-sankey";

const ALIGNMENT_MAP = {
  center: sankeyCenter,
  left: sankeyLeft,
  right: sankeyRight,
  justify: sankeyJustify,
};

const DEFAULT_COLORS = [
  "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f",
  "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac",
  "#86bcb6", "#8cd17d", "#b6992d", "#499894", "#e15759",
  "#f1ce63", "#a0cbe8", "#d37295", "#fabfd2", "#d4a6c8",
];

export class SankeyEngine {
  constructor(container, options = {}) {
    this.container = typeof container === "string" ? document.querySelector(container) : container;
    this.options = {
      width: options.width || 700,
      height: options.height || 450,
      nodeWidth: options.nodeWidth || 20,
      nodePadding: options.nodePadding || 16,
      alignment: options.alignment || "justify",
      colorScheme: options.colorScheme || DEFAULT_COLORS,
      linkOpacity: options.linkOpacity || 0.4,
      linkHoverOpacity: options.linkHoverOpacity || 0.7,
      fontSize: options.fontSize || 12,
      margin: options.margin || { top: 10, right: 120, bottom: 10, left: 10 },
      numberFormat: options.numberFormat || ",.0f",
      showValues: options.showValues !== undefined ? options.showValues : true,
      showTooltip: options.showTooltip !== undefined ? options.showTooltip : true,
    };

    this.svg = null;
    this.tooltip = null;
    this.data = null;
  }

  /**
   * Parse raw spreadsheet data into Sankey-compatible format
   * Auto-detects data structure:
   *   - 3 columns: Source, Target, Value
   *   - N columns: Multi-level flow (each row traces a path through stages)
   */
  parseData(rawData) {
    if (!rawData || rawData.length < 2) {
      throw new Error("Need at least a header row and one data row.");
    }

    const headers = rawData[0];
    const rows = rawData.slice(1).filter((row) => row.some((cell) => cell !== null && cell !== ""));

    // Detect format: if last column is numeric, treat as Source/Target/Value or multi-level with value
    const lastColValues = rows.map((r) => r[r.length - 1]);
    const lastColNumeric = lastColValues.every((v) => v !== null && v !== "" && !isNaN(Number(v)));

    if (headers.length === 3 && lastColNumeric) {
      return this._parseSourceTargetValue(headers, rows);
    } else if (headers.length >= 3 && lastColNumeric) {
      return this._parseMultiLevel(headers, rows);
    } else if (headers.length === 3) {
      // Try treating 3rd column as value anyway
      return this._parseSourceTargetValue(headers, rows);
    } else {
      // Multi-level without explicit value column — count occurrences
      return this._parseMultiLevelCount(headers, rows);
    }
  }

  _parseSourceTargetValue(headers, rows) {
    const linksMap = new Map();
    const nodeSet = new Set();

    for (const row of rows) {
      const source = String(row[0]).trim();
      const target = String(row[1]).trim();
      const value = parseFloat(row[2]) || 0;

      if (!source || !target || value <= 0) continue;

      nodeSet.add(source);
      nodeSet.add(target);

      const key = `${source}|||${target}`;
      linksMap.set(key, (linksMap.get(key) || 0) + value);
    }

    const nodes = Array.from(nodeSet).map((name) => ({ name }));
    const nodeIndex = new Map(nodes.map((n, i) => [n.name, i]));

    const links = Array.from(linksMap.entries()).map(([key, value]) => {
      const [source, target] = key.split("|||");
      return { source: nodeIndex.get(source), target: nodeIndex.get(target), value };
    });

    return { nodes, links };
  }

  _parseMultiLevel(headers, rows) {
    const stageCount = headers.length - 1; // last col is value
    const linksMap = new Map();
    const nodeSet = new Set();

    for (const row of rows) {
      const value = parseFloat(row[row.length - 1]) || 0;
      if (value <= 0) continue;

      for (let i = 0; i < stageCount - 1; i++) {
        const source = `${headers[i]}: ${String(row[i]).trim()}`;
        const target = `${headers[i + 1]}: ${String(row[i + 1]).trim()}`;

        if (!row[i] || !row[i + 1]) continue;

        nodeSet.add(source);
        nodeSet.add(target);

        const key = `${source}|||${target}`;
        linksMap.set(key, (linksMap.get(key) || 0) + value);
      }
    }

    const nodes = Array.from(nodeSet).map((name) => ({ name }));
    const nodeIndex = new Map(nodes.map((n, i) => [n.name, i]));

    const links = Array.from(linksMap.entries()).map(([key, value]) => {
      const [source, target] = key.split("|||");
      return { source: nodeIndex.get(source), target: nodeIndex.get(target), value };
    });

    return { nodes, links };
  }

  _parseMultiLevelCount(headers, rows) {
    const linksMap = new Map();
    const nodeSet = new Set();

    for (const row of rows) {
      for (let i = 0; i < headers.length - 1; i++) {
        const source = `${headers[i]}: ${String(row[i]).trim()}`;
        const target = `${headers[i + 1]}: ${String(row[i + 1]).trim()}`;

        if (!row[i] || !row[i + 1]) continue;

        nodeSet.add(source);
        nodeSet.add(target);

        const key = `${source}|||${target}`;
        linksMap.set(key, (linksMap.get(key) || 0) + 1);
      }
    }

    const nodes = Array.from(nodeSet).map((name) => ({ name }));
    const nodeIndex = new Map(nodes.map((n, i) => [n.name, i]));

    const links = Array.from(linksMap.entries()).map(([key, value]) => {
      const [source, target] = key.split("|||");
      return { source: nodeIndex.get(source), target: nodeIndex.get(target), value };
    });

    return { nodes, links };
  }

  /**
   * Render the Sankey diagram
   */
  render(parsedData) {
    this.data = parsedData;
    const { width, height, margin, nodeWidth, nodePadding, alignment, colorScheme } = this.options;

    const innerWidth = width - margin.left - margin.right;
    const innerHeight = height - margin.top - margin.bottom;

    // Clear previous render
    d3.select(this.container).selectAll("*").remove();

    // Create SVG
    this.svg = d3
      .select(this.container)
      .append("svg")
      .attr("viewBox", `0 0 ${width} ${height}`)
      .attr("width", "100%")
      .attr("height", "100%")
      .append("g")
      .attr("transform", `translate(${margin.left},${margin.top})`);

    // Create tooltip
    if (this.options.showTooltip) {
      this.tooltip = d3
        .select(this.container)
        .append("div")
        .attr("class", "sankey-tooltip")
        .style("opacity", 0);
    }

    // Set up sankey generator
    const sankeyGenerator = d3Sankey()
      .nodeId((d) => d.index)
      .nodeWidth(nodeWidth)
      .nodePadding(nodePadding)
      .nodeAlign(ALIGNMENT_MAP[alignment] || sankeyJustify)
      .extent([
        [0, 0],
        [innerWidth, innerHeight],
      ]);

    // Deep copy data for d3-sankey (it mutates in place)
    const graphData = {
      nodes: parsedData.nodes.map((d) => ({ ...d })),
      links: parsedData.links.map((d) => ({ ...d })),
    };

    const { nodes, links } = sankeyGenerator(graphData);

    // Color scale
    const color = d3.scaleOrdinal(colorScheme);
    const format = d3.format(this.options.numberFormat);

    // Draw links
    const link = this.svg
      .append("g")
      .attr("class", "sankey-links")
      .attr("fill", "none")
      .selectAll("g")
      .data(links)
      .join("g")
      .attr("class", "sankey-link");

    link
      .append("path")
      .attr("d", sankeyLinkHorizontal())
      .attr("stroke", (d) => color(d.source.name))
      .attr("stroke-width", (d) => Math.max(1, d.width))
      .attr("stroke-opacity", this.options.linkOpacity)
      .on("mouseover", (event, d) => {
        d3.select(event.currentTarget).attr("stroke-opacity", this.options.linkHoverOpacity);
        if (this.tooltip) {
          this.tooltip
            .style("opacity", 1)
            .html(
              `<strong>${this._displayName(d.source.name)}</strong> → <strong>${this._displayName(d.target.name)}</strong><br/>Value: ${format(d.value)}`
            );
        }
      })
      .on("mousemove", (event) => {
        if (this.tooltip) {
          const containerRect = this.container.getBoundingClientRect();
          this.tooltip
            .style("left", event.clientX - containerRect.left + 15 + "px")
            .style("top", event.clientY - containerRect.top - 10 + "px");
        }
      })
      .on("mouseout", (event) => {
        d3.select(event.currentTarget).attr("stroke-opacity", this.options.linkOpacity);
        if (this.tooltip) {
          this.tooltip.style("opacity", 0);
        }
      });

    // Draw nodes
    const node = this.svg
      .append("g")
      .attr("class", "sankey-nodes")
      .selectAll("g")
      .data(nodes)
      .join("g")
      .attr("class", "sankey-node");

    node
      .append("rect")
      .attr("x", (d) => d.x0)
      .attr("y", (d) => d.y0)
      .attr("height", (d) => Math.max(1, d.y1 - d.y0))
      .attr("width", (d) => d.x1 - d.x0)
      .attr("fill", (d) => color(d.name))
      .attr("stroke", "#333")
      .attr("stroke-width", 0.5)
      .attr("rx", 2)
      .on("mouseover", (event, d) => {
        if (this.tooltip) {
          const incoming = d.targetLinks.reduce((sum, l) => sum + l.value, 0);
          const outgoing = d.sourceLinks.reduce((sum, l) => sum + l.value, 0);
          let html = `<strong>${this._displayName(d.name)}</strong>`;
          if (incoming > 0) html += `<br/>Incoming: ${format(incoming)}`;
          if (outgoing > 0) html += `<br/>Outgoing: ${format(outgoing)}`;
          this.tooltip.style("opacity", 1).html(html);
        }
      })
      .on("mousemove", (event) => {
        if (this.tooltip) {
          const containerRect = this.container.getBoundingClientRect();
          this.tooltip
            .style("left", event.clientX - containerRect.left + 15 + "px")
            .style("top", event.clientY - containerRect.top - 10 + "px");
        }
      })
      .on("mouseout", () => {
        if (this.tooltip) {
          this.tooltip.style("opacity", 0);
        }
      });

    // Node labels
    node
      .append("text")
      .attr("x", (d) => (d.x0 < innerWidth / 2 ? d.x1 + 6 : d.x0 - 6))
      .attr("y", (d) => (d.y1 + d.y0) / 2)
      .attr("dy", "0.35em")
      .attr("text-anchor", (d) => (d.x0 < innerWidth / 2 ? "start" : "end"))
      .attr("font-size", this.options.fontSize)
      .attr("font-family", "Segoe UI, sans-serif")
      .attr("fill", "#333")
      .text((d) => {
        const displayName = this._displayName(d.name);
        if (this.options.showValues) {
          const total = Math.max(
            d.sourceLinks.reduce((s, l) => s + l.value, 0),
            d.targetLinks.reduce((s, l) => s + l.value, 0)
          );
          return `${displayName} (${format(total)})`;
        }
        return displayName;
      });

    return this;
  }

  /**
   * Strip stage prefix from multi-level node names for display
   */
  _displayName(name) {
    const colonIndex = name.indexOf(": ");
    return colonIndex > -1 ? name.substring(colonIndex + 2) : name;
  }

  /**
   * Export the diagram as SVG string
   */
  exportSVG() {
    const svgElement = this.container.querySelector("svg");
    if (!svgElement) return null;

    const serializer = new XMLSerializer();
    let svgString = serializer.serializeToString(svgElement);
    svgString = '<?xml version="1.0" standalone="no"?>\n' + svgString;
    return svgString;
  }

  /**
   * Export the diagram as PNG data URL
   */
  exportPNG(scale = 2) {
    return new Promise((resolve, reject) => {
      const svgString = this.exportSVG();
      if (!svgString) return reject(new Error("No SVG to export"));

      const canvas = document.createElement("canvas");
      canvas.width = this.options.width * scale;
      canvas.height = this.options.height * scale;
      const ctx = canvas.getContext("2d");
      ctx.scale(scale, scale);

      const img = new Image();
      const blob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
      const url = URL.createObjectURL(blob);

      img.onload = () => {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL("image/png"));
      };
      img.onerror = reject;
      img.src = url;
    });
  }

  /**
   * Update options and re-render
   */
  update(newOptions) {
    Object.assign(this.options, newOptions);
    if (this.data) {
      this.render(this.data);
    }
  }

  /**
   * Resize to fit container
   */
  resize() {
    const rect = this.container.getBoundingClientRect();
    this.options.width = rect.width || this.options.width;
    this.options.height = rect.height || this.options.height;
    if (this.data) {
      this.render(this.data);
    }
  }
}
