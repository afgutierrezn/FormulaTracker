/* global cytoscape, document */

// Draws the dependency graph inside #graph-container using Cytoscape.js.
// graphData is the { nodes, edges } object returned by buildGraph() in graph.js.
export function renderGraph(graphData) {
  const container = document.getElementById("graph-container");
  const warningLabel = document.getElementById("warning-label");

  // Clear whatever was drawn before
  container.innerHTML = "";

  // If there are no connections, show a plain message instead of an empty box
  if (!graphData.nodes.length) {
    container.innerHTML = '<p style="padding:16px; color:#888;">No dependencies found.</p>';
    warningLabel.style.display = "none";
    return;
  }

  // Show a warning if any formula uses dynamic references the parser can't fully trace
  const dynamicFunctions = ["INDIRECT(", "OFFSET(", "INDEX("];
  const hasDynamic = graphData.nodes.some((node) =>
    dynamicFunctions.some((fn) => node.formula.toUpperCase().includes(fn))
  );
  if (hasDynamic) {
    warningLabel.textContent =
      "Warning: some formulas use INDIRECT, OFFSET, or INDEX. These dynamic references may not be fully traced.";
    warningLabel.style.display = "block";
  } else {
    warningLabel.style.display = "none";
  }

  // Collect IDs of nodes that are targets of circular edges.
  // These are the cells at the end of a chain that feed back to an ancestor.
  const circularNodeIds = new Set(
    graphData.edges
      .filter((e) => e.label === "circular")
      .map((e) => e.target)
  );

  // Convert our nodes into Cytoscape's format.
  // Circular nodes get a second label line and a flag for red styling.
  const cyNodes = graphData.nodes.map((n) => ({
    data: {
      id: n.id,
      label: circularNodeIds.has(n.id) ? n.label + " ↩" : n.label,
      circular: circularNodeIds.has(n.id),
    },
  }));

  // Convert non-circular edges into Cytoscape's format.
  // Circular edges are dropped — the red node marking replaces them.
  const cyEdges = graphData.edges
    .filter((e) => e.label !== "circular")
    .map((e, i) => ({
      data: { id: "edge-" + i, source: e.source, target: e.target, label: "" },
    }));

  // Create the Cytoscape instance with elements and styles, but no layout yet.
  // Layout runs separately below so we can reference cy.nodes()[0] as the root.
  const cy = cytoscape({
    container: container,
    elements: [...cyNodes, ...cyEdges],

    style: [
      // Default node style
      {
        selector: "node",
        style: {
          label: "data(label)",
          "text-valign": "center",
          "text-halign": "center",
          "background-color": "#ffffff",
          "border-width": 2,
          "border-color": "#0078d7",
          "font-size": 12,
          width: "label",
          height: "label",
          padding: 8,
          shape: "round-rectangle",
        },
      },
      // Default edge style
      {
        selector: "edge",
        style: {
          width: 2,
          "line-color": "#aaaaaa",
          "target-arrow-color": "#aaaaaa",
          "target-arrow-shape": "triangle",
          "curve-style": "bezier",
          label: "data(label)",
          "font-size": 10,
          "text-background-color": "#ffffff",
          "text-background-opacity": 1,
          "text-background-padding": "2px",
        },
      },
      // Circular reference nodes — red border and pink fill instead of a crossing line
      {
        selector: "node[?circular]",
        style: {
          "border-color": "#cc0000",
          "border-width": 3,
          "background-color": "#fff0f0",
        },
      },
    ],
  });

  // Run the layout using only non-circular edges.
  // Circular edges are excluded here so they don't distort depth — a circular
  // edge connects the root directly to a deep node and would pull it to level 1.
  // The circular edges are still drawn visually; they just don't influence positions.
  cy.elements()
    .not("edge[label = 'circular']")
    .layout({
      name: "breadthfirst",
      directed: false,
      roots: cy.nodes()[0],
      padding: 20,
      spacingFactor: 1.4,
      avoidOverlap: true,
    })
    .run();
}
