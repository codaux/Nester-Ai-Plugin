// panel.js
// Main UI logic for Nester DTF Preview panel

// UXP core (for host info, if needed)
const { host } = require("uxp");
// Illustrator host module (name may differ depending on Adobe's final API)
let illustrator;
try {
  illustrator = require("illustrator");
} catch (e) {
  // Fallback: keep undefined; status messages will show error in UI.
  illustrator = null;
}

const PREVIEW_LAYER_NAME = "NEST_PREVIEW";

function $(id) {
  return document.getElementById(id);
}

function setStatus(message) {
  const statusText = $("statusText");
  if (statusText) {
    statusText.textContent = message;
  }
}

function initUI() {
  const styleSelect = $("styleSelect");
  const blockingSelect = $("blockingSelect");
  const widthFillRange = $("widthFillRange");
  const widthFillValue = $("widthFillValue");
  const buildPreviewBtn = $("buildPreviewBtn");
  const clearPreviewBtn = $("clearPreviewBtn");

  widthFillRange.addEventListener("input", () => {
    widthFillValue.textContent = widthFillRange.value;
  });

  buildPreviewBtn.addEventListener("click", async () => {
    const style = styleSelect.value;
    const blocking = parseInt(blockingSelect.value, 10);
    const widthFill = parseInt(widthFillRange.value, 10);

    setStatus("Building preview...");

    try {
      await buildPreview({ style, blocking, widthFill });
      setStatus("Preview built successfully.");
    } catch (err) {
      console.error(err);
      setStatus(`Error: ${err.message || err}`);
    }
  });

  clearPreviewBtn.addEventListener("click", async () => {
    setStatus("Clearing preview...");
    try {
      await clearPreviewLayer();
      setStatus("Preview cleared.");
    } catch (err) {
      console.error(err);
      setStatus(`Error: ${err.message || err}`);
    }
  });

  setStatus("Ready.");
}

/**
 * Main entry for building the preview.
 * This is where, in the future, you can replace the simple grid with your real nesting engine.
 */
async function buildPreview({ style, blocking, widthFill }) {
  if (!illustrator) {
    throw new Error("Illustrator UXP API (require('illustrator')) not available.");
  }

  const app = illustrator.app;
  const doc = app.activeDocument;
  if (!doc) {
    throw new Error("No active document.");
  }

  const sel = doc.selection;
  if (!sel || sel.length === 0) {
    throw new Error("No page items selected.");
  }

  // Remove old preview layer if exists
  await deleteLayerIfExists(doc, PREVIEW_LAYER_NAME);

  // Create new preview layer
  const previewLayer = doc.layers.add();
  previewLayer.name = PREVIEW_LAYER_NAME;

  // Duplicate selection into preview layer
  const duplicates = [];
  for (let i = 0; i < sel.length; i++) {
    const item = sel[i];
    const dup = item.duplicate(previewLayer, illustrator.ElementPlacement.PLACEATBEGINNING);
    duplicates.push(dup);
  }

  // Get artboard bounds (use first artboard for simplicity)
  const artboard = doc.artboards[0];
  const artboardBounds = artboard.artboardRect; // [left, top, right, bottom]

  // Compute and apply a simple grid layout
  layoutAsGridOnArtboard(duplicates, artboardBounds, {
    style,
    blocking,
    widthFill
  });

  // Panel stays open automatically.
}

/**
 * Clear the NEST_PREVIEW layer, if it exists.
 */
async function clearPreviewLayer() {
  if (!illustrator) {
    throw new Error("Illustrator UXP API (require('illustrator')) not available.");
  }

  const app = illustrator.app;
  const doc = app.activeDocument;
  if (!doc) {
    return;
  }

  await deleteLayerIfExists(doc, PREVIEW_LAYER_NAME);
}

/**
 * Delete the layer with the given name if it exists.
 */
async function deleteLayerIfExists(doc, layerName) {
  const layers = doc.layers;
  for (let i = layers.length - 1; i >= 0; i--) {
    const layer = layers[i];
    if (layer.name === layerName) {
      layer.remove();
      break;
    }
  }
}

/**
 * Very simple grid layout on the active artboard.
 *
 * This is intentionally simple and modular so you can later replace it with
 * your real nesting engine:
 * - Collect bounds of duplicates
 * - Compute positions with your nesting algorithm
 * - Apply transforms in a similar loop
 */
function layoutAsGridOnArtboard(items, artboardBounds, options) {
  if (!items || items.length === 0) {
    return;
  }

  const [left, top, right, bottom] = artboardBounds;
  const artboardWidth = right - left;
  const artboardHeight = top - bottom;

  // Interpret options (you can wire these into your real algorithm later)
  const widthFill = options.widthFill || 50; // percentage 0-100
  const blocking = options.blocking || 2;

  // Simple heuristic: number of columns based on sqrt and widthFill
  const baseCols = Math.ceil(Math.sqrt(items.length));
  const cols = Math.max(1, Math.round((baseCols * widthFill) / 100));
  const rows = Math.ceil(items.length / cols);

  const margin = 10;
  const cellWidth = (artboardWidth - margin * (cols + 1)) / cols;
  const cellHeight = (artboardHeight - margin * (rows + 1)) / rows;

  // For now, we ignore item shapes and use cell centers.
  items.forEach((item, index) => {
    const col = index % cols;
    const row = Math.floor(index / cols);

    const cellLeft = left + margin + col * (cellWidth + margin);
    const cellTop = top - margin - row * (cellHeight + margin);
    const cellCenterX = cellLeft + cellWidth / 2;
    const cellCenterY = cellTop - cellHeight / 2;

    const gb = item.geometricBounds; // [l, t, r, b]
    const itemCenterX = (gb[0] + gb[2]) / 2;
    const itemCenterY = (gb[1] + gb[3]) / 2;

    const dx = cellCenterX - itemCenterX;
    const dy = cellCenterY - itemCenterY;

    // Translate the item to the new center.
    item.translate(dx, dy);
  });
}

// Initialize UI when document is ready
document.addEventListener("DOMContentLoaded", () => {
  try {
    initUI();
  } catch (e) {
    console.error(e);
    setStatus(`Initialization error: ${e.message || e}`);
  }
});

