#target illustrator

var PREVIEW_LAYER_NAME = "NEST_PREVIEW";

function _getActiveDocument() {
    if (!app.documents.length) {
        throw new Error("No active document.");
    }
    return app.activeDocument;
}

function _deleteLayerIfExists(doc, layerName) {
    var layers = doc.layers;
    for (var i = layers.length - 1; i >= 0; i--) {
        var layer = layers[i];
        if (layer.name === layerName) {
            layer.remove();
            break;
        }
    }
}

function _layoutAsGrid(items, artboardRect, style, blocking, widthFill) {
    if (!items || items.length === 0) {
        return;
    }

    var left = artboardRect[0];
    var top = artboardRect[1];
    var right = artboardRect[2];
    var bottom = artboardRect[3];

    var artboardWidth = right - left;
    var artboardHeight = top - bottom;

    var fill = widthFill || 50;
    var baseCols = Math.ceil(Math.sqrt(items.length));
    var cols = Math.max(1, Math.round((baseCols * fill) / 100));
    var rows = Math.ceil(items.length / cols);

    var margin = 10;
    var cellWidth = (artboardWidth - margin * (cols + 1)) / cols;
    var cellHeight = (artboardHeight - margin * (rows + 1)) / rows;

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var col = i % cols;
        var row = Math.floor(i / cols);

        var cellLeft = left + margin + col * (cellWidth + margin);
        var cellTop = top - margin - row * (cellHeight + margin);
        var cellCenterX = cellLeft + cellWidth / 2;
        var cellCenterY = cellTop - cellHeight / 2;

        var gb = item.geometricBounds; // [l, t, r, b]
        var itemCenterX = (gb[0] + gb[2]) / 2;
        var itemCenterY = (gb[1] + gb[3]) / 2;

        var dx = cellCenterX - itemCenterX;
        var dy = cellCenterY - itemCenterY;

        item.translate(dx, dy);
    }
}

/**
 * Main entry point called from CEP JS.
 * This is where you will later plug in the real nesting engine instead of the simple grid.
 */
function buildPreview(style, blocking, widthFill) {
    try {
        var doc = _getActiveDocument();
        var sel = doc.selection;
        if (!sel || sel.length === 0) {
            return "ERROR: No page items selected.";
        }

        _deleteLayerIfExists(doc, PREVIEW_LAYER_NAME);

        var previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;

        var duplicates = [];
        for (var i = 0; i < sel.length; i++) {
            var item = sel[i];
            var dup = item.duplicate(previewLayer, ElementPlacement.PLACEATBEGINNING);
            duplicates.push(dup);
        }

        var ab = doc.artboards[0];
        var rect = ab.artboardRect; // [left, top, right, bottom]

        _layoutAsGrid(duplicates, rect, style, blocking, widthFill);

        return "Preview built successfully.";
    } catch (e) {
        return "ERROR: " + e.message;
    }
}

function clearPreview() {
    try {
        var doc = _getActiveDocument();
        _deleteLayerIfExists(doc, PREVIEW_LAYER_NAME);
        return "Preview cleared.";
    } catch (e) {
        return "ERROR: " + e.message;
    }
}

