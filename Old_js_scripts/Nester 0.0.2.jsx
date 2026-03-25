#target illustrator
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

// ================== CONFIG ==================
var SHEET_WIDTH_IN = 23;
var SHEET_MAX_LENGTH_IN = 100;
var SPACING_IN = 0.25;
var ALLOW_ROTATION = true;
var OUTPUT_LAYER_NAME = "NEST_OUTPUT";
var CREATE_ARTBOARD = false; // v2.0 false; later can be true
// ============================================

var PT_PER_IN = 72;
var SHEET_WIDTH = SHEET_WIDTH_IN * PT_PER_IN;
var SHEET_MAX_LENGTH = SHEET_MAX_LENGTH_IN * PT_PER_IN;
var SPACING = SPACING_IN * PT_PER_IN;
var EPS = 0.01;

// ---------- Helpers ----------
function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function stripExt(name) {
    var s = String(name);
    var i = s.lastIndexOf(".");
    return (i > 0) ? s.substring(0, i) : s;
}

function getFileNameOnly(pathOrName) {
    var s = String(pathOrName);
    s = s.replace(/\\/g, "/");
    var parts = s.split("/");
    return parts[parts.length - 1];
}

function parseQtyFromName(name) {
    // v2 rule:
    // leading digits + underscore => qty
    // examples:
    // 12_logo.png => 12
    // logo.png => 1
    var base = stripExt(getFileNameOnly(name));
    var m = base.match(/^(\d+)_/);
    if (m && m[1]) {
        var q = parseInt(m[1], 10);
        return (isNaN(q) || q < 1) ? 1 : q;
    }
    return 1;
}

function getItemFileName(item) {
    try {
        if (item.file) return getFileNameOnly(item.file.fsName || item.file.name);
    } catch (e) {}
    try {
        return getFileNameOnly(item.name);
    } catch (e2) {}
    return "unknown.png";
}

function getVisibleSize(item) {
    // visibleBounds = [left, top, right, bottom]
    var vb = item.visibleBounds;
    var w = Math.abs(vb[2] - vb[0]);
    var h = Math.abs(vb[1] - vb[3]);
    return { width: w, height: h };
}

function inToStr(ptVal) {
    return (ptVal / PT_PER_IN).toFixed(2);
}

function areaOf(w, h) {
    return w * h;
}

function longSide(w, h) {
    return (w > h) ? w : h;
}

function shortSide(w, h) {
    return (w < h) ? w : h;
}

function cloneRect(r) {
    return { x: r.x, y: r.y, w: r.w, h: r.h };
}

function rectRight(r) {
    return r.x + r.w;
}

function rectBottom(r) {
    return r.y + r.h;
}

function intersects(a, b) {
    if (a.x >= b.x + b.w - EPS) return false;
    if (a.x + a.w <= b.x + EPS) return false;
    if (a.y >= b.y + b.h - EPS) return false;
    if (a.y + a.h <= b.y + EPS) return false;
    return true;
}

function containsRect(outer, inner) {
    return (
        inner.x >= outer.x - EPS &&
        inner.y >= outer.y - EPS &&
        inner.x + inner.w <= outer.x + outer.w + EPS &&
        inner.y + inner.h <= outer.y + outer.h + EPS
    );
}

function isValidRect(r) {
    return r.w > EPS && r.h > EPS;
}

function round2(n) {
    return Math.round(n * 100) / 100;
}

function ensureOutputLayer(doc, layerName) {
    var layer;
    try {
        layer = doc.layers.getByName(layerName);
        for (var i = layer.pageItems.length - 1; i >= 0; i--) {
            try { layer.pageItems[i].remove(); } catch (e) {}
        }
    } catch (e2) {
        layer = doc.layers.add();
        layer.name = layerName;
    }
    return layer;
}

// ---------- Collect ----------
function collectPlacedItems(doc) {
    var result = [];
    for (var i = 0; i < doc.placedItems.length; i++) {
        var it = doc.placedItems[i];
        var name = getItemFileName(it);
        var qty = parseQtyFromName(name);
        var sz = getVisibleSize(it);

        result.push({
            ref: it,
            name: name,
            qty: qty,
            width: sz.width,
            height: sz.height,
            area: areaOf(sz.width, sz.height),
            longSide: longSide(sz.width, sz.height)
        });
    }
    return result;
}

// ---------- Expand ----------
function expandItems(items) {
    var expanded = [];
    for (var i = 0; i < items.length; i++) {
        for (var q = 0; q < items[i].qty; q++) {
            expanded.push({
                ref: items[i].ref,
                name: items[i].name,
                width: items[i].width,
                height: items[i].height,
                area: items[i].area,
                longSide: items[i].longSide
            });
        }
    }
    return expanded;
}

// ---------- Sort ----------
function sortExpandedPieces(arr) {
    arr.sort(function(a, b) {
        // 1) long side desc
        if (b.longSide !== a.longSide) return b.longSide - a.longSide;

        // 2) area desc
        if (b.area !== a.area) return b.area - a.area;

        // 3) height desc
        if (b.height !== a.height) return b.height - a.height;

        // 4) width desc
        return b.width - a.width;
    });
}

// ---------- Free Rect / MaxRects-like ----------
function makePaddedDims(piece) {
    // spacing modeled as padding on the right/bottom of each piece
    return {
        width: piece.width + SPACING,
        height: piece.height + SPACING
    };
}

function scorePlacement(freeRect, pw, ph, currentUsedLength) {
    // Best Short Side Fit
    var leftoverHoriz = freeRect.w - pw;
    var leftoverVert = freeRect.h - ph;
    var shortFit = Math.min(leftoverHoriz, leftoverVert);
    var longFit = Math.max(leftoverHoriz, leftoverVert);
    var areaFit = (freeRect.w * freeRect.h) - (pw * ph);

    var projectedBottom = freeRect.y + ph;
    var projectedUsedLength = (projectedBottom > currentUsedLength) ? projectedBottom : currentUsedLength;

    return {
        shortFit: shortFit,
        longFit: longFit,
        areaFit: areaFit,
        projectedUsedLength: projectedUsedLength
    };
}

function isBetterScore(a, b) {
    // true if a is better than b
    if (!b) return true;

    if (a.shortFit !== b.shortFit) return a.shortFit < b.shortFit;
    if (a.areaFit !== b.areaFit) return a.areaFit < b.areaFit;
    if (a.longFit !== b.longFit) return a.longFit < b.longFit;
    if (a.projectedUsedLength !== b.projectedUsedLength) return a.projectedUsedLength < b.projectedUsedLength;

    return false;
}

function findBestPlacement(piece, freeRects, allowRotation, currentUsedLength) {
    var best = null;
    var dims = makePaddedDims(piece);

    for (var i = 0; i < freeRects.length; i++) {
        var fr = freeRects[i];

        // normal
        if (dims.width <= fr.w + EPS && dims.height <= fr.h + EPS) {
            var s1 = scorePlacement(fr, dims.width, dims.height, currentUsedLength);
            if (isBetterScore(s1, best ? best.score : null)) {
                best = {
                    x: fr.x,
                    y: fr.y,
                    w: dims.width,
                    h: dims.height,
                    artW: piece.width,
                    artH: piece.height,
                    rotated: false,
                    freeRectIndex: i,
                    score: s1
                };
            }
        }

        // rotated
        if (allowRotation && dims.height <= fr.w + EPS && dims.width <= fr.h + EPS) {
            var s2 = scorePlacement(fr, dims.height, dims.width, currentUsedLength);
            if (isBetterScore(s2, best ? best.score : null)) {
                best = {
                    x: fr.x,
                    y: fr.y,
                    w: dims.height,
                    h: dims.width,
                    artW: piece.height,
                    artH: piece.width,
                    rotated: true,
                    freeRectIndex: i,
                    score: s2
                };
            }
        }
    }

    return best;
}

function splitFreeRect(freeRect, usedRect) {
    var result = [];

    if (!intersects(freeRect, usedRect)) {
        result.push(cloneRect(freeRect));
        return result;
    }

    var frLeft = freeRect.x;
    var frTop = freeRect.y;
    var frRight = freeRect.x + freeRect.w;
    var frBottom = freeRect.y + freeRect.h;

    var usedLeft = usedRect.x;
    var usedTop = usedRect.y;
    var usedRight = usedRect.x + usedRect.w;
    var usedBottom = usedRect.y + usedRect.h;

    // Top
    if (usedTop > frTop + EPS) {
        result.push({
            x: frLeft,
            y: frTop,
            w: freeRect.w,
            h: usedTop - frTop
        });
    }

    // Bottom
    if (usedBottom < frBottom - EPS) {
        result.push({
            x: frLeft,
            y: usedBottom,
            w: freeRect.w,
            h: frBottom - usedBottom
        });
    }

    // Left
    if (usedLeft > frLeft + EPS) {
        result.push({
            x: frLeft,
            y: frTop,
            w: usedLeft - frLeft,
            h: freeRect.h
        });
    }

    // Right
    if (usedRight < frRight - EPS) {
        result.push({
            x: usedRight,
            y: frTop,
            w: frRight - usedRight,
            h: freeRect.h
        });
    }

    return result;
}

function pruneFreeRects(freeRects) {
    var pruned = [];

    // remove invalid first
    for (var i = 0; i < freeRects.length; i++) {
        if (isValidRect(freeRects[i])) pruned.push(freeRects[i]);
    }

    // remove contained rects
    for (var a = 0; a < pruned.length; a++) {
        if (!pruned[a]) continue;

        for (var b = 0; b < pruned.length; b++) {
            if (a === b || !pruned[b]) continue;

            if (containsRect(pruned[a], pruned[b])) {
                pruned[b] = null;
            }
        }
    }

    var finalRects = [];
    for (var k = 0; k < pruned.length; k++) {
        if (pruned[k] && isValidRect(pruned[k])) finalRects.push(pruned[k]);
    }

    return finalRects;
}

function placePieceAndUpdateFreeRects(freeRects, placement) {
    var usedRect = {
        x: placement.x,
        y: placement.y,
        w: placement.w,
        h: placement.h
    };

    var newFreeRects = [];
    for (var i = 0; i < freeRects.length; i++) {
        var fr = freeRects[i];
        var splitRects = splitFreeRect(fr, usedRect);
        for (var j = 0; j < splitRects.length; j++) {
            if (isValidRect(splitRects[j])) newFreeRects.push(splitRects[j]);
        }
    }

    return pruneFreeRects(newFreeRects);
}

// ---------- Nest ----------
function nestPiecesMaxRects(pieces, sheetWidth, maxLength, allowRotation) {
    var freeRects = [{
        x: 0,
        y: 0,
        w: sheetWidth,
        h: maxLength
    }];

    var placed = [];
    var unplaced = [];
    var usedLength = 0;

    for (var i = 0; i < pieces.length; i++) {
        var piece = pieces[i];
        var placement = findBestPlacement(piece, freeRects, allowRotation, usedLength);

        if (!placement) {
            unplaced.push(piece);
            continue;
        }

        placed.push({
            ref: piece.ref,
            name: piece.name,
            x: placement.x,
            y: placement.y,
            paddedW: placement.w,
            paddedH: placement.h,
            width: placement.artW,
            height: placement.artH,
            rotated: placement.rotated
        });

        freeRects = placePieceAndUpdateFreeRects(freeRects, placement);

        var pieceBottom = placement.y + placement.h;
        if (pieceBottom > usedLength) usedLength = pieceBottom;
    }

    return {
        placed: placed,
        unplaced: unplaced,
        usedLength: usedLength,
        freeRects: freeRects
    };
}

// ---------- Render ----------
function duplicateToLayer(item, targetLayer) {
    return item.duplicate(targetLayer, ElementPlacement.PLACEATEND);
}

function setItemTopLeftByVisibleBounds(item, targetLeft, targetTop) {
    var vb = item.visibleBounds; // [left, top, right, bottom]
    var dx = targetLeft - vb[0];
    var dy = targetTop - vb[1];
    item.translate(dx, dy);
}

function renderLayout(doc, layout, outputLayer) {
    // Illustrator y-axis visually goes downward if we use "top minus y"
    var docTop = 0;

    for (var i = 0; i < layout.placed.length; i++) {
        var p = layout.placed[i];
        var dup = duplicateToLayer(p.ref, outputLayer);

        if (p.rotated) {
            dup.rotate(90);
        }

        // target top-left in packing space
        var targetLeft = p.x;
        var targetTop = docTop - p.y;

        setItemTopLeftByVisibleBounds(dup, targetLeft, targetTop);
    }
}

// ---------- Optional artboard ----------
function createOutputArtboard(doc, usedLength) {
    if (!CREATE_ARTBOARD) return;

    var width = SHEET_WIDTH;
    var height = Math.max(usedLength, 72); // minimum 1 in

    // artboardRect: [left, top, right, bottom]
    var rect = [0, 0, width, -height];
    doc.artboards.add(rect);
}

// ---------- Report ----------
function buildUnplacedSummary(unplaced) {
    if (!unplaced || unplaced.length === 0) return "None";

    var counts = {};
    for (var i = 0; i < unplaced.length; i++) {
        var name = unplaced[i].name;
        if (!counts[name]) counts[name] = 0;
        counts[name]++;
    }

    var parts = [];
    for (var k in counts) {
        if (counts.hasOwnProperty(k)) {
            parts.push(k + " x " + counts[k]);
        }
    }

    return parts.join("\n");
}

// ---------- Main ----------
function main() {
    if (app.documents.length === 0) {
        alert("No open document found.");
        return;
    }

    var doc = app.activeDocument;
    var sourceItems = collectPlacedItems(doc);

    if (sourceItems.length === 0) {
        alert("No placed items found in the current document.");
        return;
    }

    var expanded = expandItems(sourceItems);
    sortExpandedPieces(expanded);

    var outputLayer = ensureOutputLayer(doc, OUTPUT_LAYER_NAME);
    var layout = nestPiecesMaxRects(expanded, SHEET_WIDTH, SHEET_MAX_LENGTH, ALLOW_ROTATION);

    renderLayout(doc, layout, outputLayer);
    createOutputArtboard(doc, layout.usedLength);

    var usedLengthNoTrailingGap = Math.max(0, layout.usedLength - SPACING);
    var usedArea = SHEET_WIDTH * Math.max(layout.usedLength, 0);
    var artArea = 0;
    var placedCount = layout.placed.length;

    for (var i = 0; i < layout.placed.length; i++) {
        artArea += layout.placed[i].width * layout.placed[i].height;
    }

    var wastePct = 0;
    if (usedArea > EPS) {
        wastePct = ((usedArea - artArea) / usedArea) * 100;
    }

    var msg =
        "WeMust NESTER v2.0\n\n" +
        "Source items: " + sourceItems.length + "\n" +
        "Expanded pieces: " + expanded.length + "\n" +
        "Placed: " + placedCount + "\n" +
        "Unplaced: " + layout.unplaced.length + "\n" +
        "Used length: " + inToStr(usedLengthNoTrailingGap) + " in\n" +
        "Sheet width: " + SHEET_WIDTH_IN + " in\n" +
        "Max length: " + SHEET_MAX_LENGTH_IN + " in\n" +
        "Approx waste: " + round2(wastePct) + "%\n\n" +
        "Unplaced summary:\n" + buildUnplacedSummary(layout.unplaced);

    alert(msg);
}

main();