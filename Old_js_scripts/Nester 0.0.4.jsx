#target illustrator
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

// =====================================================
// WeMust NESTER v5
// Dialog-based rebuild loop + block-aware nesting
// =====================================================

var PT_PER_IN = 72;
var EPS = 0.01;

var OUTPUT_LAYER_NAME = "NEST_BUILD";
var OUTPUT_LAYER_PREFIX = "NEST_BUILD";

var DEFAULTS = {
    sheetWidthIn: 23,
    maxLengthIn: 100,
    spacingIn: 0.25,
    mode: "Balanced",               // Efficient | Balanced | Clean
    maxBlocksPerFile: 2,            // 1..4
    maxBlockAspectRatio: 3.0,       // lower = less skinny strips
    allowItemRotationInBlock: false,
    allowBlockRotationOnSheet: true,
    previousOutputAction: "Remove", // Remove | Hide | Keep
    hideSourceLayersAfterBuild: true
};

// =====================================================
// Basic helpers
// =====================================================

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
    var vb = item.visibleBounds; // [left, top, right, bottom]
    var w = Math.abs(vb[2] - vb[0]);
    var h = Math.abs(vb[1] - vb[3]);
    return { width: w, height: h };
}

function inToPt(v) { return v * PT_PER_IN; }
function ptToIn(v) { return v / PT_PER_IN; }
function inToStr(vPt) { return (vPt / PT_PER_IN).toFixed(2); }

function areaOf(w, h) { return w * h; }
function longSide(w, h) { return (w > h) ? w : h; }
function shortSide(w, h) { return (w < h) ? w : h; }

function cloneRect(r) {
    return { x: r.x, y: r.y, w: r.w, h: r.h };
}

function startsWith(str, prefix) {
    return String(str).indexOf(prefix) === 0;
}

function round2(n) {
    return Math.round(n * 100) / 100;
}

function uniquePushPartition(list, arr) {
    var key = arr.join("-");
    for (var i = 0; i < list.length; i++) {
        if (list[i].join("-") === key) return;
    }
    list.push(arr);
}

// =====================================================
// Geometry
// =====================================================

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

// =====================================================
// Dialog
// =====================================================

function showSettingsDialog(current) {
    var dlg = new Window("dialog", "WeMust NESTER v5");
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var p1 = dlg.add("panel", undefined, "Sheet");
    p1.orientation = "column";
    p1.alignChildren = "left";

    var g1 = p1.add("group");
    g1.add("statictext", undefined, "Width (in):");
    var etWidth = g1.add("edittext", undefined, String(current.sheetWidthIn));
    etWidth.characters = 7;

    var g2 = p1.add("group");
    g2.add("statictext", undefined, "Max length (in):");
    var etLength = g2.add("edittext", undefined, String(current.maxLengthIn));
    etLength.characters = 7;

    var g3 = p1.add("group");
    g3.add("statictext", undefined, "Spacing (in):");
    var etSpacing = g3.add("edittext", undefined, String(current.spacingIn));
    etSpacing.characters = 7;

    var p2 = dlg.add("panel", undefined, "Packing");
    p2.orientation = "column";
    p2.alignChildren = "left";

    var g4 = p2.add("group");
    g4.add("statictext", undefined, "Mode:");
    var ddMode = g4.add("dropdownlist", undefined, ["Efficient", "Balanced", "Clean"]);
    ddMode.selection = (current.mode === "Efficient") ? 0 : (current.mode === "Clean" ? 2 : 1);

    var g5 = p2.add("group");
    g5.add("statictext", undefined, "Max blocks/file:");
    var ddBlocks = g5.add("dropdownlist", undefined, ["1", "2", "3", "4"]);
    ddBlocks.selection = Math.max(0, Math.min(3, current.maxBlocksPerFile - 1));

    var g6 = p2.add("group");
    g6.add("statictext", undefined, "Max block ratio:");
    var etAspect = g6.add("edittext", undefined, String(current.maxBlockAspectRatio));
    etAspect.characters = 7;

    var cbItemRot = p2.add("checkbox", undefined, "Allow item rotation inside block");
    cbItemRot.value = current.allowItemRotationInBlock;

    var cbBlockRot = p2.add("checkbox", undefined, "Allow whole-block rotation on sheet");
    cbBlockRot.value = current.allowBlockRotationOnSheet;

    var p3 = dlg.add("panel", undefined, "Output");
    p3.orientation = "column";
    p3.alignChildren = "left";

    var g7 = p3.add("group");
    g7.add("statictext", undefined, "Previous output:");
    var ddPrev = g7.add("dropdownlist", undefined, ["Remove", "Hide", "Keep"]);
    ddPrev.selection = (current.previousOutputAction === "Hide") ? 1 : (current.previousOutputAction === "Keep" ? 2 : 0);

    var cbHideSource = p3.add("checkbox", undefined, "Hide source layers after build");
    cbHideSource.value = current.hideSourceLayersAfterBuild;

    var note = dlg.add("statictext", undefined, "Tip: Clean + 1 or 2 blocks/file + ratio 2.4–3.0 usually gives tidier results.");
    note.characters = 70;

    var btns = dlg.add("group");
    btns.alignment = "right";
    btns.add("button", undefined, "Build", {name: "ok"});
    btns.add("button", undefined, "Cancel", {name: "cancel"});

    if (dlg.show() != 1) return null;

    function parseNum(txt, fallback) {
        var v = parseFloat(String(txt).replace(",", "."));
        return (isNaN(v) || v <= 0) ? fallback : v;
    }

    return {
        sheetWidthIn: parseNum(etWidth.text, current.sheetWidthIn),
        maxLengthIn: parseNum(etLength.text, current.maxLengthIn),
        spacingIn: parseNum(etSpacing.text, current.spacingIn),
        mode: ddMode.selection ? ddMode.selection.text : current.mode,
        maxBlocksPerFile: ddBlocks.selection ? parseInt(ddBlocks.selection.text, 10) : current.maxBlocksPerFile,
        maxBlockAspectRatio: parseNum(etAspect.text, current.maxBlockAspectRatio),
        allowItemRotationInBlock: cbItemRot.value,
        allowBlockRotationOnSheet: cbBlockRot.value,
        previousOutputAction: ddPrev.selection ? ddPrev.selection.text : current.previousOutputAction,
        hideSourceLayersAfterBuild: cbHideSource.value
    };
}

// =====================================================
// Layer helpers
// =====================================================

function isOutputLayerName(name) {
    return startsWith(name, OUTPUT_LAYER_PREFIX);
}

function findLayerByName(doc, name) {
    try { return doc.layers.getByName(name); } catch (e) {}
    return null;
}

function prepareOutputLayer(doc, action) {
    var existing = findLayerByName(doc, OUTPUT_LAYER_NAME);

    if (existing) {
        if (action === "Remove") {
            try { existing.remove(); } catch (e1) {}
        } else if (action === "Hide") {
            try { existing.visible = false; } catch (e2) {}
            try { existing.name = OUTPUT_LAYER_NAME + "_OLD_" + (new Date().getTime()); } catch (e3) {}
        } else if (action === "Keep") {
            try { existing.name = OUTPUT_LAYER_NAME + "_OLD_" + (new Date().getTime()); } catch (e4) {}
        }
    }

    var layer = doc.layers.add();
    layer.name = OUTPUT_LAYER_NAME;
    layer.visible = true;
    return layer;
}

function hideSourceLayers(doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        var lyr = doc.layers[i];
        if (isOutputLayerName(lyr.name)) {
            try { lyr.visible = true; } catch (e1) {}
        } else {
            try { lyr.visible = false; } catch (e2) {}
        }
    }
}

// =====================================================
// Collect source items
// IMPORTANT: excludes output layers so rebuild does not re-read old builds
// =====================================================

function collectPlacedItems(doc) {
    var result = [];

    for (var i = 0; i < doc.placedItems.length; i++) {
        var it = doc.placedItems[i];

        try {
            if (it.layer && isOutputLayerName(it.layer.name)) {
                continue;
            }
        } catch (e0) {}

        var name = getItemFileName(it);
        var qty = parseQtyFromName(name);
        var sz = getVisibleSize(it);

        result.push({
            id: "SRC_" + i,
            ref: it,
            name: name,
            qty: qty,
            width: sz.width,
            height: sz.height,
            area: areaOf(sz.width, sz.height),
            longSide: longSide(sz.width, sz.height),
            shortSide: shortSide(sz.width, sz.height)
        });
    }

    return result;
}

// =====================================================
// Mode weights
// =====================================================

function getModeWeights(mode) {
    if (mode === "Efficient") {
        return {
            planBlockCount: 12000,
            planEmptyCells: 2400,
            planAspectPenalty: 10000,
            planBlockArea: 1,
            planLongSide: 4,
            planRotPenalty: 300,

            nestShortFit: 900,
            nestAreaFit: 6,
            nestLongFit: 2,
            nestProjectedLen: 70,
            nestFragmentPenalty: 0,
            nestAlignBonus: 10
        };
    } else if (mode === "Clean") {
        return {
            planBlockCount: 24000,
            planEmptyCells: 2800,
            planAspectPenalty: 26000,
            planBlockArea: 1,
            planLongSide: 10,
            planRotPenalty: 1100,

            nestShortFit: 500,
            nestAreaFit: 4,
            nestLongFit: 4,
            nestProjectedLen: 120,
            nestFragmentPenalty: 10,
            nestAlignBonus: 120
        };
    }

    return {
        planBlockCount: 17000,
        planEmptyCells: 2600,
        planAspectPenalty: 17000,
        planBlockArea: 1,
        planLongSide: 7,
        planRotPenalty: 650,

        nestShortFit: 750,
        nestAreaFit: 5,
        nestLongFit: 3,
        nestProjectedLen: 90,
        nestFragmentPenalty: 3,
        nestAlignBonus: 55
    };
}

// =====================================================
// Partition generation
// =====================================================

function balancedPartition(qty, parts) {
    var arr = [];
    var base = Math.floor(qty / parts);
    var rem = qty % parts;
    for (var i = 0; i < parts; i++) {
        arr.push(base + (i < rem ? 1 : 0));
    }
    return arr;
}

function shiftedPartition(qty, parts, shiftAmount) {
    var arr = balancedPartition(qty, parts);
    var moved = 0;

    while (moved < shiftAmount) {
        var last = arr.length - 1;
        if (arr[last] <= 1) break;
        arr[0] += 1;
        arr[last] -= 1;
        moved++;
    }

    arr.sort(function(a, b) { return b - a; });
    return arr;
}

function generatePartitionCandidates(qty, maxBlocks) {
    var out = [];
    var maxP = Math.min(maxBlocks, qty);

    for (var parts = 1; parts <= maxP; parts++) {
        uniquePushPartition(out, balancedPartition(qty, parts));
        if (parts >= 2) {
            uniquePushPartition(out, shiftedPartition(qty, parts, 1));
            if (qty >= 6) uniquePushPartition(out, shiftedPartition(qty, parts, 2));
        }
    }

    return out;
}

// =====================================================
// Block planning
// =====================================================

function chooseBestGridForCount(src, count, settings) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);
    var spacing = inToPt(settings.spacingIn);
    var weights = getModeWeights(settings.mode);
    var maxAspect = settings.maxBlockAspectRatio;

    var candidates = [];
    var orientations = [{rotated: false, cellW: src.width, cellH: src.height}];

    if (settings.allowItemRotationInBlock) {
        orientations.push({rotated: true, cellW: src.height, cellH: src.width});
    }

    for (var o = 0; o < orientations.length; o++) {
        var ori = orientations[o];

        for (var cols = 1; cols <= count; cols++) {
            var rows = Math.ceil(count / cols);
            var emptyCells = rows * cols - count;

            var blockW = cols * ori.cellW + (cols - 1) * spacing;
            var blockH = rows * ori.cellH + (rows - 1) * spacing;

            if (blockW > sheetWidth + EPS) continue;
            if (blockH > maxLength + EPS) continue;

            var ratio = longSide(blockW, blockH) / Math.max(shortSide(blockW, blockH), EPS);
            var excessAspect = Math.max(0, ratio - maxAspect);
            var aspectPenalty = excessAspect * excessAspect;

            var score =
                (emptyCells * weights.planEmptyCells) +
                (aspectPenalty * weights.planAspectPenalty) +
                ((blockW * blockH) * weights.planBlockArea) +
                (longSide(blockW, blockH) * weights.planLongSide) +
                ((ori.rotated ? 1 : 0) * weights.planRotPenalty);

            candidates.push({
                count: count,
                cols: cols,
                rows: rows,
                emptyCells: emptyCells,
                rotatedInside: ori.rotated,
                cellW: ori.cellW,
                cellH: ori.cellH,
                blockW: blockW,
                blockH: blockH,
                ratio: ratio,
                score: score
            });
        }
    }

    if (candidates.length === 0) return null;

    candidates.sort(function(a, b) {
        if (a.score !== b.score) return a.score - b.score;
        if (a.emptyCells !== b.emptyCells) return a.emptyCells - b.emptyCells;
        return (a.blockW * a.blockH) - (b.blockW * b.blockH);
    });

    return candidates[0];
}

function buildBestBlockPlanForSource(src, settings) {
    var partitions = generatePartitionCandidates(src.qty, settings.maxBlocksPerFile);
    var weights = getModeWeights(settings.mode);

    var bestPlan = null;

    for (var p = 0; p < partitions.length; p++) {
        var part = partitions[p];
        var layouts = [];
        var failed = false;

        for (var i = 0; i < part.length; i++) {
            var grid = chooseBestGridForCount(src, part[i], settings);
            if (!grid) {
                failed = true;
                break;
            }
            layouts.push(grid);
        }

        if (failed) continue;

        var totalArea = 0;
        var totalEmpty = 0;
        var totalAspectPenalty = 0;
        var totalRotPenalty = 0;
        var maxLong = 0;

        for (var j = 0; j < layouts.length; j++) {
            totalArea += layouts[j].blockW * layouts[j].blockH;
            totalEmpty += layouts[j].emptyCells;
            var ratio = layouts[j].ratio;
            var excessAspect = Math.max(0, ratio - settings.maxBlockAspectRatio);
            totalAspectPenalty += excessAspect * excessAspect;
            if (layouts[j].rotatedInside) totalRotPenalty += 1;
            var ls = longSide(layouts[j].blockW, layouts[j].blockH);
            if (ls > maxLong) maxLong = ls;
        }

        var planScore =
            (layouts.length * weights.planBlockCount) +
            (totalEmpty * weights.planEmptyCells) +
            (totalAspectPenalty * weights.planAspectPenalty) +
            (totalArea * weights.planBlockArea) +
            (maxLong * weights.planLongSide) +
            (totalRotPenalty * weights.planRotPenalty);

        if (!bestPlan || planScore < bestPlan.score) {
            bestPlan = {
                source: src,
                score: planScore,
                partition: part,
                layouts: layouts
            };
        }
    }

    return bestPlan;
}

function buildAllBlocksFromSources(sources, settings) {
    var blocks = [];
    var failedNames = [];

    for (var i = 0; i < sources.length; i++) {
        var src = sources[i];
        var plan = buildBestBlockPlanForSource(src, settings);

        if (!plan) {
            failedNames.push(src.name);
            continue;
        }

        for (var b = 0; b < plan.layouts.length; b++) {
            var g = plan.layouts[b];
            blocks.push({
                sourceId: src.id,
                sourceRef: src.ref,
                sourceName: src.name,

                count: g.count,
                cols: g.cols,
                rows: g.rows,
                emptyCells: g.emptyCells,

                rotatedInside: g.rotatedInside,
                cellW: g.cellW,
                cellH: g.cellH,

                blockW: g.blockW,
                blockH: g.blockH,
                area: g.blockW * g.blockH,
                longSide: longSide(g.blockW, g.blockH),
                shortSide: shortSide(g.blockW, g.blockH),
                ratio: g.ratio
            });
        }
    }

    if (settings.mode === "Efficient") {
        blocks.sort(function(a, b) {
            if (b.longSide !== a.longSide) return b.longSide - a.longSide;
            return b.area - a.area;
        });
    } else if (settings.mode === "Clean") {
        blocks.sort(function(a, b) {
            if (a.ratio !== b.ratio) return a.ratio - b.ratio; // less strip-like first
            if (b.blockH !== a.blockH) return b.blockH - a.blockH;
            return b.area - a.area;
        });
    } else {
        blocks.sort(function(a, b) {
            if (b.area !== a.area) return b.area - a.area;
            return b.longSide - a.longSide;
        });
    }

    return {
        blocks: blocks,
        failedNames: failedNames
    };
}

// =====================================================
// MaxRects-like nesting for blocks
// =====================================================

function estimateSplitCount(freeRect, usedRect) {
    var count = 0;

    var frLeft = freeRect.x;
    var frTop = freeRect.y;
    var frRight = freeRect.x + freeRect.w;
    var frBottom = freeRect.y + freeRect.h;

    var usedLeft = usedRect.x;
    var usedTop = usedRect.y;
    var usedRight = usedRect.x + usedRect.w;
    var usedBottom = usedRect.y + usedRect.h;

    if (usedTop > frTop + EPS) count++;
    if (usedBottom < frBottom - EPS) count++;
    if (usedLeft > frLeft + EPS) count++;
    if (usedRight < frRight - EPS) count++;

    return count;
}

function estimateAlignmentBonus(freeRect, pw, ph, placedSoFar, sheetWidth) {
    var bonus = 0;
    var x = freeRect.x;
    var y = freeRect.y;
    var right = x + pw;
    var bottom = y + ph;

    if (Math.abs(x) < EPS) bonus += 2;
    if (Math.abs(y) < EPS) bonus += 2;
    if (Math.abs(right - sheetWidth) < EPS) bonus += 2;

    for (var i = 0; i < placedSoFar.length; i++) {
        var p = placedSoFar[i];
        if (Math.abs(x - p.x) < EPS) bonus += 1;
        if (Math.abs(y - p.y) < EPS) bonus += 1;
        if (Math.abs(right - (p.x + p.paddedW)) < EPS) bonus += 1;
        if (Math.abs(bottom - (p.y + p.paddedH)) < EPS) bonus += 1;
    }
    return bonus;
}

function computePlacementScore(freeRect, pw, ph, currentUsedLength, placedSoFar, settings) {
    var weights = getModeWeights(settings.mode);

    var leftoverHoriz = freeRect.w - pw;
    var leftoverVert = freeRect.h - ph;
    var shortFit = Math.min(leftoverHoriz, leftoverVert);
    var longFit = Math.max(leftoverHoriz, leftoverVert);
    var areaFit = (freeRect.w * freeRect.h) - (pw * ph);

    var projectedBottom = freeRect.y + ph;
    var projectedUsedLength = (projectedBottom > currentUsedLength) ? projectedBottom : currentUsedLength;

    var fragPenalty = estimateSplitCount(freeRect, {
        x: freeRect.x,
        y: freeRect.y,
        w: pw,
        h: ph
    });

    var alignBonus = estimateAlignmentBonus(freeRect, pw, ph, placedSoFar, inToPt(settings.sheetWidthIn));

    var total =
        (shortFit * weights.nestShortFit) +
        (areaFit * weights.nestAreaFit) +
        (longFit * weights.nestLongFit) +
        (projectedUsedLength * weights.nestProjectedLen) +
        (fragPenalty * weights.nestFragmentPenalty) -
        (alignBonus * weights.nestAlignBonus);

    return { total: total };
}

function findBestBlockPlacement(block, freeRects, currentUsedLength, settings, placedSoFar) {
    var spacing = inToPt(settings.spacingIn);
    var orientations = [{
        rotatedOnSheet: false,
        w: block.blockW,
        h: block.blockH
    }];

    if (settings.allowBlockRotationOnSheet) {
        orientations.push({
            rotatedOnSheet: true,
            w: block.blockH,
            h: block.blockW
        });
    }

    var best = null;

    for (var o = 0; o < orientations.length; o++) {
        var ori = orientations[o];
        var paddedW = ori.w + spacing;
        var paddedH = ori.h + spacing;

        for (var i = 0; i < freeRects.length; i++) {
            var fr = freeRects[i];

            if (paddedW <= fr.w + EPS && paddedH <= fr.h + EPS) {
                var score = computePlacementScore(fr, paddedW, paddedH, currentUsedLength, placedSoFar, settings);
                if (!best || score.total < best.score.total) {
                    best = {
                        x: fr.x,
                        y: fr.y,
                        w: paddedW,
                        h: paddedH,
                        rotatedOnSheet: ori.rotatedOnSheet,
                        score: score
                    };
                }
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

    if (usedTop > frTop + EPS) {
        result.push({ x: frLeft, y: frTop, w: freeRect.w, h: usedTop - frTop });
    }
    if (usedBottom < frBottom - EPS) {
        result.push({ x: frLeft, y: usedBottom, w: freeRect.w, h: frBottom - usedBottom });
    }
    if (usedLeft > frLeft + EPS) {
        result.push({ x: frLeft, y: frTop, w: usedLeft - frLeft, h: freeRect.h });
    }
    if (usedRight < frRight - EPS) {
        result.push({ x: usedRight, y: frTop, w: frRight - usedRight, h: freeRect.h });
    }

    return result;
}

function pruneFreeRects(freeRects) {
    var pruned = [];

    for (var i = 0; i < freeRects.length; i++) {
        if (isValidRect(freeRects[i])) pruned.push(freeRects[i]);
    }

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

function placeAndUpdateFreeRects(freeRects, placement) {
    var usedRect = { x: placement.x, y: placement.y, w: placement.w, h: placement.h };
    var newRects = [];

    for (var i = 0; i < freeRects.length; i++) {
        var splitRects = splitFreeRect(freeRects[i], usedRect);
        for (var j = 0; j < splitRects.length; j++) {
            if (isValidRect(splitRects[j])) newRects.push(splitRects[j]);
        }
    }
    return pruneFreeRects(newRects);
}

function nestBlocks(blocks, settings) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);

    var freeRects = [{ x: 0, y: 0, w: sheetWidth, h: maxLength }];
    var placed = [];
    var unplaced = [];
    var usedLength = 0;

    for (var i = 0; i < blocks.length; i++) {
        var block = blocks[i];
        var p = findBestBlockPlacement(block, freeRects, usedLength, settings, placed);

        if (!p) {
            unplaced.push(block);
            continue;
        }

        placed.push({
            block: block,
            x: p.x,
            y: p.y,
            paddedW: p.w,
            paddedH: p.h,
            rotatedOnSheet: p.rotatedOnSheet
        });

        freeRects = placeAndUpdateFreeRects(freeRects, p);

        var bottom = p.y + p.h;
        if (bottom > usedLength) usedLength = bottom;
    }

    return {
        placed: placed,
        unplaced: unplaced,
        usedLength: usedLength
    };
}

// =====================================================
// Rendering
// =====================================================

function duplicateToLayer(item, targetLayer) {
    return item.duplicate(targetLayer, ElementPlacement.PLACEATEND);
}

function setItemTopLeftByVisibleBounds(item, targetLeft, targetTop) {
    var vb = item.visibleBounds;
    var dx = targetLeft - vb[0];
    var dy = targetTop - vb[1];
    item.translate(dx, dy);
}

function placeSingleCopy(refItem, targetLayer, x, y, rotationDeg) {
    var dup = duplicateToLayer(refItem, targetLayer);
    if (rotationDeg !== 0) dup.rotate(rotationDeg);

    var targetLeft = x;
    var targetTop = -y;
    setItemTopLeftByVisibleBounds(dup, targetLeft, targetTop);
    return dup;
}

function getLocalCellPosition(block, index, spacing) {
    var row = Math.floor(index / block.cols);
    var col = index % block.cols;
    return {
        x: col * (block.cellW + spacing),
        y: row * (block.cellH + spacing)
    };
}

function transformLocalForBlockRotation(localX, localY, cellW, cellH, blockW, blockH, rotatedOnSheet) {
    if (!rotatedOnSheet) {
        return { x: localX, y: localY };
    }

    // rotate whole block 90° clockwise around top-left
    return {
        x: blockH - localY - cellH,
        y: localX
    };
}

function renderPlacedBlocks(layout, outputLayer, settings) {
    var spacing = inToPt(settings.spacingIn);

    for (var i = 0; i < layout.placed.length; i++) {
        var pb = layout.placed[i];
        var b = pb.block;

        for (var n = 0; n < b.count; n++) {
            var local = getLocalCellPosition(b, n, spacing);
            var tr = transformLocalForBlockRotation(
                local.x,
                local.y,
                b.cellW,
                b.cellH,
                b.blockW,
                b.blockH,
                pb.rotatedOnSheet
            );

            var itemX = pb.x + tr.x;
            var itemY = pb.y + tr.y;

            var rotationDeg = 0;
            if (b.rotatedInside) rotationDeg += 90;
            if (pb.rotatedOnSheet) rotationDeg += 90;
            rotationDeg = rotationDeg % 360;

            placeSingleCopy(b.sourceRef, outputLayer, itemX, itemY, rotationDeg);
        }
    }
}

// =====================================================
// Summary
// =====================================================

function countPlacedCopies(layout) {
    var total = 0;
    for (var i = 0; i < layout.placed.length; i++) total += layout.placed[i].block.count;
    return total;
}

function countUnplacedCopies(layout) {
    var total = 0;
    for (var i = 0; i < layout.unplaced.length; i++) total += layout.unplaced[i].count;
    return total;
}

function buildSummary(sources, blocksInfo, layout, settings) {
    var usedNoTrailingGap = Math.max(0, layout.usedLength - inToPt(settings.spacingIn));
    var lines = [];
    lines.push("WeMust NESTER v5");
    lines.push("Mode: " + settings.mode);
    lines.push("Max blocks/file: " + settings.maxBlocksPerFile);
    lines.push("Max block ratio: " + settings.maxBlockAspectRatio);
    lines.push("Placed copies: " + countPlacedCopies(layout));
    lines.push("Unplaced copies: " + countUnplacedCopies(layout));
    lines.push("Used length: " + inToStr(usedNoTrailingGap) + " in");
    lines.push("Output layer: " + OUTPUT_LAYER_NAME);

    if (blocksInfo.failedNames.length > 0) {
        lines.push("Some items could not form valid blocks: " + blocksInfo.failedNames.length);
    }

    return lines.join("\n");
}

// =====================================================
// Build pipeline
// =====================================================

function buildOnce(doc, settings) {
    var sources = collectPlacedItems(doc);

    if (sources.length === 0) {
        throw new Error("No placed items found in the current document.");
    }

    var outputLayer = prepareOutputLayer(doc, settings.previousOutputAction);
    var blocksInfo = buildAllBlocksFromSources(sources, settings);

    if (blocksInfo.blocks.length === 0) {
        throw new Error("No valid block plans could be built.");
    }

    var layout = nestBlocks(blocksInfo.blocks, settings);
    renderPlacedBlocks(layout, outputLayer, settings);

    if (settings.hideSourceLayersAfterBuild) {
        hideSourceLayers(doc);
    }

    return {
        sources: sources,
        blocksInfo: blocksInfo,
        layout: layout
    };
}

// =====================================================
// Main loop
// =====================================================

function main() {
    if (app.documents.length === 0) {
        alert("No open document found.");
        return;
    }

    var doc = app.activeDocument;
    var current = {
        sheetWidthIn: DEFAULTS.sheetWidthIn,
        maxLengthIn: DEFAULTS.maxLengthIn,
        spacingIn: DEFAULTS.spacingIn,
        mode: DEFAULTS.mode,
        maxBlocksPerFile: DEFAULTS.maxBlocksPerFile,
        maxBlockAspectRatio: DEFAULTS.maxBlockAspectRatio,
        allowItemRotationInBlock: DEFAULTS.allowItemRotationInBlock,
        allowBlockRotationOnSheet: DEFAULTS.allowBlockRotationOnSheet,
        previousOutputAction: DEFAULTS.previousOutputAction,
        hideSourceLayersAfterBuild: DEFAULTS.hideSourceLayersAfterBuild
    };

    while (true) {
        var settings = showSettingsDialog(current);
        if (!settings) break;

        current = settings;

        try {
            var result = buildOnce(doc, settings);
            var summary = buildSummary(result.sources, result.blocksInfo, result.layout, settings);
            var again = confirm(summary + "\n\nBuild another variation?");
            if (!again) break;
        } catch (err) {
            var retry = confirm("Error:\n" + err + "\n\nTry again with new settings?");
            if (!retry) break;
        }
    }
}

main();