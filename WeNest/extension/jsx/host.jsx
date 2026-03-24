#target illustrator

// =====================================================
// WeMust NESTER v6
// Human-like production nesting
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
    optimizePreset: "Auto",         // Auto | Compact | Balanced | Ordered
    searchEffort: "Normal",         // Normal | High
    mode: "Production",             // kept for legacy ScriptUI dialog
    orderDiscipline: 45,            // 0..100 (0=aggressive, 100=ordered)
    maxBlocksPerFile: 2,            // 1..4
    maxBlockAspectRatio: 3.0,       // lower = less skinny strips
    allowItemRotationInBlock: true,
    allowBlockRotationOnSheet: true,
    hideSourceLayersAfterBuild: true,

    // v6
    useDominantShelf: true,
    dominantShelfMaxRows: 6,
    widthFillPriority: 92,          // internally kept high to fill width
    sameSourceProximity: 55,        // derived from orderDiscipline
    cavityPenalty: 68               // derived from orderDiscipline
};

// =====================================================
// Basic helpers
// =====================================================

function trimStr(s) { return String(s).replace(/^\s+|\s+$/g, ""); }

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

function getParentFolderName(pathOrName) {
    var s = trimStr(pathOrName || "");
    if (!s) return "";

    s = s.replace(/\\/g, "/");
    if (s.charAt(s.length - 1) === "/") {
        s = s.substring(0, s.length - 1);
    }

    var parts = s.split("/");
    return (parts.length > 1) ? parts[parts.length - 2] : "";
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

function getItemFilePath(item) {
    try {
        if (item.file && item.file.fsName) return String(item.file.fsName);
    } catch (e) {}
    return "";
}

function roundForKey(v) {
    return Math.round(v * 1000) / 1000;
}

function buildSourceKey(name, width, height, ordinal) {
    return [
        String(name),
        roundForKey(width),
        roundForKey(height),
        ordinal
    ].join("|");
}

function getVisibleSize(item) {
    var vb = item.visibleBounds; // [left, top, right, bottom]
    var w = Math.abs(vb[2] - vb[0]);
    var h = Math.abs(vb[1] - vb[3]);
    return { width: w, height: h };
}

function inToPt(v) { return v * PT_PER_IN; }
function inToStr(vPt) { return (vPt / PT_PER_IN).toFixed(2); }

function areaOf(w, h) { return w * h; }
function longSide(w, h) { return (w > h) ? w : h; }
function shortSide(w, h) { return (w < h) ? w : h; }

function cloneRect(r) { return { x: r.x, y: r.y, w: r.w, h: r.h }; }

function startsWith(str, prefix) { return String(str).indexOf(prefix) === 0; }

function uniquePushPartition(list, arr) {
    var key = arr.join("-");
    for (var i = 0; i < list.length; i++) {
        if (list[i].join("-") === key) return;
    }
    list.push(arr);
}

function clamp(v, minV, maxV) {
    return Math.max(minV, Math.min(maxV, v));
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
    var dlg = new Window("dialog", "WeMust NESTER v6");
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
    var ddMode = g4.add("dropdownlist", undefined, ["Efficient", "Balanced", "Clean", "Production"]);
    ddMode.selection = (current.mode === "Efficient") ? 0 : (current.mode === "Clean" ? 2 : (current.mode === "Production" ? 3 : 1));

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

    var p3 = dlg.add("panel", undefined, "Production Bias");
    p3.orientation = "column";
    p3.alignChildren = "left";

    var cbDomShelf = p3.add("checkbox", undefined, "Use dominant item shelf");
    cbDomShelf.value = current.useDominantShelf;

    var g7 = p3.add("group");
    g7.add("statictext", undefined, "Dominant shelf max rows:");
    var ddRows = g7.add("dropdownlist", undefined, ["2", "3", "4", "5", "6", "7", "8"]);
    ddRows.selection = Math.max(0, Math.min(6, current.dominantShelfMaxRows - 2));

    var g8 = p3.add("group");
    g8.add("statictext", undefined, "Width fill priority:");
    var etWidthBias = g8.add("edittext", undefined, String(current.widthFillPriority));
    etWidthBias.characters = 5;

    var g9 = p3.add("group");
    g9.add("statictext", undefined, "Same-source proximity:");
    var etProx = g9.add("edittext", undefined, String(current.sameSourceProximity));
    etProx.characters = 5;

    var g10 = p3.add("group");
    g10.add("statictext", undefined, "Cavity penalty:");
    var etCavity = g10.add("edittext", undefined, String(current.cavityPenalty));
    etCavity.characters = 5;
var note = dlg.add("statictext", undefined, "Tip: Production + dominant shelf + width fill 70–90 often gives the most human-like result.");
    note.characters = 80;

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
        orderDiscipline: current.orderDiscipline,
        maxBlocksPerFile: ddBlocks.selection ? parseInt(ddBlocks.selection.text, 10) : current.maxBlocksPerFile,
        maxBlockAspectRatio: parseNum(etAspect.text, current.maxBlockAspectRatio),
        allowItemRotationInBlock: cbItemRot.value,
        allowBlockRotationOnSheet: cbBlockRot.value,
        hideSourceLayersAfterBuild: current.hideSourceLayersAfterBuild,
        useDominantShelf: cbDomShelf.value,
        dominantShelfMaxRows: ddRows.selection ? parseInt(ddRows.selection.text, 10) : current.dominantShelfMaxRows,
        widthFillPriority: clamp(parseNum(etWidthBias.text, current.widthFillPriority), 0, 100),
        sameSourceProximity: clamp(parseNum(etProx.text, current.sameSourceProximity), 0, 100),
        cavityPenalty: clamp(parseNum(etCavity.text, current.cavityPenalty), 0, 100)
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

function prepareOutputLayer(doc) {
    var existing = findLayerByName(doc, OUTPUT_LAYER_NAME);
    if (existing) {
        try { existing.remove(); } catch (e1) {}
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
// =====================================================

function collectPlacedItems(doc, quantityOverrides) {
    var result = [];
    var ordinal = 0;

    for (var i = 0; i < doc.placedItems.length; i++) {
        var it = doc.placedItems[i];

        try {
            if (it.layer && isOutputLayerName(it.layer.name)) continue;
        } catch (e0) {}

        var name = getItemFileName(it);
        var filePath = getItemFilePath(it);
        var sz = getVisibleSize(it);
        var key = buildSourceKey(name, sz.width, sz.height, ordinal);
        var baseQty = parseQtyFromName(name);
        var overrideQty = quantityOverrides && quantityOverrides.hasOwnProperty(key)
            ? quantityOverrides[key]
            : null;
        var qty = (overrideQty === null || overrideQty === undefined)
            ? baseQty
            : Math.max(1, overrideQty);

        result.push({
            id: "SRC_" + ordinal,
            key: key,
            ref: it,
            name: name,
            filePath: filePath,
            baseQty: baseQty,
            qty: qty,
            width: sz.width,
            height: sz.height,
            area: areaOf(sz.width, sz.height),
            longSide: longSide(sz.width, sz.height),
            shortSide: shortSide(sz.width, sz.height)
        });
        ordinal += 1;
    }

    return result;
}

// =====================================================
// Weights
// =====================================================

function getModeWeights(settings) {
    var orderRaw = parseFloat(settings.orderDiscipline);
    if (isNaN(orderRaw)) orderRaw = DEFAULTS.orderDiscipline;
    var order = clamp(orderRaw, 0, 100) / 100.0;

    return {
        // Block planning (smaller = better)
        planBlockCount: 2200 + (order * 1800),
        planEmptyCells: 2200 + (order * 700),
        planAspectPenalty: 11000 + (order * 16000),
        planBlockArea: 1,
        planLongSide: 6 + (order * 5),
        planRotPenalty: 300 + (order * 700),

        // Nesting objective:
        // 1) minimize projected length aggressively
        // 2) maximize width fill
        // 3) reduce disorder/fragments based on order control
        nestProjectedLen: 1600 + (order * 600),
        nestWidthWaste: 25000,
        nestXOffset: 1200 + (order * 3800),
        nestFragmentPenalty: 300 + (order * 900),
        nestCavityPenalty: 1500 + (order * 1500),
        nestAlignBonus: 150 + (order * 450),
        nestSourceBonus: 22000 + (order * 38000)
    };
}

// =====================================================
// Partition generation
// =====================================================

function balancedPartition(qty, parts) {
    var arr = [];
    var base = Math.floor(qty / parts);
    var rem = qty % parts;
    for (var i = 0; i < parts; i++) arr.push(base + (i < rem ? 1 : 0));
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
            if (qty >= 12) uniquePushPartition(out, shiftedPartition(qty, parts, 3));
            if (qty >= 18) uniquePushPartition(out, shiftedPartition(qty, parts, 4));
        }
    }

    return out;
}

// =====================================================
// Block planning
// =====================================================

function chooseBestGridForCount(src, count, settings, favorWide) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);
    var spacing = inToPt(settings.spacingIn);
    var weights = getModeWeights(settings);
    var maxAspect = settings.maxBlockAspectRatio;
    var widthBias = settings.widthFillPriority / 100.0;

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

            if (favorWide && rows > settings.dominantShelfMaxRows) continue;

            var blockW = cols * ori.cellW + (cols - 1) * spacing;
            var blockH = rows * ori.cellH + (rows - 1) * spacing;

            if (blockW > sheetWidth + EPS) continue;
            if (blockH > maxLength + EPS) continue;

            var ratio = longSide(blockW, blockH) / Math.max(shortSide(blockW, blockH), EPS);
            var excessAspect = Math.max(0, ratio - maxAspect);
            var aspectPenalty = excessAspect * excessAspect;

            var widthFill = blockW / Math.max(sheetWidth, EPS);
            var widthWaste = 1 - widthFill;
            var widthBonus = favorWide ? widthFill * 17000 * widthBias : widthFill * 9000 * widthBias;

            var widePenalty = favorWide ? (Math.max(0, rows - settings.dominantShelfMaxRows) * 10000) : 0;
            var heightPenalty = blockH * 2200;

            var score =
                (emptyCells * weights.planEmptyCells) +
                (aspectPenalty * weights.planAspectPenalty) +
                (widthWaste * 45000) +
                (heightPenalty) +
                ((blockW * blockH) * weights.planBlockArea) +
                (longSide(blockW, blockH) * weights.planLongSide) +
                ((ori.rotated ? 1 : 0) * weights.planRotPenalty) +
                widePenalty -
                widthBonus;

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
                widthFill: widthFill,
                score: score
            });
        }
    }

    if (candidates.length === 0) return null;

    candidates.sort(function(a, b) {
        if (a.score !== b.score) return a.score - b.score;
        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        if (a.emptyCells !== b.emptyCells) return a.emptyCells - b.emptyCells;
        return (a.blockW * a.blockH) - (b.blockW * b.blockH);
    });

    return candidates[0];
}

function buildCandidateBlockPlansForSource(src, settings, isDominant) {
    var dynamicMaxBlocks = settings.maxBlocksPerFile;
    if (settings.searchEffort === "High") dynamicMaxBlocks = Math.max(dynamicMaxBlocks, 4);
    else dynamicMaxBlocks = Math.max(dynamicMaxBlocks, 3);
    dynamicMaxBlocks = Math.min(dynamicMaxBlocks, src.qty);

    var keepTop = (settings.searchEffort === "High") ? 3 : 2;
    var partitions = generatePartitionCandidates(src.qty, dynamicMaxBlocks);
    var weights = getModeWeights(settings);
    var plans = [];
    var seen = {};

    for (var p = 0; p < partitions.length; p++) {
        var part = partitions[p];
        var layouts = [];
        var failed = false;

        for (var i = 0; i < part.length; i++) {
            var grid = chooseBestGridForCount(src, part[i], settings, isDominant && settings.useDominantShelf);
            if (!grid) {
                failed = true;
                break;
            }
            layouts.push(grid);
        }

        if (failed) continue;

        var totalArea = 0;
        var totalBlockHeight = 0;
        var totalEmpty = 0;
        var totalAspectPenalty = 0;
        var totalRotPenalty = 0;
        var maxLong = 0;
        var totalWidthFill = 0;
        var shapeKeyParts = [];

        for (var j = 0; j < layouts.length; j++) {
            totalArea += layouts[j].blockW * layouts[j].blockH;
            totalEmpty += layouts[j].emptyCells;
            var ratio = layouts[j].ratio;
            var excessAspect = Math.max(0, ratio - settings.maxBlockAspectRatio);
            totalAspectPenalty += excessAspect * excessAspect;
            if (layouts[j].rotatedInside) totalRotPenalty += 1;
            totalBlockHeight += layouts[j].blockH;
            var ls = longSide(layouts[j].blockW, layouts[j].blockH);
            if (ls > maxLong) maxLong = ls;
            totalWidthFill += layouts[j].widthFill;
            shapeKeyParts.push([layouts[j].count, layouts[j].cols, layouts[j].rows, layouts[j].rotatedInside ? 1 : 0].join(':'));
        }

        var planKey = shapeKeyParts.join('|');
        if (seen[planKey]) continue;
        seen[planKey] = true;

        var widthFillBias = settings.widthFillPriority / 100.0;
        var dominantBonus = (isDominant && settings.useDominantShelf) ? (totalWidthFill * 18000 * widthFillBias) : 0;

        var planScore =
            (layouts.length * weights.planBlockCount) +
            (totalEmpty * weights.planEmptyCells) +
            (totalAspectPenalty * weights.planAspectPenalty) +
            (totalBlockHeight * 1800) +
            (totalArea * weights.planBlockArea) +
            (maxLong * weights.planLongSide) +
            (totalRotPenalty * weights.planRotPenalty) -
            dominantBonus;

        plans.push({
            source: src,
            score: planScore,
            partition: part,
            layouts: layouts,
            isDominant: isDominant
        });
    }

    plans.sort(function(a, b) {
        if (a.score !== b.score) return a.score - b.score;
        if (a.layouts.length !== b.layouts.length) return a.layouts.length - b.layouts.length;
        return a.partition.length - b.partition.length;
    });

    return plans.slice(0, keepTop);
}
function detectDominantSource(sources) {
    if (!sources || sources.length === 0) return null;
    var sorted = sources.slice(0);
    sorted.sort(function(a, b) {
        if (b.qty !== a.qty) return b.qty - a.qty;
        return b.area - a.area;
    });
    return sorted[0].id;
}

function sortBlocksForSolve(blocks, settings) {
    var orderRaw = parseFloat(settings.orderDiscipline);
    if (isNaN(orderRaw)) orderRaw = DEFAULTS.orderDiscipline;
    var orderBias = clamp(orderRaw, 0, 100) / 100.0;

    blocks.sort(function(a, b) {
        if (a.isDominant !== b.isDominant) return a.isDominant ? -1 : 1;

        if (orderBias > 0.55 && a.sourceId !== b.sourceId) {
            if (a.sourceName < b.sourceName) return -1;
            if (a.sourceName > b.sourceName) return 1;
        }

        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        if (a.blockH !== b.blockH) return a.blockH - b.blockH;
        if (a.ratio !== b.ratio) return a.ratio - b.ratio;
        return b.area - a.area;
    });
}

function materializeBlocksFromPlans(selectedPlans, settings, dominantId, failedNames) {
    var blocks = [];
    var blockUidCounter = 1;

    for (var i = 0; i < selectedPlans.length; i++) {
        var plan = selectedPlans[i];
        if (!plan) continue;

        for (var b = 0; b < plan.layouts.length; b++) {
            var g = plan.layouts[b];
            blocks.push({
                uid: "B" + (blockUidCounter++),
                sourceId: plan.source.id,
                sourceKey: plan.source.key,
                sourceRef: plan.source.ref,
                sourceName: plan.source.name,
                sourceBaseQty: plan.source.baseQty,
                sourceQty: plan.source.qty,
                sourceWidth: plan.source.width,
                sourceHeight: plan.source.height,
                sourceArea: plan.source.area,
                sourceLongSide: plan.source.longSide,
                sourceShortSide: plan.source.shortSide,
                planScore: plan.score,
                planBlockCount: plan.layouts.length,

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
                ratio: g.ratio,
                widthFill: g.widthFill,
                isDominant: plan.isDominant
            });
        }
    }

    sortBlocksForSolve(blocks, settings);

    return {
        blocks: blocks,
        failedNames: failedNames ? failedNames.slice(0) : [],
        dominantId: dominantId,
        selectedPlans: selectedPlans.slice(0)
    };
}

function collectSourcePlanSets(sources, settings) {
    var failedNames = [];
    var dominantId = settings.useDominantShelf ? detectDominantSource(sources) : null;
    var planSets = [];

    for (var i = 0; i < sources.length; i++) {
        var src = sources[i];
        var isDominant = (src.id === dominantId);
        var plans = buildCandidateBlockPlansForSource(src, settings, isDominant);

        if (!plans || plans.length === 0) {
            failedNames.push(src.name);
            continue;
        }

        planSets.push({
            source: src,
            isDominant: isDominant,
            plans: plans
        });
    }

    return {
        planSets: planSets,
        failedNames: failedNames,
        dominantId: dominantId
    };
}

function buildPlanCombinations(planSets, settings) {
    if (!planSets || planSets.length === 0) return [];

    var beamWidth = (settings.searchEffort === "High") ? 4 : 2;
    var partials = [{ plans: [], heuristicScore: 0, totalLayouts: 0 }];

    for (var i = 0; i < planSets.length; i++) {
        var next = [];
        var plans = planSets[i].plans;

        for (var p = 0; p < partials.length; p++) {
            for (var j = 0; j < plans.length; j++) {
                next.push({
                    plans: partials[p].plans.concat([plans[j]]),
                    heuristicScore: partials[p].heuristicScore + plans[j].score,
                    totalLayouts: partials[p].totalLayouts + plans[j].layouts.length
                });
            }
        }

        next.sort(function(a, b) {
            if (a.heuristicScore !== b.heuristicScore) return a.heuristicScore - b.heuristicScore;
            return a.totalLayouts - b.totalLayouts;
        });

        if (next.length > beamWidth) next = next.slice(0, beamWidth);
        partials = next;
    }

    return partials;
}

function buildAllBlocksFromSources(sources, settings) {
    var planInfo = collectSourcePlanSets(sources, settings);
    var planCombos = buildPlanCombinations(planInfo.planSets, settings);

    if (planCombos.length === 0) {
        return {
            blocks: [],
            failedNames: planInfo.failedNames,
            dominantId: planInfo.dominantId,
            selectedPlans: []
        };
    }

    return materializeBlocksFromPlans(
        planCombos[0].plans,
        settings,
        planInfo.dominantId,
        planInfo.failedNames
    );
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

function estimateSourceProximityBonus(freeRect, pw, ph, placedSoFar, sourceId) {
    var bonus = 0;
    var x = freeRect.x;
    var y = freeRect.y;
    var cx = x + pw / 2;
    var cy = y + ph / 2;

    for (var i = 0; i < placedSoFar.length; i++) {
        var p = placedSoFar[i];
        if (p.block.sourceId !== sourceId) continue;

        var pcx = p.x + p.paddedW / 2;
        var pcy = p.y + p.paddedH / 2;
        var dx = cx - pcx;
        var dy = cy - pcy;
        var dist = Math.sqrt(dx * dx + dy * dy);

        bonus += 1 / Math.max(dist, 1);
    }
    return bonus;
}

function estimateWidthFillBonus(freeRect, pw, sheetWidth) {
    var fillRatio = pw / Math.max(sheetWidth, EPS);
    return fillRatio;
}

function estimateCavityPenalty(freeRect, pw, ph) {
    var rightW = freeRect.w - pw;
    var bottomH = freeRect.h - ph;

    var penalty = 0;

    if (rightW > EPS && freeRect.h > EPS) {
        var ratio1 = longSide(rightW, freeRect.h) / Math.max(shortSide(rightW, freeRect.h), EPS);
        penalty += ratio1;
    }

    if (freeRect.w > EPS && bottomH > EPS) {
        var ratio2 = longSide(freeRect.w, bottomH) / Math.max(shortSide(freeRect.w, bottomH), EPS);
        penalty += ratio2;
    }

    return penalty;
}

function blockFitsFreeRect(block, rect, settings) {
    var spacing = inToPt(settings.spacingIn);
    var options = [
        { w: block.blockW + spacing, h: block.blockH + spacing }
    ];

    if (settings.allowBlockRotationOnSheet) {
        options.push({ w: block.blockH + spacing, h: block.blockW + spacing });
    }

    for (var i = 0; i < options.length; i++) {
        if (options[i].w <= rect.w + EPS && options[i].h <= rect.h + EPS) {
            return options[i];
        }
    }

    return null;
}

function estimateFutureFitBonus(nextFreeRects, remainingBlocks, settings) {
    if (!remainingBlocks || remainingBlocks.length === 0) return 0;

    var checkedBlocks = Math.min(remainingBlocks.length, 6);
    var bestBonus = 0;
    var totalBonus = 0;
    var sheetWidth = inToPt(settings.sheetWidthIn);

    for (var i = 0; i < checkedBlocks; i++) {
        var block = remainingBlocks[i];
        var blockBest = 0;

        for (var j = 0; j < nextFreeRects.length; j++) {
            var fr = nextFreeRects[j];
            var fit = blockFitsFreeRect(block, fr, settings);
            if (!fit) continue;

            var widthSlack = Math.max(0, fr.w - fit.w);
            var snugWidth = 1 - clamp(widthSlack / Math.max(sheetWidth, EPS), 0, 1);
            var topBias = 1 / (1 + (fr.y / (PT_PER_IN * 8)));
            var corridorBias = (fr.x + fr.w >= sheetWidth - EPS) ? 1.2 : 1.0;
            var match = (block.area * 0.015) * (0.45 + 0.35 * snugWidth + 0.20 * topBias) * corridorBias;

            if (match > blockBest) blockBest = match;
        }

        totalBonus += blockBest;
        if (blockBest > bestBonus) bestBonus = blockBest;
    }

    return (bestBonus * 2.2) + (totalBonus * 0.55);
}

function computePlacementScore(freeRect, pw, ph, currentUsedLength, placedSoFar, settings, block, futureFitBonus) {
    var weights = getModeWeights(settings);
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var projectedBottom = freeRect.y + ph;
    var projectedUsedLength = (projectedBottom > currentUsedLength) ? projectedBottom : currentUsedLength;
    var projectedUsedIn = projectedUsedLength / PT_PER_IN;

    var fragPenalty = estimateSplitCount(freeRect, {
        x: freeRect.x,
        y: freeRect.y,
        w: pw,
        h: ph
    });

    var alignBonus = estimateAlignmentBonus(freeRect, pw, ph, placedSoFar, sheetWidth);
    var sourceBonus = estimateSourceProximityBonus(freeRect, pw, ph, placedSoFar, block.sourceId);
    var widthFill = estimateWidthFillBonus(freeRect, pw, sheetWidth);
    var widthWaste = 1 - widthFill;
    var cavityPenalty = estimateCavityPenalty(freeRect, pw, ph);
    var xOffsetRatio = Math.abs(freeRect.x) / Math.max(sheetWidth, EPS);

    var widthBias = settings.widthFillPriority / 100.0;
    var proxBias = settings.sameSourceProximity / 100.0;
    var cavityBias = settings.cavityPenalty / 100.0;
    var orderRaw = parseFloat(settings.orderDiscipline);
    if (isNaN(orderRaw)) orderRaw = DEFAULTS.orderDiscipline;
    var orderBias = clamp(orderRaw, 0, 100) / 100.0;

    var total =
        (projectedUsedIn * weights.nestProjectedLen) +
        (widthWaste * weights.nestWidthWaste * (0.7 + 0.3 * widthBias)) +
        (xOffsetRatio * weights.nestXOffset) +
        (fragPenalty * weights.nestFragmentPenalty) -
        (alignBonus * weights.nestAlignBonus * (0.5 + 0.5 * orderBias)) -
        (futureFitBonus || 0) -
        (sourceBonus * weights.nestSourceBonus * proxBias) +
        (cavityPenalty * weights.nestCavityPenalty * cavityBias);

    return {
        total: total,
        projectedUsedIn: projectedUsedIn,
        widthWaste: widthWaste
    };
}

function findBestBlockPlacement(block, freeRects, currentUsedLength, settings, placedSoFar, remainingBlocks) {
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
                var placement = {
                    x: fr.x,
                    y: fr.y,
                    w: paddedW,
                    h: paddedH,
                    rotatedOnSheet: ori.rotatedOnSheet
                };
                var nextFreeRects = placeAndUpdateFreeRects(freeRects, placement);
                var futureFitBonus = estimateFutureFitBonus(nextFreeRects, remainingBlocks, settings);
                var score = computePlacementScore(
                    fr,
                    paddedW,
                    paddedH,
                    currentUsedLength,
                    placedSoFar,
                    settings,
                    block,
                    futureFitBonus
                );

                if (!best) {
                    best = {
                        x: fr.x,
                        y: fr.y,
                        w: paddedW,
                        h: paddedH,
                        rotatedOnSheet: ori.rotatedOnSheet,
                        score: score
                    };
                    continue;
                }

                var betterByLength = score.projectedUsedIn < (best.score.projectedUsedIn - 1e-6);
                var tieOnLength = Math.abs(score.projectedUsedIn - best.score.projectedUsedIn) <= 1e-6;
                var betterByWidth = tieOnLength && (score.widthWaste < best.score.widthWaste - 1e-6);
                var tieOnWidth = tieOnLength && Math.abs(score.widthWaste - best.score.widthWaste) <= 1e-6;
                var betterByScore = tieOnWidth && (score.total < best.score.total);

                if (betterByLength || betterByWidth || betterByScore) {
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

    if (usedTop > frTop + EPS) result.push({ x: frLeft, y: frTop, w: freeRect.w, h: usedTop - frTop });
    if (usedBottom < frBottom - EPS) result.push({ x: frLeft, y: usedBottom, w: freeRect.w, h: frBottom - usedBottom });
    if (usedLeft > frLeft + EPS) result.push({ x: frLeft, y: frTop, w: usedLeft - frLeft, h: freeRect.h });
    if (usedRight < frRight - EPS) result.push({ x: usedRight, y: frTop, w: frRight - usedRight, h: freeRect.h });

    return result;
}

function canMergeHorizontally(a, b) {
    if (Math.abs(a.y - b.y) > EPS) return false;
    if (Math.abs(a.h - b.h) > EPS) return false;

    var aRight = a.x + a.w;
    var bRight = b.x + b.w;
    return Math.abs(aRight - b.x) <= EPS || Math.abs(bRight - a.x) <= EPS;
}

function canMergeVertically(a, b) {
    if (Math.abs(a.x - b.x) > EPS) return false;
    if (Math.abs(a.w - b.w) > EPS) return false;

    var aBottom = a.y + a.h;
    var bBottom = b.y + b.h;
    return Math.abs(aBottom - b.y) <= EPS || Math.abs(bBottom - a.y) <= EPS;
}

function mergeTwoRects(a, b) {
    var left = Math.min(a.x, b.x);
    var top = Math.min(a.y, b.y);
    var right = Math.max(a.x + a.w, b.x + b.w);
    var bottom = Math.max(a.y + a.h, b.y + b.h);
    return {
        x: left,
        y: top,
        w: right - left,
        h: bottom - top
    };
}

function mergeAdjacentFreeRects(freeRects) {
    var rects = [];
    for (var i = 0; i < freeRects.length; i++) {
        if (isValidRect(freeRects[i])) rects.push(cloneRect(freeRects[i]));
    }

    var changed = true;
    while (changed) {
        changed = false;

        for (var a = 0; a < rects.length; a++) {
            for (var b = a + 1; b < rects.length; b++) {
                if (canMergeHorizontally(rects[a], rects[b]) || canMergeVertically(rects[a], rects[b])) {
                    rects[a] = mergeTwoRects(rects[a], rects[b]);
                    rects.splice(b, 1);
                    changed = true;
                    break;
                }
            }
            if (changed) break;
        }
    }

    return rects;
}

function pruneFreeRects(freeRects) {
    var pruned = [];

    for (var i = 0; i < freeRects.length; i++) {
        if (isValidRect(freeRects[i])) pruned.push(cloneRect(freeRects[i]));
    }

    pruned = mergeAdjacentFreeRects(pruned);

    var changed = true;
    while (changed) {
        changed = false;

        for (var a = 0; a < pruned.length; a++) {
            if (!pruned[a]) continue;

            for (var b = 0; b < pruned.length; b++) {
                if (a === b || !pruned[b]) continue;

                if (containsRect(pruned[a], pruned[b])) {
                    pruned[b] = null;
                    changed = true;
                }
            }
        }

        var compacted = [];
        for (var k = 0; k < pruned.length; k++) {
            if (pruned[k] && isValidRect(pruned[k])) compacted.push(pruned[k]);
        }
        pruned = mergeAdjacentFreeRects(compacted);
    }

    return pruned;
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

function nestBlocksSinglePass(blocks, settings) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);

    var freeRects = [{ x: 0, y: 0, w: sheetWidth, h: maxLength }];
    var placed = [];
    var unplaced = [];
    var usedLength = 0;

    for (var i = 0; i < blocks.length; i++) {
        var block = blocks[i];
        var remainingBlocks = blocks.slice(i + 1);
        var p = findBestBlockPlacement(block, freeRects, usedLength, settings, placed, remainingBlocks);

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
        usedLength: usedLength,
        freeRects: freeRects
    };
}

function cloneFreeRectList(list) {
    var out = [];
    for (var i = 0; i < list.length; i++) out.push(cloneRect(list[i]));
    return out;
}

function clonePlacedList(list) {
    var out = [];
    for (var i = 0; i < list.length; i++) {
        var p = list[i];
        out.push({
            block: p.block,
            x: p.x,
            y: p.y,
            paddedW: p.paddedW,
            paddedH: p.paddedH,
            rotatedOnSheet: p.rotatedOnSheet
        });
    }
    return out;
}

function backfillUnplaced(layout, settings) {
    if (!layout || !layout.unplaced || layout.unplaced.length === 0) return layout;

    var placed = clonePlacedList(layout.placed || []);
    var unplaced = [];
    for (var i = 0; i < layout.unplaced.length; i++) unplaced.push(layout.unplaced[i]);
    var freeRects = cloneFreeRectList(layout.freeRects || []);
    var usedLength = layout.usedLength || 0;

    var improved = true;
    while (improved) {
        improved = false;

        for (var idx = 0; idx < unplaced.length; idx++) {
            var block = unplaced[idx];
            var remainingBlocks = unplaced.slice(0, idx).concat(unplaced.slice(idx + 1));
            var p = findBestBlockPlacement(block, freeRects, usedLength, settings, placed, remainingBlocks);
            if (!p) continue;

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

            unplaced.splice(idx, 1);
            improved = true;
            break;
        }
    }

    return {
        placed: placed,
        unplaced: unplaced,
        usedLength: usedLength,
        freeRects: freeRects,
        meta: layout.meta
    };
}

function sortPlacedReadingOrder(list) {
    list.sort(function(a, b) {
        if (Math.abs(a.y - b.y) > EPS) return a.y - b.y;
        if (Math.abs(a.x - b.x) > EPS) return a.x - b.x;
        var aArea = a.paddedW * a.paddedH;
        var bArea = b.paddedW * b.paddedH;
        return aArea - bArea;
    });
}

function rebuildLayoutFromPlaced(placed, unplaced, settings, meta) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);
    var sortedPlaced = clonePlacedList(placed || []);
    var freeRects = [{ x: 0, y: 0, w: sheetWidth, h: maxLength }];
    var usedLength = 0;

    sortPlacedReadingOrder(sortedPlaced);

    for (var i = 0; i < sortedPlaced.length; i++) {
        var p = sortedPlaced[i];
        freeRects = placeAndUpdateFreeRects(freeRects, {
            x: p.x,
            y: p.y,
            w: p.paddedW,
            h: p.paddedH
        });

        var bottom = p.y + p.paddedH;
        if (bottom > usedLength) usedLength = bottom;
    }

    return {
        placed: sortedPlaced,
        unplaced: unplaced ? unplaced.slice(0) : [],
        usedLength: usedLength,
        freeRects: freeRects,
        meta: meta
    };
}

function buildSourceRecordFromBlock(block) {
    return {
        id: block.sourceId,
        key: block.sourceKey,
        ref: block.sourceRef,
        name: block.sourceName,
        baseQty: block.sourceBaseQty,
        qty: block.sourceQty,
        width: block.sourceWidth,
        height: block.sourceHeight,
        area: block.sourceArea,
        longSide: block.sourceLongSide,
        shortSide: block.sourceShortSide
    };
}

function buildSingleSourceOrderVariants(blocks, settings) {
    var variants = buildBlockOrderVariants(blocks, settings);

    var smallFirst = cloneBlockOrder(blocks);
    smallFirst.sort(function(a, b) {
        if (a.count !== b.count) return a.count - b.count;
        if (a.area !== b.area) return a.area - b.area;
        return a.blockW - b.blockW;
    });
    variants.push(smallFirst);

    var largeFirst = cloneBlockOrder(blocks);
    largeFirst.sort(function(a, b) {
        if (b.count !== a.count) return b.count - a.count;
        if (b.area !== a.area) return b.area - a.area;
        return b.blockW - a.blockW;
    });
    variants.push(largeFirst);

    return variants;
}

function tryRepackSingleSource(layout, settings, sourceId) {
    var fixedPlaced = [];
    var sampleBlock = null;

    for (var i = 0; i < layout.placed.length; i++) {
        var p = layout.placed[i];
        if (p.block.sourceId === sourceId) {
            if (!sampleBlock) sampleBlock = p.block;
        } else {
            fixedPlaced.push(p);
        }
    }

    if (!sampleBlock) return null;

    var src = buildSourceRecordFromBlock(sampleBlock);
    var candidatePlans = buildCandidateBlockPlansForSource(src, settings, sampleBlock.isDominant);
    if (!candidatePlans || candidatePlans.length === 0) return null;

    var base = rebuildLayoutFromPlaced(fixedPlaced, layout.unplaced, settings, layout.meta);
    var bestLayout = null;
    var bestScore = null;

    for (var pIdx = 0; pIdx < candidatePlans.length; pIdx++) {
        var plan = candidatePlans[pIdx];
        var candidateBlocksInfo = materializeBlocksFromPlans([plan], settings, null, []);
        var planBlocks = candidateBlocksInfo.blocks;
        var variants = buildSingleSourceOrderVariants(planBlocks, settings);

        for (var v = 0; v < variants.length; v++) {
            var order = cloneBlockOrder(variants[v]);
            var placed = clonePlacedList(fixedPlaced);
            var freeRects = cloneFreeRectList(base.freeRects);
            var usedLength = base.usedLength;
            var failed = false;

            for (var i2 = 0; i2 < order.length; i2++) {
                var block = order[i2];
                var remainingBlocks = order.slice(i2 + 1);
                var placement = findBestBlockPlacement(block, freeRects, usedLength, settings, placed, remainingBlocks);
                if (!placement) {
                    failed = true;
                    break;
                }

                placed.push({
                    block: block,
                    x: placement.x,
                    y: placement.y,
                    paddedW: placement.w,
                    paddedH: placement.h,
                    rotatedOnSheet: placement.rotatedOnSheet
                });

                freeRects = placeAndUpdateFreeRects(freeRects, placement);
                var bottom = placement.y + placement.h;
                if (bottom > usedLength) usedLength = bottom;
            }

            if (failed) continue;

            var candidate = rebuildLayoutFromPlaced(placed, layout.unplaced, settings, layout.meta);
            if (!isLayoutSpacingValid(candidate, settings)) continue;

            var score = scoreLayout(candidate, settings);
            if (!bestScore || score.total < bestScore.total) {
                bestLayout = candidate;
                bestScore = score;
            }
        }
    }

    return bestLayout;
}

function optimizeSourceRepack(layout, settings) {
    if (!layout || !layout.placed || layout.placed.length < 2) return layout;

    var current = rebuildLayoutFromPlaced(layout.placed, layout.unplaced, settings, layout.meta);
    var currentScore = scoreLayout(current, settings);
    var sourceCounts = {};
    var sourceIds = [];
    var passes = (settings.searchEffort === "High") ? 3 : 2;

    for (var i = 0; i < current.placed.length; i++) {
        var sourceId = current.placed[i].block.sourceId;
        if (!sourceCounts[sourceId]) {
            sourceCounts[sourceId] = 0;
            sourceIds.push(sourceId);
        }
        sourceCounts[sourceId] += 1;
    }

    sourceIds.sort(function(a, b) {
        return sourceCounts[b] - sourceCounts[a];
    });

    for (var pass = 0; pass < passes; pass++) {
        var improved = false;

        for (var s = 0; s < sourceIds.length; s++) {
            var sid = sourceIds[s];
            if (sourceCounts[sid] < 2) continue;

            var candidate = tryRepackSingleSource(current, settings, sid);
            if (!candidate) continue;

            var candidateScore = scoreLayout(candidate, settings);
            var betterByLength = candidateScore.usedNoTrailingGap < (currentScore.usedNoTrailingGap - EPS);
            var tieOnLength = Math.abs(candidateScore.usedNoTrailingGap - currentScore.usedNoTrailingGap) <= EPS;
            var betterByUtil = tieOnLength && (candidateScore.utilization > currentScore.utilization + 1e-6);
            var betterByTotal = candidateScore.total < (currentScore.total - 1e-3);

            if (betterByLength || betterByUtil || betterByTotal) {
                current = candidate;
                currentScore = candidateScore;
                improved = true;
            }
        }

        if (!improved) break;
    }

    current.meta = layout.meta;
    return current;
}

function getRequiredPlacedRect(placedBlock, settings) {
    var spacing = inToPt(settings.spacingIn);
    var w = placedBlock.rotatedOnSheet ? placedBlock.block.blockH : placedBlock.block.blockW;
    var h = placedBlock.rotatedOnSheet ? placedBlock.block.blockW : placedBlock.block.blockH;
    return {
        x: placedBlock.x,
        y: placedBlock.y,
        w: w + spacing,
        h: h + spacing
    };
}

function isLayoutSpacingValid(layout, settings) {
    if (!layout || !layout.placed) return true;

    var requiredRects = [];
    for (var i = 0; i < layout.placed.length; i++) {
        var p = layout.placed[i];
        var req = getRequiredPlacedRect(p, settings);

        // Guard: no placement may use less than configured spacing.
        if (p.paddedW + EPS < req.w || p.paddedH + EPS < req.h) return false;
        requiredRects.push(req);
    }

    for (var a = 0; a < requiredRects.length; a++) {
        for (var b = a + 1; b < requiredRects.length; b++) {
            if (intersects(requiredRects[a], requiredRects[b])) return false;
        }
    }

    return true;
}

function cloneBlockOrder(blocks) {
    var out = [];
    for (var i = 0; i < blocks.length; i++) out.push(blocks[i]);
    return out;
}

function buildBlockOrderVariants(blocks, settings) {
    var variants = [];
    var orderRaw = parseFloat(settings.orderDiscipline);
    if (isNaN(orderRaw)) orderRaw = DEFAULTS.orderDiscipline;
    var orderBias = clamp(orderRaw, 0, 100) / 100.0;

    var base = cloneBlockOrder(blocks);
    variants.push(base);

    var vWidthFirst = cloneBlockOrder(blocks);
    vWidthFirst.sort(function(a, b) {
        if (a.isDominant !== b.isDominant) return a.isDominant ? -1 : 1;
        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        if (a.blockH !== b.blockH) return a.blockH - b.blockH;
        return b.area - a.area;
    });
    variants.push(vWidthFirst);

    var vHeightFirst = cloneBlockOrder(blocks);
    vHeightFirst.sort(function(a, b) {
        if (a.isDominant !== b.isDominant) return a.isDominant ? -1 : 1;
        if (a.blockH !== b.blockH) return a.blockH - b.blockH;
        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        return b.area - a.area;
    });
    variants.push(vHeightFirst);

    var vMinorityFirst = cloneBlockOrder(blocks);
    vMinorityFirst.sort(function(a, b) {
        if (a.isDominant !== b.isDominant) return a.isDominant ? 1 : -1;
        if (b.longSide !== a.longSide) return b.longSide - a.longSide;
        return b.area - a.area;
    });
    variants.push(vMinorityFirst);

    var dominant = [];
    var minority = [];
    for (var i = 0; i < blocks.length; i++) {
        if (blocks[i].isDominant) dominant.push(blocks[i]);
        else minority.push(blocks[i]);
    }
    dominant.sort(function(a, b) {
        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        return a.blockH - b.blockH;
    });
    minority.sort(function(a, b) {
        if (b.longSide !== a.longSide) return b.longSide - a.longSide;
        return b.area - a.area;
    });

    var vInterleave = [];
    var maxLen = Math.max(dominant.length, minority.length);
    for (var j = 0; j < maxLen; j++) {
        if (j < dominant.length) vInterleave.push(dominant[j]);
        if (j < minority.length) vInterleave.push(minority[j]);
    }
    if (vInterleave.length > 0) variants.push(vInterleave);

    if (orderBias > 0.4) {
        var vGrouped = cloneBlockOrder(blocks);
        vGrouped.sort(function(a, b) {
            if (a.sourceName < b.sourceName) return -1;
            if (a.sourceName > b.sourceName) return 1;
            if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
            return a.blockH - b.blockH;
        });
        variants.push(vGrouped);
    }

    return variants;
}
function scoreLayout(layout, settings) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var usedNoTrailingGap = Math.max(0, layout.usedLength - inToPt(settings.spacingIn));
    var usedArea = 0;
    var sourceSwitches = 0;
    var prevSource = null;
    var i;

    for (i = 0; i < layout.placed.length; i++) {
        var p = layout.placed[i];
        usedArea += p.block.area;
        if (prevSource !== null && prevSource !== p.block.sourceId) sourceSwitches += 1;
        prevSource = p.block.sourceId;
    }

    var placedCopyCount = countPlacedCopies(layout);
    var unplacedCopyCount = countUnplacedCopies(layout);
    var capacityArea = sheetWidth * Math.max(layout.usedLength, EPS);
    var utilization = (capacityArea > EPS) ? (usedArea / capacityArea) : 0;
    utilization = clamp(utilization, 0, 1);

    var orderRaw = parseFloat(settings.orderDiscipline);
    if (isNaN(orderRaw)) orderRaw = DEFAULTS.orderDiscipline;
    var orderBias = clamp(orderRaw, 0, 100) / 100.0;

    // Strong lexicographic-style objective in one scalar:
    // unplaced copies >> used length >> width utilization >> ordering switches.
    return {
        total:
            (unplacedCopyCount * 1000000000) +
            (usedNoTrailingGap * 10000) +
            ((1 - utilization) * 120000) +
            (sourceSwitches * orderBias * 3000),
        usedNoTrailingGap: usedNoTrailingGap,
        utilization: utilization,
        placedCopyCount: placedCopyCount,
        unplacedCopyCount: unplacedCopyCount,
        sourceSwitches: sourceSwitches
    };
}

function prioritizeUnplaced(order, layout) {
    if (!layout || !layout.unplaced || layout.unplaced.length === 0) return cloneBlockOrder(order);

    var map = {};
    var i;
    for (i = 0; i < layout.unplaced.length; i++) {
        map[layout.unplaced[i].uid] = true;
    }

    var front = [];
    var back = [];
    for (i = 0; i < order.length; i++) {
        if (map[order[i].uid]) front.push(order[i]);
        else back.push(order[i]);
    }
    return front.concat(back);
}

function perturbOrder(order, seed) {
    var out = cloneBlockOrder(order);
    if (out.length < 3) return out;

    var a = seed % out.length;
    var b = (seed * 7 + 3) % out.length;
    if (a === b) b = (b + 1) % out.length;

    var temp = out[a];
    out[a] = out[b];
    out[b] = temp;

    return out;
}

function orderSignature(order) {
    var ids = [];
    for (var i = 0; i < order.length; i++) ids.push(order[i].uid);
    return ids.join("|");
}

function pushUniqueOrderVariant(list, order, seen) {
    var sig = orderSignature(order);
    if (seen[sig]) return;
    seen[sig] = true;
    list.push(order);
}

function collectAllBlocksFromLayout(layout) {
    var out = [];
    var i;

    if (layout && layout.placed) {
        for (i = 0; i < layout.placed.length; i++) out.push(layout.placed[i].block);
    }
    if (layout && layout.unplaced) {
        for (i = 0; i < layout.unplaced.length; i++) out.push(layout.unplaced[i]);
    }

    return out;
}

function buildCompactionOrderVariants(layout, settings) {
    var seen = {};
    var variants = [];
    var placed = clonePlacedList(layout.placed || []);
    var tail = layout.unplaced ? cloneBlockOrder(layout.unplaced) : [];

    placed.sort(function(a, b) {
        if (Math.abs(a.y - b.y) > EPS) return a.y - b.y;
        if (Math.abs(a.x - b.x) > EPS) return a.x - b.x;
        return b.block.area - a.block.area;
    });

    var reading = [];
    var i;
    for (i = 0; i < placed.length; i++) reading.push(placed[i].block);
    for (i = 0; i < tail.length; i++) reading.push(tail[i]);
    pushUniqueOrderVariant(variants, reading, seen);

    var widthFirst = cloneBlockOrder(reading);
    widthFirst.sort(function(a, b) {
        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        if (a.blockH !== b.blockH) return a.blockH - b.blockH;
        return b.area - a.area;
    });
    pushUniqueOrderVariant(variants, widthFirst, seen);

    var areaFirst = cloneBlockOrder(reading);
    areaFirst.sort(function(a, b) {
        if (b.area !== a.area) return b.area - a.area;
        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        return a.blockH - b.blockH;
    });
    pushUniqueOrderVariant(variants, areaFirst, seen);

    if (settings.searchEffort === "High") {
        var mixed = perturbOrder(widthFirst, 17);
        pushUniqueOrderVariant(variants, mixed, seen);
    }

    return variants;
}

function compactLayoutGreedy(layout, settings) {
    if (!layout) return layout;

    var allBlocks = collectAllBlocksFromLayout(layout);
    if (allBlocks.length < 2) return layout;

    var variants = buildCompactionOrderVariants(layout, settings);
    var bestLayout = rebuildLayoutFromPlaced(layout.placed, layout.unplaced, settings, layout.meta);
    var bestScore = scoreLayout(bestLayout, settings);

    for (var i = 0; i < variants.length; i++) {
        var candidate = nestBlocksSinglePass(variants[i], settings);
        candidate = backfillUnplaced(candidate, settings);
        if (!isLayoutSpacingValid(candidate, settings)) continue;

        var candidateScore = scoreLayout(candidate, settings);
        var betterByLength = candidateScore.usedNoTrailingGap < (bestScore.usedNoTrailingGap - EPS);
        var tieOnLength = Math.abs(candidateScore.usedNoTrailingGap - bestScore.usedNoTrailingGap) <= EPS;
        var betterByUtil = tieOnLength && (candidateScore.utilization > bestScore.utilization + 1e-6);
        var betterByTotal = candidateScore.total < (bestScore.total - 1e-3);

        if (betterByLength || betterByUtil || betterByTotal) {
            bestLayout = candidate;
            bestScore = candidateScore;
        }
    }

    bestLayout.meta = layout.meta;
    return bestLayout;
}

function nestBlocks(blocks, settings) {
    var variants = buildBlockOrderVariants(blocks, settings);
    var bestLayout = null;
    var bestScore = null;
    var bestMeta = null;
    var maxVariants = (settings.searchEffort === "High") ? 4 : 3;

    if (variants.length > maxVariants) variants = variants.slice(0, maxVariants);

    for (var v = 0; v < variants.length; v++) {
        var layout = nestBlocksSinglePass(variants[v], settings);
        layout = backfillUnplaced(layout, settings);
        layout = compactLayoutGreedy(layout, settings);
        layout = backfillUnplaced(layout, settings);

        if (!isLayoutSpacingValid(layout, settings)) continue;

        var score = scoreLayout(layout, settings);
        if (!bestScore || score.total < bestScore.total) {
            bestLayout = layout;
            bestScore = score;
            bestMeta = { variantIndex: v, step: 0 };
        }
    }

    if (bestLayout) {
        bestLayout.meta = {
            score: bestScore,
            variantIndex: bestMeta.variantIndex,
            step: bestMeta.step
        };
        return bestLayout;
    }

    var fallback = nestBlocksSinglePass(blocks, settings);
    fallback = backfillUnplaced(fallback, settings);
    fallback = compactLayoutGreedy(fallback, settings);
    fallback = backfillUnplaced(fallback, settings);
    if (!isLayoutSpacingValid(fallback, settings)) {
        throw new Error("Internal spacing guard failed. No valid layout found with requested spacing.");
    }
    return fallback;
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

function placeSingleCopy(refItem, targetLayer, x, y, rotationDeg, sourceKey) {
    var dup = duplicateToLayer(refItem, targetLayer);
    if (rotationDeg !== 0) dup.rotate(rotationDeg);

    var targetLeft = x;
    var targetTop = -y;
    setItemTopLeftByVisibleBounds(dup, targetLeft, targetTop);

    try { dup.note = "NESTER_SOURCE_KEY=" + String(sourceKey || ""); } catch (e1) {}
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
    if (!rotatedOnSheet) return { x: localX, y: localY };

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

            placeSingleCopy(b.sourceRef, outputLayer, itemX, itemY, rotationDeg, b.sourceKey);
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

function buildSummary(result, settings) {
    var usedNoTrailingGap = Math.max(0, result.layout.usedLength - inToPt(settings.spacingIn));
    var orderShown = clamp(_numOr(settings.orderDiscipline, DEFAULTS.orderDiscipline), 0, 100);
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var util = 0;
    if (usedNoTrailingGap > EPS) {
        var areaSum = 0;
        for (var i = 0; i < result.layout.placed.length; i++) areaSum += result.layout.placed[i].block.area;
        util = clamp(areaSum / Math.max(sheetWidth * usedNoTrailingGap, EPS), 0, 1);
    }
    var lines = [];
    lines.push("WeMust NESTER v6");
    lines.push("Objective: Min length + Max width fill");
    if (settings.optimizePreset) lines.push("Preset: " + settings.optimizePreset);
    if (settings.searchEffort) lines.push("Solver effort: " + settings.searchEffort);
    lines.push("Order control (resolved): " + orderShown);
    lines.push("Placed copies: " + countPlacedCopies(result.layout));
    lines.push("Unplaced copies: " + countUnplacedCopies(result.layout));
    lines.push("Used length: " + inToStr(usedNoTrailingGap) + " in");
    lines.push("Width utilization: " + (util * 100).toFixed(1) + "%");
    lines.push("Item rotation in block: " + (settings.allowItemRotationInBlock ? "On" : "Off"));
    lines.push("Block rotation on sheet: " + (settings.allowBlockRotationOnSheet ? "On" : "Off"));
    if (result.layout.meta && result.layout.meta.score) {
        lines.push("Search variant: " + result.layout.meta.variantIndex + ", step: " + result.layout.meta.step);
    }
    if (result.layout.meta && result.layout.meta.autoTuning) {
        lines.push("Auto candidates: " + result.layout.meta.autoTuning.candidateCount);
    }
    return lines.join("\n");
}

function buildSourceQuantitySummary(result) {
    var out = [];
    var byKey = {};
    var i;

    if (!result || !result.sources) return out;

    for (i = 0; i < result.sources.length; i++) {
        var src = result.sources[i];
        byKey[src.key] = {
            id: src.id,
            key: src.key,
            name: src.name,
            filePath: src.filePath || "",
            detectedQty: src.baseQty,
            requestedQty: src.qty,
            placedQty: 0,
            unplacedQty: 0,
            widthIn: parseFloat(inToStr(src.width)),
            heightIn: parseFloat(inToStr(src.height)),
            dimensionsText: inToStr(src.width) + " x " + inToStr(src.height)
        };
        out.push(byKey[src.key]);
    }

    if (result.layout && result.layout.placed) {
        for (i = 0; i < result.layout.placed.length; i++) {
            var placed = result.layout.placed[i];
            if (byKey[placed.block.sourceKey]) {
                byKey[placed.block.sourceKey].placedQty += placed.block.count;
            }
        }
    }

    if (result.layout && result.layout.unplaced) {
        for (i = 0; i < result.layout.unplaced.length; i++) {
            var unplaced = result.layout.unplaced[i];
            if (byKey[unplaced.sourceKey]) {
                byKey[unplaced.sourceKey].unplacedQty += unplaced.count;
            }
        }
    }

    return out;
}

// =====================================================
// Build pipeline
// =====================================================

function cloneSettings(settings) {
    var out = {};
    for (var k in settings) {
        if (settings.hasOwnProperty(k)) out[k] = settings[k];
    }
    return out;
}

function solveFromSources(sources, settings) {
    var planInfo = collectSourcePlanSets(sources, settings);
    var planCombos = buildPlanCombinations(planInfo.planSets, settings);

    if (planCombos.length === 0) {
        throw new Error("No valid block plans could be built.");
    }

    var bestBlocksInfo = null;
    var bestLayout = null;
    var bestScore = null;
    var bestCombo = null;
    var bestComboIndex = -1;

    for (var i = 0; i < planCombos.length; i++) {
        var combo = planCombos[i];
        var blocksInfo = materializeBlocksFromPlans(
            combo.plans,
            settings,
            planInfo.dominantId,
            planInfo.failedNames
        );

        if (blocksInfo.blocks.length === 0) continue;

        var layout = nestBlocks(blocksInfo.blocks, settings);
        var score = scoreLayout(layout, settings);

        if (!bestScore || score.total < bestScore.total || (score.total === bestScore.total && combo.heuristicScore < bestCombo.heuristicScore)) {
            bestBlocksInfo = blocksInfo;
            bestLayout = layout;
            bestScore = score;
            bestCombo = combo;
            bestComboIndex = i;
        }
    }

    if (!bestBlocksInfo || !bestLayout || !bestCombo) {
        throw new Error("No valid block plans could be built.");
    }

    bestLayout.meta = bestLayout.meta || {};
    bestLayout.meta.planComboIndex = bestComboIndex;
    bestLayout.meta.planComboCount = planCombos.length;
    bestLayout.meta.planComboHeuristic = bestCombo.heuristicScore;

    return {
        sources: sources,
        blocksInfo: bestBlocksInfo,
        layout: bestLayout
    };
}

function renderSolution(doc, solution, settings) {
    var outputLayer = prepareOutputLayer(doc);
    renderPlacedBlocks(solution.layout, outputLayer, settings);

    if (settings.hideSourceLayersAfterBuild) {
        hideSourceLayers(doc);
    }
}

function buildCandidateSettings(baseSettings) {
    var out = [];
    var preset = baseSettings.optimizePreset || "Auto";
    var effort = baseSettings.searchEffort || "Normal";

    function addCandidate(s) {
        var key = [
            s.orderDiscipline,
            s.maxBlocksPerFile,
            s.maxBlockAspectRatio.toFixed(2),
            s.allowBlockRotationOnSheet ? 1 : 0
        ].join("|");

        for (var i = 0; i < out.length; i++) {
            var existing = out[i];
            var k2 = [
                existing.orderDiscipline,
                existing.maxBlocksPerFile,
                existing.maxBlockAspectRatio.toFixed(2),
                existing.allowBlockRotationOnSheet ? 1 : 0
            ].join("|");
            if (k2 === key) return;
        }
        out.push(s);
    }

    if (preset === "Auto") {
        var orders = (effort === "High")
            ? [22, 46, 70]
            : [32, 60];

        for (var oi = 0; oi < orders.length; oi++) {
            var c = cloneSettings(baseSettings);
            c.orderDiscipline = orders[oi];
            applyDerivedTuning(c);
            addCandidate(c);

            if (effort === "High" && oi === 1) {
                var c2 = cloneSettings(c);
                c2.maxBlocksPerFile = clamp(c2.maxBlocksPerFile + 1, 2, 4);
                c2.maxBlockAspectRatio = c2.maxBlockAspectRatio + 0.25;
                addCandidate(c2);
            }
        }
    } else {
        var single = cloneSettings(baseSettings);
        applyDerivedTuning(single);
        addCandidate(single);

        if (effort === "High") {
            var alt = cloneSettings(single);
            alt.maxBlocksPerFile = clamp(alt.maxBlocksPerFile + 1, 2, 4);
            alt.maxBlockAspectRatio = alt.maxBlockAspectRatio + 0.2;
            addCandidate(alt);
        }
    }

    return out;
}

function buildOnce(doc, settings) {
    var sources = collectPlacedItems(doc, settings.quantityOverrides);

    if (sources.length === 0) {
        throw new Error("No placed items found in the current document.");
    }

    var candidates = buildCandidateSettings(settings);
    if (candidates.length === 0) {
        throw new Error("No candidate settings generated.");
    }

    var bestSolution = null;
    var bestSettings = null;
    var bestScore = null;

    for (var i = 0; i < candidates.length; i++) {
        var cSettings = candidates[i];
        var solved = solveFromSources(sources, cSettings);
        var s = scoreLayout(solved.layout, cSettings);

        if (!bestScore || s.total < bestScore.total) {
            bestSolution = solved;
            bestSettings = cSettings;
            bestScore = s;
        }
    }

    if (!bestSolution || !bestSettings) {
        throw new Error("Solver could not produce a valid layout.");
    }

    renderSolution(doc, bestSolution, bestSettings);
    bestSolution.layout.meta = bestSolution.layout.meta || {};
    bestSolution.layout.meta.autoTuning = {
        candidateCount: candidates.length,
        selectedOrder: bestSettings.orderDiscipline,
        selectedMaxBlocks: bestSettings.maxBlocksPerFile,
        selectedAspect: bestSettings.maxBlockAspectRatio
    };

    return {
        sources: bestSolution.sources,
        blocksInfo: bestSolution.blocksInfo,
        layout: bestSolution.layout,
        usedSettings: bestSettings
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
        optimizePreset: DEFAULTS.optimizePreset,
        searchEffort: DEFAULTS.searchEffort,
        mode: DEFAULTS.mode,
        orderDiscipline: DEFAULTS.orderDiscipline,
        maxBlocksPerFile: DEFAULTS.maxBlocksPerFile,
        maxBlockAspectRatio: DEFAULTS.maxBlockAspectRatio,
        allowItemRotationInBlock: DEFAULTS.allowItemRotationInBlock,
        allowBlockRotationOnSheet: DEFAULTS.allowBlockRotationOnSheet,
        hideSourceLayersAfterBuild: DEFAULTS.hideSourceLayersAfterBuild,
        useDominantShelf: DEFAULTS.useDominantShelf,
        dominantShelfMaxRows: DEFAULTS.dominantShelfMaxRows,
        widthFillPriority: DEFAULTS.widthFillPriority,
        sameSourceProximity: DEFAULTS.sameSourceProximity,
        cavityPenalty: DEFAULTS.cavityPenalty
    };

    while (true) {
        var settings = showSettingsDialog(current);
        if (!settings) break;

        current = settings;

        try {
            var result = buildOnce(doc, settings);
            var summary = buildSummary(result, result.usedSettings || settings);
            var again = confirm(summary + "\n\nBuild another variation?");
            if (!again) break;
        } catch (err) {
            var retry = confirm("Error:\n" + err + "\n\nTry again with new settings?");
            if (!retry) break;
        }
    }
}

function nesterRunMain() {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        main();
        return "OK";
    } catch (e) {
        return "ERROR: " + e;
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}

function _copyDefaults() {
    return {
        sheetWidthIn: DEFAULTS.sheetWidthIn,
        maxLengthIn: DEFAULTS.maxLengthIn,
        spacingIn: DEFAULTS.spacingIn,
        optimizePreset: DEFAULTS.optimizePreset,
        searchEffort: DEFAULTS.searchEffort,
        mode: DEFAULTS.mode,
        orderDiscipline: DEFAULTS.orderDiscipline,
        maxBlocksPerFile: DEFAULTS.maxBlocksPerFile,
        maxBlockAspectRatio: DEFAULTS.maxBlockAspectRatio,
        allowItemRotationInBlock: DEFAULTS.allowItemRotationInBlock,
        allowBlockRotationOnSheet: DEFAULTS.allowBlockRotationOnSheet,
        hideSourceLayersAfterBuild: DEFAULTS.hideSourceLayersAfterBuild,
        useDominantShelf: DEFAULTS.useDominantShelf,
        dominantShelfMaxRows: DEFAULTS.dominantShelfMaxRows,
        widthFillPriority: DEFAULTS.widthFillPriority,
        sameSourceProximity: DEFAULTS.sameSourceProximity,
        cavityPenalty: DEFAULTS.cavityPenalty
    };
}

function _numOr(v, fallback) {
    var n = parseFloat(v);
    return isNaN(n) ? fallback : n;
}

function _intOr(v, fallback) {
    var n = parseInt(v, 10);
    return isNaN(n) ? fallback : n;
}

function _boolOr(v, fallback) {
    if (v === true || v === false) return v;
    if (v === 1 || v === "1" || v === "true") return true;
    if (v === 0 || v === "0" || v === "false") return false;
    return fallback;
}

function _oneOfOr(v, allowed, fallback) {
    var s = String(v);
    for (var i = 0; i < allowed.length; i++) {
        if (allowed[i] === s) return s;
    }
    return fallback;
}

function _jsonEscape(str) {
    return String(str)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, "\\\"")
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n");
}

function _jsonStringify(value) {
    if (typeof JSON !== "undefined" && JSON.stringify) {
        return JSON.stringify(value);
    }

    if (value === null || value === undefined) return "null";
    if (typeof value === "string") return "\"" + _jsonEscape(value) + "\"";
    if (typeof value === "number" || typeof value === "boolean") return String(value);

    if (value instanceof Array) {
        var arrParts = [];
        for (var i = 0; i < value.length; i++) arrParts.push(_jsonStringify(value[i]));
        return "[" + arrParts.join(",") + "]";
    }

    var objParts = [];
    for (var k in value) {
        if (!value.hasOwnProperty(k)) continue;
        objParts.push("\"" + _jsonEscape(k) + "\":" + _jsonStringify(value[k]));
    }
    return "{" + objParts.join(",") + "}";
}

function _jsonParse(str) {
    if (typeof JSON !== "undefined" && JSON.parse) {
        return JSON.parse(String(str));
    }
    return eval("(" + str + ")");
}

function applyDerivedTuning(s) {
    var orderBias = clamp(_numOr(s.orderDiscipline, DEFAULTS.orderDiscipline), 0, 100) / 100.0;
    s.orderDiscipline = clamp(_numOr(s.orderDiscipline, DEFAULTS.orderDiscipline), 0, 100);

    // Keep legacy knobs internal and derive them from compact user controls.
    s.mode = "Production";
    s.maxBlocksPerFile = (s.orderDiscipline < 25) ? 3 : 2;
    s.maxBlockAspectRatio = 2.2 + ((1 - orderBias) * 1.2); // 2.2..3.4

    s.useDominantShelf = true;
    s.dominantShelfMaxRows = clamp(Math.round(5 + (orderBias * 2)), 5, 7);
    s.widthFillPriority = 92;
    s.sameSourceProximity = clamp(20 + (s.orderDiscipline * 0.7), 0, 100);
    s.cavityPenalty = clamp(55 + (s.orderDiscipline * 0.35), 0, 100);
}

function normalizeQuantityOverrides(raw) {
    var out = {};
    if (!raw || typeof raw !== "object") return out;

    for (var key in raw) {
        if (!raw.hasOwnProperty(key)) continue;
        var qty = _intOr(raw[key], NaN);
        if (isNaN(qty)) continue;
        out[key] = Math.max(1, qty);
    }

    return out;
}

function normalizeSettingsFromPanel(input) {
    var s = _copyDefaults();
    if (!input || typeof input !== "object") return s;

    s.sheetWidthIn = Math.max(0.01, _numOr(input.sheetWidthIn, s.sheetWidthIn));
    s.maxLengthIn = Math.max(0.01, _numOr(input.maxLengthIn, s.maxLengthIn));
    s.spacingIn = Math.max(0, _numOr(input.spacingIn, s.spacingIn));

    s.optimizePreset = _oneOfOr(input.optimizePreset, ["Auto", "Compact", "Balanced", "Ordered"], s.optimizePreset);
    s.searchEffort = _oneOfOr(input.searchEffort, ["Normal", "High"], s.searchEffort);
    s.allowItemRotationInBlock = _boolOr(input.allowItemRotationInBlock, s.allowItemRotationInBlock);

    if (s.optimizePreset === "Compact") s.orderDiscipline = 22;
    else if (s.optimizePreset === "Balanced") s.orderDiscipline = 48;
    else if (s.optimizePreset === "Ordered") s.orderDiscipline = 78;
    else s.orderDiscipline = clamp(_numOr(input.orderDiscipline, s.orderDiscipline), 0, 100);

    s.allowBlockRotationOnSheet = _boolOr(input.allowBlockRotationOnSheet, s.allowBlockRotationOnSheet);
    s.hideSourceLayersAfterBuild = DEFAULTS.hideSourceLayersAfterBuild;
    s.quantityOverrides = normalizeQuantityOverrides(input.quantityOverrides);

    applyDerivedTuning(s);

    return s;
}

function nesterGetDefaultSettings() {
    return _jsonStringify({
        sheetWidthIn: DEFAULTS.sheetWidthIn,
        maxLengthIn: DEFAULTS.maxLengthIn,
        spacingIn: DEFAULTS.spacingIn,
        optimizePreset: DEFAULTS.optimizePreset,
        searchEffort: DEFAULTS.searchEffort,
        allowItemRotationInBlock: DEFAULTS.allowItemRotationInBlock,
        allowBlockRotationOnSheet: DEFAULTS.allowBlockRotationOnSheet,
        hideSourceLayersAfterBuild: DEFAULTS.hideSourceLayersAfterBuild,
    });
}

function getSourceFolderLabel(sources) {
    if (!sources || !sources.length) return "";

    var folderName = "";

    for (var i = 0; i < sources.length; i++) {
        var candidate = getParentFolderName(sources[i].filePath || "");
        if (!candidate) continue;

        if (!folderName) {
            folderName = candidate;
            continue;
        }

        if (folderName !== candidate) {
            return "MixedFolders";
        }
    }

    return folderName;
}

function buildOutputBoundsText(doc) {
    if (!doc) return "";

    var layer = findLayerByName(doc, OUTPUT_LAYER_NAME);
    if (!layer || layer.pageItems.length === 0) return "";

    var left = null;
    var top = null;
    var right = null;
    var bottom = null;

    for (var i = 0; i < layer.pageItems.length; i++) {
        var item = layer.pageItems[i];
        var vb;
        try { vb = item.visibleBounds; } catch (e1) { vb = null; }
        if (!vb) continue;

        if (left === null || vb[0] < left) left = vb[0];
        if (top === null || vb[1] > top) top = vb[1];
        if (right === null || vb[2] > right) right = vb[2];
        if (bottom === null || vb[3] < bottom) bottom = vb[3];
    }

    if (left === null || top === null || right === null || bottom === null) return "";

    var width = Math.abs(right - left);
    var height = Math.abs(top - bottom);
    return inToStr(width) + "x" + inToStr(height) + "in";
}

function nesterSelectPlacedCopyBySourceKey(sourceKey) {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        if (app.documents.length === 0) {
            return _jsonStringify({ ok: false, error: "No open document found." });
        }

        var doc = app.activeDocument;
        var layer = findLayerByName(doc, OUTPUT_LAYER_NAME);
        if (!layer || layer.pageItems.length === 0) {
            return _jsonStringify({ ok: false, error: "No built output found." });
        }

        var needle = "NESTER_SOURCE_KEY=" + String(sourceKey || "");
        var found = null;

        for (var i = 0; i < layer.pageItems.length; i++) {
            var item = layer.pageItems[i];
            try {
                if (item.note === needle) {
                    found = item;
                    break;
                }
            } catch (e1) {}
        }

        if (!found) {
            return _jsonStringify({ ok: false, error: "No matching placed copy found." });
        }

        doc.selection = null;
        doc.selection = [found];
        app.redraw();
        return _jsonStringify({ ok: true });
    } catch (e) {
        return _jsonStringify({ ok: false, error: String(e) });
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}

function nesterRunWithSettings(settingsJson) {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        if (app.documents.length === 0) {
            return _jsonStringify({ ok: false, error: "No open document found." });
        }

        var parsed = null;
        if (settingsJson !== undefined && settingsJson !== null && String(settingsJson) !== "") {
            parsed = _jsonParse(settingsJson);
        }

        var settings = normalizeSettingsFromPanel(parsed);
        var result = buildOnce(app.activeDocument, settings);
        var effective = result.usedSettings || settings;
        var usedNoTrailingGap = Math.max(0, result.layout.usedLength - inToPt(effective.spacingIn));

        return _jsonStringify({
            ok: true,
            summary: buildSummary(result, effective),
            usedLengthIn: parseFloat(inToStr(usedNoTrailingGap)),
            placedCopies: countPlacedCopies(result.layout),
            unplacedCopies: countUnplacedCopies(result.layout),
            searchMeta: result.layout.meta ? result.layout.meta : null,
            sourceItems: buildSourceQuantitySummary(result),
            outputBoundsText: buildOutputBoundsText(app.activeDocument),
            sourceFolderLabel: getSourceFolderLabel(result.sources),
            outputSizeText: buildOutputBoundsText(app.activeDocument)
        });
    } catch (e) {
        return _jsonStringify({ ok: false, error: String(e) });
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}






























