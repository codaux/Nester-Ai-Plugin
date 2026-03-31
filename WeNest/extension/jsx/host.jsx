#target illustrator

// =====================================================
// WeMust NESTER v7
// Staged production solver rewrite
//
// Developer note:
// The solver is now organized as deterministic staged candidates:
// 1) collect and normalize sources
// 2) generate source block plans
// 3) build global candidate plan sets
// 4) run primary block layout
// 5) run compaction pass
// 6) run adaptive block splitting pass
// 7) run hole filling pass
// 8) run local repair pass
// 9) validate spacing / overlap / bounds
// 10) score all candidates and choose the winner
//
// The stable CEP contract is preserved. The host now returns richer
// structured debug data under searchMeta plus top-level aliases.
// =====================================================

var PT_PER_IN = 72;
var EPS = 0.01;

var OUTPUT_LAYER_NAME = "NEST_BUILD";
var OUTPUT_LAYER_PREFIX = "NEST_BUILD";
var OUTPUT_META_NOTE_PREFIX = "NESTER_META=";
var OUTPUT_FOLDER_LABEL_META = "FOLDER_LABEL";
var LAST_SOURCE_FOLDER_PATHS = [];
var DIGITS_TYPE_ARABIC = 1684627826;
var CHARACTER_DIRECTION_LTR = 1278366308;

var DEFAULTS = {
    sheetWidthIn: 23,
    maxLengthIn: 100,
    spacingIn: 0.25,
    optimizePreset: "Auto",
    searchEffort: "Normal",
    allowItemRotationInBlock: true,
    allowBlockRotationOnSheet: true,
    hideSourceLayersAfterBuild: true
};

var BLOCK_UID_COUNTER = 1;

var PERSONALITY_LIBRARY = {
    Compact: {
        name: "Compact",
        orderDiscipline: 24,
        widthFillPriority: 96,
        continuityWeight: 24,
        cavityWeight: 92,
        fragmentationWeight: 64,
        holeFillAggressiveness: 95,
        splitAggressiveness: 92,
        compactionAggressiveness: 90,
        maxBlocksPerSource: 5,
        maxBlockAspectRatio: 4.6,
        dominantShelfRows: 7,
        baselineLike: false
    },
    Balanced: {
        name: "Balanced",
        orderDiscipline: 52,
        widthFillPriority: 86,
        continuityWeight: 56,
        cavityWeight: 76,
        fragmentationWeight: 58,
        holeFillAggressiveness: 62,
        splitAggressiveness: 58,
        compactionAggressiveness: 64,
        maxBlocksPerSource: 4,
        maxBlockAspectRatio: 3.6,
        dominantShelfRows: 6,
        baselineLike: false
    },
    Ordered: {
        name: "Ordered",
        orderDiscipline: 84,
        widthFillPriority: 72,
        continuityWeight: 92,
        cavityWeight: 78,
        fragmentationWeight: 88,
        holeFillAggressiveness: 34,
        splitAggressiveness: 28,
        compactionAggressiveness: 38,
        maxBlocksPerSource: 3,
        maxBlockAspectRatio: 2.7,
        dominantShelfRows: 5,
        baselineLike: false
    },
    Baseline: {
        name: "Baseline",
        orderDiscipline: 62,
        widthFillPriority: 82,
        continuityWeight: 68,
        cavityWeight: 55,
        fragmentationWeight: 48,
        holeFillAggressiveness: 8,
        splitAggressiveness: 0,
        compactionAggressiveness: 12,
        maxBlocksPerSource: 3,
        maxBlockAspectRatio: 3.4,
        dominantShelfRows: 6,
        baselineLike: true
    }
};

// =====================================================
// Basic helpers
// =====================================================

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function normalizeAsciiDigits(value) {
    return String(value || "")
        .replace(/[\u0660-\u0669]/g, function(ch) { return String(ch.charCodeAt(0) - 1632); })
        .replace(/[\u06F0-\u06F9]/g, function(ch) { return String(ch.charCodeAt(0) - 1776); });
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

function getParentFolderName(pathOrName) {
    var s = trimStr(pathOrName || "");
    if (!s) return "";
    s = s.replace(/\\/g, "/");
    if (s.charAt(s.length - 1) === "/") s = s.substring(0, s.length - 1);
    var parts = s.split("/");
    return (parts.length > 1) ? parts[parts.length - 2] : "";
}

function getParentFolderPath(pathOrName) {
    var s = trimStr(pathOrName || "");
    if (!s) return "";

    try {
        var f = new File(s);
        if (f && f.parent) return String(f.parent.fsName || f.parent.fullName || "");
    } catch (e1) {}

    s = s.replace(/\\/g, "/");
    if (s.charAt(s.length - 1) === "/") s = s.substring(0, s.length - 1);
    var cut = s.lastIndexOf("/");
    return (cut > 0) ? s.substring(0, cut) : "";
}

function getPreferredFolderLabel(pathOrName) {
    var s = normalizeAsciiDigits(trimStr(pathOrName || ""));
    if (!s) return "";
    s = s.replace(/\\/g, "/");
    if (s.charAt(s.length - 1) === "/") s = s.substring(0, s.length - 1);

    var parts = s.split("/");
    if (parts.length < 3) return getParentFolderName(s);

    for (var i = parts.length - 3; i >= 0; i--) {
        var part = trimStr(parts[i] || "");
        if (isAcceptedSourceFolderLabel(part)) return part;
    }

    return getParentFolderName(s);
}

function isAcceptedSourceFolderLabel(label) {
    var value = normalizeAsciiDigits(trimStr(label || ""));
    if (!value) return false;
    return /^(?:.+_\d{4}|.+_\d{6}|\d{5})$/.test(value);
}

function sanitizeFolderPaths(paths) {
    var out = [];
    var seen = {};
    if (!paths || !paths.length) return out;

    for (var i = 0; i < paths.length; i++) {
        var value = trimStr(paths[i] || "");
        if (!value) continue;

        var seenKey = String(value).toLowerCase();
        if (seen[seenKey]) continue;
        seen[seenKey] = true;
        out.push(value);
    }

    return out;
}

function setLastSourceFolderPathsFromSources(sources) {
    LAST_SOURCE_FOLDER_PATHS = sanitizeFolderPaths(getSourceFolderPathsFromSources(sources || []));
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
    } catch (e1) {}
    try {
        return getFileNameOnly(item.name);
    } catch (e2) {}
    return "unknown.png";
}

function getItemFilePath(item) {
    try {
        if (item.file && item.file.fsName) return String(item.file.fsName);
    } catch (e1) {}
    return "";
}

function sanitizeExportBaseName(name) {
    var s = trimStr(name || "");
    s = s.replace(/\.png$/i, "");
    s = s.replace(/[\r\n\t]+/g, " ");
    s = s.replace(/[<>:"\/\\|?*]/g, "_");
    s = s.replace(/\s+/g, " ");
    s = s.replace(/[. ]+$/g, "");
    s = trimStr(s);
    return s || "NESTER_EXPORT";
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
    var vb = item.visibleBounds;
    var w = Math.abs(vb[2] - vb[0]);
    var h = Math.abs(vb[1] - vb[3]);
    return { width: w, height: h };
}

function inToPt(v) { return v * PT_PER_IN; }
function inToStr(vPt) { return (vPt / PT_PER_IN).toFixed(2); }
function areaOf(w, h) { return w * h; }
function longSide(w, h) { return (w > h) ? w : h; }
function shortSide(w, h) { return (w < h) ? w : h; }

function clamp(v, minV, maxV) {
    return Math.max(minV, Math.min(maxV, v));
}

function startsWith(str, prefix) {
    return String(str).indexOf(prefix) === 0;
}

function getOutputMetaNote(kind) {
    return OUTPUT_META_NOTE_PREFIX + String(kind || "");
}

function isOutputMetaItem(item, kind) {
    var note = "";
    try { note = String(item.note || ""); } catch (e1) { note = ""; }
    if (!note) return false;
    return kind ? note === getOutputMetaNote(kind) : startsWith(note, OUTPUT_META_NOTE_PREFIX);
}

function cloneRect(r) {
    return { x: r.x, y: r.y, w: r.w, h: r.h };
}

function cloneRectList(list) {
    var out = [];
    for (var i = 0; i < list.length; i++) out.push(cloneRect(list[i]));
    return out;
}

function cloneArray(list) {
    var out = [];
    for (var i = 0; i < list.length; i++) out.push(list[i]);
    return out;
}

function cloneSettings(settings) {
    var out = {};
    for (var key in settings) {
        if (settings.hasOwnProperty(key)) out[key] = settings[key];
    }
    return out;
}

function nextBlockUid() {
    var id = "B" + BLOCK_UID_COUNTER;
    BLOCK_UID_COUNTER += 1;
    return id;
}

function resetBlockUidCounter() {
    BLOCK_UID_COUNTER = 1;
}

function uniquePushPartition(list, arr) {
    var key = arr.join("-");
    for (var i = 0; i < list.length; i++) {
        if (list[i].join("-") === key) return;
    }
    list.push(arr);
}

function sortNumericDesc(a, b) {
    return b - a;
}

// =====================================================
// JSON helpers
// =====================================================

function _jsonEscape(str) {
    return String(str)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, "\\\"")
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n");
}

function _jsonStringify(value) {
    if (typeof JSON !== "undefined" && JSON.stringify) return JSON.stringify(value);

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
    if (typeof JSON !== "undefined" && JSON.parse) return JSON.parse(String(str));
    return eval("(" + str + ")");
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
    return r && r.w > EPS && r.h > EPS;
}

function rectArea(r) { return r.w * r.h; }
function rectCenterX(r) { return r.x + (r.w / 2); }
function rectCenterY(r) { return r.y + (r.h / 2); }

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
    return { x: left, y: top, w: right - left, h: bottom - top };
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

// =====================================================
// Layer helpers
// =====================================================

function isOutputLayerName(name) {
    return startsWith(name, OUTPUT_LAYER_PREFIX);
}

function findLayerByName(doc, name) {
    try { return doc.layers.getByName(name); } catch (e1) {}
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
// Source normalization
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
        var overrideQty = quantityOverrides && quantityOverrides.hasOwnProperty(key) ? quantityOverrides[key] : null;
        var qty = (overrideQty === null || overrideQty === undefined) ? baseQty : Math.max(1, overrideQty);

        result.push({
            id: "SRC_" + ordinal,
            ordinal: ordinal,
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

function collectAndNormalizeSources(doc, settings) {
    var sources = collectPlacedItems(doc, settings.quantityOverrides);
    if (!sources || sources.length === 0) throw new Error("No placed items found in the current document.");
    return sources;
}

// =====================================================
// Personality + effort config
// =====================================================

function clonePersonalityTemplate(name) {
    var src = PERSONALITY_LIBRARY[name];
    var out = {};
    for (var k in src) {
        if (src.hasOwnProperty(k)) out[k] = src[k];
    }
    return out;
}

function createPersonality(name, variantTag) {
    var p = clonePersonalityTemplate(name);
    p.variantTag = variantTag || "base";
    p.strategyName = (variantTag && variantTag !== "base") ? (name + ":" + variantTag) : name;

    if (variantTag === "wide") {
        p.widthFillPriority = clamp(p.widthFillPriority + 6, 0, 100);
        p.cavityWeight = clamp(p.cavityWeight + 6, 0, 100);
        p.maxBlockAspectRatio += 0.3;
    } else if (variantTag === "dense") {
        p.splitAggressiveness = clamp(p.splitAggressiveness + 10, 0, 100);
        p.holeFillAggressiveness = clamp(p.holeFillAggressiveness + 10, 0, 100);
        p.compactionAggressiveness = clamp(p.compactionAggressiveness + 8, 0, 100);
    } else if (variantTag === "strict") {
        p.orderDiscipline = clamp(p.orderDiscipline + 8, 0, 100);
        p.continuityWeight = clamp(p.continuityWeight + 8, 0, 100);
        p.maxBlockAspectRatio = Math.max(2.2, p.maxBlockAspectRatio - 0.25);
    }

    return p;
}

function getEffortConfig(searchEffort) {
    if (searchEffort === "High") {
        return {
            name: "High",
            plansPerSource: 4,
            comboLimit: 7,
            orderVariantLimit: 6,
            compactionPasses: 2,
            splitPasses: 2,
            holeAttemptsPerHole: 4,
            repairSweeps: 2,
            holeScanLimit: 10,
            primaryPlanLimit: 8,
            futureFitCheckCount: 8
        };
    }

    return {
        name: "Normal",
        plansPerSource: 2,
        comboLimit: 3,
        orderVariantLimit: 3,
        compactionPasses: 1,
        splitPasses: 1,
        holeAttemptsPerHole: 2,
        repairSweeps: 1,
        holeScanLimit: 6,
        primaryPlanLimit: 4,
        futureFitCheckCount: 5
    };
}

function buildStrategyCandidates(settings) {
    var strategies = [];

    function pushStrategy(presetName, personalityName, variantTag, isBaseline) {
        strategies.push({
            presetName: presetName,
            personalityName: personalityName,
            personality: createPersonality(personalityName, variantTag),
            isBaseline: isBaseline === true
        });
    }

    if (settings.optimizePreset === "Auto") {
        pushStrategy("Auto", "Baseline", "base", true);
        pushStrategy("Auto", "Compact", "base", false);
        pushStrategy("Auto", "Balanced", "base", false);
        pushStrategy("Auto", "Ordered", "base", false);

        if (settings.searchEffort === "High") {
            pushStrategy("Auto", "Compact", "wide", false);
            pushStrategy("Auto", "Balanced", "dense", false);
            pushStrategy("Auto", "Ordered", "strict", false);
        }
    } else {
        pushStrategy(settings.optimizePreset, "Baseline", "base", true);
        pushStrategy(settings.optimizePreset, settings.optimizePreset, "base", false);

        if (settings.searchEffort === "High") {
            if (settings.optimizePreset === "Compact") pushStrategy(settings.optimizePreset, "Compact", "wide", false);
            else if (settings.optimizePreset === "Balanced") pushStrategy(settings.optimizePreset, "Balanced", "dense", false);
            else if (settings.optimizePreset === "Ordered") pushStrategy(settings.optimizePreset, "Ordered", "strict", false);
        }
    }

    return strategies;
}

// =====================================================
// Block templates and planning
// =====================================================

function createBlockTemplateFromGrid(src, grid, planKind, planScore) {
    return {
        sourceId: src.id,
        sourceOrdinal: src.ordinal,
        sourceKey: src.key,
        sourceRef: src.ref,
        sourceName: src.name,
        sourceBaseQty: src.baseQty,
        sourceQty: src.qty,
        sourceWidth: src.width,
        sourceHeight: src.height,
        sourceArea: src.area,
        sourceLongSide: src.longSide,
        sourceShortSide: src.shortSide,
        count: grid.count,
        cols: grid.cols,
        rows: grid.rows,
        emptyCells: grid.emptyCells,
        rotatedInside: grid.rotatedInside,
        cellW: grid.cellW,
        cellH: grid.cellH,
        blockW: grid.blockW,
        blockH: grid.blockH,
        area: grid.blockW * grid.blockH,
        longSide: longSide(grid.blockW, grid.blockH),
        shortSide: shortSide(grid.blockW, grid.blockH),
        ratio: grid.ratio,
        widthFill: grid.widthFill,
        planKind: planKind || "primary",
        sourcePlanScore: planScore || 0
    };
}

function instantiateBlockFromTemplate(template) {
    var out = {};
    for (var key in template) {
        if (template.hasOwnProperty(key)) out[key] = template[key];
    }
    out.uid = nextBlockUid();
    return out;
}

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
        moved += 1;
    }

    arr.sort(sortNumericDesc);
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
        }
    }

    return out;
}

function chooseBestGridForCount(src, count, settings, personality, favorWide) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);
    var spacing = inToPt(settings.spacingIn);
    var maxAspect = personality.maxBlockAspectRatio;
    var widthBias = personality.widthFillPriority / 100.0;
    var continuityBias = personality.continuityWeight / 100.0;

    var candidates = [];
    var orientations = [{ rotatedInside: false, cellW: src.width, cellH: src.height }];

    if (settings.allowItemRotationInBlock) {
        orientations.push({ rotatedInside: true, cellW: src.height, cellH: src.width });
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
            var widthFill = blockW / Math.max(sheetWidth, EPS);
            var heightFill = blockH / Math.max(maxLength, EPS);
            var aspectPenalty = Math.max(0, ratio - maxAspect);
            var widePenalty = favorWide && rows > personality.dominantShelfRows
                ? (rows - personality.dominantShelfRows) * 10000
                : 0;

            var score =
                (emptyCells * (1400 + (continuityBias * 600))) +
                (aspectPenalty * aspectPenalty * 19000) +
                ((1 - widthFill) * 25000 * (0.55 + (widthBias * 0.45))) +
                (heightFill * 3200) +
                (ratio * 420) +
                (ori.rotatedInside ? 180 : 0) +
                widePenalty;

            candidates.push({
                count: count,
                cols: cols,
                rows: rows,
                emptyCells: emptyCells,
                rotatedInside: ori.rotatedInside,
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

function detectDominantSourceId(sources) {
    if (!sources || sources.length === 0) return null;
    var best = sources[0];
    for (var i = 1; i < sources.length; i++) {
        var src = sources[i];
        if (src.qty > best.qty) best = src;
        else if (src.qty === best.qty && src.area > best.area) best = src;
    }
    return best.id;
}

function buildSourcePlanCacheForSource(src, settings, personality, isDominant) {
    var countCache = {};
    for (var count = 1; count <= Math.max(1, src.qty); count++) {
        var grid = chooseBestGridForCount(src, count, settings, personality, isDominant);
        if (grid) countCache[count] = createBlockTemplateFromGrid(src, grid, count === 1 ? "piece" : "count", 0);
    }

    return {
        source: src,
        isDominant: isDominant,
        bestBlockByCount: countCache
    };
}

function buildSourcePlanCandidates(src, sourceCache, settings, personality, effortConfig) {
    var maxBlocks = Math.min(src.qty, personality.maxBlocksPerSource + (effortConfig.name === "High" ? 1 : 0));
    if (maxBlocks < 1) maxBlocks = 1;

    var partitions = generatePartitionCandidates(src.qty, maxBlocks);
    var plans = [];
    var seen = {};
    var widthBias = personality.widthFillPriority / 100.0;
    var continuityBias = personality.continuityWeight / 100.0;

    for (var p = 0; p < partitions.length; p++) {
        var part = partitions[p];
        var layouts = [];
        var failed = false;
        var shapeKeyParts = [];
        var totalArea = 0;
        var totalEmpty = 0;
        var totalAspect = 0;
        var totalWidthFill = 0;
        var maxHeight = 0;

        for (var i = 0; i < part.length; i++) {
            var template = sourceCache.bestBlockByCount[part[i]];
            if (!template) {
                failed = true;
                break;
            }

            layouts.push(template);
            shapeKeyParts.push([template.count, template.cols, template.rows, template.rotatedInside ? 1 : 0].join(":"));
            totalArea += template.area;
            totalEmpty += template.emptyCells;
            totalAspect += Math.max(0, template.ratio - personality.maxBlockAspectRatio);
            totalWidthFill += template.widthFill;
            if (template.blockH > maxHeight) maxHeight = template.blockH;
        }

        if (failed) continue;

        var planKey = shapeKeyParts.join("|");
        if (seen[planKey]) continue;
        seen[planKey] = true;

        var planScore =
            (layouts.length * (1850 + (continuityBias * 950))) +
            (totalEmpty * 1650) +
            (totalAspect * totalAspect * 14000) +
            (maxHeight * 2.2) +
            ((1 - (totalWidthFill / Math.max(layouts.length, 1))) * 26000 * (0.6 + (widthBias * 0.4))) +
            (totalArea * 0.00001);

        plans.push({
            source: src,
            planScore: planScore,
            partition: cloneArray(part),
            templates: cloneArray(layouts),
            isDominant: sourceCache.isDominant
        });
    }

    plans.sort(function(a, b) {
        if (a.planScore !== b.planScore) return a.planScore - b.planScore;
        return a.templates.length - b.templates.length;
    });

    return plans.slice(0, effortConfig.plansPerSource);
}

function generateSourceBlockPlans(sources, personality, effortConfig, settings, debugRoot) {
    var dominantId = detectDominantSourceId(sources);
    var bySource = {};
    var sourcePlanSets = [];
    var failedSources = [];
    var sourcePlanSummary = [];

    for (var i = 0; i < sources.length; i++) {
        var src = sources[i];
        var cache = buildSourcePlanCacheForSource(src, settings, personality, src.id === dominantId);
        bySource[src.id] = cache;

        var plans = buildSourcePlanCandidates(src, cache, settings, personality, effortConfig);
        if (!plans || plans.length === 0) {
            failedSources.push(src.name);
            continue;
        }

        sourcePlanSets.push({
            source: src,
            plans: plans,
            isDominant: cache.isDominant
        });

        var topScores = [];
        for (var j = 0; j < plans.length; j++) topScores.push(Math.round(plans[j].planScore));

        sourcePlanSummary.push({
            sourceId: src.id,
            sourceKey: src.key,
            sourceName: src.name,
            qty: src.qty,
            isDominant: cache.isDominant,
            planCount: plans.length,
            topPlanScores: topScores
        });
    }

    debugRoot.sourcePlanSummary = sourcePlanSummary;
    debugRoot.failedSources = failedSources;

    return {
        dominantId: dominantId,
        bySource: bySource,
        sourcePlanSets: sourcePlanSets,
        failedSources: failedSources
    };
}

function buildGlobalCandidatePlanSets(sourcePlanData, personality, effortConfig, settings, debugRoot) {
    var planSets = sourcePlanData.sourcePlanSets;
    if (!planSets || planSets.length === 0) return [];

    var beam = [{ plans: [], heuristicScore: 0, totalBlockCount: 0 }];

    for (var i = 0; i < planSets.length; i++) {
        var next = [];
        var localPlans = planSets[i].plans;

        for (var b = 0; b < beam.length; b++) {
            for (var p = 0; p < localPlans.length; p++) {
                next.push({
                    plans: beam[b].plans.concat([localPlans[p]]),
                    heuristicScore: beam[b].heuristicScore + localPlans[p].planScore,
                    totalBlockCount: beam[b].totalBlockCount + localPlans[p].templates.length
                });
            }
        }

        next.sort(function(a, b) {
            if (a.heuristicScore !== b.heuristicScore) return a.heuristicScore - b.heuristicScore;
            return a.totalBlockCount - b.totalBlockCount;
        });

        if (next.length > effortConfig.comboLimit) next = next.slice(0, effortConfig.comboLimit);
        beam = next;
    }

    var summaries = [];
    for (var s = 0; s < beam.length; s++) {
        summaries.push({
            comboIndex: s,
            heuristicScore: Math.round(beam[s].heuristicScore),
            totalBlockCount: beam[s].totalBlockCount
        });
    }
    debugRoot.globalCandidatePlanSets = summaries;

    return beam;
}

function materializeBlocksFromSelectedPlans(planBundle, sourcePlanData) {
    var blocks = [];
    for (var i = 0; i < planBundle.plans.length; i++) {
        var plan = planBundle.plans[i];
        for (var j = 0; j < plan.templates.length; j++) {
            var block = instantiateBlockFromTemplate(plan.templates[j]);
            block.isDominant = plan.isDominant;
            block.sourcePlanScore = plan.planScore;
            blocks.push(block);
        }
    }
    return blocks;
}

// =====================================================
// Ordering
// =====================================================

function compareStrings(a, b) {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
}

function buildBlockOrderVariants(blocks, personality, effortConfig, mode) {
    var variants = [];
    var seen = {};
    var orderBias = personality.orderDiscipline / 100.0;

    function pushVariant(order, label) {
        var sig = [];
        for (var i = 0; i < order.length; i++) sig.push(order[i].uid);
        sig = sig.join("|");
        if (seen[sig]) return;
        seen[sig] = true;
        variants.push({ label: label, order: order });
    }

    var base = cloneArray(blocks);
    base.sort(function(a, b) {
        if (personality.baselineLike) {
            if (b.area !== a.area) return b.area - a.area;
            if (b.longSide !== a.longSide) return b.longSide - a.longSide;
            return compareStrings(a.sourceName, b.sourceName);
        }

        if (mode === "repair") {
            if (a.count !== b.count) return a.count - b.count;
            if (a.area !== b.area) return a.area - b.area;
            return compareStrings(a.sourceName, b.sourceName);
        }

        if (mode === "compaction") {
            if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
            if (a.blockH !== b.blockH) return a.blockH - b.blockH;
            return b.area - a.area;
        }

        if (orderBias > 0.72) {
            var sourceCmp = compareStrings(a.sourceName, b.sourceName);
            if (sourceCmp !== 0) return sourceCmp;
        }

        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        if (b.area !== a.area) return b.area - a.area;
        return a.blockH - b.blockH;
    });
    pushVariant(base, "base");

    var widthFirst = cloneArray(blocks);
    widthFirst.sort(function(a, b) {
        if (b.widthFill !== a.widthFill) return b.widthFill - a.widthFill;
        if (a.blockH !== b.blockH) return a.blockH - b.blockH;
        return b.area - a.area;
    });
    pushVariant(widthFirst, "widthFirst");

    var areaFirst = cloneArray(blocks);
    areaFirst.sort(function(a, b) {
        if (b.area !== a.area) return b.area - a.area;
        if (b.count !== a.count) return b.count - a.count;
        return compareStrings(a.sourceName, b.sourceName);
    });
    pushVariant(areaFirst, "areaFirst");

    if (!personality.baselineLike) {
        var sourceGrouped = cloneArray(blocks);
        sourceGrouped.sort(function(a, b) {
            var sourceCmp = compareStrings(a.sourceName, b.sourceName);
            if (sourceCmp !== 0) return sourceCmp;
            if (b.count !== a.count) return b.count - a.count;
            return b.area - a.area;
        });
        pushVariant(sourceGrouped, "sourceGrouped");
    }

    if (mode === "repair" || personality.splitAggressiveness > 40) {
        var smallFirst = cloneArray(blocks);
        smallFirst.sort(function(a, b) {
            if (a.count !== b.count) return a.count - b.count;
            if (a.area !== b.area) return a.area - b.area;
            return compareStrings(a.sourceName, b.sourceName);
        });
        pushVariant(smallFirst, "smallFirst");
    }

    if (mode === "compaction" || personality.compactionAggressiveness > 55) {
        var interleave = [];
        var grouped = {};
        var keys = [];
        var gi;
        for (gi = 0; gi < blocks.length; gi++) {
            var srcId = blocks[gi].sourceId;
            if (!grouped[srcId]) {
                grouped[srcId] = [];
                keys.push(srcId);
            }
            grouped[srcId].push(blocks[gi]);
        }
        keys.sort();
        var active = true;
        while (active) {
            active = false;
            for (gi = 0; gi < keys.length; gi++) {
                if (grouped[keys[gi]].length > 0) {
                    interleave.push(grouped[keys[gi]].shift());
                    active = true;
                }
            }
        }
        if (interleave.length > 0) pushVariant(interleave, "interleave");
    }

    if (variants.length > effortConfig.orderVariantLimit) variants = variants.slice(0, effortConfig.orderVariantLimit);
    return variants;
}

// =====================================================
// Placement
// =====================================================

function countPlacedCopies(layoutState) {
    var total = 0;
    if (!layoutState || !layoutState.placed) return 0;
    for (var i = 0; i < layoutState.placed.length; i++) total += layoutState.placed[i].block.count;
    return total;
}

function countUnplacedCopies(layoutState) {
    var total = 0;
    if (!layoutState || !layoutState.unplaced) return 0;
    for (var i = 0; i < layoutState.unplaced.length; i++) total += layoutState.unplaced[i].count;
    return total;
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
            rotatedOnSheet: p.rotatedOnSheet,
            stageName: p.stageName,
            placementReason: p.placementReason || ""
        });
    }
    return out;
}

function cloneLayoutState(layoutState) {
    return {
        placed: clonePlacedList(layoutState.placed || []),
        unplaced: cloneArray(layoutState.unplaced || []),
        freeRects: cloneRectList(layoutState.freeRects || []),
        usedLength: layoutState.usedLength || 0,
        metrics: layoutState.metrics ? cloneSettings(layoutState.metrics) : null,
        debug: layoutState.debug || null,
        validation: layoutState.validation ? cloneSettings(layoutState.validation) : null,
        score: layoutState.score ? cloneSettings(layoutState.score) : null,
        strategyName: layoutState.strategyName || "",
        stageName: layoutState.stageName || ""
    };
}

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

    if (usedTop > frTop + EPS) count += 1;
    if (usedBottom < frBottom - EPS) count += 1;
    if (usedLeft > frLeft + EPS) count += 1;
    if (usedRight < frRight - EPS) count += 1;
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
    var cx = freeRect.x + (pw / 2);
    var cy = freeRect.y + (ph / 2);

    for (var i = 0; i < placedSoFar.length; i++) {
        var p = placedSoFar[i];
        if (p.block.sourceId !== sourceId) continue;
        var pcx = p.x + (p.paddedW / 2);
        var pcy = p.y + (p.paddedH / 2);
        var dx = cx - pcx;
        var dy = cy - pcy;
        var dist = Math.sqrt(dx * dx + dy * dy);
        bonus += 1 / Math.max(dist, 1);
    }

    return bonus;
}

function estimateCavityPenaltyForPlacement(freeRect, pw, ph) {
    var rightW = freeRect.w - pw;
    var bottomH = freeRect.h - ph;
    var penalty = 0;

    if (rightW > EPS && freeRect.h > EPS) {
        penalty += longSide(rightW, freeRect.h) / Math.max(shortSide(rightW, freeRect.h), EPS);
    }
    if (freeRect.w > EPS && bottomH > EPS) {
        penalty += longSide(freeRect.w, bottomH) / Math.max(shortSide(freeRect.w, bottomH), EPS);
    }

    return penalty;
}

function getPlacementFootprint(x, y, contentW, contentH, settings) {
    var spacing = inToPt(settings.spacingIn);
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);
    var rightEdge = x + contentW;
    var bottomEdge = y + contentH;
    var rightGap = (rightEdge >= sheetWidth - EPS) ? 0 : spacing;
    var bottomGap = (bottomEdge >= maxLength - EPS) ? 0 : spacing;

    return {
        x: x,
        y: y,
        contentW: contentW,
        contentH: contentH,
        w: contentW + rightGap,
        h: contentH + bottomGap,
        rightGap: rightGap,
        bottomGap: bottomGap
    };
}

function blockFitsRect(block, rect, settings) {
    var options = [{ rotatedOnSheet: false, contentW: block.blockW, contentH: block.blockH }];

    if (settings.allowBlockRotationOnSheet) {
        options.push({ rotatedOnSheet: true, contentW: block.blockH, contentH: block.blockW });
    }

    for (var i = 0; i < options.length; i++) {
        var footprint = getPlacementFootprint(rect.x, rect.y, options[i].contentW, options[i].contentH, settings);
        if (footprint.w <= rect.w + EPS && footprint.h <= rect.h + EPS) {
            return {
                rotatedOnSheet: options[i].rotatedOnSheet,
                contentW: options[i].contentW,
                contentH: options[i].contentH,
                w: footprint.w,
                h: footprint.h,
                rightGap: footprint.rightGap,
                bottomGap: footprint.bottomGap
            };
        }
    }
    return null;
}

function estimateFutureFitBonus(nextFreeRects, remainingBlocks, settings, effortConfig) {
    if (!remainingBlocks || remainingBlocks.length === 0) return 0;

    var checkedBlocks = Math.min(remainingBlocks.length, effortConfig.futureFitCheckCount);
    var bestBonus = 0;
    var totalBonus = 0;
    var sheetWidth = inToPt(settings.sheetWidthIn);

    for (var i = 0; i < checkedBlocks; i++) {
        var block = remainingBlocks[i];
        var blockBest = 0;

        for (var j = 0; j < nextFreeRects.length; j++) {
            var fr = nextFreeRects[j];
            var fit = blockFitsRect(block, fr, settings);
            if (!fit) continue;

            var widthSlack = Math.max(0, fr.w - fit.w);
            var snugWidth = 1 - clamp(widthSlack / Math.max(sheetWidth, EPS), 0, 1);
            var topBias = 1 / (1 + (fr.y / (PT_PER_IN * 8)));
            var match = (block.area * 0.015) * (0.55 + (0.45 * snugWidth)) * topBias;
            if (match > blockBest) blockBest = match;
        }

        totalBonus += blockBest;
        if (blockBest > bestBonus) bestBonus = blockBest;
    }

    return (bestBonus * 2.1) + (totalBonus * 0.4);
}

function computePlacementScore(freeRect, pw, ph, currentUsedLength, placedSoFar, settings, block, remainingBlocks, personality, effortConfig) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var projectedBottom = freeRect.y + ph;
    var projectedUsedLength = projectedBottom > currentUsedLength ? projectedBottom : currentUsedLength;
    var projectedUsedIn = projectedUsedLength / PT_PER_IN;

    var nextFreeRects = placeAndUpdateFreeRects([cloneRect(freeRect)], {
        x: freeRect.x,
        y: freeRect.y,
        w: pw,
        h: ph
    });

    var futureFitBonus = estimateFutureFitBonus(nextFreeRects, remainingBlocks, settings, effortConfig);
    var fragPenalty = estimateSplitCount(freeRect, { x: freeRect.x, y: freeRect.y, w: pw, h: ph });
    var alignBonus = estimateAlignmentBonus(freeRect, pw, ph, placedSoFar, sheetWidth);
    var sourceBonus = estimateSourceProximityBonus(freeRect, pw, ph, placedSoFar, block.sourceId);
    var widthFill = pw / Math.max(sheetWidth, EPS);
    var cavityPenalty = estimateCavityPenaltyForPlacement(freeRect, pw, ph);
    var widthBias = personality.widthFillPriority / 100.0;
    var continuityBias = personality.continuityWeight / 100.0;
    var cavityBias = personality.cavityWeight / 100.0;
    var fragmentBias = personality.fragmentationWeight / 100.0;

    var total =
        (projectedUsedIn * 1000) +
        ((1 - widthFill) * 24000 * (0.55 + (widthBias * 0.45))) +
        (fragPenalty * 650 * fragmentBias) +
        (cavityPenalty * 1200 * cavityBias) -
        (alignBonus * 180) -
        (sourceBonus * 16000 * continuityBias) -
        futureFitBonus;

    return {
        total: total,
        projectedUsedLength: projectedUsedLength,
        projectedUsedIn: projectedUsedIn,
        widthFill: widthFill,
        cavityPenalty: cavityPenalty
    };
}

function findBestPlacementForBlock(block, freeRects, currentUsedLength, placedSoFar, settings, personality, effortConfig, remainingBlocks, preferredRect) {
    var orientations = [{ rotatedOnSheet: false, w: block.blockW, h: block.blockH }];

    if (settings.allowBlockRotationOnSheet) {
        orientations.push({ rotatedOnSheet: true, w: block.blockH, h: block.blockW });
    }

    var best = null;

    for (var o = 0; o < orientations.length; o++) {
        var ori = orientations[o];

        for (var i = 0; i < freeRects.length; i++) {
            var fr = freeRects[i];
            if (preferredRect && (
                Math.abs(fr.x - preferredRect.x) > EPS ||
                Math.abs(fr.y - preferredRect.y) > EPS ||
                Math.abs(fr.w - preferredRect.w) > EPS ||
                Math.abs(fr.h - preferredRect.h) > EPS
            )) continue;

            var footprint = getPlacementFootprint(fr.x, fr.y, ori.w, ori.h, settings);
            var paddedW = footprint.w;
            var paddedH = footprint.h;
            if (paddedW > fr.w + EPS || paddedH > fr.h + EPS) continue;

            var score = computePlacementScore(
                fr,
                paddedW,
                paddedH,
                currentUsedLength,
                placedSoFar,
                settings,
                block,
                remainingBlocks,
                personality,
                effortConfig
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

            var betterByLength = score.projectedUsedLength < (best.score.projectedUsedLength - EPS);
            var tieOnLength = Math.abs(score.projectedUsedLength - best.score.projectedUsedLength) <= EPS;
            var betterByFill = tieOnLength && (score.widthFill > best.score.widthFill + 1e-6);
            var tieOnFill = tieOnLength && Math.abs(score.widthFill - best.score.widthFill) <= 1e-6;
            var betterByScore = tieOnFill && (score.total < best.score.total);

            if (betterByLength || betterByFill || betterByScore) {
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

    return best;
}

function buildLayoutState(placed, unplaced, freeRects, usedLength, strategyName, stageName, debugRoot) {
    return {
        placed: placed || [],
        unplaced: unplaced || [],
        freeRects: freeRects || [],
        usedLength: usedLength || 0,
        metrics: null,
        debug: debugRoot || null,
        validation: null,
        score: null,
        strategyName: strategyName || "",
        stageName: stageName || ""
    };
}

function nestBlocksByOrder(order, strategyName, stageName, settings, personality, effortConfig) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);
    var freeRects = [{ x: 0, y: 0, w: sheetWidth, h: maxLength }];
    var placed = [];
    var unplaced = [];
    var usedLength = 0;

    for (var i = 0; i < order.length; i++) {
        var block = order[i];
        var remaining = order.slice(i + 1);
        var placement = findBestPlacementForBlock(block, freeRects, usedLength, placed, settings, personality, effortConfig, remaining, null);
        if (!placement) {
            unplaced.push(block);
            continue;
        }

        placed.push({
            block: block,
            x: placement.x,
            y: placement.y,
            paddedW: placement.w,
            paddedH: placement.h,
            rotatedOnSheet: placement.rotatedOnSheet,
            stageName: stageName,
            placementReason: "primary"
        });

        freeRects = placeAndUpdateFreeRects(freeRects, placement);
        if (placement.y + placement.h > usedLength) usedLength = placement.y + placement.h;
    }

    return buildLayoutState(placed, unplaced, freeRects, usedLength, strategyName, stageName, null);
}

function runBasicBackfill(layoutState, settings, personality, effortConfig, stageName) {
    var current = cloneLayoutState(layoutState);
    var improved = true;

    while (improved) {
        improved = false;
        current.unplaced.sort(function(a, b) {
            if (a.count !== b.count) return a.count - b.count;
            if (a.area !== b.area) return a.area - b.area;
            return compareStrings(a.sourceName, b.sourceName);
        });

        for (var i = 0; i < current.unplaced.length; i++) {
            var block = current.unplaced[i];
            var remaining = current.unplaced.slice(0, i).concat(current.unplaced.slice(i + 1));
            var placement = findBestPlacementForBlock(block, current.freeRects, current.usedLength, current.placed, settings, personality, effortConfig, remaining, null);
            if (!placement) continue;

            current.placed.push({
                block: block,
                x: placement.x,
                y: placement.y,
                paddedW: placement.w,
                paddedH: placement.h,
                rotatedOnSheet: placement.rotatedOnSheet,
                stageName: stageName,
                placementReason: "backfill"
            });

            current.freeRects = placeAndUpdateFreeRects(current.freeRects, placement);
            if (placement.y + placement.h > current.usedLength) current.usedLength = placement.y + placement.h;
            current.unplaced.splice(i, 1);
            improved = true;
            break;
        }
    }

    current.stageName = stageName;
    return current;
}

// =====================================================
// Metrics / validation / scoring
// =====================================================

function getRequiredPlacedRect(placedBlock, settings) {
    var w = placedBlock.rotatedOnSheet ? placedBlock.block.blockH : placedBlock.block.blockW;
    var h = placedBlock.rotatedOnSheet ? placedBlock.block.blockW : placedBlock.block.blockH;
    return getPlacementFootprint(placedBlock.x, placedBlock.y, w, h, settings);
}

function buildLooseFitCandidates(layoutState, sourcePlanCache) {
    var out = [];
    var seenSource = {};

    for (var i = 0; i < layoutState.unplaced.length; i++) {
        var block = layoutState.unplaced[i];
        if (seenSource[block.sourceId]) continue;
        seenSource[block.sourceId] = true;

        var sourceCache = sourcePlanCache.bySource[block.sourceId];
        if (sourceCache && sourceCache.bestBlockByCount[1]) out.push(sourceCache.bestBlockByCount[1]);
    }

    return out;
}

function analyzeWidthUtilizationBands(layoutState, settings) {
    var bandCount = 4;
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var usedLength = Math.max(layoutState.usedLength, inToPt(settings.spacingIn));
    var bands = [];

    for (var i = 0; i < bandCount; i++) {
        var bandTop = (usedLength / bandCount) * i;
        var bandBottom = (usedLength / bandCount) * (i + 1);
        var area = 0;

        for (var j = 0; j < layoutState.placed.length; j++) {
            var p = layoutState.placed[j];
            var rect = getRequiredPlacedRect(p, settings);
            var overlapTop = Math.max(rect.y, bandTop);
            var overlapBottom = Math.min(rect.y + rect.h, bandBottom);
            if (overlapBottom <= overlapTop + EPS) continue;
            area += rect.w * (overlapBottom - overlapTop);
        }

        bands.push(clamp(area / Math.max(sheetWidth * (bandBottom - bandTop), EPS), 0, 1));
    }

    return bands;
}

function analyzeFreeRectangles(layoutState, settings, sourcePlanCache) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var usedLength = Math.max(layoutState.usedLength, inToPt(settings.spacingIn));
    var capacityArea = Math.max(sheetWidth * usedLength, EPS);
    var looseFits = buildLooseFitCandidates(layoutState, sourcePlanCache);

    var largeRectCount = 0;
    var usableHoleArea = 0;
    var centralHoleArea = 0;
    var centralHoleScore = 0;
    var multiCopyHoleCount = 0;
    var skinnyFragmentArea = 0;
    var largestUsableRect = null;
    var largestCentralRect = null;

    for (var i = 0; i < layoutState.freeRects.length; i++) {
        var rect = layoutState.freeRects[i];
        var area = rectArea(rect);
        var usable = false;
        var multiCopy = false;
        var cx = rectCenterX(rect);
        var cy = rectCenterY(rect);

        if (area >= capacityArea * 0.01) largeRectCount += 1;

        for (var b = 0; b < layoutState.unplaced.length; b++) {
            if (blockFitsRect(layoutState.unplaced[b], rect, settings)) {
                usable = true;
                if (layoutState.unplaced[b].count >= 2 || layoutState.unplaced[b].planKind !== "piece") multiCopy = true;
                break;
            }
        }

        if (!usable) {
            for (b = 0; b < looseFits.length; b++) {
                if (blockFitsRect(looseFits[b], rect, settings)) {
                    usable = true;
                    if (rect.w >= (looseFits[b].blockW * 2)) multiCopy = true;
                    break;
                }
            }
        }

        if (rect.w < inToPt(1.0) || rect.h < inToPt(0.7)) skinnyFragmentArea += area;

        var central = (
            cx >= (sheetWidth * 0.25) &&
            cx <= (sheetWidth * 0.75) &&
            cy >= (usedLength * 0.2) &&
            cy <= (usedLength * 0.8)
        );

        if (usable) {
            usableHoleArea += area;
            if (!largestUsableRect || area > rectArea(largestUsableRect)) largestUsableRect = cloneRect(rect);
            if (multiCopy) multiCopyHoleCount += 1;
        }

        if (central && usable) {
            centralHoleArea += area;
            centralHoleScore += (area / capacityArea) * 100;
            if (!largestCentralRect || area > rectArea(largestCentralRect)) largestCentralRect = cloneRect(rect);
        }
    }

    return {
        largeRectCount: largeRectCount,
        usableHoleArea: usableHoleArea,
        centralHoleArea: centralHoleArea,
        centralHoleScore: centralHoleScore,
        multiCopyHoleCount: multiCopyHoleCount,
        skinnyFragmentArea: skinnyFragmentArea,
        largestUsableRect: largestUsableRect,
        largestCentralRect: largestCentralRect,
        widthUtilizationBands: analyzeWidthUtilizationBands(layoutState, settings)
    };
}

function sortPlacedReadingOrder(list) {
    list.sort(function(a, b) {
        if (Math.abs(a.y - b.y) > EPS) return a.y - b.y;
        if (Math.abs(a.x - b.x) > EPS) return a.x - b.x;
        return compareStrings(a.block.sourceName, b.block.sourceName);
    });
}

function evaluateSanityMetrics(layoutState, settings, sourcePlanCache) {
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var usedLength = Math.max(layoutState.usedLength, EPS);
    var usedArea = 0;
    var reading = clonePlacedList(layoutState.placed || []);
    var sourceSwitches = 0;
    var sourceFragments = {};
    var prevSource = null;
    var holeMetrics = analyzeFreeRectangles(layoutState, settings, sourcePlanCache);

    sortPlacedReadingOrder(reading);

    for (var i = 0; i < reading.length; i++) {
        usedArea += reading[i].block.area;
        if (prevSource !== null && prevSource !== reading[i].block.sourceId) sourceSwitches += 1;
        if (!sourceFragments[reading[i].block.sourceId]) sourceFragments[reading[i].block.sourceId] = 0;
        if (prevSource !== reading[i].block.sourceId) sourceFragments[reading[i].block.sourceId] += 1;
        prevSource = reading[i].block.sourceId;
    }

    var fragmentCount = 0;
    for (var key in sourceFragments) {
        if (sourceFragments.hasOwnProperty(key)) fragmentCount += sourceFragments[key];
    }

    var utilization = clamp(usedArea / Math.max(sheetWidth * usedLength, EPS), 0, 1);
    var bands = holeMetrics.widthUtilizationBands;
    var bandAvg = 0;
    var bandVariance = 0;
    for (i = 0; i < bands.length; i++) bandAvg += bands[i];
    bandAvg = bandAvg / Math.max(bands.length, 1);
    for (i = 0; i < bands.length; i++) bandVariance += Math.abs(bands[i] - bandAvg);

    var largestCentralRectArea = holeMetrics.largestCentralRect ? rectArea(holeMetrics.largestCentralRect) : 0;
    var capacityArea = Math.max(sheetWidth * usedLength, EPS);

    return {
        unplacedCopies: countUnplacedCopies(layoutState),
        usedLength: usedLength,
        utilization: utilization,
        largestCentralCavityRatio: largestCentralRectArea / capacityArea,
        usableHoleAreaRatio: holeMetrics.usableHoleArea / capacityArea,
        fragmentationScore: ((layoutState.freeRects.length * 0.04) + (holeMetrics.skinnyFragmentArea / capacityArea)),
        sourceFragmentationScore: ((sourceSwitches * 0.08) + (fragmentCount * 0.05)),
        denseRegionGapScore: ((holeMetrics.centralHoleScore * 0.01) + (bandVariance * 0.6)),
        widthWasteScore: (1 - utilization),
        sourceSwitches: sourceSwitches,
        sourceFragments: fragmentCount,
        holeMetrics: holeMetrics
    };
}

function validateLayoutState(layoutState, settings) {
    var valid = true;
    var reasons = [];
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var maxLength = inToPt(settings.maxLengthIn);
    var requiredRects = [];

    for (var i = 0; i < layoutState.placed.length; i++) {
        var p = layoutState.placed[i];
        var req = getRequiredPlacedRect(p, settings);
        requiredRects.push(req);

        if (req.x < -EPS || req.y < -EPS) {
            valid = false;
            reasons.push("Negative placement bounds");
            break;
        }
        if (req.x + req.w > sheetWidth + EPS) {
            valid = false;
            reasons.push("Sheet width exceeded");
            break;
        }
        if (req.y + req.h > maxLength + EPS) {
            valid = false;
            reasons.push("Sheet length exceeded");
            break;
        }
        if (p.paddedW + EPS < req.w || p.paddedH + EPS < req.h) {
            valid = false;
            reasons.push("Configured spacing lost");
            break;
        }
    }

    if (valid) {
        for (var a = 0; a < requiredRects.length; a++) {
            for (var b = a + 1; b < requiredRects.length; b++) {
                if (intersects(requiredRects[a], requiredRects[b])) {
                    valid = false;
                    reasons.push("Placed blocks overlap");
                    break;
                }
            }
            if (!valid) break;
        }
    }

    return { valid: valid, reasons: reasons, hardReject: !valid };
}

function scoreCandidateLayout(layoutState, settings, personality) {
    var sanity = layoutState.metrics ? layoutState.metrics.sanity : null;
    var usedNoTrailingGap = Math.max(0, layoutState.usedLength - inToPt(settings.spacingIn));
    var utilization = sanity ? sanity.utilization : 0;
    var unplacedPenalty = sanity ? sanity.unplacedCopies : countUnplacedCopies(layoutState);
    var lengthPenalty = Math.round(usedNoTrailingGap * 1000);
    var utilizationPenalty = Math.round((1 - utilization) * 100000);
    var cavityPenalty = sanity
        ? Math.round(
            (sanity.largestCentralCavityRatio * 100000) +
            (sanity.usableHoleAreaRatio * 50000) +
            (sanity.denseRegionGapScore * 12000)
        )
        : 0;
    var continuityPenalty = sanity ? Math.round(sanity.sourceFragmentationScore * 12000) : 0;
    var fragmentationPenalty = sanity ? Math.round((sanity.fragmentationScore * 12000) + (sanity.widthWasteScore * 6000)) : 0;
    var bonuses = (personality && personality.name === "Ordered") ? 50 : 0;

    var total =
        (unplacedPenalty * 1000000000000) +
        (lengthPenalty * 1000000) +
        (utilizationPenalty * 1000) +
        (cavityPenalty * 10) +
        continuityPenalty +
        fragmentationPenalty -
        bonuses;

    return {
        total: total,
        unplacedPenalty: unplacedPenalty,
        lengthPenalty: lengthPenalty,
        utilizationPenalty: utilizationPenalty,
        cavityPenalty: cavityPenalty,
        continuityPenalty: continuityPenalty,
        fragmentationPenalty: fragmentationPenalty,
        bonuses: bonuses,
        rankTuple: [
            unplacedPenalty,
            lengthPenalty,
            utilizationPenalty,
            cavityPenalty,
            continuityPenalty,
            fragmentationPenalty
        ],
        usedNoTrailingGap: usedNoTrailingGap,
        utilization: utilization
    };
}

function decorateLayoutState(layoutState, settings, sourcePlanCache, personality) {
    layoutState.validation = validateLayoutState(layoutState, settings);
    layoutState.metrics = { sanity: evaluateSanityMetrics(layoutState, settings, sourcePlanCache) };
    layoutState.score = scoreCandidateLayout(layoutState, settings, personality);
    return layoutState;
}

function compareScoreObjects(a, b) {
    if (!a && !b) return 0;
    if (!a) return 1;
    if (!b) return -1;
    if (a.total < b.total) return -1;
    if (a.total > b.total) return 1;
    return 0;
}

// =====================================================
// Split fallback helpers
// =====================================================

function buildFallbackPartitions(count, splitAggressiveness) {
    var out = [];
    if (count <= 1) return out;

    uniquePushPartition(out, [Math.ceil(count / 2), Math.floor(count / 2)]);

    if (count >= 4) {
        var thirds = [];
        var thirdSize = Math.ceil(count / 3);
        var remaining = count;
        while (remaining > 0) {
            thirds.push(Math.min(thirdSize, remaining));
            remaining -= Math.min(thirdSize, remaining);
        }
        thirds.sort(sortNumericDesc);
        uniquePushPartition(out, thirds);
    }

    if (count >= 6 || splitAggressiveness >= 70) {
        var quarterSize = Math.max(1, Math.ceil(count / 4));
        var quarters = [];
        remaining = count;
        while (remaining > 0) {
            quarters.push(Math.min(quarterSize, remaining));
            remaining -= Math.min(quarterSize, remaining);
        }
        quarters.sort(sortNumericDesc);
        uniquePushPartition(out, quarters);
    }

    if (count >= 8 && splitAggressiveness >= 80) {
        var equalTwo = [];
        for (var i = 0; i < count; i += 2) equalTwo.push(Math.min(2, count - i));
        equalTwo.sort(sortNumericDesc);
        uniquePushPartition(out, equalTwo);
    }

    var loose = [];
    for (i = 0; i < count; i++) loose.push(1);
    uniquePushPartition(out, loose);

    return out;
}

function splitBlockIntoFallbackOptions(block, sourcePlanCache, personality) {
    var sourceCache = sourcePlanCache.bySource[block.sourceId];
    if (!sourceCache) return [];

    var partitions = buildFallbackPartitions(block.count, personality.splitAggressiveness);
    var options = [];
    var seen = {};

    for (var p = 0; p < partitions.length; p++) {
        var partition = partitions[p];
        var blocks = [];
        var valid = true;

        for (var i = 0; i < partition.length; i++) {
            var template = sourceCache.bestBlockByCount[partition[i]];
            if (!template) {
                valid = false;
                break;
            }
            blocks.push(instantiateBlockFromTemplate(template));
        }

        if (!valid) continue;

        var key = partition.join("+");
        if (seen[key]) continue;
        seen[key] = true;

        options.push({
            label: "split:" + key,
            partition: partition,
            blocks: blocks
        });
    }

    return options;
}

function removeBlockByUid(blocks, uid) {
    var out = [];
    for (var i = 0; i < blocks.length; i++) {
        if (blocks[i].uid !== uid) out.push(blocks[i]);
    }
    return out;
}

function applyReplacementOption(baseState, targetBlock, replacementOption, settings, personality, effortConfig, stageName, preferredRect) {
    var state = cloneLayoutState(baseState);
    var replacementBlocks = cloneArray(replacementOption.blocks);
    var pendingUnplaced = removeBlockByUid(state.unplaced, targetBlock.uid);
    var firstPlaced = false;

    replacementBlocks.sort(function(a, b) {
        if (preferredRect) {
            var fitA = blockFitsRect(a, preferredRect, settings) ? 1 : 0;
            var fitB = blockFitsRect(b, preferredRect, settings) ? 1 : 0;
            if (fitA !== fitB) return fitB - fitA;
        }
        if (a.count !== b.count) return a.count - b.count;
        return a.area - b.area;
    });

    for (var i = 0; i < replacementBlocks.length; i++) {
        var block = replacementBlocks[i];
        var remaining = replacementBlocks.slice(i + 1).concat(pendingUnplaced);
        var placement = findBestPlacementForBlock(
            block,
            state.freeRects,
            state.usedLength,
            state.placed,
            settings,
            personality,
            effortConfig,
            remaining,
            (!firstPlaced && preferredRect) ? preferredRect : null
        );

        if (!placement) {
            pendingUnplaced.push(block);
            continue;
        }

        firstPlaced = true;
        state.placed.push({
            block: block,
            x: placement.x,
            y: placement.y,
            paddedW: placement.w,
            paddedH: placement.h,
            rotatedOnSheet: placement.rotatedOnSheet,
            stageName: stageName,
            placementReason: replacementOption.label
        });
        state.freeRects = placeAndUpdateFreeRects(state.freeRects, placement);
        if (placement.y + placement.h > state.usedLength) state.usedLength = placement.y + placement.h;
    }

    state.unplaced = pendingUnplaced;
    state.stageName = stageName;
    return state;
}

// =====================================================
// Solver stages
// =====================================================

function runPrimaryBlockLayout(candidatePlan, personality, effortConfig, settings, sourcePlanData, debugRoot) {
    var blocks = materializeBlocksFromSelectedPlans(candidatePlan, sourcePlanData);
    var variants = buildBlockOrderVariants(blocks, personality, effortConfig, "primary");
    var best = null;
    var variantSummary = [];

    for (var i = 0; i < variants.length; i++) {
        var variant = variants[i];
        var nested = nestBlocksByOrder(variant.order, personality.strategyName, "primary", settings, personality, effortConfig);
        nested = runBasicBackfill(nested, settings, personality, effortConfig, "primaryBackfill");
        nested.debug = debugRoot;
        decorateLayoutState(nested, settings, sourcePlanData, personality);

        variantSummary.push({
            label: variant.label,
            placedCopies: countPlacedCopies(nested),
            unplacedCopies: countUnplacedCopies(nested),
            totalScore: nested.score.total
        });

        if (!best || compareScoreObjects(nested.score, best.score) < 0) best = nested;
    }

    debugRoot.blockOrderVariantsConsidered = variantSummary;
    best.strategyName = personality.strategyName;
    return best;
}

function repackAllBlocks(layoutState, personality, effortConfig, settings, sourcePlanData, mode, debugRoot) {
    var allBlocks = [];
    var i;

    for (i = 0; i < layoutState.placed.length; i++) allBlocks.push(layoutState.placed[i].block);
    for (i = 0; i < layoutState.unplaced.length; i++) allBlocks.push(layoutState.unplaced[i]);

    var variants = buildBlockOrderVariants(allBlocks, personality, effortConfig, mode);
    var best = decorateLayoutState(cloneLayoutState(layoutState), settings, sourcePlanData, personality);

    for (i = 0; i < variants.length; i++) {
        var variant = variants[i];
        var nested = nestBlocksByOrder(variant.order, personality.strategyName, mode, settings, personality, effortConfig);
        nested = runBasicBackfill(nested, settings, personality, effortConfig, mode + "Backfill");
        decorateLayoutState(nested, settings, sourcePlanData, personality);
        if (compareScoreObjects(nested.score, best.score) < 0) best = nested;
    }

    best.debug = debugRoot;
    best.strategyName = personality.strategyName;
    best.stageName = mode;
    return best;
}

function runCompactionPass(layoutState, personality, effortConfig, settings, sourcePlanData, debugRoot) {
    var current = decorateLayoutState(cloneLayoutState(layoutState), settings, sourcePlanData, personality);
    var accepted = 0;

    for (var pass = 0; pass < effortConfig.compactionPasses; pass++) {
        var candidate = repackAllBlocks(current, personality, effortConfig, settings, sourcePlanData, "compaction", debugRoot);
        if (compareScoreObjects(candidate.score, current.score) < 0) {
            current = candidate;
            accepted += 1;
        } else {
            break;
        }
    }

    debugRoot.compactionStats = { passes: effortConfig.compactionPasses, accepted: accepted };
    current.stageName = "compaction";
    return current;
}

function runAdaptiveBlockSplittingPass(layoutState, sourcePlanCache, personality, effortConfig, settings, debugRoot) {
    var current = decorateLayoutState(cloneLayoutState(layoutState), settings, sourcePlanCache, personality);
    var stats = { attempted: 0, accepted: 0, blocksSplit: 0, copiesRecovered: 0 };

    if (personality.baselineLike || personality.splitAggressiveness <= 0) {
        debugRoot.splitStats = stats;
        return current;
    }

    for (var pass = 0; pass < effortConfig.splitPasses; pass++) {
        var improved = false;

        current.unplaced.sort(function(a, b) {
            if (b.count !== a.count) return b.count - a.count;
            if (b.area !== a.area) return b.area - a.area;
            return compareStrings(a.sourceName, b.sourceName);
        });

        for (var i = 0; i < current.unplaced.length; i++) {
            var target = current.unplaced[i];
            if (target.count <= 1) continue;

            var options = splitBlockIntoFallbackOptions(target, sourcePlanCache, personality);
            for (var o = 0; o < options.length; o++) {
                stats.attempted += 1;
                var candidate = applyReplacementOption(current, target, options[o], settings, personality, effortConfig, "adaptiveSplit", null);
                decorateLayoutState(candidate, settings, sourcePlanCache, personality);

                if (candidate.validation.hardReject) continue;
                if (compareScoreObjects(candidate.score, current.score) < 0) {
                    stats.accepted += 1;
                    stats.blocksSplit += 1;
                    stats.copiesRecovered += Math.max(0, countPlacedCopies(candidate) - countPlacedCopies(current));
                    current = candidate;
                    improved = true;
                    break;
                }
            }

            if (improved) break;
        }

        if (!improved) break;
    }

    debugRoot.splitStats = stats;
    current.stageName = "adaptiveSplit";
    return current;
}

function buildPrioritizedHoles(layoutState, settings, sourcePlanCache) {
    var holeInfo = analyzeFreeRectangles(layoutState, settings, sourcePlanCache);
    var holes = [];
    var sheetWidth = inToPt(settings.sheetWidthIn);
    var usedLength = Math.max(layoutState.usedLength, EPS);

    for (var i = 0; i < layoutState.freeRects.length; i++) {
        var rect = layoutState.freeRects[i];
        var area = rectArea(rect);
        var cx = rectCenterX(rect);
        var cy = rectCenterY(rect);
        var dx = Math.abs(cx - (sheetWidth / 2)) / Math.max(sheetWidth / 2, EPS);
        var dy = Math.abs(cy - (usedLength / 2)) / Math.max(usedLength / 2, EPS);
        var centrality = 1 - clamp((dx + dy) / 2, 0, 1);

        holes.push({
            rect: rect,
            area: area,
            centrality: centrality,
            fitValue: area * (0.55 + (0.45 * centrality))
        });
    }

    holes.sort(function(a, b) {
        if (b.fitValue !== a.fitValue) return b.fitValue - a.fitValue;
        return b.area - a.area;
    });

    return { holes: holes, metrics: holeInfo };
}

function runHoleFillingPass(layoutState, sourcePlanCache, personality, effortConfig, settings, debugRoot) {
    var current = decorateLayoutState(cloneLayoutState(layoutState), settings, sourcePlanCache, personality);
    var stats = {
        holesAnalyzed: 0,
        attempted: 0,
        accepted: 0,
        filledCopies: 0,
        acceptedReasons: []
    };

    if (personality.baselineLike || personality.holeFillAggressiveness <= 0) {
        debugRoot.holeFillStats = stats;
        return current;
    }

    var loopGuard = 0;
    while (loopGuard < effortConfig.holeScanLimit) {
        loopGuard += 1;
        var holeInfo = buildPrioritizedHoles(current, settings, sourcePlanCache);
        var holes = holeInfo.holes;
        var improved = false;

        for (var h = 0; h < holes.length && h < effortConfig.holeScanLimit; h++) {
            var hole = holes[h];
            stats.holesAnalyzed += 1;

            for (var u = 0; u < current.unplaced.length; u++) {
                var target = current.unplaced[u];
                var options = splitBlockIntoFallbackOptions(target, sourcePlanCache, personality);
                var tries = 0;

                for (var o = 0; o < options.length; o++) {
                    if (tries >= effortConfig.holeAttemptsPerHole) break;
                    tries += 1;
                    stats.attempted += 1;

                    var candidate = applyReplacementOption(current, target, options[o], settings, personality, effortConfig, "holeFill", hole.rect);
                    decorateLayoutState(candidate, settings, sourcePlanCache, personality);
                    if (candidate.validation.hardReject) continue;

                    var placedGain = countPlacedCopies(candidate) - countPlacedCopies(current);
                    var oldCentral = current.metrics.sanity.largestCentralCavityRatio;
                    var newCentral = candidate.metrics.sanity.largestCentralCavityRatio;
                    var centralImproved = newCentral <= oldCentral + 1e-6;

                    if (placedGain > 0 || (centralImproved && compareScoreObjects(candidate.score, current.score) < 0)) {
                        stats.accepted += 1;
                        stats.filledCopies += Math.max(0, placedGain);
                        stats.acceptedReasons.push(options[o].label + "@" + Math.round(hole.area));
                        current = candidate;
                        improved = true;
                        break;
                    }
                }

                if (improved) break;
            }

            if (improved) break;
        }

        if (!improved) break;
    }

    debugRoot.holeFillStats = stats;
    current.stageName = "holeFill";
    return current;
}

function runLocalRepairPass(layoutState, sourcePlanCache, personality, effortConfig, settings, debugRoot) {
    var current = decorateLayoutState(cloneLayoutState(layoutState), settings, sourcePlanCache, personality);
    var accepted = 0;

    for (var sweep = 0; sweep < effortConfig.repairSweeps; sweep++) {
        var candidate = repackAllBlocks(current, personality, effortConfig, settings, sourcePlanCache, "repair", debugRoot);
        if (compareScoreObjects(candidate.score, current.score) < 0) {
            current = candidate;
            accepted += 1;
        } else {
            break;
        }
    }

    debugRoot.localRepairStats = { sweeps: effortConfig.repairSweeps, accepted: accepted };
    current.stageName = "localRepair";
    return current;
}

// =====================================================
// Candidate solving
// =====================================================

function buildStageRecord(label, layoutState) {
    return {
        stage: label,
        placedCopies: countPlacedCopies(layoutState),
        unplacedCopies: countUnplacedCopies(layoutState),
        usedLengthIn: parseFloat(inToStr(layoutState.usedLength)),
        totalScore: layoutState.score ? layoutState.score.total : null,
        validation: layoutState.validation ? layoutState.validation.valid : null
    };
}

function buildDebugRoot(strategy, effortConfig) {
    return {
        requestedPreset: strategy.presetName,
        selectedPresetPersonality: strategy.personalityName,
        selectedSearchEffort: effortConfig.name,
        strategyName: strategy.personality.strategyName,
        isBaseline: strategy.isBaseline,
        sourcePlanSummary: [],
        globalCandidatePlanSets: [],
        blockOrderVariantsConsidered: [],
        stageRecords: [],
        splitStats: null,
        holeFillStats: null,
        localRepairStats: null,
        compactionStats: null,
        rejectedReasons: []
    };
}

function solveBaselineLayout(sources, settings, strategy, effortConfig) {
    var personality = strategy.personality;
    var debugRoot = buildDebugRoot(strategy, effortConfig);
    var sourcePlanData = generateSourceBlockPlans(sources, personality, effortConfig, settings, debugRoot);
    var planSets = buildGlobalCandidatePlanSets(sourcePlanData, personality, effortConfig, settings, debugRoot);

    if (planSets.length === 0) throw new Error("No valid baseline block plans could be built.");

    var baselinePlan = planSets[0];
    var blocks = materializeBlocksFromSelectedPlans(baselinePlan, sourcePlanData);
    var order = buildBlockOrderVariants(blocks, personality, effortConfig, "baseline")[0];
    var layoutState = nestBlocksByOrder(order.order, personality.strategyName, "baseline", settings, personality, effortConfig);
    layoutState = runBasicBackfill(layoutState, settings, personality, effortConfig, "baselineBackfill");
    layoutState.debug = debugRoot;
    decorateLayoutState(layoutState, settings, sourcePlanData, personality);

    debugRoot.stageRecords.push(buildStageRecord("baseline", layoutState));

    return {
        layoutState: layoutState,
        sourcePlanData: sourcePlanData,
        debugRoot: debugRoot,
        personality: personality,
        strategyName: personality.strategyName,
        presetName: strategy.presetName,
        personalityName: strategy.personalityName
    };
}

function solveSmartStrategy(sources, settings, strategy, effortConfig) {
    var personality = strategy.personality;
    var debugRoot = buildDebugRoot(strategy, effortConfig);
    var sourcePlanData = generateSourceBlockPlans(sources, personality, effortConfig, settings, debugRoot);
    var planSets = buildGlobalCandidatePlanSets(sourcePlanData, personality, effortConfig, settings, debugRoot);

    if (planSets.length === 0) throw new Error("No valid block plans could be built.");

    var best = null;

    for (var i = 0; i < planSets.length && i < effortConfig.primaryPlanLimit; i++) {
        var candidatePlan = planSets[i];
        var primary = runPrimaryBlockLayout(candidatePlan, personality, effortConfig, settings, sourcePlanData, debugRoot);
        debugRoot.stageRecords.push(buildStageRecord("primary#" + i, primary));

        var compacted = runCompactionPass(primary, personality, effortConfig, settings, sourcePlanData, debugRoot);
        debugRoot.stageRecords.push(buildStageRecord("compaction#" + i, compacted));

        var split = runAdaptiveBlockSplittingPass(compacted, sourcePlanData, personality, effortConfig, settings, debugRoot);
        debugRoot.stageRecords.push(buildStageRecord("split#" + i, split));

        var holeFill = runHoleFillingPass(split, sourcePlanData, personality, effortConfig, settings, debugRoot);
        debugRoot.stageRecords.push(buildStageRecord("holeFill#" + i, holeFill));

        var repaired = runLocalRepairPass(holeFill, sourcePlanData, personality, effortConfig, settings, debugRoot);
        debugRoot.stageRecords.push(buildStageRecord("repair#" + i, repaired));

        if (!best || compareScoreObjects(repaired.score, best.layoutState.score) < 0) {
            best = {
                layoutState: repaired,
                sourcePlanData: sourcePlanData,
                debugRoot: debugRoot,
                personality: personality,
                strategyName: personality.strategyName,
                presetName: strategy.presetName,
                personalityName: strategy.personalityName
            };
        }
    }

    return best;
}

function summarizeCandidateComparison(candidate, baseline) {
    if (!candidate || !baseline) return {};
    return {
        candidatePlaced: countPlacedCopies(candidate.layoutState),
        candidateUnplaced: countUnplacedCopies(candidate.layoutState),
        baselinePlaced: countPlacedCopies(baseline.layoutState),
        baselineUnplaced: countUnplacedCopies(baseline.layoutState),
        candidateLengthIn: parseFloat(inToStr(candidate.layoutState.score.usedNoTrailingGap)),
        baselineLengthIn: parseFloat(inToStr(baseline.layoutState.score.usedNoTrailingGap))
    };
}

function rejectAgainstBaseline(candidate, baseline, settings) {
    if (!candidate || !baseline) return null;
    var candidatePlaced = countPlacedCopies(candidate.layoutState);
    var baselinePlaced = countPlacedCopies(baseline.layoutState);

    if (candidatePlaced < baselinePlaced) return "Baseline placed more copies.";

    if (candidatePlaced === baselinePlaced) {
        var lengthGap = candidate.layoutState.score.usedNoTrailingGap - baseline.layoutState.score.usedNoTrailingGap;
        if (lengthGap > inToPt(0.35)) return "Baseline used materially shorter length.";
    }

    if (candidate.layoutState.validation && candidate.layoutState.validation.hardReject) return "Validation failed.";
    return null;
}

function finalizeWinner(candidates, baseline, settings, effortConfig) {
    var accepted = [];
    var rejected = [];
    var best = baseline;

    for (var i = 0; i < candidates.length; i++) {
        var candidate = candidates[i];
        var reason = rejectAgainstBaseline(candidate, baseline, settings);
        if (reason) {
            candidate.debugRoot.rejectedReasons.push(reason);
            rejected.push({ strategyName: candidate.strategyName, reason: reason });
            continue;
        }
        accepted.push(candidate);
        if (!best || compareScoreObjects(candidate.layoutState.score, best.layoutState.score) < 0) best = candidate;
    }

    if (!best) best = baseline;

    var bestDebug = best.debugRoot;
    bestDebug.candidateComparison = [];
    for (i = 0; i < candidates.length; i++) {
        bestDebug.candidateComparison.push({
            strategyName: candidates[i].strategyName,
            comparison: summarizeCandidateComparison(candidates[i], baseline),
            rejectedReasons: candidates[i].debugRoot.rejectedReasons
        });
    }
    bestDebug.rejectedCandidates = rejected;

    return best;
}

function buildDebugSummaryText(result) {
    var layoutState = result.layoutState;
    var score = layoutState.score;
    var sanity = layoutState.metrics.sanity;
    var splitStats = result.debugRoot.splitStats || { attempted: 0, accepted: 0 };
    var holeStats = result.debugRoot.holeFillStats || { attempted: 0, accepted: 0, filledCopies: 0 };

    var lines = [];
    lines.push("Strategy: " + result.strategyName);
    lines.push("Preset: " + result.presetName);
    lines.push("Personality: " + result.personalityName);
    lines.push("Placed copies: " + countPlacedCopies(layoutState));
    lines.push("Unplaced copies: " + countUnplacedCopies(layoutState));
    lines.push("Used length: " + inToStr(score.usedNoTrailingGap) + " in");
    lines.push("Utilization: " + (score.utilization * 100).toFixed(1) + "%");
    lines.push("Central cavity ratio: " + (sanity.largestCentralCavityRatio * 100).toFixed(1) + "%");
    lines.push("Source fragmentation: " + sanity.sourceFragmentationScore.toFixed(2));
    lines.push("Split accepted: " + splitStats.accepted + "/" + splitStats.attempted);
    lines.push("Hole fill accepted: " + holeStats.accepted + "/" + holeStats.attempted + " (" + holeStats.filledCopies + " copies)");
    return lines.join("\n");
}

function solveFromSources(sources, settings) {
    resetBlockUidCounter();

    var effortConfig = getEffortConfig(settings.searchEffort);
    var strategies = buildStrategyCandidates(settings);
    var baselineResult = null;
    var smartCandidates = [];

    for (var i = 0; i < strategies.length; i++) {
        var strategy = strategies[i];
        if (strategy.isBaseline) baselineResult = solveBaselineLayout(sources, settings, strategy, effortConfig);
        else smartCandidates.push(solveSmartStrategy(sources, settings, strategy, effortConfig));
    }

    if (!baselineResult) throw new Error("Baseline solver did not run.");

    var winner = finalizeWinner(smartCandidates, baselineResult, settings, effortConfig);
    winner.debugRoot.candidateCount = strategies.length;
    winner.debugRoot.winningStrategyName = winner.strategyName;
    winner.debugRoot.winningPreset = winner.personalityName;
    winner.debugRoot.winningSearchEffort = effortConfig.name;
    winner.debugRoot.winningScoreBreakdown = winner.layoutState.score;
    winner.debugRoot.sanityMetrics = winner.layoutState.metrics.sanity;
    winner.debugRoot.holeFillStats = winner.debugRoot.holeFillStats || { attempted: 0, accepted: 0, filledCopies: 0 };
    winner.debugRoot.splitStats = winner.debugRoot.splitStats || { attempted: 0, accepted: 0, blocksSplit: 0, copiesRecovered: 0 };
    winner.debugRoot.debugSummaryText = buildDebugSummaryText(winner);

    return {
        sources: sources,
        layout: winner.layoutState,
        usedSettings: settings,
        debugMeta: winner.debugRoot,
        winningStrategyName: winner.strategyName,
        winningPreset: winner.personalityName,
        winningSearchEffort: effortConfig.name,
        winningScoreBreakdown: winner.layoutState.score,
        sanityMetrics: winner.layoutState.metrics.sanity,
        holeFillStats: winner.debugRoot.holeFillStats,
        splitStats: winner.debugRoot.splitStats,
        candidateCount: strategies.length
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

function setItemBottomRightByVisibleBounds(item, targetRight, targetBottom) {
    var vb = item.visibleBounds;
    var dx = targetRight - vb[2];
    var dy = targetBottom - vb[3];
    item.translate(dx, dy);
}

function setItemTopRightByVisibleBounds(item, targetRight, targetTop) {
    var vb = item.visibleBounds;
    var dx = targetRight - vb[2];
    var dy = targetTop - vb[1];
    item.translate(dx, dy);
}

function getLayerVisibleBounds(layer, ignoreMetaItems) {
    if (!layer || layer.pageItems.length === 0) return null;

    var left = null;
    var top = null;
    var right = null;
    var bottom = null;

    for (var i = 0; i < layer.pageItems.length; i++) {
        var item = layer.pageItems[i];
        if (ignoreMetaItems && isOutputMetaItem(item)) continue;

        var vb;
        try { vb = item.visibleBounds; } catch (e1) { vb = null; }
        if (!vb) continue;

        if (left === null || vb[0] < left) left = vb[0];
        if (top === null || vb[1] > top) top = vb[1];
        if (right === null || vb[2] > right) right = vb[2];
        if (bottom === null || vb[3] < bottom) bottom = vb[3];
    }

    if (left === null || top === null || right === null || bottom === null) return null;

    return {
        left: left,
        top: top,
        right: right,
        bottom: bottom,
        width: Math.abs(right - left),
        height: Math.abs(top - bottom)
    };
}

function tryAssignThinTextFont(charAttrs) {
    if (!charAttrs || !app.textFonts) return false;

    var preferredFonts = [
        "Montserrat-Light",
        "Montserrat Light",
        "MyriadPro-Light",
        "HelveticaNeue-Light",
        "Aptos-Light",
        "ArialMT"
    ];

    for (var i = 0; i < preferredFonts.length; i++) {
        try {
            charAttrs.textFont = app.textFonts.getByName(preferredFonts[i]);
            return true;
        } catch (e1) {}
    }

    return false;
}

function forceFolderLabelDigitsLtr(tf) {
    if (!tf || !tf.textRange) return;

    try {
        if (typeof LanguageType !== "undefined" && LanguageType.ENGLISHUSA !== undefined) {
            tf.textRange.characterAttributes.language = LanguageType.ENGLISHUSA;
        }
    } catch (e1) {}

    try { tf.textRange.characterAttributes.digitsType = DIGITS_TYPE_ARABIC; } catch (e2) {}
    try { tf.textRange.characterAttributes.characterDirection = CHARACTER_DIRECTION_LTR; } catch (e3) {}
    try { tf.textRange.characterAttributes.keyboardDirection = CHARACTER_DIRECTION_LTR; } catch (e4) {}
}

function addSourceFolderLabelText(doc, outputLayer, folderLabel) {
    if (!doc || !outputLayer) return null;

    var textValue = normalizeAsciiDigits(trimStr(folderLabel || ""));
    if (!textValue) return null;

    var bounds = getLayerVisibleBounds(outputLayer, true);
    if (!bounds) return null;

    var marginX = 4;
    var marginBottom = inToPt(0.25);
    var minFontSize = 11;
    var fontSize = 11;
    var availableWidth = Math.max(bounds.width - (marginX * 2), 24);

    var tf = null;
    try { tf = doc.textFrames.add(); } catch (e1) { tf = null; }
    if (!tf) return null;

    tf.contents = textValue;
    try { tf.move(outputLayer, ElementPlacement.PLACEATEND); } catch (e2) {}
    try { tf.name = "NESTER_FOLDER_LABEL"; } catch (e3) {}
    try { tf.note = getOutputMetaNote(OUTPUT_FOLDER_LABEL_META); } catch (e4) {}

    try {
        var attrs = tf.textRange.characterAttributes;
        attrs.size = fontSize;
        attrs.stroked = false;
        attrs.filled = true;
        tryAssignThinTextFont(attrs);

        var fill = new GrayColor();
        fill.gray = 100;
        attrs.fillColor = fill;
    } catch (e5) {}

    forceFolderLabelDigitsLtr(tf);

    for (var pass = 0; pass < 4; pass++) {
        var vb = null;
        try { vb = tf.visibleBounds; } catch (e6) { vb = null; }
        if (!vb) break;

        var width = Math.abs(vb[2] - vb[0]);
        if (width <= availableWidth + EPS || fontSize <= minFontSize + EPS) break;

        fontSize = Math.max(minFontSize, fontSize * (availableWidth / Math.max(width, EPS)));
        try { tf.textRange.characterAttributes.size = fontSize; } catch (e7) { break; }
    }

    try {
        // Put the top edge of the full text box at least 0.25in below the lowest artwork edge.
        setItemTopRightByVisibleBounds(tf, bounds.right - marginX, bounds.bottom - marginBottom);
    } catch (e8) {}

    try {
        var outlined = tf.createOutline();
        if (outlined) {
            try { outlined.name = "NESTER_FOLDER_LABEL"; } catch (e9) {}
            try { outlined.note = getOutputMetaNote(OUTPUT_FOLDER_LABEL_META); } catch (e10) {}
            return outlined;
        }
    } catch (e11) {}

    return tf;
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

function renderPlacedBlocks(layoutState, outputLayer, settings) {
    var spacing = inToPt(settings.spacingIn);

    for (var i = 0; i < layoutState.placed.length; i++) {
        var pb = layoutState.placed[i];
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

function renderSolution(doc, result, settings) {
    var outputLayer = prepareOutputLayer(doc);
    renderPlacedBlocks(result.layout, outputLayer, settings);
    addSourceFolderLabelText(doc, outputLayer, getSourceFolderLabel(result.sources));
    if (settings.hideSourceLayersAfterBuild) hideSourceLayers(doc);
}

// =====================================================
// Summary + bridge payload
// =====================================================

function buildSummary(result, settings) {
    var lines = [];
    var score = result.layout.score;
    var sanity = result.sanityMetrics;
    var splitStats = result.splitStats || { accepted: 0, attempted: 0 };
    var holeStats = result.holeFillStats || { accepted: 0, attempted: 0, filledCopies: 0 };

    lines.push("WeMust NESTER v7");
    lines.push("Requested preset: " + settings.optimizePreset);
    lines.push("Winning strategy: " + result.winningStrategyName);
    lines.push("Winning personality: " + result.winningPreset);
    lines.push("Solver effort: " + result.winningSearchEffort);
    lines.push("Candidate count: " + result.candidateCount);
    lines.push("Placed copies: " + countPlacedCopies(result.layout));
    lines.push("Unplaced copies: " + countUnplacedCopies(result.layout));
    lines.push("Used length: " + inToStr(score.usedNoTrailingGap) + " in");
    lines.push("Width utilization: " + (score.utilization * 100).toFixed(1) + "%");
    lines.push("Central cavity ratio: " + (sanity.largestCentralCavityRatio * 100).toFixed(1) + "%");
    lines.push("Split recovery: " + splitStats.accepted + "/" + splitStats.attempted);
    lines.push("Hole fill recovery: " + holeStats.accepted + "/" + holeStats.attempted + " (" + holeStats.filledCopies + " copies)");
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
            if (byKey[placed.block.sourceKey]) byKey[placed.block.sourceKey].placedQty += placed.block.count;
        }
    }

    if (result.layout && result.layout.unplaced) {
        for (i = 0; i < result.layout.unplaced.length; i++) {
            var unplaced = result.layout.unplaced[i];
            if (byKey[unplaced.sourceKey]) byKey[unplaced.sourceKey].unplacedQty += unplaced.count;
        }
    }

    return out;
}

function getSourceFolderLabel(sources) {
    if (!sources || !sources.length) return "";

    var folderNames = [];
    var seen = {};

    for (var i = 0; i < sources.length; i++) {
        var candidate = getPreferredFolderLabel(sources[i].filePath || "");
        if (!candidate) return "MixedFolders";
        if (!isAcceptedSourceFolderLabel(candidate)) return "MixedFolders";
        if (!seen[candidate]) {
            seen[candidate] = true;
            folderNames.push(candidate);
        }
    }

    if (!folderNames.length) return "";
    if (folderNames.length === 1) return folderNames[0];
    return folderNames.join("_AND_");
}

function buildOutputBoundsText(doc) {
    if (!doc) return "";

    var layer = findLayerByName(doc, OUTPUT_LAYER_NAME);
    if (!layer || layer.pageItems.length === 0) return "";

    var bounds = getLayerVisibleBounds(layer, false);
    if (!bounds) return "";

    return inToStr(bounds.width) + "x" + inToStr(bounds.height) + "in";
}

function buildCurrentOutputContext(doc) {
    var sources = [];
    try { sources = collectPlacedItems(doc, null); } catch (e1) { sources = []; }

    return {
        outputBoundsText: buildOutputBoundsText(doc),
        sourceFolderLabel: normalizeAsciiDigits(getSourceFolderLabel(sources))
    };
}

function getSourceFolderPathsFromSources(sources) {
    var out = [];
    var seen = {};
    if (!sources || !sources.length) return out;

    for (var i = 0; i < sources.length; i++) {
        var folderPath = getParentFolderPath(sources[i].filePath || "");
        if (!folderPath) continue;
        var seenKey = String(folderPath).toLowerCase();
        if (seen[seenKey]) continue;
        seen[seenKey] = true;
        out.push(folderPath);
    }

    return out;
}

function getSourceFolderPathsFromDocument(doc) {
    var out = [];
    var seen = {};
    if (!doc || !doc.placedItems || doc.placedItems.length === 0) return out;

    for (var i = 0; i < doc.placedItems.length; i++) {
        var item = doc.placedItems[i];

        try {
            if (item.layer && isOutputLayerName(item.layer.name)) continue;
        } catch (e1) {}

        var folderPath = getParentFolderPath(getItemFilePath(item));
        if (!folderPath) continue;

        var seenKey = String(folderPath).toLowerCase();
        if (seen[seenKey]) continue;
        seen[seenKey] = true;
        out.push(folderPath);
    }

    return out;
}

function createPngImageCaptureOptions300Dpi() {
    var options = new ImageCaptureOptions();
    options.antiAliasing = true;
    options.transparency = true;
    options.matte = false;
    options.resolution = 300;
    return options;
}

function withIsolatedOutputLayer(doc, callback) {
    if (!doc || !callback) return null;

    var snapshots = [];
    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        snapshots.push({ layer: layer, visible: layer.visible });
    }

    try {
        for (var j = 0; j < snapshots.length; j++) {
            try {
                snapshots[j].layer.visible = isOutputLayerName(snapshots[j].layer.name);
            } catch (e1) {}
        }
        return callback();
    } finally {
        for (var k = 0; k < snapshots.length; k++) {
            try { snapshots[k].layer.visible = snapshots[k].visible; } catch (e2) {}
        }
    }
}

function exportOutputLayerPngToFolders(doc, exportBaseName, folderPaths) {
    var outputLayer = findLayerByName(doc, OUTPUT_LAYER_NAME);
    if (!outputLayer || outputLayer.pageItems.length === 0) {
        throw new Error("No NEST_BUILD output found to export.");
    }

    var captureBounds = getLayerVisibleBounds(outputLayer, false);
    if (!captureBounds) {
        throw new Error("Could not resolve output bounds for export.");
    }

    var exportedPaths = [];
    var failedPaths = [];
    var overwrittenCount = 0;
    var options = createPngImageCaptureOptions300Dpi();
    var clipBounds = [captureBounds.left, captureBounds.top, captureBounds.right, captureBounds.bottom];
    var exportTargets = [];

    for (var i = 0; i < folderPaths.length; i++) {
        var folderPath = folderPaths[i];
        var folder = new Folder(folderPath);
        if (!folder.exists) {
            failedPaths.push(folderPath + " | Folder not found");
            continue;
        }
        exportTargets.push({
            folder: folder,
            file: new File(folder.fsName + "/" + exportBaseName + ".png")
        });
    }

    if (!exportTargets.length) {
        return {
            exportedPaths: exportedPaths,
            failedPaths: failedPaths,
            overwrittenCount: overwrittenCount
        };
    }

    withIsolatedOutputLayer(doc, function () {
        var primaryTarget = exportTargets[0];

        try {
            if (primaryTarget.file.exists) {
                try { primaryTarget.file.remove(); } catch (e1) {}
                overwrittenCount += 1;
            }
            doc.imageCapture(primaryTarget.file, clipBounds, options);
            exportedPaths.push(primaryTarget.file.fsName);
        } catch (e2) {
            failedPaths.push(primaryTarget.folder.fsName + " | " + String(e2));
            return;
        }

        for (var j = 1; j < exportTargets.length; j++) {
            var target = exportTargets[j];
            try {
                if (target.file.exists) {
                    try { target.file.remove(); } catch (e3) {}
                    overwrittenCount += 1;
                }

                if (!primaryTarget.file.copy(target.file.fsName)) {
                    throw new Error("Could not copy exported PNG.");
                }
                exportedPaths.push(target.file.fsName);
            } catch (e4) {
                failedPaths.push(target.folder.fsName + " | " + String(e4));
            }
        }
    });

    return {
        exportedPaths: exportedPaths,
        failedPaths: failedPaths,
        overwrittenCount: overwrittenCount
    };
}

function buildSearchMeta(result) {
    return {
        candidateCount: result.candidateCount,
        winningStrategyName: result.winningStrategyName,
        winningPreset: result.winningPreset,
        winningSearchEffort: result.winningSearchEffort,
        winningScoreBreakdown: result.winningScoreBreakdown,
        sanityMetrics: result.sanityMetrics,
        holeFillStats: result.holeFillStats,
        splitStats: result.splitStats,
        debugSummaryText: result.debugMeta.debugSummaryText,
        selectedPresetPersonality: result.debugMeta.selectedPresetPersonality,
        selectedSearchEffort: result.debugMeta.selectedSearchEffort,
        sourcePlanSummary: result.debugMeta.sourcePlanSummary,
        globalCandidatePlanSets: result.debugMeta.globalCandidatePlanSets,
        blockOrderVariantsConsidered: result.debugMeta.blockOrderVariantsConsidered,
        stageRecords: result.debugMeta.stageRecords,
        candidateComparison: result.debugMeta.candidateComparison,
        rejectedCandidates: result.debugMeta.rejectedCandidates,
        failedSources: result.debugMeta.failedSources
    };
}

function buildOnce(doc, settings) {
    var sources = collectAndNormalizeSources(doc, settings);
    var result = solveFromSources(sources, settings);
    renderSolution(doc, result, settings);
    return result;
}

// =====================================================
// Selection / navigation
// =====================================================

function focusViewOnItem(doc, item) {
    if (!doc || !item) return;

    var view = null;
    try { view = doc.activeView; } catch (e1) { view = null; }
    if (!view) return;

    var vb = null;
    try { vb = item.visibleBounds; } catch (e2) { vb = null; }
    if (!vb || vb.length < 4) return;

    var left = Math.min(vb[0], vb[2]);
    var right = Math.max(vb[0], vb[2]);
    var top = Math.max(vb[1], vb[3]);
    var bottom = Math.min(vb[1], vb[3]);
    var width = Math.max(right - left, EPS);
    var height = Math.max(top - bottom, EPS);

    try {
        view.centerPoint = [(left + right) / 2, (top + bottom) / 2];
    } catch (e3) {}

    try {
        var viewBounds = view.bounds;
        if (viewBounds && viewBounds.length >= 4 && view.zoom) {
            var viewWidth = Math.abs(viewBounds[2] - viewBounds[0]);
            var viewHeight = Math.abs(viewBounds[1] - viewBounds[3]);
            if (viewWidth > EPS && viewHeight > EPS) {
                var fitScale = Math.min(viewWidth / width, viewHeight / height);
                // if (isFinite(fitScale) && fitScale > 0) view.zoom = Math.max(0.05, view.zoom * fitScale * 0.82);
            }
        }
    } catch (e4) {}
}

function findPlacedCopyBySourceKey(doc, sourceKey) {
    var layer = findLayerByName(doc, OUTPUT_LAYER_NAME);
    if (!layer || layer.pageItems.length === 0) return null;

    var needle = "NESTER_SOURCE_KEY=" + String(sourceKey || "");
    for (var i = 0; i < layer.pageItems.length; i++) {
        var item = layer.pageItems[i];
        try {
            if (item.note === needle) return item;
        } catch (e1) {}
    }
    return null;
}

function nesterSelectPlacedCopyBySourceKey(sourceKey) {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        if (app.documents.length === 0) return _jsonStringify({ ok: false, error: "No open document found." });

        var found = findPlacedCopyBySourceKey(app.activeDocument, sourceKey);
        if (!found) return _jsonStringify({ ok: false, error: "No matching placed copy found." });

        app.activeDocument.selection = null;
        app.activeDocument.selection = [found];
        app.redraw();
        return _jsonStringify({ ok: true });
    } catch (e) {
        return _jsonStringify({ ok: false, error: String(e) });
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}

function nesterGoToPlacedCopyBySourceKey(sourceKey) {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        if (app.documents.length === 0) return _jsonStringify({ ok: false, error: "No open document found." });

        var found = findPlacedCopyBySourceKey(app.activeDocument, sourceKey);
        if (!found) return _jsonStringify({ ok: false, error: "No matching placed copy found." });

        app.activeDocument.selection = null;
        app.activeDocument.selection = [found];
        focusViewOnItem(app.activeDocument, found);
        app.redraw();
        return _jsonStringify({ ok: true });
    } catch (e) {
        return _jsonStringify({ ok: false, error: String(e) });
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}

// =====================================================
// Panel settings normalization
// =====================================================

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
    var s = cloneSettings(DEFAULTS);
    if (!input || typeof input !== "object") return s;

    s.sheetWidthIn = Math.max(0.01, _numOr(input.sheetWidthIn, s.sheetWidthIn));
    s.maxLengthIn = Math.max(0.01, _numOr(input.maxLengthIn, s.maxLengthIn));
    s.spacingIn = Math.max(0, _numOr(input.spacingIn, s.spacingIn));
    s.optimizePreset = _oneOfOr(input.optimizePreset, ["Auto", "Compact", "Balanced", "Ordered"], s.optimizePreset);
    s.searchEffort = _oneOfOr(input.searchEffort, ["Normal", "High"], s.searchEffort);
    s.allowItemRotationInBlock = _boolOr(input.allowItemRotationInBlock, s.allowItemRotationInBlock);
    s.allowBlockRotationOnSheet = _boolOr(input.allowBlockRotationOnSheet, s.allowBlockRotationOnSheet);
    s.hideSourceLayersAfterBuild = DEFAULTS.hideSourceLayersAfterBuild;
    s.quantityOverrides = normalizeQuantityOverrides(input.quantityOverrides);
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
        hideSourceLayersAfterBuild: DEFAULTS.hideSourceLayersAfterBuild
    });
}

function nesterGetOutputContext() {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        if (app.documents.length === 0) {
            return _jsonStringify({ ok: false, error: "No open document found." });
        }

        var context = buildCurrentOutputContext(app.activeDocument);
        return _jsonStringify({
            ok: true,
            outputBoundsText: context.outputBoundsText,
            sourceFolderLabel: context.sourceFolderLabel
        });
    } catch (e) {
        return _jsonStringify({ ok: false, error: String(e) });
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}

function nesterExportOutputPngToSourceFolders(payloadJson) {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        if (app.documents.length === 0) {
            return _jsonStringify({ ok: false, error: "No open document found." });
        }

        var payload = null;
        if (payloadJson !== undefined && payloadJson !== null && String(payloadJson) !== "") {
            payload = _jsonParse(payloadJson);
        }

        var doc = app.activeDocument;
        var folderPaths = sanitizeFolderPaths(LAST_SOURCE_FOLDER_PATHS);
        var seenFolderPaths = {};

        for (var s = 0; s < folderPaths.length; s++) {
            seenFolderPaths[String(folderPaths[s]).toLowerCase()] = true;
        }

        if (payload && payload.folderPaths && payload.folderPaths.length) {
            for (var i = 0; i < payload.folderPaths.length; i++) {
                var folderPath = trimStr(payload.folderPaths[i] || "");
                if (!folderPath) continue;

                var seenKey = String(folderPath).toLowerCase();
                if (seenFolderPaths[seenKey]) continue;
                seenFolderPaths[seenKey] = true;
                folderPaths.push(folderPath);
            }
        }

        if (!folderPaths.length) folderPaths = getSourceFolderPathsFromDocument(doc);
        if (!folderPaths.length) {
            return _jsonStringify({ ok: false, error: "No source folders found from placed files." });
        }

        var exportBaseName = sanitizeExportBaseName(payload && payload.fileName ? payload.fileName : "");
        var exportResult = exportOutputLayerPngToFolders(doc, exportBaseName, folderPaths);

        return _jsonStringify({
            ok: true,
            fileName: exportBaseName,
            exportedPaths: exportResult.exportedPaths,
            failedPaths: exportResult.failedPaths,
            overwrittenCount: exportResult.overwrittenCount
        });
    } catch (e) {
        return _jsonStringify({ ok: false, error: String(e) });
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}

// =====================================================
// Development helpers
// =====================================================

function _testSplitFallbacks() {
    var partitions = buildFallbackPartitions(12, 90);
    return {
        count: partitions.length,
        first: partitions.length ? partitions[0].join("+") : ""
    };
}

function _testFreeRectMerging() {
    var input = [
        { x: 0, y: 0, w: 10, h: 10 },
        { x: 10, y: 0, w: 10, h: 10 }
    ];
    var merged = mergeAdjacentFreeRects(input);
    return {
        inputCount: input.length,
        mergedCount: merged.length,
        mergedWidth: merged.length ? merged[0].w : 0
    };
}

function _testScoreOrdering() {
    var a = { total: 100 };
    var b = { total: 200 };
    return {
        compareAB: compareScoreObjects(a, b),
        compareBA: compareScoreObjects(b, a),
        compareAA: compareScoreObjects(a, a)
    };
}

function nesterRunSolverTestHarness() {
    return _jsonStringify({
        ok: true,
        tests: {
            splitFallbacks: _testSplitFallbacks(),
            freeRectMerging: _testFreeRectMerging(),
            scoreOrdering: _testScoreOrdering()
        }
    });
}

// =====================================================
// ScriptUI entry point kept for manual host execution
// =====================================================

function nesterRunMain() {
    var previousLevel = app.userInteractionLevel;
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        if (app.documents.length === 0) return "ERROR: No open document found.";

        var settings = cloneSettings(DEFAULTS);
        var result = buildOnce(app.activeDocument, settings);
        return buildSummary(result, settings);
    } catch (e) {
        return "ERROR: " + e;
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}

// =====================================================
// CEP bridge entry point
// =====================================================

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
        setLastSourceFolderPathsFromSources(result.sources);
        var score = result.layout.score;
        var summary = buildSummary(result, settings);
        var searchMeta = buildSearchMeta(result);

        return _jsonStringify({
            ok: true,
            summary: summary,
            usedLengthIn: parseFloat(inToStr(score.usedNoTrailingGap)),
            placedCopies: countPlacedCopies(result.layout),
            unplacedCopies: countUnplacedCopies(result.layout),
            searchMeta: searchMeta,
            sourceItems: buildSourceQuantitySummary(result),
            sourceFolderPaths: getSourceFolderPathsFromSources(result.sources),
            outputBoundsText: buildOutputBoundsText(app.activeDocument),
            sourceFolderLabel: normalizeAsciiDigits(getSourceFolderLabel(result.sources)),
            outputSizeText: buildOutputBoundsText(app.activeDocument),
            candidateCount: result.candidateCount,
            winningStrategyName: result.winningStrategyName,
            winningPreset: result.winningPreset,
            winningSearchEffort: result.winningSearchEffort,
            winningScoreBreakdown: result.winningScoreBreakdown,
            sanityMetrics: result.sanityMetrics,
            holeFillStats: result.holeFillStats,
            splitStats: result.splitStats,
            debugSummaryText: result.debugMeta.debugSummaryText
        });
    } catch (e) {
        return _jsonStringify({ ok: false, error: String(e) });
    } finally {
        try { app.userInteractionLevel = previousLevel; } catch (_e) {}
    }
}
