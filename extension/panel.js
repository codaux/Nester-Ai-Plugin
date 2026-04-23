(function () {
  "use strict";

  var today = new Date().getDate();
  if (!isFinite(today) || today < 1 || today > 31) today = 1;

  var STORAGE_KEY = "nester.settings.v5";
  var STORAGE_SCHEMA_VERSION = 2;
  var formEl = document.getElementById("settings-form");
  var runBtn = document.getElementById("btn-run");
  var exportBtn = document.getElementById("btn-export");
  var runActionEl = runBtn ? runBtn.parentNode : null;
  var exportActionEl = exportBtn ? exportBtn.parentNode : null;
  var exportFolderBtn = document.getElementById("btn-export-folder");
  var regenerateNameBtn = document.getElementById("btn-regenerate-name");
  var resetBtn = document.getElementById("btn-reset");
  var inventoryListEl = document.getElementById("inventory-list");
  var inventoryEmptyEl = document.getElementById("inventory-empty");
  var inventoryMetaEl = document.getElementById("inventory-meta");
  var orderEmailBtn = document.getElementById("btn-order-email");
  var orderEmailWrapEl = document.getElementById("order-email-wrap");
  var orderEmailBackdropEl = document.getElementById("order-email-backdrop");
  var orderEmailCloseBtn = document.getElementById("btn-order-email-close");
  var orderEmailTextEl = document.getElementById("order-email-text");
  var resultNameField = document.getElementById("result-name-field");
  var panelFeedbackEl = document.getElementById("panel-feedback");
  var solverModeNoteEl = document.getElementById("solverModeNote");

  var RESULT_NAME_SIZE_TOKEN = "{SIZE}";
  var DEFAULTS = {
    sheetWidthIn: 23,
    maxLengthIn: 100,
    spacingIn: 0.25,
    optimizePreset: "Auto",
    searchEffort: "High",
    solverWidthFillBias: 78,
    solverOrderBias: 58,
    solverHoleRepairBias: 72,
    solverSearchDepth: 68,
    allowItemRotationInBlock: true,
    allowBlockRotationOnSheet: true,
    legacySolverEnabled: false,
    outputDate: today,
    tagUSA: false,
    tagRUSH: false,
    tagCUT: false,
    tagOverNight: false,
    tagPurolator: false,
    tagGLS: false,
    tagINTERNAL: false,
    tagREPRINT: false,
    outputPcCount: 1,
    outputChokeCount: 6,
    outputPartCount: 0
  };

  var FIELD_IDS = [
    "sheetWidthIn",
    "maxLengthIn",
    "spacingIn",
    "solverWidthFillBias",
    "solverOrderBias",
    "solverHoleRepairBias",
    "solverSearchDepth",
    "optimizePreset",
    "searchEffort",
    "allowItemRotationInBlock",
    "allowBlockRotationOnSheet",
    "legacySolverEnabled",
    "outputDate",
    "tagUSA",
    "tagRUSH",
    "tagCUT",
    "tagOverNight",
    "tagPurolator",
    "tagGLS",
    "tagINTERNAL",
    "tagREPRINT",
    "outputPcCount",
    "outputChokeCount",
    "outputPartCount"
  ];
  var STEPPER_FIELD_LIMITS = {
    outputDate: { min: 1, max: 31 },
    outputPcCount: { min: 1, max: 9 },
    outputChokeCount: { min: 1, max: 8 },
    outputPartCount: { min: 0, max: 9 }
  };
  var TUNING_SLIDER_IDS = [
    "solverWidthFillBias",
    "solverOrderBias",
    "solverHoleRepairBias",
    "solverSearchDepth"
  ];
  var AUTO_NEST_FIELD_IDS = {
    sheetWidthIn: true,
    maxLengthIn: true,
    spacingIn: true,
    solverWidthFillBias: true,
    solverOrderBias: true,
    solverHoleRepairBias: true,
    solverSearchDepth: true,
    allowItemRotationInBlock: true,
    allowBlockRotationOnSheet: true,
    legacySolverEnabled: true
  };
  var NAMING_RESETTABLE_IDS = {
    outputPcCount: true,
    outputChokeCount: true,
    outputPartCount: true
  };

  var state = {
    quantityOverrides: {},
    lastSourceItems: [],
    lastSourceFolderPaths: [],
    extraExportFolderPath: "",
    resultSizeText: "",
    outputBoundsText: "",
    sourceFolderLabel: "",
    resultNameManualOverride: false,
    resultNameTemplateText: "",
    orderEmailText: ""
  };
  var resultNameFeedbackTimer = null;
  var resultNameClickTimer = null;
  var panelFeedbackTimer = null;
  var panelFeedbackHideTimer = null;
  var autoNestTimer = null;
  var autoNestQueued = false;
  var busyMode = null;
  var AUTO_NEST_DELAY_MS = 500;
  var suspendInventoryInputSync = false;

  function byId(id) {
    return document.getElementById(id);
  }

  function createBubbledEvent(name) {
    if (typeof Event === "function") return new Event(name, { bubbles: true });
    var evt = document.createEvent("Event");
    evt.initEvent(name, true, false);
    return evt;
  }

  function hasCepBridge() {
    return typeof window.__adobe_cep__ !== "undefined";
  }

  function isPreviewMode() {
    return !hasCepBridge();
  }

  function notifyPreviewMode(message) {
    setPanelFeedback(message || "Preview mode only. CEP bridge is unavailable here.", "warning", 2800);
  }
  function escapeForJsxString(s) {
    return String(s)
      .replace(/\\/g, "\\\\")
      .replace(/"/g, '\\"')
      .replace(/\r/g, "\\r")
      .replace(/\n/g, "\\n");
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function normalizeMultilineText(value) {
    return String(value || "").replace(/\r\n?/g, "\n");
  }

  function toFileUrl(filePath) {
    if (!filePath) return "";
    var normalized = String(filePath).replace(/\\/g, "/");
    if (/^file:/i.test(normalized)) return encodeURI(normalized);
    if (normalized.charAt(0) !== "/") normalized = "/" + normalized;
    return encodeURI("file://" + normalized);
  }

  function parseJsonSafe(s) {
    try {
      return JSON.parse(s);
    } catch (_e) {
      return null;
    }
  }

  function merge(target, source) {
    var out = {};
    var key;
    for (key in target) {
      if (Object.prototype.hasOwnProperty.call(target, key)) out[key] = target[key];
    }
    if (!source || typeof source !== "object") return out;
    for (key in source) {
      if (Object.prototype.hasOwnProperty.call(source, key)) out[key] = source[key];
    }
    return out;
  }

  function migrateStoredSettings(stored) {
    if (!stored || typeof stored !== "object") return null;

    var next = merge({}, stored);
    var schemaVersion = Number(next.storageSchemaVersion) || 0;

    if (schemaVersion < 2) {
      next.allowItemRotationInBlock = true;
      next.allowBlockRotationOnSheet = true;
    }

    next.storageSchemaVersion = STORAGE_SCHEMA_VERSION;
    return next;
  }

  function sanitizeQty(value, fallback) {
    var parsed = Number(value);
    if (!isFinite(parsed) || parsed < 1) return fallback;
    return Math.max(1, Math.round(parsed));
  }

  function sanitizeCountInRange(value, minValue, maxValue, fallback) {
    var parsed = Number(value);
    if (!isFinite(parsed)) return fallback;
    parsed = Math.round(parsed);
    if (parsed < minValue) return minValue;
    if (parsed > maxValue) return maxValue;
    return parsed;
  }

  function sanitizeSliderValue(value, fallback) {
    var parsed = Number(value);
    if (!isFinite(parsed)) return fallback;
    if (parsed < 0) return 0;
    if (parsed > 100) return 100;
    return Math.round(parsed);
  }

  function findStepButton(node) {
    while (node && node !== formEl) {
      if (node.className && String(node.className).indexOf("number-stepper-btn") !== -1) return node;
      node = node.parentNode;
    }
    return null;
  }

  function stepNamingNumberField(input, direction) {
    if (!input || !input.id) return;
    var limits = STEPPER_FIELD_LIMITS[input.id];
    var nextValue;
    if (!limits) return;

    var currentValue = sanitizeCountInRange(input.value, limits.min, limits.max, DEFAULTS[input.id]);
    if (input.id === "outputDate") {
      nextValue = currentValue + direction;
      if (nextValue > limits.max) nextValue = limits.min;
      else if (nextValue < limits.min) nextValue = limits.max;
    } else {
      nextValue = sanitizeCountInRange(currentValue + direction, limits.min, limits.max, currentValue);
    }
    if (String(input.value) === String(nextValue)) return;

    input.value = String(nextValue);
    input.dispatchEvent(createBubbledEvent("input"));
  }

  function findInventoryStepButton(node) {
    while (node && node !== inventoryListEl) {
      if (node.className && String(node.className).indexOf("inventory-stepper-btn") !== -1) return node;
      node = node.parentNode;
    }
    return null;
  }

  function stepInventoryQuantityInput(input, direction) {
    var currentValue;
    var nextValue;
    if (!input || !input.getAttribute("data-source-key")) return;

    currentValue = sanitizeQty(input.value, 1);
    nextValue = sanitizeQty(currentValue + direction, currentValue);
    if (String(input.value) === String(nextValue)) return;

    input.value = String(nextValue);
    input.dispatchEvent(createBubbledEvent("input"));
  }

  function sanitizeQuantityOverrides(raw) {
    var out = {};
    var key;
    if (!raw || typeof raw !== "object") return out;
    for (key in raw) {
      if (!Object.prototype.hasOwnProperty.call(raw, key)) continue;
      out[key] = sanitizeQty(raw[key], 1);
    }
    return out;
  }

  function sanitizeSourceItems(items) {
    var out = [];
    if (!items || !items.length) return out;

    for (var i = 0; i < items.length; i++) {
      var item = items[i] || {};
      if (!item.key || !item.name) continue;
      out.push({
        id: item.id || ("row-" + i),
        key: String(item.key),
        name: String(item.name),
        filePath: item.filePath ? String(item.filePath) : "",
        thumbnailUrl: toFileUrl(item.filePath),
        detectedQty: sanitizeQty(item.detectedQty, 1),
        requestedQty: sanitizeQty(item.requestedQty, 1),
        placedQty: Math.max(0, Number(item.placedQty) || 0),
        unplacedQty: Math.max(0, Number(item.unplacedQty) || 0),
        dimensionsText: item.dimensionsText ? String(item.dimensionsText) : "",
        dimensionsTitle: item.dimensionsTitle ? String(item.dimensionsTitle) : "",
        dimensionStatus: item.dimensionStatus ? String(item.dimensionStatus) : "none"
      });
    }

    return out;
  }

  function sanitizeFolderPaths(paths) {
    var out = [];
    var seen = {};
    if (!paths || !paths.length) return out;

    for (var i = 0; i < paths.length; i++) {
      var value = paths[i] ? String(paths[i]) : "";
      var key = value.toLowerCase();
      if (!value || seen[key]) continue;
      seen[key] = true;
      out.push(value);
    }

    return out;
  }

  function readStoredSettings() {
    try {
      var raw = window.localStorage.getItem(STORAGE_KEY);
      return raw ? migrateStoredSettings(parseJsonSafe(raw)) : null;
    } catch (_e) {
      return null;
    }
  }

  function writeStoredSettings(settings) {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    } catch (_e) {}
  }

  function buildStoredState(settings) {
    var out = merge(settings, {
      quantityOverrides: state.quantityOverrides,
      lastSourceItems: state.lastSourceItems,
      lastSourceFolderPaths: state.lastSourceFolderPaths,
      extraExportFolderPath: state.extraExportFolderPath,
      resultSizeText: state.resultSizeText,
      outputBoundsText: state.outputBoundsText,
      sourceFolderLabel: state.sourceFolderLabel,
      resultNameManualOverride: state.resultNameManualOverride,
      resultNameTemplateText: state.resultNameTemplateText
    });
    out.storageSchemaVersion = STORAGE_SCHEMA_VERSION;
    delete out.orderEmailText;
    return out;
  }

  function evalHost(script, callback) {
    if (!hasCepBridge()) return false;
    window.__adobe_cep__.evalScript(script, function (result) {
      if (callback) callback(result);
    });
    return true;
  }

  function alertUnexpectedHostResponse(rawResult) {
    window.alert("Unexpected host response.\n\nRaw response:\n" + String(rawResult || ""));
  }

  function setFormValues(settings) {
    for (var i = 0; i < FIELD_IDS.length; i++) {
      var id = FIELD_IDS[i];
      var el = byId(id);
      if (!el) continue;
      if (el.type === "checkbox") el.checked = Boolean(settings[id]);
      else if (id === "outputDate") el.value = String(sanitizeCountInRange(settings[id], 1, 31, DEFAULTS.outputDate));
      else if (id === "outputPcCount") el.value = String(sanitizeCountInRange(settings[id], 1, 9, DEFAULTS.outputPcCount));
      else if (id === "outputChokeCount") el.value = String(sanitizeCountInRange(settings[id], 1, 8, DEFAULTS.outputChokeCount));
      else if (id === "outputPartCount") el.value = String(sanitizeCountInRange(settings[id], 0, 9, DEFAULTS.outputPartCount));
      else if (el.type === "range") el.value = String(sanitizeSliderValue(settings[id], DEFAULTS[id]));
      else el.value = String(settings[id]);
    }
    updateAllSolverSliderValues();
    updateLegacySolverUi();
    updateNamingFieldVisualStates();
  }

  function syncOverridesFromInputs() {
    var inputs = inventoryListEl.querySelectorAll("[data-source-key]");
    var next = {};
    for (var i = 0; i < inputs.length; i++) {
      var input = inputs[i];
      var key = input.getAttribute("data-source-key");
      if (!key) continue;
      next[key] = sanitizeQty(input.value, 1);
    }
    state.quantityOverrides = next;
  }

  function getFormValues() {
    var settings = {};
    for (var i = 0; i < FIELD_IDS.length; i++) {
      var id = FIELD_IDS[i];
      var el = byId(id);
      if (!el) continue;

      if (el.type === "checkbox") settings[id] = Boolean(el.checked);
      else if (id === "outputDate") settings[id] = sanitizeCountInRange(el.value, 1, 31, DEFAULTS.outputDate);
      else if (id === "outputPcCount") settings[id] = sanitizeCountInRange(el.value, 1, 9, DEFAULTS.outputPcCount);
      else if (id === "outputChokeCount") settings[id] = sanitizeCountInRange(el.value, 1, 8, DEFAULTS.outputChokeCount);
      else if (id === "outputPartCount") settings[id] = sanitizeCountInRange(el.value, 0, 9, DEFAULTS.outputPartCount);
      else if (el.type === "range") settings[id] = sanitizeSliderValue(el.value, DEFAULTS[id]);
      else if (el.type === "number") settings[id] = Number(el.value);
      else settings[id] = el.value;
    }

    if (!suspendInventoryInputSync) syncOverridesFromInputs();
    settings.quantityOverrides = sanitizeQuantityOverrides(state.quantityOverrides);
    settings.orderEmailText = state.orderEmailText;
    return settings;
  }

  function isNamingFieldId(id) {
    return id === "outputDate" ||
      id === "tagUSA" ||
      id === "tagRUSH" ||
      id === "tagCUT" ||
      id === "tagOverNight" ||
      id === "tagPurolator" ||
      id === "tagGLS" ||
      id === "tagINTERNAL" ||
      id === "tagREPRINT" ||
      id === "outputPcCount" ||
      id === "outputChokeCount" ||
      id === "outputPartCount";
  }

  function isAutoNestFieldId(id) {
    return Boolean(id && AUTO_NEST_FIELD_IDS[id]);
  }

  function updateSolverSliderValue(id) {
    var input = byId(id);
    var output = byId(id + "Value");
    var fallback = Object.prototype.hasOwnProperty.call(DEFAULTS, id) ? DEFAULTS[id] : 0;
    var value;
    if (!input || !output) return;
    value = sanitizeSliderValue(input.value, fallback);
    input.value = String(value);
    output.value = String(value);
    output.textContent = String(value);
  }

  function updateAllSolverSliderValues() {
    for (var i = 0; i < TUNING_SLIDER_IDS.length; i++) updateSolverSliderValue(TUNING_SLIDER_IDS[i]);
  }

  function setLegacyLockedState(inputId, locked) {
    var input = byId(inputId);
    var wrapper = input ? input.parentNode : null;
    if (!input) return;
    while (wrapper && wrapper.tagName !== "LABEL") wrapper = wrapper.parentNode;
    input.disabled = Boolean(locked);
    if (wrapper) wrapper.classList.toggle("legacy-locked", Boolean(locked));
  }

  function updateLegacySolverUi() {
    var legacyToggle = byId("legacySolverEnabled");
    var legacyEnabled = Boolean(legacyToggle && legacyToggle.checked);

    if (solverModeNoteEl) {
      solverModeNoteEl.textContent = legacyEnabled
        ? "Legacy JS 0.2 method"
        : "Quick human bias";
    }

    for (var i = 0; i < TUNING_SLIDER_IDS.length; i++) {
      setLegacyLockedState(TUNING_SLIDER_IDS[i], legacyEnabled);
    }

    setLegacyLockedState("allowBlockRotationOnSheet", legacyEnabled);
  }

  function updateNamingFieldVisualState(id) {
    var input = byId(id);
    var wrapper;
    var currentValue;
    var defaultValue;
    if (!input || !NAMING_RESETTABLE_IDS[id]) return;
    wrapper = input.parentNode;
    if (!wrapper || String(wrapper.className).indexOf("number-field") === -1) return;
    currentValue = sanitizeCountInRange(input.value, Number(input.min || 0), Number(input.max || 100), DEFAULTS[id]);
    defaultValue = DEFAULTS[id];
    wrapper.classList.toggle("is-default", currentValue === defaultValue);
    wrapper.classList.toggle("is-dirty", currentValue !== defaultValue);
  }

  function updateNamingFieldVisualStates() {
    updateNamingFieldVisualState("outputPcCount");
    updateNamingFieldVisualState("outputChokeCount");
    updateNamingFieldVisualState("outputPartCount");
  }

  function resetNamingFieldToDefault(id) {
    var input = byId(id);
    if (!input || !Object.prototype.hasOwnProperty.call(DEFAULTS, id)) return;
    input.value = String(DEFAULTS[id]);
    input.dispatchEvent(createBubbledEvent("input"));
  }

  function buildOutputPrefixSegment(settings) {
    var prefixParts = [
      "UVM",
      String(sanitizeCountInRange(settings.outputDate, 1, 31, DEFAULTS.outputDate))
    ];

    if (settings.tagUSA) prefixParts.push("USA");
    if (settings.tagRUSH) prefixParts.push("RUSH");
    if (settings.tagOverNight) prefixParts.push("OverNight");
    if (settings.tagCUT) prefixParts.push("CUT");
    if (settings.tagPurolator) prefixParts.push("Purolator");
    if (settings.tagGLS) prefixParts.push("GLS");

    var partCount = sanitizeCountInRange(settings.outputPartCount, 0, 9, DEFAULTS.outputPartCount);
    if (partCount > 0) prefixParts.push("P" + String(partCount));

    return prefixParts.join("_");
  }

  function buildOutputSizeSegment(boundsText) {
    if (!boundsText) return "";
    return String(boundsText);
  }

  function composeOutputNameText(boundsText, folderLabel, settings) {
    if (!boundsText) return "";

    var pcCount = sanitizeCountInRange(settings.outputPcCount, 1, 9, DEFAULTS.outputPcCount);
    var chokeCount = sanitizeCountInRange(settings.outputChokeCount, 1, 8, DEFAULTS.outputChokeCount);

    var parts = [
      buildOutputPrefixSegment(settings),
      buildOutputSizeSegment(boundsText),
      String(pcCount) + (pcCount > 1 ? "pcs" : "pc")
    ];

    if (chokeCount !== 6) parts.push(String(chokeCount) + "Choke");

    if (folderLabel) parts.push(String(folderLabel));
    if (settings.tagINTERNAL) parts.push("INTERNAL");
    if (settings.tagREPRINT) parts.push("REPRINT");
    return parts.join("__");
  }

  function buildAutoResultName(settings) {
    return composeOutputNameText(state.outputBoundsText, state.sourceFolderLabel, settings);
  }

  function getPathLeaf(path) {
    var normalized = String(path || "").replace(/[\\/]+/g, "/").replace(/\/$/, "");
    if (!normalized) return "";
    var parts = normalized.split("/");
    return parts[parts.length - 1] || normalized;
  }

  function updateExtraExportFolderButton() {
    if (!exportFolderBtn) return;
    var hasExtraFolder = Boolean(state.extraExportFolderPath);
    if (exportFolderBtn.classList) exportFolderBtn.classList.toggle("is-selected", hasExtraFolder);

    if (!hasExtraFolder) {
      exportFolderBtn.title = "Choose extra export folder";
      return;
    }

    exportFolderBtn.title = "Extra export folder: " + String(state.extraExportFolderPath) + "\nClick to choose a different folder.";
  }

  function chooseExtraExportFolder() {
    if (!hasCepBridge() || !window.cep || !window.cep.fs || typeof window.cep.fs.showOpenDialogEx !== "function") {
      setPanelFeedback("Modern folder picker is not available in this host.", "warning", 3600);
      return;
    }

    var selected = null;
    try {
      selected = window.cep.fs.showOpenDialogEx(
        false,
        true,
        "Select extra export folder",
        state.extraExportFolderPath || ""
      );
    } catch (_e) {
      selected = null;
    }

    var pickedPath = "";
    if (selected && typeof selected.data === "string") pickedPath = selected.data;
    else if (selected && selected.data && selected.data.length) pickedPath = String(selected.data[0] || "");
    if (!pickedPath) return;

    state.extraExportFolderPath = pickedPath;
    updateExtraExportFolderButton();
    writeStoredSettings(buildStoredState(getFormValues()));
    setPanelFeedback("Extra export folder set to " + getPathLeaf(pickedPath) + ".", "success", 2200);
  }

  function setBusy(mode) {
    busyMode = mode || null;
    var isNestBusy = mode === "nest";
    var isExportBusy = mode === "export";
    var isBusy = isNestBusy || isExportBusy;

    if (runActionEl) runActionEl.classList.toggle("is-busy", isNestBusy);
    runBtn.disabled = isBusy;
    runBtn.textContent = isNestBusy ? "NESTING..." : "NEST";
    runBtn.setAttribute("aria-busy", isNestBusy ? "true" : "false");

    if (exportBtn) {
      if (exportActionEl) exportActionEl.classList.toggle("is-busy", isExportBusy);
      exportBtn.disabled = isBusy || !getResultNameText();
      exportBtn.textContent = isExportBusy ? "Exporting..." : "EXPORT";
      exportBtn.setAttribute("aria-busy", isExportBusy ? "true" : "false");
    }
    if (exportFolderBtn) exportFolderBtn.disabled = isBusy;
    if (orderEmailBtn) orderEmailBtn.disabled = isBusy;
    if (regenerateNameBtn) regenerateNameBtn.disabled = isBusy || !state.outputBoundsText;

    resetBtn.disabled = isBusy;
  }

  function cancelAutoNest() {
    autoNestQueued = false;
    if (!autoNestTimer) return;
    window.clearTimeout(autoNestTimer);
    autoNestTimer = null;
  }

  function setOrderEmailVisibility(isVisible) {
    if (orderEmailWrapEl) orderEmailWrapEl.hidden = !isVisible;
    if (orderEmailBtn) orderEmailBtn.classList.toggle("is-selected", Boolean(state.orderEmailText));
  }

  function clearPendingOrderState() {
    state.orderEmailText = "";
    if (orderEmailTextEl) orderEmailTextEl.value = "";
    setOrderEmailVisibility(false);
  }

  function flushAutoNest() {
    var settings;
    autoNestTimer = null;
    if (!autoNestQueued) return;
    if (busyMode) {
      autoNestTimer = window.setTimeout(flushAutoNest, AUTO_NEST_DELAY_MS);
      return;
    }

    autoNestQueued = false;
    settings = getFormValues();
    writeStoredSettings(buildStoredState(settings));
    runNester(settings);
  }

  function scheduleAutoNest() {
    autoNestQueued = true;
    if (autoNestTimer) window.clearTimeout(autoNestTimer);
    autoNestTimer = window.setTimeout(flushAutoNest, AUTO_NEST_DELAY_MS);
  }

  function syncQuantityOverridesFromResult(items) {
    var next = {};
    for (var i = 0; i < items.length; i++) {
      next[items[i].key] = sanitizeQty(items[i].requestedQty, items[i].detectedQty);
    }
    state.quantityOverrides = next;
    suspendInventoryInputSync = false;
  }

  function getInventoryMetaSummary(items) {
    var fileCount = items.length;
    var placed = 0;
    var requested = 0;

    for (var i = 0; i < items.length; i++) {
      placed += items[i].placedQty;
      requested += items[i].requestedQty;
    }

    return {
      text: fileCount + " files | " + placed + "/" + requested + " placed",
      isIncomplete: requested > 0 && placed < requested
    };
  }

  function getResultNameText() {
    return state.resultSizeText ? String(state.resultSizeText) : "";
  }

  function updateResultNameRegenerateButton() {
    if (!regenerateNameBtn) return;

    regenerateNameBtn.disabled = !state.outputBoundsText;
    regenerateNameBtn.classList.toggle("is-active", Boolean(state.resultNameManualOverride && state.outputBoundsText));
  }

  function syncResultNameFieldHeight() {
    if (!resultNameField) return;
    resultNameField.style.height = "auto";
    resultNameField.style.height = Math.max(resultNameField.scrollHeight, 44) + "px";
  }

  function setResultSizeText(text, options) {
    options = options || {};
    state.resultSizeText = text ? String(text) : "";
    if (!options.preserveManualState) {
      state.resultNameManualOverride = Boolean(options.manualOverride);
    }
    if (Object.prototype.hasOwnProperty.call(options, "templateText")) {
      state.resultNameTemplateText = options.templateText ? String(options.templateText) : "";
    } else if (!state.resultNameManualOverride && !options.preserveTemplateState) {
      state.resultNameTemplateText = "";
    }

    if (resultNameFeedbackTimer) {
      window.clearTimeout(resultNameFeedbackTimer);
      resultNameFeedbackTimer = null;
    }

    if (resultNameField) {
      resultNameField.classList.remove("is-copied");
      resultNameField.value = state.resultSizeText || "UVM__--";
      resultNameField.title = state.resultSizeText || "UVM__--";
      syncResultNameFieldHeight();
    }

    updateResultNameRegenerateButton();
    if (exportBtn) exportBtn.disabled = !state.resultSizeText;
  }

  function setPanelFeedback(message, tone, timeoutMs) {
    if (!panelFeedbackEl) return;

    if (panelFeedbackTimer) {
      window.clearTimeout(panelFeedbackTimer);
      panelFeedbackTimer = null;
    }
    if (panelFeedbackHideTimer) {
      window.clearTimeout(panelFeedbackHideTimer);
      panelFeedbackHideTimer = null;
    }

    if (!message) {
      panelFeedbackEl.className = "panel-feedback";
      panelFeedbackHideTimer = window.setTimeout(function () {
        panelFeedbackEl.hidden = true;
        panelFeedbackEl.textContent = "";
        panelFeedbackEl.className = "panel-feedback";
        panelFeedbackHideTimer = null;
      }, 220);
      return;
    }

    panelFeedbackEl.hidden = false;
    panelFeedbackEl.textContent = String(message);
    panelFeedbackEl.className = "panel-feedback" + (tone ? (" is-" + tone) : "");
    window.requestAnimationFrame(function () {
      if (!panelFeedbackEl.hidden) panelFeedbackEl.className += " is-visible";
    });

    if (timeoutMs && timeoutMs > 0) {
      panelFeedbackTimer = window.setTimeout(function () {
        panelFeedbackEl.className = "panel-feedback";
        panelFeedbackHideTimer = window.setTimeout(function () {
          panelFeedbackEl.hidden = true;
          panelFeedbackEl.textContent = "";
          panelFeedbackEl.className = "panel-feedback";
          panelFeedbackHideTimer = null;
        }, 220);
        panelFeedbackTimer = null;
      }, timeoutMs);
    }
  }

  function applyOutputContext(context) {
    if (!context || typeof context !== "object") return;
    state.outputBoundsText = context.outputBoundsText ? String(context.outputBoundsText) : "";
    state.sourceFolderLabel = context.sourceFolderLabel ? String(context.sourceFolderLabel) : "";
    updateResultNameRegenerateButton();
  }

  function applyAutoResultName(settings, context) {
    if (context) applyOutputContext(context);

    if (!state.outputBoundsText) {
      setResultSizeText("", { manualOverride: false, templateText: "" });
      return "";
    }

    var nextText = buildAutoResultName(settings);
    setResultSizeText(nextText, { manualOverride: false, templateText: "" });
    return nextText;
  }

  function applyResultNameTemplate(templateText, settings) {
    var safeTemplate = String(templateText || "");
    if (!safeTemplate) return "";

    var sizeText = buildOutputSizeSegment(state.outputBoundsText);
    if (!sizeText || safeTemplate.indexOf(RESULT_NAME_SIZE_TOKEN) === -1) return safeTemplate;
    return safeTemplate.split(RESULT_NAME_SIZE_TOKEN).join(sizeText);
  }

  function applyManualResultName(settings, templateText) {
    var safeTemplate = sanitizeEditableResultName(templateText);
    var rendered = sanitizeEditableResultName(applyResultNameTemplate(safeTemplate, settings));
    setResultSizeText(rendered, { manualOverride: true, templateText: safeTemplate });
    return rendered;
  }

  function injectSizeToken(text, sizeText) {
    var value = String(text || "");
    var currentSize = String(sizeText || "");
    if (!value || !currentSize) return value;

    var segments = value.split("__");
    for (var i = 0; i < segments.length; i++) {
      if (segments[i] === currentSize) {
        segments[i] = RESULT_NAME_SIZE_TOKEN;
        return segments.join("__");
      }
    }

    var matchIndex = value.indexOf(currentSize);
    if (matchIndex === -1) return value;

    return value.slice(0, matchIndex) + RESULT_NAME_SIZE_TOKEN + value.slice(matchIndex + currentSize.length);
  }

  function buildEditableAutoResultTemplate(settings) {
    return injectSizeToken(buildAutoResultName(settings), buildOutputSizeSegment(state.outputBoundsText));
  }

  function getEditableResultNameTemplate(settings) {
    var currentTemplate = state.resultNameManualOverride
      ? (state.resultNameTemplateText || state.resultSizeText || "")
      : buildEditableAutoResultTemplate(settings);

    if (currentTemplate.indexOf(RESULT_NAME_SIZE_TOKEN) !== -1) return currentTemplate;
    return injectSizeToken(currentTemplate, buildOutputSizeSegment(state.outputBoundsText));
  }

  function refreshDisplayedResultName(settings, context) {
    if (context) applyOutputContext(context);

    if (!state.outputBoundsText) {
      if (state.resultNameManualOverride) {
        setResultSizeText(state.resultSizeText || "", {
          manualOverride: true,
          templateText: state.resultNameTemplateText || state.resultSizeText || ""
        });
        return state.resultSizeText || "";
      }

      setResultSizeText("", { manualOverride: false, templateText: "" });
      return "";
    }

    if (state.resultNameManualOverride) {
      return applyManualResultName(settings, state.resultNameTemplateText || state.resultSizeText || "");
    }

    return applyAutoResultName(settings);
  }

  function sanitizeEditableResultName(text) {
    return String(text || "")
      .replace(/[\r\n\t]+/g, " ")
      .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
      .replace(/\s+/g, " ")
      .replace(/[. ]+$/g, "")
      .trim();
  }

  function isEditingResultName() {
    return Boolean(resultNameField && resultNameField.classList.contains("is-editing"));
  }

  function closeResultNameEditUi() {
    if (!resultNameField) return;
    resultNameField.readOnly = true;
    resultNameField.classList.remove("is-editing");
    resultNameField.removeAttribute("data-edit-start-rendered");
    resultNameField.removeAttribute("data-edit-start-template");
    resultNameField.removeAttribute("data-edit-start-manual");
  }

  function cancelResultNameEditMode() {
    if (!resultNameField || !isEditingResultName()) return;

    var startRendered = resultNameField.getAttribute("data-edit-start-rendered") || state.resultSizeText || "";
    var startTemplate = resultNameField.getAttribute("data-edit-start-template") || "";
    var startManual = resultNameField.getAttribute("data-edit-start-manual") === "true";

    setResultSizeText(startRendered, {
      manualOverride: startManual,
      templateText: startManual ? (startTemplate || startRendered) : ""
    });
    closeResultNameEditUi();
  }

  function exitResultNameEditMode() {
    if (!resultNameField || !isEditingResultName()) return;

    var settings = getFormValues();
    var nextTemplate = sanitizeEditableResultName(resultNameField.value);
    var autoTemplate = sanitizeEditableResultName(buildEditableAutoResultTemplate(settings));

    if (!nextTemplate) {
      refreshDisplayedResultName(settings);
      state.resultNameManualOverride = false;
      state.resultNameTemplateText = "";
      if (state.outputBoundsText) applyAutoResultName(settings);
      else setResultSizeText("", { manualOverride: false, templateText: "" });
    } else if (nextTemplate === autoTemplate) {
      applyAutoResultName(settings);
    } else {
      applyManualResultName(settings, nextTemplate);
    }

    closeResultNameEditUi();
    writeStoredSettings(buildStoredState(settings));
  }

  function enterResultNameEditMode() {
    if (!resultNameField) return;
    if (resultNameClickTimer) {
      window.clearTimeout(resultNameClickTimer);
      resultNameClickTimer = null;
    }

    var settings = getFormValues();
    var editableTemplate = getEditableResultNameTemplate(settings);
    var renderedStart = resultNameField.value || "";

    resultNameField.setAttribute("data-edit-start-rendered", renderedStart);
    resultNameField.setAttribute("data-edit-start-template", state.resultNameTemplateText || editableTemplate || renderedStart);
    resultNameField.setAttribute("data-edit-start-manual", state.resultNameManualOverride ? "true" : "false");
    resultNameField.value = editableTemplate || renderedStart;
    resultNameField.readOnly = false;
    resultNameField.classList.add("is-editing");
    resultNameField.focus();
    resultNameField.select();
    syncResultNameFieldHeight();
  }

  function refreshOutputContext(callback) {
    if (isPreviewMode()) {
      if (callback) callback({ ok: false, preview: true, error: "CEP bridge not found." });
      return;
    }
    evalHost("nesterGetOutputContext()", function (result) {
      var parsed = parseJsonSafe(result);
      if (parsed && parsed.ok) applyOutputContext(parsed);
      if (callback) callback(parsed);
    });
  }

  function refreshOutputNameFromInputs() {
    refreshDisplayedResultName(getFormValues());
  }

  function copyTextToClipboard(text) {
    if (!text) return;

    function fallbackCopy() {
      var textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.setAttribute("readonly", "readonly");
      textarea.style.position = "absolute";
      textarea.style.left = "-9999px";
      document.body.appendChild(textarea);
      textarea.select();
      try { document.execCommand("copy"); } catch (_e) {}
      document.body.removeChild(textarea);
    }

    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).catch(function () {
          fallbackCopy();
        });
        return;
      }
    } catch (_e2) {}

    fallbackCopy();
  }

  function flashResultSizeCopied() {
    if (!resultNameField || !getResultNameText()) return;
    if (resultNameFeedbackTimer) window.clearTimeout(resultNameFeedbackTimer);
    resultNameField.classList.add("is-copied");
    resultNameFeedbackTimer = window.setTimeout(function () {
      resultNameField.classList.remove("is-copied");
      resultNameFeedbackTimer = null;
    }, 1100);
  }

  function findCardElement(node) {
    while (node && node !== inventoryListEl) {
      if (node.className && String(node.className).indexOf("inventory-card") !== -1) return node;
      node = node.parentNode;
    }
    return null;
  }

  function goToPlacedCopy(sourceKey) {
    if (!sourceKey) return;
    if (isPreviewMode()) {
      notifyPreviewMode();
      return;
    }
    var script = 'nesterGoToPlacedCopyBySourceKey("' + escapeForJsxString(sourceKey) + '")';
    evalHost(script, function (_result) {});
  }

  function renderInventory(items) {
    state.lastSourceItems = sanitizeSourceItems(items);

    if (!state.lastSourceItems.length) {
      inventoryListEl.hidden = true;
      inventoryListEl.innerHTML = "";
      inventoryEmptyEl.hidden = false;
      inventoryMetaEl.textContent = "Run NEST once to load items";
      inventoryMetaEl.classList.remove("inventory-meta-warning");
      return;
    }

    var html = [];

    for (var i = 0; i < state.lastSourceItems.length; i++) {
      var item = state.lastSourceItems[i];
      var qtyValue = Object.prototype.hasOwnProperty.call(state.quantityOverrides, item.key)
        ? state.quantityOverrides[item.key]
        : item.detectedQty;
      var placedLabel = item.unplacedQty > 0
        ? (item.placedQty + "/" + item.requestedQty)
        : String(item.placedQty);
      var dimensionsClass = "inventory-dimensions";
      if (item.dimensionStatus === "error") dimensionsClass += " is-error";
      var dimensionsTitle = item.dimensionsTitle || item.dimensionsText;

      html.push(
        '<div class="inventory-card" data-source-key="' + escapeHtml(item.key) + '">' +
          '<div class="inventory-thumb">' +
            (item.thumbnailUrl
              ? '<img class="inventory-thumb-img" src="' + escapeHtml(item.thumbnailUrl) + '" alt="' + escapeHtml(item.name) + '" />'
              : '<span class="inventory-thumb-fallback">PNG</span>') +
          '</div>' +
          '<div class="inventory-main">' +
            '<div class="inventory-name" title="' + escapeHtml(item.name) + '">' +
              '<strong>' + escapeHtml(item.name) + '</strong>' +
            '</div>' +
            '<div class="inventory-controls">' +
              '<span class="inventory-nameqty">' + escapeHtml(String(item.detectedQty)) + '</span>' +
              '<span class="inventory-arrow">&#x1F87A;</span>' +
              '<div class="inventory-input-wrap">' +
                '<input class="inventory-input" type="number" min="1" step="1" value="' + qtyValue + '" data-source-key="' + escapeHtml(item.key) + '" />' +
                '<div class="inventory-stepper">' +
                  '<button class="inventory-stepper-btn" type="button" data-step-dir="1" aria-label="Increase quantity"></button>' +
                  '<button class="inventory-stepper-btn" type="button" data-step-dir="-1" aria-label="Decrease quantity"></button>' +
                '</div>' +
              '</div>' +
              '<span class="inventory-arrow">&#x1F87A;</span>' +
              '<span class="inventory-placed">' + escapeHtml(placedLabel) + '</span>' +
              '<span class="' + dimensionsClass + '" title="' + escapeHtml(dimensionsTitle) + '">' + escapeHtml(item.dimensionsText) + '</span>' +
            '</div>' +
          '</div>' +
        '</div>'
      );
    }

    inventoryListEl.innerHTML = html.join("");
    inventoryListEl.hidden = false;
    inventoryEmptyEl.hidden = true;
    var metaSummary = getInventoryMetaSummary(state.lastSourceItems);
    inventoryMetaEl.textContent = metaSummary.text;
    inventoryMetaEl.classList.toggle("inventory-meta-warning", metaSummary.isIncomplete);
  }

  function refreshResultNameFromHostForAction(callback) {
    if (isPreviewMode()) {
      notifyPreviewMode();
      if (callback) callback(false);
      return;
    }
    var settings = getFormValues();
    refreshOutputContext(function (parsed) {
      if (!parsed) {
        alertUnexpectedHostResponse("null");
        if (callback) callback(false);
        return;
      }
      if (!parsed.ok) {
        window.alert("Error:\n" + parsed.error);
        if (callback) callback(false);
        return;
      }

      refreshDisplayedResultName(settings, parsed);
      writeStoredSettings(buildStoredState(settings));
      if (callback) callback(true, settings);
    });
  }

  function prepareResultNameForAction(callback) {
    exitResultNameEditMode();
    refreshResultNameFromHostForAction(function (ok, settings) {
      if (!ok) {
        if (callback) callback(false);
        return;
      }

      var text = getResultNameText();
      if (!text) {
        if (callback) callback(false, settings, "");
        return;
      }

      copyTextToClipboard(text);
      flashResultSizeCopied();
      if (callback) callback(true, settings, text);
    });
  }

  function runExport(settings) {
    if (isPreviewMode()) {
      notifyPreviewMode();
      return;
    }
    settings = settings || getFormValues();
    var exportName = getResultNameText();
    if (!exportName) {
      window.alert("No output name available to export.");
      return;
    }

    var payload = JSON.stringify({
      fileName: exportName,
      folderPaths: sanitizeFolderPaths(
        state.lastSourceFolderPaths.concat(state.extraExportFolderPath ? [state.extraExportFolderPath] : [])
      )
    });
    var script = 'nesterExportOutputPngToSourceFolders("' + escapeForJsxString(payload) + '")';

    setPanelFeedback("");
    setBusy("export");
    evalHost(script, function (result) {
      setBusy(null);

      var parsed = parseJsonSafe(result);
      if (!parsed) {
        alertUnexpectedHostResponse(result);
        return;
      }
      if (!parsed.ok) {
        window.alert("Error:\n" + parsed.error);
        return;
      }

      writeStoredSettings(buildStoredState(settings));

      if (parsed.failedPaths && parsed.failedPaths.length) {
        setPanelFeedback(
          "Export completed with " + String(parsed.failedPaths.length) + " failed path(s).",
          "warning",
          9000
        );
        var summary = [
          "Export completed with warnings.",
          "Exported to " + String((parsed.exportedPaths && parsed.exportedPaths.length) || 0) + " folder(s)."
        ];
        if (parsed.overwrittenCount) summary.push("Overwritten: " + String(parsed.overwrittenCount));
        summary.push("Failed:");
        summary.push(parsed.failedPaths.join("\n"));
        window.alert(summary.join("\n"));
        return;
      }

      var message = "Exported PNG to " + String((parsed.exportedPaths && parsed.exportedPaths.length) || 0) + " folder.";
      if ((parsed.exportedPaths && parsed.exportedPaths.length) !== 1) message = message.replace(" folder.", " folders.");
      if (parsed.overwrittenCount) message += " Replaced " + String(parsed.overwrittenCount) + " existing file.";
      if (parsed.overwrittenCount > 1) message = message.replace(" existing file.", " existing files.");
      setPanelFeedback(message, "success", 5200);
    });
  }

  function runNester(settings) {
    if (isPreviewMode()) {
      notifyPreviewMode();
      return;
    }
    var hadPendingOrderText = Boolean(state.orderEmailText);
    var payload = JSON.stringify(settings);
    var script = 'nesterRunWithSettings("' + escapeForJsxString(payload) + '")';

    setBusy("nest");
    evalHost(script, function (result) {
      setBusy(null);
      var parsed = parseJsonSafe(result);
      if (!parsed) {
        alertUnexpectedHostResponse(result);
        return;
      }
      if (!parsed.ok) {
        window.alert("Error:\n" + parsed.error);
        return;
      }

      if (parsed.sourceItems && parsed.sourceItems.length) {
        syncQuantityOverridesFromResult(parsed.sourceItems);
        renderInventory(parsed.sourceItems);
      }
      state.lastSourceFolderPaths = sanitizeFolderPaths(parsed.sourceFolderPaths);

      applyOutputContext(parsed);
      refreshDisplayedResultName(settings);
      if (hadPendingOrderText) clearPendingOrderState();
      writeStoredSettings(buildStoredState(settings));
    });
  }

  function initWithDefaults(hostDefaults) {
    var combinedDefaults = merge(DEFAULTS, hostDefaults);
    var stored = readStoredSettings() || {};
    var startSettings = merge(combinedDefaults, stored);

    state.quantityOverrides = sanitizeQuantityOverrides(stored.quantityOverrides);
    state.lastSourceItems = sanitizeSourceItems(stored.lastSourceItems);
    state.lastSourceFolderPaths = sanitizeFolderPaths(stored.lastSourceFolderPaths);
    state.extraExportFolderPath = stored.extraExportFolderPath ? String(stored.extraExportFolderPath) : "";
    state.outputBoundsText = stored.outputBoundsText ? String(stored.outputBoundsText) : "";
    state.sourceFolderLabel = stored.sourceFolderLabel ? String(stored.sourceFolderLabel) : "";
    state.resultSizeText = stored.resultSizeText ? String(stored.resultSizeText) : "";
    state.resultNameManualOverride = Boolean(stored.resultNameManualOverride);
    state.resultNameTemplateText = stored.resultNameTemplateText ? String(stored.resultNameTemplateText) : "";
    state.orderEmailText = "";

    setFormValues(startSettings);
    updateExtraExportFolderButton();
    renderInventory(state.lastSourceItems);
    if (orderEmailTextEl) orderEmailTextEl.value = "";
    setOrderEmailVisibility(false);
    if (state.outputBoundsText) refreshDisplayedResultName(startSettings);
    else {
      setResultSizeText(state.resultSizeText, {
        manualOverride: state.resultNameManualOverride,
        templateText: state.resultNameTemplateText || (state.resultNameManualOverride ? state.resultSizeText : "")
      });
    }

    inventoryListEl.addEventListener("input", function (evt) {
      var target = evt.target;
      if (!target || !target.getAttribute("data-source-key")) return;
      suspendInventoryInputSync = false;
      state.quantityOverrides[target.getAttribute("data-source-key")] = sanitizeQty(target.value, 1);
      writeStoredSettings(buildStoredState(getFormValues()));
      scheduleAutoNest();
    });
    inventoryListEl.addEventListener("mousedown", function (evt) {
      if (findInventoryStepButton(evt.target)) evt.preventDefault();
    });
    inventoryListEl.addEventListener("click", function (evt) {
      var target = evt.target;
      var button = findInventoryStepButton(target);
      var input;
      var direction;
      if (!target) return;
      if (button) {
        direction = Number(button.getAttribute("data-step-dir"));
        if (direction !== 1 && direction !== -1) return;
        input = button.parentNode ? button.parentNode.parentNode.querySelector(".inventory-input") : null;
        if (!input) return;
        stepInventoryQuantityInput(input, direction);
        return;
      }
      if (target.tagName === "INPUT") return;
      var card = findCardElement(target);
      if (!card) return;
      goToPlacedCopy(card.getAttribute("data-source-key"));
    });

    function handleFieldUiUpdate(evt) {
      var target = evt.target;
      if (!target || !target.id) return;

      if (target.type === "range") updateSolverSliderValue(target.id);
      if (target.id === "legacySolverEnabled") updateLegacySolverUi();
      if (NAMING_RESETTABLE_IDS[target.id]) updateNamingFieldVisualState(target.id);

      if (isNamingFieldId(target.id)) {
        refreshOutputNameFromInputs();
        writeStoredSettings(buildStoredState(getFormValues()));
        return;
      }

      if (isAutoNestFieldId(target.id)) {
        writeStoredSettings(buildStoredState(getFormValues()));
        scheduleAutoNest();
      }
    }

    formEl.addEventListener("mousedown", function (evt) {
      if (findStepButton(evt.target)) evt.preventDefault();
    });
    formEl.addEventListener("click", function (evt) {
      var resetTarget = evt.target && evt.target.getAttribute ? evt.target.getAttribute("data-default-target") : "";
      var button = findStepButton(evt.target);
      var direction;
      var input;
      if (resetTarget) {
        evt.preventDefault();
        resetNamingFieldToDefault(resetTarget);
        return;
      }
      if (!button) return;

      direction = Number(button.getAttribute("data-step-dir"));
      if (direction !== 1 && direction !== -1) return;

      input = byId(button.getAttribute("data-step-target"));
      if (!input) return;

      stepNamingNumberField(input, direction);
    });
    formEl.addEventListener("input", handleFieldUiUpdate);
    formEl.addEventListener("change", handleFieldUiUpdate);

    if (orderEmailBtn) {
      orderEmailBtn.addEventListener("click", function () {
        var willShow = !orderEmailWrapEl || Boolean(orderEmailWrapEl.hidden);
        setOrderEmailVisibility(willShow);
        if (willShow && orderEmailTextEl) orderEmailTextEl.focus();
      });
    }

    if (orderEmailBackdropEl) {
      orderEmailBackdropEl.addEventListener("click", function () {
        setOrderEmailVisibility(false);
      });
    }

    if (orderEmailCloseBtn) {
      orderEmailCloseBtn.addEventListener("click", function () {
        setOrderEmailVisibility(false);
      });
    }

    if (orderEmailTextEl) {
      orderEmailTextEl.addEventListener("input", function () {
        state.orderEmailText = normalizeMultilineText(orderEmailTextEl.value);
        setOrderEmailVisibility(!orderEmailWrapEl || !orderEmailWrapEl.hidden);
        suspendInventoryInputSync = true;
        state.quantityOverrides = {};
        writeStoredSettings(buildStoredState(getFormValues()));
        scheduleAutoNest();
      });
      orderEmailTextEl.addEventListener("paste", function () {
        window.setTimeout(function () {
          state.orderEmailText = normalizeMultilineText(orderEmailTextEl.value);
          suspendInventoryInputSync = true;
          state.quantityOverrides = {};
          writeStoredSettings(buildStoredState(getFormValues()));
          scheduleAutoNest();
          setOrderEmailVisibility(false);
        }, 0);
      });
    }

    if (resultNameField) {
      resultNameField.addEventListener("click", function () {
        if (isEditingResultName()) return;
        if (resultNameClickTimer) window.clearTimeout(resultNameClickTimer);
        resultNameClickTimer = window.setTimeout(function () {
          resultNameClickTimer = null;
          prepareResultNameForAction(function (_ok) {});
        }, 260);
      });

      resultNameField.addEventListener("dblclick", function () {
        enterResultNameEditMode();
      });

      resultNameField.addEventListener("input", function () {
        syncResultNameFieldHeight();
      });

      resultNameField.addEventListener("blur", function () {
        exitResultNameEditMode();
      });

      resultNameField.addEventListener("keydown", function (evt) {
        if (!isEditingResultName()) return;

        if (evt.key === "Enter") {
          evt.preventDefault();
          resultNameField.blur();
          return;
        }

        if (evt.key === "Escape") {
          evt.preventDefault();
          cancelResultNameEditMode();
        }
      });
    }

    document.addEventListener("keydown", function (evt) {
      if (evt.key === "Escape" && orderEmailWrapEl && !orderEmailWrapEl.hidden) {
        setOrderEmailVisibility(false);
      }
    });

    formEl.addEventListener("submit", function (evt) {
      evt.preventDefault();
      cancelAutoNest();
      var settings = getFormValues();
      writeStoredSettings(buildStoredState(settings));
      runNester(settings);
    });

    if (exportBtn) {
      exportBtn.addEventListener("click", function () {
        prepareResultNameForAction(function (ok, settings) {
          if (!ok) return;
          runExport(settings);
        });
      });
    }
    if (exportFolderBtn) {
      exportFolderBtn.addEventListener("click", function () {
        chooseExtraExportFolder();
      });
    }
    if (regenerateNameBtn) {
      regenerateNameBtn.addEventListener("click", function () {
        exitResultNameEditMode();
        state.resultNameManualOverride = false;
        state.resultNameTemplateText = "";
        refreshDisplayedResultName(getFormValues());
        writeStoredSettings(buildStoredState(getFormValues()));
      });
    }

    resetBtn.addEventListener("click", function () {
      cancelAutoNest();
      exitResultNameEditMode();
      setFormValues(combinedDefaults);
      state.quantityOverrides = {};
      clearPendingOrderState();
      renderInventory(state.lastSourceItems);
      state.resultNameManualOverride = false;
      state.resultNameTemplateText = "";
      updateExtraExportFolderButton();
      if (state.outputBoundsText) refreshDisplayedResultName(combinedDefaults);
      else setResultSizeText("", { manualOverride: false, templateText: "" });
      writeStoredSettings(buildStoredState(combinedDefaults));
    });
  }

  if (isPreviewMode()) {
    initWithDefaults(DEFAULTS);
    return;
  }

  evalHost("nesterGetDefaultSettings()", function (raw) {
    var parsed = parseJsonSafe(raw);
    if (!parsed || parsed.ok === false) {
      initWithDefaults(DEFAULTS);
      return;
    }
    initWithDefaults(parsed);
  });
})();


