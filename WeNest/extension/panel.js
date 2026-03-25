(function () {
  "use strict";

  var today = new Date().getDate();
  if (!isFinite(today) || today < 1 || today > 31) today = 1;

  var STORAGE_KEY = "nester.settings.v5";
  var formEl = document.getElementById("settings-form");
  var runBtn = document.getElementById("btn-run");
  var resetBtn = document.getElementById("btn-reset");
  var inventoryListEl = document.getElementById("inventory-list");
  var inventoryEmptyEl = document.getElementById("inventory-empty");
  var inventoryMetaEl = document.getElementById("inventory-meta");
  var resultSizeBtn = document.getElementById("btn-copy-output-size");

  var DEFAULTS = {
    sheetWidthIn: 23,
    maxLengthIn: 100,
    spacingIn: 0.25,
    optimizePreset: "Auto",
    searchEffort: "Normal",
    allowItemRotationInBlock: true,
    allowBlockRotationOnSheet: true,
    outputDate: today,
    tagUSA: false,
    tagRUSH: false,
    tagCUT: false,
    tagOverNight: false,
    tagPurolator: false,
    tagGLS: false,
    outputPcCount: 1,
    outputChokeCount: 6
  };

  var FIELD_IDS = [
    "sheetWidthIn",
    "maxLengthIn",
    "spacingIn",
    "optimizePreset",
    "searchEffort",
    "allowItemRotationInBlock",
    "allowBlockRotationOnSheet",
    "outputDate",
    "tagUSA",
    "tagRUSH",
    "tagCUT",
    "tagOverNight",
    "tagPurolator",
    "tagGLS",
    "outputPcCount",
    "outputChokeCount"
  ];

  var state = {
    quantityOverrides: {},
    lastSourceItems: [],
    resultSizeText: "",
    outputBoundsText: "",
    sourceFolderLabel: ""
  };
  var resultSizeFeedbackTimer = null;

  function byId(id) {
    return document.getElementById(id);
  }

  function hasCepBridge() {
    return typeof window.__adobe_cep__ !== "undefined";
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
        dimensionsText: item.dimensionsText ? String(item.dimensionsText) : ""
      });
    }

    return out;
  }

  function readStoredSettings() {
    try {
      var raw = window.localStorage.getItem(STORAGE_KEY);
      return raw ? parseJsonSafe(raw) : null;
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
      resultSizeText: state.resultSizeText,
      outputBoundsText: state.outputBoundsText,
      sourceFolderLabel: state.sourceFolderLabel
    });
    return out;
  }

  function evalHost(script, callback) {
    if (!hasCepBridge()) {
      window.alert("CEP bridge not found.");
      return;
    }
    window.__adobe_cep__.evalScript(script, function (result) {
      if (callback) callback(result);
    });
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
      else el.value = String(settings[id]);
    }
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
      else if (el.type === "number") settings[id] = Number(el.value);
      else settings[id] = el.value;
    }

    syncOverridesFromInputs();
    settings.quantityOverrides = sanitizeQuantityOverrides(state.quantityOverrides);
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
      id === "outputPcCount" ||
      id === "outputChokeCount";
  }

  function composeOutputNameText(boundsText, folderLabel, settings) {
    if (!boundsText) return "";

    var prefixParts = [
      "UVM",
      String(sanitizeCountInRange(settings.outputDate, 1, 31, DEFAULTS.outputDate))
    ];
    var pcCount = sanitizeCountInRange(settings.outputPcCount, 1, 9, DEFAULTS.outputPcCount);
    var chokeCount = sanitizeCountInRange(settings.outputChokeCount, 1, 8, DEFAULTS.outputChokeCount);

    if (settings.tagUSA) prefixParts.push("USA");
    if (settings.tagRUSH) prefixParts.push("RUSH");
    if (settings.tagOverNight) prefixParts.push("OverNight");
    if (settings.tagCUT) prefixParts.push("CUT");
    if (settings.tagPurolator) prefixParts.push("Purolator");
    if (settings.tagGLS) prefixParts.push("GLS");

    var parts = [
      prefixParts.join("_"),
      String(boundsText),
      String(pcCount) + (pcCount > 1 ? "pcs" : "pc")
    ];

    if (chokeCount !== 6) parts.push(String(chokeCount) + "Choke");

    if (folderLabel) parts.push(String(folderLabel));
    return parts.join("__");
  }

  function setBusy(isBusy) {
    runBtn.disabled = isBusy;
    resetBtn.disabled = isBusy;
    if (resultSizeBtn) resultSizeBtn.disabled = isBusy || !state.resultSizeText;
    runBtn.textContent = isBusy ? "Building..." : "NEST";
  }

  function syncQuantityOverridesFromResult(items) {
    var next = {};
    for (var i = 0; i < items.length; i++) {
      next[items[i].key] = sanitizeQty(items[i].requestedQty, items[i].detectedQty);
    }
    state.quantityOverrides = next;
  }

  function buildInventoryMeta(items) {
    var fileCount = items.length;
    var placed = 0;
    var requested = 0;

    for (var i = 0; i < items.length; i++) {
      placed += items[i].placedQty;
      requested += items[i].requestedQty;
    }

    return fileCount + " files | " + placed + "/" + requested + " placed";
  }

  function setResultSizeText(text) {
    state.resultSizeText = text ? String(text) : "";
    if (resultSizeFeedbackTimer) {
      window.clearTimeout(resultSizeFeedbackTimer);
      resultSizeFeedbackTimer = null;
    }
    resultSizeBtn.classList.remove("is-copied");
    resultSizeBtn.textContent = state.resultSizeText || "UVM__--";
    resultSizeBtn.disabled = !state.resultSizeText;
  }

  function refreshOutputNameFromInputs() {
    if (!state.outputBoundsText) {
      setResultSizeText("");
      return;
    }
    setResultSizeText(composeOutputNameText(state.outputBoundsText, state.sourceFolderLabel, getFormValues()));
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
    if (!resultSizeBtn || !state.resultSizeText) return;
    if (resultSizeFeedbackTimer) window.clearTimeout(resultSizeFeedbackTimer);
    resultSizeBtn.classList.add("is-copied");
    resultSizeBtn.textContent = state.resultSizeText + "  Copied";
    resultSizeFeedbackTimer = window.setTimeout(function () {
      resultSizeBtn.classList.remove("is-copied");
      resultSizeBtn.textContent = state.resultSizeText || "UVM__--";
      resultSizeFeedbackTimer = null;
    }, 1100);
  }

  function findCardElement(node) {
    while (node && node !== inventoryListEl) {
      if (node.className && String(node.className).indexOf("inventory-card") !== -1) return node;
      node = node.parentNode;
    }
    return null;
  }

  function selectPlacedCopy(sourceKey) {
    if (!sourceKey) return;
    var script = 'nesterSelectPlacedCopyBySourceKey("' + escapeForJsxString(sourceKey) + '")';
    evalHost(script, function (_result) {});
  }

  function renderInventory(items) {
    state.lastSourceItems = sanitizeSourceItems(items);

    if (!state.lastSourceItems.length) {
      inventoryListEl.hidden = true;
      inventoryListEl.innerHTML = "";
      inventoryEmptyEl.hidden = false;
      inventoryMetaEl.textContent = "Run NEST once to load items";
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
              '</div>' +
              '<span class="inventory-arrow">&#x1F87A;</span>' +
              '<span class="inventory-placed">' + escapeHtml(placedLabel) + '</span>' +
              '<span class="inventory-dimensions">' + escapeHtml(item.dimensionsText) + '</span>' +
            '</div>' +
          '</div>' +
        '</div>'
      );
    }

    inventoryListEl.innerHTML = html.join("");
    inventoryListEl.hidden = false;
    inventoryEmptyEl.hidden = true;
    inventoryMetaEl.textContent = buildInventoryMeta(state.lastSourceItems);
  }

  function runNester(settings) {
    var payload = JSON.stringify(settings);
    var script = 'nesterRunWithSettings("' + escapeForJsxString(payload) + '")';

    setBusy(true);
    evalHost(script, function (result) {
      setBusy(false);
      var parsed = parseJsonSafe(result);
      if (!parsed) {
        window.alert("Unexpected host response.");
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

      state.outputBoundsText = parsed.outputBoundsText ? String(parsed.outputBoundsText) : "";
      state.sourceFolderLabel = parsed.sourceFolderLabel ? String(parsed.sourceFolderLabel) : "";
      setResultSizeText(
        composeOutputNameText(state.outputBoundsText, state.sourceFolderLabel, settings) ||
        (parsed.outputSizeText ? String(parsed.outputSizeText) : "")
      );
      writeStoredSettings(buildStoredState(settings));
    });
  }

  function initWithDefaults(hostDefaults) {
    var combinedDefaults = merge(DEFAULTS, hostDefaults);
    var stored = readStoredSettings() || {};
    var startSettings = merge(combinedDefaults, stored);

    state.quantityOverrides = sanitizeQuantityOverrides(stored.quantityOverrides);
    state.lastSourceItems = sanitizeSourceItems(stored.lastSourceItems);
    state.outputBoundsText = stored.outputBoundsText ? String(stored.outputBoundsText) : "";
    state.sourceFolderLabel = stored.sourceFolderLabel ? String(stored.sourceFolderLabel) : "";
    state.resultSizeText = stored.resultSizeText ? String(stored.resultSizeText) : "";

    setFormValues(startSettings);
    renderInventory(state.lastSourceItems);
    if (state.outputBoundsText) refreshOutputNameFromInputs();
    else setResultSizeText(state.resultSizeText);

    inventoryListEl.addEventListener("input", function (evt) {
      var target = evt.target;
      if (!target || !target.getAttribute("data-source-key")) return;
      state.quantityOverrides[target.getAttribute("data-source-key")] = sanitizeQty(target.value, 1);
      writeStoredSettings(buildStoredState(getFormValues()));
    });
    inventoryListEl.addEventListener("click", function (evt) {
      var target = evt.target;
      if (!target) return;
      if (target.tagName === "INPUT") return;
      var card = findCardElement(target);
      if (!card) return;
      selectPlacedCopy(card.getAttribute("data-source-key"));
    });

    function handleNamingFieldUpdate(evt) {
      var target = evt.target;
      if (!target || !target.id || !isNamingFieldId(target.id)) return;
      refreshOutputNameFromInputs();
      writeStoredSettings(buildStoredState(getFormValues()));
    }

    formEl.addEventListener("input", handleNamingFieldUpdate);
    formEl.addEventListener("change", handleNamingFieldUpdate);

    if (resultSizeBtn) {
      resultSizeBtn.addEventListener("click", function () {
        if (!state.resultSizeText) return;
        copyTextToClipboard(state.resultSizeText);
        flashResultSizeCopied();
      });
    }

    formEl.addEventListener("submit", function (evt) {
      evt.preventDefault();
      var settings = getFormValues();
      writeStoredSettings(buildStoredState(settings));
      runNester(settings);
    });

    resetBtn.addEventListener("click", function () {
      setFormValues(combinedDefaults);
      state.quantityOverrides = {};
      renderInventory(state.lastSourceItems);
      if (state.outputBoundsText) refreshOutputNameFromInputs();
      else setResultSizeText("");
      writeStoredSettings(buildStoredState(combinedDefaults));
    });
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





