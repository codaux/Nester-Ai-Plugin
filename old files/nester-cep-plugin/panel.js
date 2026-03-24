(function () {
  'use strict';

  function $(id) {
    return document.getElementById(id);
  }

  function evalExtScript(code, callback) {
    if (window.__adobe_cep__ && typeof window.__adobe_cep__.evalScript === 'function') {
      window.__adobe_cep__.evalScript(code, callback || function () {});
    } else {
      console.warn('CEP host not available.');
      if (callback) {
        callback('CEP host not available.');
      }
    }
  }

  function setStatus(message) {
    var statusText = $('statusText');
    if (statusText) {
      statusText.textContent = message;
    }
  }

  function initUI() {
    var styleSelect = $('styleSelect');
    var blockingSelect = $('blockingSelect');
    var widthFillRange = $('widthFillRange');
    var widthFillValue = $('widthFillValue');
    var buildPreviewBtn = $('buildPreviewBtn');
    var clearPreviewBtn = $('clearPreviewBtn');

    widthFillRange.addEventListener('input', function () {
      widthFillValue.textContent = widthFillRange.value;
    });

    buildPreviewBtn.addEventListener('click', function () {
      var style = styleSelect.value;
      var blocking = parseInt(blockingSelect.value, 10);
      var widthFill = parseInt(widthFillRange.value, 10);

      setStatus('Building preview...');

      // NOTE: this is the main hook to the future nesting engine.
      // For now it calls a simple grid layout implemented in JSX.
      var script =
        'buildPreview("' +
        style +
        '",' +
        blocking +
        ',' +
        widthFill +
        ');';

      evalExtScript(script, function (result) {
        if (typeof result === 'string' && result.indexOf('ERROR:') === 0) {
          setStatus(result);
        } else {
          setStatus(result || 'Preview built.');
        }
      });
    });

    clearPreviewBtn.addEventListener('click', function () {
      setStatus('Clearing preview...');
      evalExtScript('clearPreview();', function (result) {
        setStatus(result || 'Preview cleared.');
      });
    });

    setStatus('Ready.');
  }

  document.addEventListener('DOMContentLoaded', function () {
    initUI();
  });
})();

