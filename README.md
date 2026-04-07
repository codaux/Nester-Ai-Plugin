# Nester CEP Panel for Adobe Illustrator 2026

A custom CEP panel for Adobe Illustrator 2026 that runs the Nester layout workflow directly inside Illustrator.

This repository packages the panel UI, CEP manifest, install script, and ExtendScript host logic required to run the tool as an unsigned development extension on Windows.

## Screenshots

![Nester panel overview](./Screenshot/Screenshot%202026-03-23%20000810.png)

![Nester inventory and naming workflow](./Screenshot/Screenshot%202026-03-23%20000844.png)

## What It Does

- Adds a `Nester` panel under `Window > Extensions (Legacy)` in Illustrator.
- Lets you run the nesting workflow directly from a compact panel UI.
- Sends panel settings to the ExtendScript host through `nesterRunWithSettings(...)`.
- Supports sheet controls such as width, max length, spacing, preset selection, and solver effort.
- Includes quantity editing, naming helpers, output size copy, and source item inspection inside the panel.
- Uses the full host-side nesting logic from `extension/jsx/host.jsx`.

## Project Structure

```text
WeNest/
  extension/
    CSXS/
      manifest.xml
    jsx/
      host.jsx
    icons/
    index.html
    panel.js
    styles.css
  Screenshot/
  install-dev.ps1
  README.md
```

## Requirements

- Windows
- Adobe Illustrator 2026
- PowerShell
- Permission to install an unsigned CEP extension in your user profile

## Install on Windows

1. Close Illustrator completely.
2. Open PowerShell in the project folder.
3. Run:

   ```powershell
   cd WeNest
   .\install-dev.ps1
   ```

4. Start Illustrator.
5. Open the panel from:

   `Window > Extensions (Legacy) > Nester`

## What the Install Script Does

The included `install-dev.ps1` script:

- Copies the contents of `extension/` into your CEP extensions directory.
- Installs the panel under the default extension ID `com.nester.ai.cepstarter`.
- Enables `PlayerDebugMode=1` for `CSXS.8` through `CSXS.13`.
- Refuses to run if Illustrator is still open.

By default, the extension is installed to:

`%APPDATA%\Adobe\CEP\extensions\com.nester.ai.cepstarter`

Running the script again updates the installed extension in place.

## Development Notes

- The panel UI lives in `extension/index.html`, `extension/styles.css`, and `extension/panel.js`.
- The Illustrator-side execution logic lives in `extension/jsx/host.jsx`.
- The CEP manifest is located at `extension/CSXS/manifest.xml`.
- The current workflow is aimed at local development and unsigned CEP testing.

## Troubleshooting

- If the panel does not appear, make sure Illustrator was fully closed before running `install-dev.ps1`.
- Verify the installed manifest exists under `%APPDATA%\Adobe\CEP\extensions\com.nester.ai.cepstarter\CSXS\manifest.xml`.
- Confirm the panel is being opened from `Window > Extensions (Legacy)`.
- Restart Illustrator after every install or update.
- If CEP still fails to load, check `%TEMP%` for CEP-related logs.

## Notes

This repository is currently structured as a development CEP extension rather than a signed production release.

