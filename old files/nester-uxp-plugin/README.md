# Nester DTF Preview – Illustrator UXP Prototype

This is a minimal UXP-based prototype panel for Adobe Illustrator to test
a persistent-panel workflow for a future DTF nesting tool.

## Features

- Persistent panel with simple UI:
  - Style: Tight / Balanced / Production
  - Blocking: 1 / 2 / 3
  - Width Fill: 0–100 slider
  - Build Preview button
  - Clear Preview button
  - Status text area
- On **Build Preview**:
  - Reads the current selection in Illustrator
  - Deletes any existing `NEST_PREVIEW` layer
  - Creates a new `NEST_PREVIEW` layer
  - Duplicates selected items into this layer
  - Arranges them in a very simple grid across the first artboard
  - Keeps the panel open and updates status text
- On **Clear Preview**:
  - Deletes the `NEST_PREVIEW` layer if it exists
  - Keeps the panel open

The grid layout is intentionally simple and modular so that it can later be
replaced with a real nesting engine.

## Folder Structure

```text
nester-uxp-plugin/
  manifest.json
  index.html
  panel.js
  styles.css
  README.md
```

## Installing and Running with Adobe UXP Developer Tool (Windows)

1. **Copy the folder**

   Place the `nester-uxp-plugin` folder somewhere on disk, e.g.:

   ```text
   C:\My Projects\Nester Ai Plugin\nester-uxp-plugin
   ```

2. **Open the UXP Developer Tool**

   - Launch the Adobe UXP Developer Tool.
   - Go to the **Plugins** tab.

3. **Add the plugin**

   - Click **Add Plugin** (or **Add Existing Plugin**).
   - Browse to the `nester-uxp-plugin` folder and select it (the folder that contains `manifest.json`).

4. **Load the plugin**

   - After it appears in the list, select `Nester DTF Preview`.
   - Click **Load** / **Run** for that plugin entry.

5. **Open Illustrator and the panel**

   - Start Adobe Illustrator (or bring it to the foreground).
   - In Illustrator, open the panel via the menu (exact menu labels may vary slightly by version):
     - `Window` → `Plugins` / `Extensions` / `UXP Plugins` → `Nester DTF Preview`
   - The panel should appear as a persistent panel.

6. **Test the prototype**

   - Create a document and some artwork.
   - Select some page items.
   - In the `Nester DTF Preview` panel:
     - Adjust **Style**, **Blocking**, and **Width Fill** as desired.
     - Click **Build Preview**:
       - A `NEST_PREVIEW` layer should be created.
       - Duplicates of the selection should be laid out in a simple grid.
       - The status text should update (e.g. “Preview built successfully.”).
     - Click **Clear Preview**:
       - The `NEST_PREVIEW` layer should be deleted if present.
       - The status text should update.

## Where to Plug In the Real Nesting Engine

- The current implementation uses a simple grid layout:
  - See `panel.js` → `layoutAsGridOnArtboard(items, artboardBounds, options)`.
- To integrate a real nesting engine:
  1. Keep the **selection duplication** and **layer management** logic as is.
  2. Replace the body of `layoutAsGridOnArtboard` with:
     - Calls to your nesting engine
     - Computation of final positions and rotations per duplicate
     - Application of transforms on each item

This keeps the panel workflow and rebuild cycle the same, while swapping out
only the layout engine.

