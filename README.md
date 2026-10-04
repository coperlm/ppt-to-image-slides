# PowerPoint to Image Slides Tool

This tool can export each slide of a PowerPoint file as an image, then create a new PPT file composed entirely of images (as backgrounds).

It avoids format corruption during transmission due to different fonts or platforms.

Limitations of existing solutions:
 - Default PPT export has low resolution issues
 - Exported PDFs are clear enough but difficult to play normally like the original PDF (limited to platforms if using WPS)

Suitable for academic PPT presentations (does not support animation effects, only static pages can be preserved)

This program is compatible with Office version PPT on Windows, not yet adapted for WPS and other operating systems (the exported PPT can be used on any platform)

V2.0.0 Update: Previously, images could be dragged and modified arbitrarily. The latest version sets them as backgrounds, making them less likely to be modified. If you need the old version, please download v1.0.0

## 🚀 Quick Start

```
pip install -r requirements.txt
py main.py
```

Fully packaged exe files are available in Releases, no Python environment required, ready to use

## 📋 System Requirements

- **Operating System**: Windows 7/8/10/11
- **Python**: 3.6 or higher
- **Office Software**: Microsoft PowerPoint 2010 or higher
- **Dependencies**: pywin32, Pillow (automatically installed); tkinterdnd2 (optional, for drag-and-drop)

## 🌐 Web Version (no install, any OS)

A pure-frontend port lives in [`web/`](web): it renders each slide inside the browser and rebuilds the deck with those images as true slide backgrounds — the same output model as the desktop tool.

- **Runs entirely in your browser.** The file is never uploaded; parsing, rendering and packing all happen locally.
- Works on Windows / macOS / Linux / mobile — no PowerPoint, no Python, nothing to install.
- Legacy `.ppt` (97-2003 binary) is not supported and cannot be: it is an OLE2 compound document, not a ZIP of XML, and no browser-side engine can lay it out. Open it in PowerPoint/WPS and **Save As `.pptx`** first — one click, then the whole pipeline works.
- Deployed to GitHub Pages by GitHub Actions ([`.github/workflows/deploy-pages.yml`](.github/workflows/deploy-pages.yml)) on every push to `main`.

Run it locally:

```
cd web
npm install --include=dev
npm run dev        # http://127.0.0.1:5174
npm test           # unit tests (vitest + jsdom)
npm run test:e2e   # end-to-end in headless Chromium (Playwright)
```

Fidelity evaluation tooling (reference renders + corpus checklist) lives in [`web/tools/fidelity/`](web/tools/fidelity/corpus.md).

How it differs from the desktop tool:

| | Desktop (`main.py`) | Web (`web/`) |
|---|---|---|
| Rendering engine | Real PowerPoint via COM | `pptx-preview` in the browser |
| Input | `.ppt` and `.pptx` | unencrypted `.pptx` only |
| Speaker notes | Preserved (the original file is reused as a template) | Preserved as plain text (formatting is dropped) |
| Fidelity | Exact | Depends on the fonts installed on your machine |
| Requires | Windows + Office | Any modern browser |

Keep the tab in the foreground while converting: browsers suspend rendering in background tabs, so the conversion pauses until you switch back.

## Changelog

### 2026/10/4

#### web
- feat: Pure-frontend web version in `web/`, built and deployed to GitHub Pages by GitHub Actions
- feat: Quality presets defined by output long edge (3840 / 2560 / 1920 px), with `dpr` pinned so output resolution no longer varies with the visitor's display scaling
- feat: Structural self-check that resolves each slide's background `r:embed` through its rels to a real JPEG/PNG part inside the package, and blocks the download when page count or backgrounds do not match

### 2025/9/19

#### v2.2.0
- feat: Discontinued PNG usage, switched to JPEG format images, significantly reducing generated file size by approximately 6-8 times

### 2025/9/18

#### v2.1.0
- feat: File drag and drop functionality
- fix: Path truncation issue

### 2025/9/17 

#### v2.0.1
- fix: Fixed blank page issue, changed from deletion to creation

#### v2.0.0 
- feat: Changed image format to background images

### 2025/7/3
- Packaged as ready-to-use exe file for convenience

### 2025/6/30

#### v1.0.0 

- Initial version release

## 📄 License

This project is licensed under the MIT License. See the LICENSE file for details.
