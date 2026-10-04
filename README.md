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

## Changelog

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
