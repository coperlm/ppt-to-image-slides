# PowerPoint转图片幻灯片工具

该工具可以将PowerPoint文件的每一页导出为图片，然后创建一个新的由纯图片组成的PPT文件

避免了传输过程中由于字体或者平台的不同，导致格式错乱

默认导出PPT方案具有分辨率过低等问题，而导出PDF虽然足够清晰但是难以直接和原PDF
一样播放

适用于学术PPT（不支持动画效果，仅能保留静态页面）

已完整打包Releases的exe文件，无需python环境，点击即用

适用于windows下的office版本PPT，暂未适配WPS和其他操作系统

## 🚀 快速开始

```
pip install requirements.txt
py main.py
```

## 📋 系统要求

- **操作系统**：Windows 7/8/10/11
- **Python**：3.6或更高版本
- **办公软件**：Microsoft PowerPoint 2010或更高版本
- **依赖包**：pywin32, python-pptx, Pillow（自动安装）

## 📄 许可证

本项目采用 MIT 许可证。详情请查看 LICENSE 文件。
