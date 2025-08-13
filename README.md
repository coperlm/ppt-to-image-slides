# PowerPoint转图片幻灯片工具

该工具可以将PowerPoint文件的每一页导出为图片，然后创建一个新的由纯图片组成的PPT文件

避免了传输过程中由于字体或者平台的不同，导致格式错乱

默认导出PPT方案具有分辨率过低等问题，而导出PDF虽然足够清晰但是难以直接和原PDF
一样播放

适用于学术PPT（不支持动画效果，仅能保留静态页面）

已完整打包Releases的exe文件，无需python环境，点击即用

适用于windows下的office版本PPT，暂未适配WPS和其他操作系统

## ✨ 功能特点

1. **🖥️ 图形化界面**：简单易用的GUI界面，支持文件拖拽选择
2. **🎯 智能DPI检测**：自动检测导出图片的实际DPI，确保尺寸精确
3. **📐 页面尺寸保持**：新PPT自动使用与原PPT相同的页面尺寸
4. **🖼️ 多格式支持**：支持PNG和JPG格式
5. **📏 智能适配**：图片自动适应幻灯片尺寸，填满页面
6. **🌏 中文支持**：完美支持中文路径和包含空格的路径
7. **📁 自动管理**：自动创建输出目录和清理临时文件
8. **⚡ 实时反馈**：实时日志显示，直观的转换进度
9. **🔧 双重模式**：提供图形化界面和命令行两种使用方式

## 🚀 快速开始

### 第一步：安装依赖
双击运行 `install_dependencies.bat` 自动安装所需的Python包。

### 第二步：启动程序
双击运行 `start.bat` 启动图形化界面。

### 第三步：转换PPT
1. 点击"浏览..."选择输入PPT文件
2. 设置输出文件路径（可自动生成）
3. 选择图片格式（推荐PNG）
4. 点击"开始转换"

就这么简单！🎉

## 📋 系统要求

- **操作系统**：Windows 7/8/10/11
- **Python**：3.6或更高版本
- **办公软件**：Microsoft PowerPoint 2010或更高版本
- **依赖包**：pywin32, python-pptx, Pillow（自动安装）

## 📁 文件说明

### 🎯 主要文件
- **`start.bat`** - 🚀 **一键启动脚本（推荐）**
- **`install_dependencies.bat`** - 📦 **依赖包安装脚本**
- **`ppt_to_image_slides_gui.py`** - 🖥️ 图形化界面程序
- **`ppt_to_image_slides.py`** - ⌨️ 命令行程序

### 🔧 辅助文件
- **`advanced_menu.bat`** - 高级选项菜单
- **`test_dependencies.py`** - 依赖验证脚本
- **`requirements.txt`** - 依赖包列表
- **`examples.py`** - 使用示例

### 📚 文档文件
- **`README.md`** - 使用说明（本文件）
- **`CHANGELOG.md`** - 版本更新日志
- **`DEPENDENCIES.md`** - 详细依赖说明
- **`SUMMARY.md`** - 项目总结

## 💡 使用教程

### 🖥️ 图形化界面（推荐新手）

**方法一：最简单**
1. 双击 `start.bat` 文件
2. 等待图形界面启动

**方法二：手动启动**
```bash
python ppt_to_image_slides_gui.py
```

**使用步骤：**
1. **选择输入文件**：点击第一个"浏览..."按钮选择PPT文件
2. **设置输出路径**：会自动生成，也可手动修改
3. **选择图片格式**：建议使用PNG获得最佳质量
4. **开始转换**：点击"开始转换"按钮
5. **查看进度**：在日志区域查看实时进度
6. **完成提示**：转换完成后会弹出成功提示

### ⌨️ 命令行界面（高级用户）

**基本语法：**
```bash
python ppt_to_image_slides.py [输入文件] [选项]
```

**常用示例：**
```bash
# 最简单用法
python ppt_to_image_slides.py presentation.pptx

# 指定输出文件和格式
python ppt_to_image_slides.py "我的演示文稿.pptx" -o "输出文件.pptx" -f PNG

# 使用临时目录
python ppt_to_image_slides.py input.pptx -t temp_images

# 查看所有选项
python ppt_to_image_slides.py -h
```

**参数说明：**
- `input`: 输入PPT文件路径（必需）
- `-o, --output`: 输出PPT文件路径（可选）
- `-f, --format`: 图片格式PNG/JPG（可选，默认PNG）
- `-t, --temp-dir`: 临时目录（可选）

## 🔧 高级功能

### 批处理脚本
- **`start.bat`** - 直接启动图形界面
- **`install_dependencies.bat`** - 安装/更新依赖包
- **`advanced_menu.bat`** - 显示所有功能菜单

### 依赖管理
```bash
# 测试依赖是否正确安装
python test_dependencies.py

# 查看使用示例
python examples.py
```

## ❗ 注意事项

- **操作系统**：仅支持Windows系统
- **软件要求**：需要安装Microsoft PowerPoint（2010或更高版本）
- **权限要求**：首次运行可能需要管理员权限
- **网络连接**：安装依赖时需要网络连接

## 🐛 常见问题

### 安装问题
- **Python未找到**：确保Python已安装并添加到PATH环境变量
- **依赖安装失败**：尝试使用管理员权限运行`install_dependencies.bat`
- **pywin32安装问题**：可以尝试使用conda安装：`conda install pywin32`

### 运行问题
- **PowerPoint COM接口失败**：确保PowerPoint已正确安装且可以正常启动
- **文件路径错误**：确保文件路径正确，支持中文和空格
- **转换失败**：检查输入PPT文件是否可以正常打开

### 性能问题
- **转换速度慢**：大文件或复杂动画会影响转换速度
- **内存占用高**：处理大型PPT时会占用较多内存

## 🎯 最佳实践

1. **文件命名**：避免使用特殊字符，中文和空格是支持的
2. **文件位置**：建议将PPT文件放在用户有写权限的目录
3. **格式选择**：PNG质量更好，JPG文件更小
4. **备份原文件**：转换前建议备份原始PPT文件

## 📞 技术支持

如果遇到问题：
1. 首先运行 `test_dependencies.py` 检查环境
2. 查看 `CHANGELOG.md` 了解版本更新
3. 参考 `DEPENDENCIES.md` 获取详细技术信息

## 📄 许可证

本项目采用 MIT 许可证。详情请查看 LICENSE 文件。
