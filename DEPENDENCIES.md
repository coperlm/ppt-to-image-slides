# 依赖包说明

本项目需要以下Python包：

## 核心依赖

### pywin32 (>=305)
- **用途**: Windows COM接口支持
- **功能**: 与Microsoft PowerPoint进行通信
- **重要性**: 必需，用于读取PPT文件和导出幻灯片

### python-pptx (>=0.6.21)
- **用途**: PowerPoint文件操作
- **功能**: 创建新的PPT文件，插入图片
- **重要性**: 必需，用于生成最终的图片PPT文件

### Pillow (>=9.0.0)
- **用途**: 图片处理
- **功能**: 读取图片尺寸，进行缩放计算
- **重要性**: 必需，用于图片尺寸适配

## 标准库依赖

以下是Python标准库中的模块，无需额外安装：

- **tkinter**: 图形用户界面（GUI）支持
- **os**: 文件系统操作
- **sys**: 系统功能
- **threading**: 多线程支持（用于GUI响应性）
- **tempfile**: 临时文件管理
- **shutil**: 文件操作工具
- **pathlib**: 现代路径操作
- **argparse**: 命令行参数解析

## 安装方法

### 方法1：使用requirements.txt（推荐）
```bash
pip install -r requirements.txt
```

### 方法2：单独安装
```bash
pip install pywin32 python-pptx Pillow
```

### 方法3：指定版本安装
```bash
pip install pywin32>=305 python-pptx>=0.6.21 Pillow>=9.0.0
```

## 兼容性说明

- **操作系统**: Windows（因为需要PowerPoint COM接口）
- **Python版本**: 3.6+（推荐3.8+）
- **PowerPoint版本**: Microsoft Office 2010或更高版本
- **图形界面**: 需要支持tkinter（大多数Python安装都包含）

## 常见问题

### 1. pywin32安装失败
```bash
# 尝试使用conda安装
conda install pywin32

# 或者升级pip后重试
python -m pip install --upgrade pip
pip install pywin32
```

### 2. tkinter不可用
- Windows: 重新安装Python，确保勾选"tcl/tk and IDLE"选项
- Linux: `sudo apt-get install python3-tk`

### 3. PowerPoint COM接口错误
- 确保已安装Microsoft PowerPoint
- 以管理员权限运行命令提示符，执行一次脚本
- 检查Windows防火墙或杀毒软件设置

## 验证安装

运行依赖测试脚本：
```bash
python test_dependencies.py
```

如果所有测试通过，说明环境配置正确。
