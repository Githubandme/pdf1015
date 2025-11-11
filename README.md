# PDF 页面缩放工具

这个 Python 程序可以批量处理 PDF 文件，将页面按比例缩放到指定尺寸范围内。

## 功能特点

- 批量处理多个 PDF 文件
- 按比例缩放页面，保持原有宽高比
- 确保缩放后的页面宽度不超过 10cm，高度不超过 15cm
- 支持任意尺寸的原始 PDF 页面
- 提供多种用户界面选择

### 缩放规则

程序按照以下规则进行缩放：

1. 按比例缩放页面
2. 缩放后宽度不超过 10cm，高度不超过 15cm
3. 具体示例：
   - 100*150cm 转换后是 10*15cm
   - 100*1500cm 转换后为 1*15cm
   - 1000*150cm 转换后为 10*1.5cm

## 安装依赖

在运行程序之前，请确保已安装所需依赖：

```bash
pip install -r requirements.txt
```

或者手动安装：

```bash
pip install PyPDF2 reportlab ttkbootstrap
```

## 使用方法

程序提供三个版本：

### 1. 命令行版本

```bash
python pdf_resizer.py
```

按照提示操作即可。

### 2. 图形界面版本（传统）

```bash
python pdf_resizer_gui.py
```

图形界面版本具有以下特点：
- 直观的文件夹选择界面
- 可视化参数设置
- 实时处理日志显示
- 一键创建示例文件

### 3. 现代化界面版本（推荐）

```bash
python pdf_resizer_modern.py
```

现代化界面版本具有以下增强功能：
- 使用 ttkbootstrap 创建的现代化界面
- 多种主题可选（cosmo, flatly, litera 等）
- 更美观的控件和布局
- 状态栏显示当前操作状态
- 卡片式信息展示

## 打包为可执行文件

程序可以打包为 Windows 可执行文件 (.exe)，便于在没有 Python 环境的计算机上运行。

### 打包方法

1. 安装打包工具：
   ```bash
   pip install pyinstaller
   ```

2. 使用提供的脚本打包：
   - 运行 `python build.py` 并根据提示选择要打包的版本
   - 或者直接运行对应的批处理文件：
     - `build_modern.bat` - 打包现代化GUI版本
     - `build_gui.bat` - 打包传统GUI版本
     - `build_cli.bat` - 打包命令行版本
     - `build_all.bat` - 打包所有版本

3. 打包完成后，可执行文件将位于 `dist` 文件夹中

### 打包配置说明

- 排除了不必要的依赖（如 matplotlib, numpy, scipy 等），减小打包文件大小
- 包含了必要的依赖（PyPDF2, reportlab, ttkbootstrap）
- GUI版本不显示控制台窗口，命令行版本显示控制台窗口

## 示例

程序支持创建示例文件以便测试，在运行时选择创建示例文件即可。

## 技术说明

- 使用 `PyPDF2` 库处理 PDF 文件
- 使用 `reportlab` 库创建示例文件
- 使用 `ttkbootstrap` 库创建现代化界面
- 页面尺寸单位转换：1cm = 28.3464567 points