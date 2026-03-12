# Excel工具 Windows 打包说明

## 📋 打包前准备

### 1. 系统要求
- Windows 10/11 (推荐)
- Python 3.8+ (推荐 3.9-3.11)
- 至少 2GB 可用磁盘空间

### 2. 需要传输的文件
将以下文件从 Mac 传输到 Windows 系统：

**必需文件：**
- `excel_tool.py` - 主程序文件
- `start_excel_tool.py` - 启动脚本
- `requirements.txt` - 依赖列表
- `build_config.py` - 打包配置脚本
- `install_dependencies.py` - 依赖安装脚本（如果有）

**可选文件：**
- `resource/` 目录 - 测试数据文件
- `doc/` 目录 - 文档文件

## 🚀 打包步骤

### 方法一：使用 auto-py-to-exe（推荐新手）

#### 1. 安装依赖
```bash
# 创建虚拟环境（推荐）
python -m venv excel_tool_env
excel_tool_env\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

#### 2. 启动 auto-py-to-exe
```bash
auto-py-to-exe
```

#### 3. 配置参数
在 auto-py-to-exe 界面中设置：

**基本设置：**
- Script Location: 选择 `start_excel_tool.py`
- Onefile: 建议选择 "One Directory"
- Console Window: 选择 "Window Based (hide the console)"

**高级设置：**
- Output Directory: 选择输出目录
- Additional Files: 无需添加
- Hidden Imports: 添加以下模块
  ```
  PyQt6.QtCore
  PyQt6.QtWidgets
  PyQt6.QtGui
  pandas._libs.tslibs.timedeltas
  pandas._libs.tslibs.np_datetime
  pandas._libs.tslibs.nattype
  numpy.core._methods
  xlrd.xldate
  openpyxl.cell
  psutil._psutil_windows
  chardet
  ```

#### 4. 开始打包
点击 "CONVERT .PY TO .EXE" 按钮开始打包

### 方法二：使用 PyInstaller 命令行

#### 1. 安装依赖
```bash
pip install -r requirements.txt
```

#### 2. 生成打包配置
```bash
python build_config.py
```

#### 3. 执行打包
```bash
# 使用生成的命令（从 build_config.py 输出复制）
pyinstaller --windowed --clean --noconfirm --name "Excel数据匹配工具" --hidden-import PyQt6.QtCore --hidden-import PyQt6.QtWidgets --hidden-import PyQt6.QtGui --hidden-import pandas._libs.tslibs.timedeltas --hidden-import pandas._libs.tslibs.np_datetime --hidden-import pandas._libs.tslibs.nattype --hidden-import pandas._libs.properties --hidden-import numpy.core._methods --hidden-import numpy.lib.format --hidden-import xlrd.xldate --hidden-import openpyxl.cell --hidden-import openpyxl.workbook --hidden-import psutil._psutil_windows --hidden-import chardet --exclude-module matplotlib --exclude-module tkinter --exclude-module test --exclude-module unittest --exclude-module pydoc --exclude-module doctest start_excel_tool.py

# 或使用 spec 文件
pyinstaller "Excel数据匹配工具.spec"
```

## 📁 输出文件

打包完成后，在 `dist` 目录下会生成：
- `Excel数据匹配工具.exe` - 可执行文件

## ⚠️ 常见问题及解决方案

### 1. 打包失败：缺少模块
**解决方案：** 在 Hidden Imports 中添加缺少的模块

### 2. exe 文件过大（>200MB）
**解决方案：**
- 使用虚拟环境，只安装必需依赖
- 在 excludes 中添加不需要的模块
- 考虑使用 `--exclude-module` 排除大型库

### 3. 运行时出现 DLL 错误
**解决方案：**
- 确保 Windows 系统已安装 Visual C++ Redistributable
- 使用 `--collect-all` 参数包含所有相关文件

### 4. PyQt6 相关错误
**解决方案：**
- 确保所有 PyQt6 子模块都在 Hidden Imports 中
- 添加 `--collect-submodules PyQt6`

### 5. 大文件处理表现不稳定
**解决方案：**
- 优先使用 `One Directory` 打包方式，减少启动解包开销
- 确保包含 `chardet`，避免 CSV 编码检测在打包后失效

## 🧪 测试建议

1. **功能测试：** 在不同 Windows 版本上测试所有功能
2. **性能测试：** 测试大文件处理性能
3. **兼容性测试：** 在没有安装 Python 的机器上测试
4. **路径测试：** 测试中文路径和特殊字符路径

## 📊 优化建议

### 减小文件大小
```bash
# 使用 UPX 压缩（可选）
pip install upx-ucl
pyinstaller --upx-dir /path/to/upx your_spec_file.spec
```

### 提高启动速度
- 使用 `--onedir` 而不是 `--onefile`（文件大但启动快）
- 优化导入语句，延迟加载大型库

## 🔧 调试技巧

### 查看详细错误信息
```bash
# 临时启用控制台窗口
pyinstaller --onefile --console start_excel_tool.py
```

### 检查依赖关系
```bash
# 分析依赖
pyi-archive_viewer Excel数据匹配工具.exe
```

## 📝 版本信息

- 工具版本：v2.0 (性能优化版)
- 支持的文件格式：Excel (.xlsx, .xls), CSV
- 主要功能：数据匹配、列对比、结果导出
- 性能特性：异步预览、流式导出、低内存占用、大文件支持

---

**注意：** 首次打包可能需要较长时间（10-30分钟），请耐心等待。建议在网络良好的环境下进行打包，确保所有依赖都能正确下载。
