# Excel 表格数据匹配工具

一个面向日常表格处理的桌面工具，用来把两个 Excel 或 CSV 文件按指定列进行匹配，并导出新的结果文件。

![匹配设置界面](docs/images/match-step.png)

## 主要功能

- 支持 Excel 和 CSV 文件。
- 支持选择工作表并预览数据。
- 支持一组或多组匹配列。
- 支持选择最终要导出的列。
- 支持 CSV 编码切换。
- 支持按匹配列去重或保留全部结果。
- 支持 macOS 和 Windows 打包使用。

## 使用流程

1. 在第一步选择“匹配源文件”和“被匹配文件”。
2. 选择两边用于匹配的列。
3. 点击“下一步”。
4. 选择最终要导出的列和输出文件名。
5. 点击“开始匹配”生成结果文件。

![输出设置界面](docs/images/output-step.png)

## 下载

最新版本发布后，可以从下面地址下载：

- [下载 macOS 版本](https://github.com/keru-s/excel-table-data-matcher/releases/latest/download/Excel-data-match-tool-macOS.zip)
- [下载 Windows 版本](https://github.com/keru-s/excel-table-data-matcher/releases/latest/download/Excel-data-match-tool-Windows.zip)

如果链接暂时不可用，说明还没有创建正式版本。推送版本标签后，系统会自动打包并生成下载文件。

## Windows 使用说明

1. 下载 `Excel-data-match-tool-Windows.zip`。
2. 右键压缩包，选择“全部解压”。
3. 打开解压后的文件夹。
4. 双击 `Excel数据匹配工具.exe` 启动软件。
5. 不要删除 `_internal` 文件夹，它是软件运行所需文件。

如果 Windows 提示“无法确认发布者”或“Windows 已保护你的电脑”，这是因为当前版本没有代码签名。点击“更多信息”，再选择“仍要运行”。

如果双击没有反应，请确认：

- 已经完整解压，不是在压缩包里直接运行。
- `Excel数据匹配工具.exe` 和 `_internal` 文件夹在同一个目录。
- 没有被杀毒软件隔离。

## 本地运行

需要先安装 Python 3.10 或更高版本。

```bash
uv venv
uv pip install -r requirements.txt
uv run python start_excel_tool.py
```

## 本地测试

```bash
uv run python -m unittest tests/test_excel_core.py
uv run python -m unittest tests/test_gui_regression.py
```

## 发布版本

创建并推送版本标签后，GitHub Actions 会自动打包 macOS 和 Windows 版本，并发布到 GitHub Release。

```bash
git tag v1.0.0
git push origin v1.0.0
```
