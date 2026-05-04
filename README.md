# LaTeX 导出 DOCX 桌面工具

这是一个用于毕业论文项目的 Python Tkinter 桌面工具。它会选择本地 LaTeX 项目文件夹，识别主 `.tex` 文件，并通过 Pandoc 导出 Word `.docx`。

## 功能

- 选择论文项目文件夹。
- 自动扫描可能的主 `.tex` 文件，优先选择包含 `\documentclass` 和 `\begin{document}` 的文件。
- 支持手动选择主 `.tex`、输出 `.docx`、可选 `reference.docx`、`.bib`、`.csl`。
- 启动时检测 Pandoc；如果未安装，会尝试通过 `pypandoc` 自动下载。
- 转换时以项目文件夹作为工作目录，尽量保留图片、`\input`、`\include` 等相对路径。
- 自动识别 TJUThesis 本科模板，预处理 `\makecover`、`\makeabstract`、`\figuremacro`、`\tablemacro` 等模板宏，补入封面、摘要和参考文献插入点。
- TJUThesis 模式下会执行 Word 后处理：插入可更新目录字段、套用天津大学模板关键样式、移动参考文献条目到“参考文献”标题下。
- 自动发现常用默认文件：项目中的 `reference.bib`、桌面 `编写latex` 目录下的 GB/T 7714 CSL 和天津大学 Word 模板。
- 导出文件统一保存到项目目录下的 `docx导出/`，Pandoc 日志保存到 `docx导出/logs/`，方便和 LaTeX 源文件分开管理。

## 安装与运行

```bash
python3 -m pip install -e .
python3 -m latex_docx_converter
```

也可以使用命令入口：

```bash
latex-docx-converter
```

如果自动下载 Pandoc 失败，可以手动安装 Pandoc：

- macOS: `brew install pandoc`
- Windows / Linux: 查看 https://pandoc.org/installing.html

## 使用建议

1. 选择包含论文主文件、图片、章节文件和参考文献文件的项目文件夹。
2. 确认主 `.tex` 文件是否正确。
3. 如需控制 Word 样式，选择一个 `reference.docx`；TJUThesis 项目通常会自动发现天津大学模板。
4. 如需参考文献，选择 `.bib` 文件；如需指定引用格式，选择 `.csl` 文件。若未选择，工具会尝试自动使用 GB/T 7714-2015 顺序编码制 CSL。
5. 点击“开始导出 DOCX”。

导出完成后，文件默认位于：

```text
论文项目/docx导出/主文件名.docx
论文项目/docx导出/logs/主文件名-pandoc-时间.log
```

## 已知限制

Pandoc 可以处理大量常规 LaTeX 内容，但对复杂宏包、自定义命令、交叉引用、特殊排版环境和院校模板中的深度格式不保证完全还原。推荐先用 `reference.docx` 控制 Word 样式，再对生成的 Word 文件做少量人工校对。

TJUThesis 兼容模式会把封面、摘要、目录、章标题和参考文献做预处理与 Word 后处理。目录是 Word TOC 字段，打开 Word 后可右键更新目录或按提示更新域。`\includepdf{独创性声明.pdf}` 这类 PDF 页面不能由 Pandoc 直接嵌入 DOCX，工具会在导出的 Word 中加入“独创性声明”提示页，最终声明页建议在 Word 中按学校模板插入或核对。

如果日志中出现 `Missing bibliography entries`，说明正文引用了 `.bib` 中不存在的 citation key。补齐 `.bib` 后重新导出即可。

## 测试

```bash
PYTHONPATH=src python3 -m unittest discover -s tests
```
