# Paper Format Skill

一句话介绍：将纯文本论文草稿一键排版为符合期刊标准格式的 Word 文档（支持等宽双栏版式）。

---

## Features

- **自动识别论文结构**：中文标题、英文标题、摘要（中英文）、关键词、引言、正文章节、参考文献
- **精确格式控制**：基于实测期刊论文提取的格式规范（字号、字体、行距、段前段后距、首行缩进）
- **等宽双栏排版**：引言开始至文档末尾自动设置为期刊标准等宽双栏，栏间距 0.85 cm；标题/摘要/关键词保持单栏
- **跨平台兼容**：路径优先读环境变量，支持 Windows / macOS / Linux
- **零配置运行**：默认模板即为可用的完整样式定义

---

## Demo

输入纯文本草稿：

```
基于声振图像的转辙机预测性维护方法
Multimodal Fusion Method for Switch Machine Prognostics Based on Acoustic-Vibration Images
关键词：转辙机；预测性维护；声振图像
Keywords: switch machine; predictive maintenance; acoustic-vibration images
摘要
本文提出一种...
Abstract
This paper proposes...
1 引言
...
参考文献
[1] 张三. 机械工程学报, 2020.
```

输出：符合排版标准的 `.docx` 文件，双栏显示。

---

## Quick Start

### 1. 准备输入文件

在任意文本编辑器中按以下顺序组织论文草稿，以换行分隔各部分：

```
中文标题（第一行）
英文标题（全文无中文，在关键词：之前）
关键词：xxx；xxx；xxx
Keywords: xxx; xxx; xxx
摘要
中文摘要正文
Abstract
英文摘要正文
1 一级章节标题
1.1 二级章节标题
正文段落...
参考文献
[1] 参考文献条目1
[2] 参考文献条目2
```

### 2. 运行格式化

```python
from paper_format import format_paper  # 如 skill 已安装为 Python 包

input_file = 'paper_input.txt'   # 你的草稿路径
output_name = '我的论文.docx'
out_path = format_paper(input_file, output_name)
print(f'已生成: {out_path}')
```

### 3. 打开输出文件

用 Word 打开 `~/Desktop/论文/output/我的论文.docx`，确认双栏显示正常。

---

## Installation

### 依赖

- Python 3.8+
- python-docx

```bash
pip install python-docx
```

### 目录结构

```
paper-format/
├── README.md
├── SKILL.md                      # Skill 定义文件（含完整执行代码）
├── 论文/
│   ├── 论文1_明杆闸阀丝杠开度识别.docx   # 默认模板（含全部样式定义）
│   ├── 论文2_变频柜预测性维护.docx       # 备选模板
│   ├── 论文3_客车栓电缆监测.docx         # 备选模板
│   └── backup/                          # 原始模板备份
```

### 配置环境变量（可选）

| 变量 | 作用 | 默认值 |
|------|------|--------|
| `PAPER_FORMAT_TEMPLATE` | 指定模板文件 | `{skill_dir}/论文/论文1_...docx` |
| `PAPER_FORMAT_OUTPUT` | 指定输出目录 | `~/Desktop/论文/output/` |
| `PAPER_FORMAT_INPUT` | 指定输入文件 | `temp/paper_input.txt` |

---

## Usage

### 作为 Claude Code Skill 使用

将 `paper-format` 目录放入 `~/.claude/skills/`，Claude Code 会自动识别。

触发词：**论文格式化**、**格式化论文**、**排版论文**、**生成论文 Word**

### 直接调用函数

```python
from docx import Document
from paper_format import format_paper

# 方式1：使用环境变量配置的输入文件
doc = Document()  # 创建空文档
out_path = format_paper(
    os.environ.get('PAPER_FORMAT_INPUT', 'paper_input.txt'),
    'output.docx'
)

# 方式2：自定义路径
import os
os.environ['PAPER_FORMAT_TEMPLATE'] = '/path/to/my_template.docx'
os.environ['PAPER_FORMAT_OUTPUT'] = '/path/to/output/'
out_path = format_paper('/path/to/input.txt', 'output.docx')
```

---

## Project Architecture

```
输入纯文本
    │
    ▼
build_blocks()          ── 解析论文结构，识别各部分类型
    │
    ▼
复制模板 DOCX
    │
    ▼
清空模板段落 / 删除残留表格
    │
    ▼
按 new_order 顺序写入新段落（逐 block 应用样式）
    │
    ▼
set_page_setup()        ── A4 页面，边距 2.15/1.95 cm
    │
    ▼
inject sectPr (XML)     ── 插入连续分节符，标记双栏节
    │
    ▼
doc.save()              ── 输出 .docx
```

**关键技术点**：

- Word 双栏通过 `<w:sectPr>` 中的 `<w:cols>` 控制，`w:num="2"`, `w:space="480"`（≈0.85 cm）
- 分节符嵌入在 intro 前一段的 `<w:pPr>` 中（连续分节，不换页）
- 样式通过 `apply_rPr()` 直接注入 XML，绕过 python-docx 样式API的默认行为

---

## Project Structure

```
paper-format/
├── README.md
├── demo                #输出示例
├── SKILL.md            # Skill 定义（含 format_paper 完整源码）
└── 论文/
    ├── 论文1_明杆闸阀丝杠开度识别.docx   # 默认模板（已注"仅供格式参考"）
    ├── 论文2_变频柜预测性维护.docx
    ├── 论文3_客车栓电缆监测.docx
    └── backup/          # 原始模板备份目录
```

---

## Roadmap

- [ ] 支持更多中文期刊模板（《计算机学报》《自动化学报》等）
- [ ] 图表标题自动编号
- [ ] 英文摘要/关键词自动生成
- [ ] 命令行界面（CLI）
- [ ] 多语言界面支持

---

## Limitations

1. **输入格式严格**：草稿必须严格按指定顺序排列，否则解析可能失败
2. **仅支持 .docx**：不输出 PDF 或 LaTeX
3. **模板依赖**：输出格式完全依赖模板中预定义的样式，不支持从零创建格式
4. **双栏为连续分节**：不支持每栏独立页眉页脚（期刊通常也只需要统一页眉）
5. **不支持图表**：草稿中的图片/表格需要手动插入到输出的 Word 文档中

---

## License

MIT License - 可自由使用、修改、分发。
