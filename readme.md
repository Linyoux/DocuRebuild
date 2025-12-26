# docx_ai_reconstruct 📄➡️🤖➡️📝

> **排版已死，AI 重生。**
> 基于“解耦与重构”范式的 AI 文档自动化重建工具链。

## 📖 背景与哲学

我们正处于一个 AI 能力爆发但工具链尚未跟上的尴尬期。目前的 LLM（如 GPT-4o, Claude 3.5）在直接修改二进制 `.docx` 文档时，往往会破坏原有格式，甚至产生幻觉。

**docx_ai_reconstruct** 提出了一种全新的工作流范式：**解耦与重构 (Decoupling & Reconstruction)**。

我们不再强求 AI 去“修改”文档，而是将文档拆解为 AI 最擅长的两个形态：

1. **纯文本骨架**：供 AI 理解逻辑。
2. **视觉参考 (Visual Reference)**：供 AI 视觉模型“看”图。

最后，由 AI 编写代码（Python/python-docx）将二者重新“缝合”，从而实现 100% 可控、格式完美的文档生成。

## ✨ 核心功能

本项目目前实现了该流水线的 **第一阶段：自动化预处理与视觉资产提取**。

* **🕵️‍♂️ 深度解包**：自动解压 `.docx` 容器，无损提取所有媒体资源（Images）。
* **🔢 自然排序**：实现了符合人类直觉的文件名排序算法（`image2.png` 排在 `image10.png` 之前），确保上下文顺序正确。
* **👁️ 视觉参考生成**：自动生成《视觉参考 PDF》。每一页包含一张大图及其文件名 ID。
* *作用*：直接投喂给 GPT-4o/Claude，让 AI 能够建立 `文件名 <-> 图片内容` 的语义索引。


* **📦 智能分块**：针对超多图文档，自动将 PDF 切分为多个部分（默认 100 页/卷），防止超出 AI 上下文限制。

## 🛠️ 安装与依赖

本项目基于 Python 开发，核心依赖为 `Pillow` (PIL)。


2. 安装依赖：
```bash
pip install pillow

```



## 🚀 快速开始

### 1. 准备文档

将你需要处理的 Word 文档（例如 `report.docx`）准备好。

### 2. 运行提取脚本

打开 `deconstruct.py`，在底部修改输入文件路径（后续版本将支持命令行参数）：

```python
if __name__ == "__main__":
    # 修改这里
    input_docx = "你的文档路径.docx" 
    output_folder = "./pipeline_output"
    ...

```

运行脚本：

```bash
python deconstruct.py

```

### 3. 查看输出

脚本将在 `pipeline_output` 目录下生成如下结构：

```text
pipeline_output/
└── 你的文档名/
    ├── media_source/       # [资源库] 提取出的所有原始图片文件
    └── visual_refs/        # [给AI看] 视觉参考 PDF (例如: Report_VisualRef_Part1.pdf)

```

## 🤖 AI 工作流指南 (The Workflow)

要完成完整的“文档重建”，请按照以下步骤操作：

**Step 1: 视觉索引 (Vision)**
将生成的 `VisualRef.pdf` 上传给 AI (GPT-4o/Claude)，并输入 Prompt：

> “这是一个文档的图片资源表。请读取每一页，提取图片 ID 和图片内容的简要描述，输出为一个 JSON 列表。”

**Step 2: 逻辑缝合 (Coding)**
将文档的纯文本内容（你需要另外提取文本，或等待本项目后续更新）和上一步得到的 JSON 投喂给 AI，并输入 Prompt：

> “请写一个 Python 脚本（使用 python-docx），根据文本内容重写文档。当你遇到图片占位符时，根据 JSON 中的描述找到对应的图片 ID，将其插入文档，并根据图片内容自动生成图注。”

**Step 3: 执行 (Execution)**
运行 AI 生成的 Python 脚本，获得完美排版的新文档。

## 📅 Roadmap

* [x] **v0.1**: `.docx` 媒体提取与《视觉参考 PDF》生成。
* [ ] **v0.2**: 文本层提取与“锚点植入” (在文本中自动插入 `<<IMG_ANCHOR_01>>` 占位符)。
* [ ] **v0.3**: 提供标准的 Prompt 模板库，方便用户直接复制给 LLM。
* [ ] **v0.4**: 命令行工具 (CLI) 支持。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！如果你对“文档即代码” (Document as Code) 有新的想法，请务必分享。

