# Role
你是一位精通 Python 自动化办公的高级工程师，擅长使用 `python-docx` 库进行文档排版与重构。

# Task
我需要你根据我提供的素材，编写一个 Python 脚本，重新构建一份 Word 文档。

# Inputs
我为你提供了两个文件：
1. **skeleton.md (文本骨架)**: 包含了文档的文字内容和图片位置锚点。
   - 锚点格式为: `> **[插入图片]** ID: <<filename.png>>`
2. **VisualRef.pdf (视觉参考)**: 包含了图片 ID 对应的实际画面。

# Requirements (必须严格遵守)

### 1. 代码逻辑
- 编写一个完整的 Python 脚本。
- 使用 `python-docx` 库。
- **必须**使用 `doc.add_picture()` 函数插入图片。
- **必须**从本地路径 `media_source/` 读取图片（例如: `doc.add_picture('media_source/image1.png')`）。

### 2. 内容处理
- 读取 `skeleton.md` 的内容（你可以将文本硬编码在代码里，或者写代码读取它，为了代码简洁，建议直接将文本内容重构进代码的 `doc.add_paragraph` 中）。
- **识别 Markdown 语法**：
  - `#` -> 转换为 Heading 1 样式。
  - `##` -> 转换为 Heading 2 样式。
  - 正文 -> 转换为 Normal 样式。
- **锚点替换**：
  - 遇到 `<<filename>>` 锚点时，不要保留文字，而是插入对应的图片。
  - **严禁修改文件名**：必须完全照搬锚点中的文件名（如 `image3.png`），否则代码运行会报错。

### 3. 视觉与排版优化 (利用你的视觉理解能力)
- 参考 `VisualRef.pdf`：
  - 如果图片是**宽图**（如架构图、流程图），请在代码中设置 `width=Inches(6)` 以撑满页面。
  - 如果图片是**小图标**或**手机截图**，请适当缩小尺寸（如 `width=Inches(2)` 或 `Inches(3)`）。
  - 为每张图片添加一个简单的图注（Caption），内容基于你在 PDF 中看到的画面。

### 4. 样式美化
- 在代码开头设置全局中文字体为 "微软雅黑" (Microsoft YaHei)，西文为 "Arial"。
- 适当调整段落间距，使文档看起来专业、透气。

# Output
请直接输出 Python 代码块。不要解释代码逻辑，直接给我可执行的代码。