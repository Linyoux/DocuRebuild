import os
import zipfile
import re
import shutil
from PIL import Image, ImageDraw, ImageFont
from math import ceil
from docx import Document
from docx.oxml.ns import qn

# ==========================================
# 模块 1: 辅助工具 (排序与图像处理)
# ==========================================

def natural_sort_key(s):
    """自然排序：确保 image2 在 image10 之前"""
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', s)]

def process_image_for_ai(img_path):
    """
    图像预处理核心：
    1. 修正透明背景 (解决黑白图在 AI 面前'隐身'的问题)
    2. 转换为 RGB
    """
    try:
        img = Image.open(img_path)
        # 如果有透明通道 (RGBA) 或 P 模式，转换为白色背景的 RGB
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            # 使用 alpha 通道作为掩码进行合成
            background.paste(img, mask=img.split()[3]) 
            return background
        else:
            return img.convert('RGB')
    except Exception as e:
        print(f"Warning: 图像处理出错 {img_path} - {e}")
        return None

def create_visual_reference_pdf(image_files, media_dir, output_pdf_path):
    """生成视觉参考 PDF (已升级：增加透明度处理)"""
    pdf_pages = []
    page_width, page_height = 595, 842 # A4
    margin = 50
    try:
        font = ImageFont.truetype("arial.ttf", 24)
    except:
        font = ImageFont.load_default()

    for filename in image_files:
        img_path = os.path.join(media_dir, filename)
        
        # --- 升级点：调用预处理函数 ---
        src_img = process_image_for_ai(img_path)
        if src_img is None: continue

        # 创建页面
        page = Image.new('RGB', (page_width, page_height), (255, 255, 255))
        draw = ImageDraw.Draw(page)
        
        # 写入 ID
        text = f"ID: {filename}"
        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        draw.text(((page_width - text_w) / 2, margin), text, fill=(0, 0, 0), font=font)
        
        # 缩放图片
        max_img_w = page_width - 2 * margin
        max_img_h = page_height - 3 * margin - text_h
        src_img.thumbnail((max_img_w, max_img_h), Image.Resampling.LANCZOS)
        
        # 居中粘贴
        img_x = int((page_width - src_img.width) / 2)
        img_y = int(margin + text_h + 20)
        page.paste(src_img, (img_x, img_y))
        
        pdf_pages.append(page)

    if pdf_pages:
        pdf_pages[0].save(output_pdf_path, "PDF", resolution=100.0, save_all=True, append_images=pdf_pages[1:])
        print(f"   --> 生成视觉参考: {os.path.basename(output_pdf_path)}")

# ==========================================
# 模块 2: 骨架提取器 (Text Skeleton)
# ==========================================

class SkeletonExtractor:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.doc = Document(docx_path)
        self.rels = self.doc.part.rels
        self.rId_to_filename = {}
        self._map_rels()

    def _map_rels(self):
        """
        建立 rId -> 实际文件名的映射表
        Word 内部通过 rId (如 rId7) 引用图片，而不是直接用文件名。
        """
        for rel in self.rels.values():
            if "image" in rel.target_ref:
                # target_ref 通常是 'media/image1.png'
                filename = os.path.basename(rel.target_ref)
                self.rId_to_filename[rel.rId] = filename

    def extract_to_markdown(self):
        """
        遍历文档段落，生成带 <<IMG_xxx>> 锚点的 Markdown
        """
        md_lines = []
        
        for para in self.doc.paragraphs:
            text = para.text.strip()
            style_name = para.style.name

            # 1. 简单的样式映射
            prefix = ""
            if style_name.startswith('Heading 1'): prefix = "# "
            elif style_name.startswith('Heading 2'): prefix = "## "
            elif style_name.startswith('Heading 3'): prefix = "### "
            elif "List" in style_name: prefix = "- "
            
            # 2. 检查段落中的 XML 是否包含图片引用 (blip)
            # 这是一个底层操作，寻找 <a:blip r:embed="rIdX">
            if 'graphicData' in para._p.xml:
                for rId, filename in self.rId_to_filename.items():
                    # 如果该段落的 XML 源码中包含这个 rId
                    if f'r:embed="{rId}"' in para._p.xml:
                        # 插入锚点 (这是给 AI 看的逻辑指针)
                        md_lines.append(f"\n> **[插入图片]** ID: <<{filename}>>\n")
            
            if text:
                md_lines.append(f"{prefix}{text}")
                
        return "\n\n".join(md_lines)

# ==========================================
# 主程序
# ==========================================

def main(input_docx, output_folder):
    if not os.path.exists(input_docx):
        print(f"❌ 错误: 找不到文件 {input_docx}")
        return

    doc_name = os.path.splitext(os.path.basename(input_docx))[0]
    base_output_dir = os.path.join(output_folder, doc_name)
    media_dir = os.path.join(base_output_dir, "media_source")
    visual_ref_dir = os.path.join(base_output_dir, "visual_refs")
    
    # 清理重建目录
    if os.path.exists(base_output_dir):
        shutil.rmtree(base_output_dir)
    os.makedirs(media_dir)
    os.makedirs(visual_ref_dir)

    print(f"🚀 开始拆解: {doc_name}")

    # Step 1: 物理提取图片 (使用 ZipFile 确保无损)
    print("   ...正在解压媒体资源")
    with zipfile.ZipFile(input_docx, 'r') as z:
        for file_info in z.infolist():
            if file_info.filename.startswith('word/media/'):
                z.extract(file_info, media_dir)
    
    # 移动文件到 media_source 根目录并清理空文件夹
    actual_media_dir = os.path.join(media_dir, 'word', 'media')
    if os.path.exists(actual_media_dir):
        files = os.listdir(actual_media_dir)
        valid_exts = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff'}
        image_files = [f for f in files if os.path.splitext(f)[1].lower() in valid_exts]
        image_files.sort(key=natural_sort_key) # 排序

        for f in image_files:
            shutil.move(os.path.join(actual_media_dir, f), os.path.join(media_dir, f))
        shutil.rmtree(os.path.join(media_dir, 'word')) # 删除空壳
    else:
        image_files = []
        print("   ⚠️ 文档中未发现图片")

    # Step 2: 生成视觉参考 PDF
    if image_files:
        print("   ...正在生成视觉参考 PDF")
        CHUNK_SIZE = 50 # 减小一点，防止 PDF 过大
        total_chunks = ceil(len(image_files) / CHUNK_SIZE)
        for i in range(total_chunks):
            chunk = image_files[i*CHUNK_SIZE : (i+1)*CHUNK_SIZE]
            pdf_path = os.path.join(visual_ref_dir, f"{doc_name}_VisualRef_Part{i+1}.pdf")
            create_visual_reference_pdf(chunk, media_dir, pdf_path)

    # Step 3: 提取文本骨架 (Markdown)
    print("   ...正在生成文本骨架 (Markdown)")
    extractor = SkeletonExtractor(input_docx)
    skeleton_md = extractor.extract_to_markdown()
    
    skeleton_path = os.path.join(base_output_dir, "skeleton.md")
    with open(skeleton_path, "w", encoding="utf-8") as f:
        f.write(f"# 文档骨架: {doc_name}\n\n")
        f.write("> 此文档由 AI 自动拆解。<<IMG_...>> 为图片占位符。\n\n")
        f.write(skeleton_md)

    print(f"✅ 任务完成! 输出目录: {base_output_dir}")
    print(f"   -> 骨架文件: skeleton.md")
    print(f"   -> 图片资源: media_source/")

if __name__ == "__main__":
    # 修改这里为你的文件名
    input_file = "input.docx" 

    main(input_file, "./pipeline_output")
