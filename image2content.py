# sudo apt update
# sudo apt install tesseract-ocr tesseract-ocr-chi-sim -y
# pip install pytesseract Pillow

import pytesseract
from PIL import Image
import os
from pathlib import Path

def images_to_single_markdown(image_folder, output_file):
    # 查找所有常用格式图片
    extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.webp')
    image_files = [f for f in os.listdir(image_folder) if f.lower().endswith(extensions)]
    image_files.sort()  # 按文件名排序

    if not image_files:
        print("未在目录下找到图片文件。")
        return

    with open(output_file, 'w', encoding='utf-8') as md:
        for idx, filename in enumerate(image_files, 1):
            img_path = os.path.join(image_folder, filename)
            print(f"[{idx}/{len(image_files)}] 正在处理: {filename}...")
            
            try:
                # 使用 Tesseract 提取文字（中英双语）
                text = pytesseract.image_to_string(Image.open(img_path), lang='chi_sim+eng')
                
                # 写入 Markdown
                md.write(f"### 图片名称: `{filename}`\n\n")
                md.write("#### 识别内容:\n")
                md.write(f"```text\n{text.strip()}\n```\n\n")
                md.write("---\n\n")
            except Exception as e:
                md.write(f"### 图片名称: `{filename}`\n\n> [错误] 无法识别该图片: {e}\n\n---\n\n")

    print(f"\n✅ 处理完成！请查看: {output_file}")

if __name__ == "__main__":
    # 指定你的图片文件夹路径
    img_dir = "../images" 
    output_md = "提升认知.md"
    
    if not os.path.exists(img_dir):
        os.makedirs(img_dir)
        print(f"请将图片放入 {img_dir} 文件夹后重新运行。")
    else:
        images_to_single_markdown(img_dir, output_md)
