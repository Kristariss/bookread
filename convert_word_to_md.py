import os
from docx import Document

def convert_heading(paragraph):
    heading_map = {'Heading 1': '#', 'Heading 2': '##', 'Heading 3': '###'}
    return f"{heading_map.get(paragraph.style.name, '')} {paragraph.text}\n"

# 新增样式转换函数
def convert_formatting(run):
    formatting = []
    if run.bold: formatting.append('**')
    if run.italic: formatting.append('_')
    return ''.join(formatting)

# 更新转换函数
def convert_docx_to_md(input_path, output_path):
    doc = Document(input_path)
    with open(output_path, 'w', encoding='utf-8') as f:
        for para in doc.paragraphs:
            if para.style.name.startswith('Heading'):
                f.write(convert_heading(para))
            elif para.text.strip():
                # 处理文本格式和列表
                md_text = para.text
                
                # 转换格式
                if para.style.name.startswith('List'):
                    md_text = f"* {md_text}"
                
                # 处理粗体斜体
                formatted_runs = [f"{convert_formatting(run)}{run.text}{convert_formatting(run)[::-1]}" 
                                for run in para.runs]
                md_text = ''.join(formatted_runs)
                
                f.write(f"{md_text}\n\n")

input_dir = os.path.expanduser('~/Desktop/颜拉图儿的笔记')
os.makedirs(f"{input_dir}/markdown_notes", exist_ok=True)

# for filename in os.listdir(input_dir):
#     if filename.endswith('.docx'):
#         output_name = filename.replace('.docx', '.md')
#         convert_docx_to_md(
#             f"{input_dir}/{filename}",
#             f"{input_dir}/markdown_notes/{output_name}"
#         )