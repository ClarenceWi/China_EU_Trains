# coder:Clarence_William
# 2025年06月30日16时34分52秒
# soforso@163.com

import pdfplumber

# 请修改为您的PDF文件路径
pdf_path = "C:\\Users\\26765\\Desktop\\sxpdf\\中亚果斯12-饲用小麦粉-报关单-进口.pdf"
output_txt = "C:\\Users\\26765\\Desktop\\sxpdf\\报关单文本.txt"

with pdfplumber.open(pdf_path) as pdf:
    with open(output_txt, "w", encoding="utf-8") as f:
        for page_num, page in enumerate(pdf.pages, 1):
            f.write(f"===== 第 {page_num} 页 =====")
            f.write(page.extract_text())
            f.write("\n\n")

print(f"PDF文本已保存至：{output_txt}")