# coder:Clarence_William
# 2025年06月30日16时27分55秒
# soforso@163.com

import pdfplumber
import pandas as pd


def extract_single_pdf(pdf_path):
    """提取单个PDF文件中的6张报关单信息"""
    data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) != 6:
                print(f"错误：{pdf_path} 包含 {len(pdf.pages)} 页，不是6页")
                return []

            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if not text:
                    print(f"警告：第 {page_num} 页无文本内容")
                    continue

                # 提取关键信息
                info = {
                    "报关单号": "",
                    "境内收发货人": "",
                    "主要商品名称": "",
                    "贸易类型": "进口",
                    "币制": "美元",
                    "金额": 0.0,
                    "汇率": 7,
                    "折合人民币金额": 0.0
                }

                # 简单提取逻辑（根据实际PDF调整）
                if "报关单号" in text:
                    info["报关单号"] = text.split("报关单号")[1].split()[0]
                if "境内收发货人" in text:
                    info["境内收发货人"] = text.split("境内收发货人")[1].split()[0]
                if "商品名称" in text:
                    info["主要商品名称"] = text.split("商品名称")[1].split()[0]
                if "金额" in text:
                    amount_str = text.split("金额")[1].split()[0].replace(",", "")
                    info["金额"] = float(amount_str) if amount_str.isdigit() else 0
                    info["折合人民币金额"] = info["金额"] * 7

                data.append(info)
                print(f"已提取第 {page_num} 页信息")

        return data
    except Exception as e:
        print(f"处理PDF时出错：{str(e)}")
        return []


# 手动指定文件路径（请修改为您的实际路径）
pdf_path = "C:\\Users\\26765\\Desktop\\sxpdf\\中亚果斯12-饲用小麦粉-报关单-进口.pdf"
output_excel = "C:\\Users\\26765\\Desktop\\sxpdf\\报关单结果.xlsx"

# 执行提取
print(f"开始处理：{pdf_path}")
all_data = extract_single_pdf(pdf_path)

if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(output_excel, index=False)
    print(f"提取完成，结果已保存至：{output_excel}")
else:
    print("提取失败，未获取到数据")