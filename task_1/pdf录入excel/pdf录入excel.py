# coder:Clarence_William
# 2025年06月30日11时16分14秒
# soforso@163.com

import pdfplumber
import pandas as pd
import re
import os
from decimal import Decimal, ROUND_HALF_UP

# 关键字段配置（根据报关单结构调整）
FIELDS_CONFIG = {
    "海关编号": r"海关编号\s*[:：]?\s*(\w+)",  # 原预录入编号
    "境内收货人": r"收货人\s*[:：]?\s*([^\n]+)",
    "运输方式": r"运输方式\s*[:：]?\s*([^\n]+)",
    "启运国": r"启运国\(地区\)\s*[:：]?\s*([^\n]+)",
    "进口日期": r"进口日期\s*[:：]?\s*(\d{8})",
    "入境口岸": r"入境口岸\s*[:：]?\s*([^\n]+)",
    "商品名称及规格型号": r"商品名称、规格型号\s*([\s\S]*?)(?:\d+\.\d{2,})",  # 匹配到价格前的内容
    "总价": r"总价\s+([\d,]+\.\d{2})"  # 匹配总价金额
}


def extract_product_name(full_desc):
    """从商品名称及规格型号中提取最像商品名称的部分"""
    # 策略1：取第一个换行符前的内容
    if '\n' in full_desc:
        return full_desc.split('\n')[0].strip()

    # 策略2：取第一个中文短语（排除规格说明）
    chinese_phrase = re.search(r"[\u4e00-\u9fa5]+[^\d]*", full_desc)
    if chinese_phrase:
        return chinese_phrase.group(0).strip()

    # 策略3：直接返回前30个字符
    return full_desc[:30].strip()


def format_currency(value):
    """格式化金额为千分位表示"""
    try:
        num = float(value.replace(',', ''))
        return "{:,.2f}".format(num)
    except:
        return value


def extract_customs_data(pdf_path):
    """从PDF提取报关单数据"""
    all_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # 增强文本提取
            text = page.extract_text(x_tolerance=3, y_tolerance=3, layout=True)
            if not text:
                continue

            # 分割单张报关单（根据实际分隔符调整）
            declarations = re.split(r"-{10,}|\*{10,}", text)  # 常见分隔符

            for dec_text in declarations:
                if not dec_text.strip():
                    continue

                record = {}
                # 提取基本字段
                for field, pattern in FIELDS_CONFIG.items():
                    match = re.search(pattern, dec_text, re.DOTALL)
                    record[field] = match.group(1).strip() if match else "N/A"

                # 特殊字段处理
                record["商品名称"] = extract_product_name(record.get("商品名称及规格型号", ""))
                record["贸易类型"] = "进口"
                record["币制"] = "美元"

                # 金额和汇率处理
                amount_str = record.get("总价", "0").replace(',', '')
                try:
                    amount = Decimal(amount_str)
                    record["金额"] = format_currency(amount_str)  # 带千分位的字符串
                    record["汇率"] = 7
                    # 计算人民币金额（四舍五入保留2位小数）
                    rmb_amount = (amount * 7).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    record["折合人民币金额"] = format_currency(str(rmb_amount))
                except:
                    record["金额"] = "N/A"
                    record["汇率"] = 7
                    record["折合人民币金额"] = "N/A"

                # 添加到结果
                all_data.append(record)

    return all_data


def process_pdfs(folder_path, output_excel):
    """批量处理PDF并输出Excel"""
    all_records = []
    processed_files = 0

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            print(f"处理文件中: {filename}")
            try:
                records = extract_customs_data(pdf_path)
                if len(records) != 6:
                    print(f"  警告: 解析到 {len(records)} 条记录 (应为6条)")
                all_records.extend(records)
                processed_files += 1
            except Exception as e:
                print(f"  处理失败: {str(e)}")

    # 创建DataFrame并保存
    if all_records:
        df = pd.DataFrame(all_records)
        # 按要求的列顺序排序
        final_columns = [
            '海关编号', '境内收货人', '商品名称', '运输方式', '贸易类型',
            '启运国', '进口日期', '入境口岸', '币制', '金额',
            '汇率', '折合人民币金额'
        ]
        df = df[final_columns]

        # 添加序号列
        df.insert(0, '序号', range(1, len(df) + 1))

        df.to_excel(output_excel, index=False, engine='openpyxl')
        print(f"\n处理完成! 共处理 {processed_files} 个PDF文件")
        print(f"生成 {len(df)} 条记录到 {output_excel}")
    else:
        print("未提取到任何有效数据")


# 使用示例
if __name__ == "__main__":
    pdf_folder = r"C:\Users\26765\Desktop\sxpdf"
    output_file = "报关单数据汇总.xlsx"

    # 创建输出目录（如果不存在）
    # os.makedirs(os.path.dirname(output_file), exist_ok=True)
    #
    process_pdfs(pdf_folder, output_file)
