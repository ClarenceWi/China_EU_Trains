# coder:Clarence_William
# 2025年06月30日15时44分27秒
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
        for page_num, page in enumerate(pdf.pages):
            # 增强文本提取
            text = page.extract_text(x_tolerance=3, y_tolerance=3, layout=True)
            ########### 添加调试输出 ###########
            print(f"\n--- 第 {page_num + 1} 页文本片段 ---")
            if text:
                print(text[:300])  # 打印前300字符便于调试
            else:
                print("无法提取文本!")

            if not text:
                continue

            # 分割单张报关单（更稳健的分割方式）
            ########### 改进分隔逻辑 ###########
            # 尝试多种分隔方式
            declarations = re.split(r"-{20,}|报关单|申报单位", text)
            # 如果分割失败（只有1个元素），尝试按固定位置分割
            if len(declarations) <= 1:
                declarations = [text[i:i + 2000] for i in range(0, len(text), 2000)]  # 每2000字符分割

            print(f"发现 {len(declarations)} 张报关单")

            for dec_idx, dec_text in enumerate(declarations):
                if not dec_text.strip():
                    continue

                record = {}
                # 提取基本字段
                ########### 增强字段提取 ###########
                for field, pattern in FIELDS_CONFIG.items():
                    match = re.search(pattern, dec_text, re.DOTALL)
                    if match:
                        record[field] = match.group(1).strip()
                        # 调试输出
                        print(f"  字段 [{field}] 匹配成功: {record[field][:30]}")
                    else:
                        record[field] = "N/A"
                        # 调试输出
                        print(f"  字段 [{field}] 匹配失败!")
                        # 打印部分文本帮助调试
                        print("    文本片段:", dec_text[:200].replace('\n', ' '))

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
                print(f"报关单 #{dec_idx + 1} 解析完成")

    return all_data


def process_pdfs(folder_path, output_excel):
    """批量处理PDF并输出Excel"""
    all_records = []
    processed_files = 0

    ########### 确保目录存在 ###########
    if not os.path.exists(folder_path):
        print(f"错误: 目录不存在 - {folder_path}")
        return

    print(f"开始扫描目录: {folder_path}")
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    print(f"找到 {len(pdf_files)} 个PDF文件")

    for filename in pdf_files:
        pdf_path = os.path.join(folder_path, filename)
        print(f"\n处理文件中: {filename}")
        try:
            records = extract_customs_data(pdf_path)
            if len(records) != 6:
                print(f"  警告: 解析到 {len(records)} 条记录 (应为6条)")
            all_records.extend(records)
            processed_files += 1
            print(f"  成功提取 {len(records)} 条记录")
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

        ########### 确保列存在 ###########
        # 添加缺失的列
        for col in final_columns:
            if col not in df.columns:
                df[col] = "N/A"
                print(f"警告: 列 '{col}' 缺失，已创建空列")

        df = df[final_columns]

        # 添加序号列
        df.insert(0, '序号', range(1, len(df) + 1))

        df.to_excel(output_excel, index=False, engine='openpyxl')
        print(f"\n处理完成! 共处理 {processed_files} 个PDF文件")
        print(f"生成 {len(df)} 条记录到 {output_excel}")
    else:
        print("\n错误: 未提取到任何有效数据!")
        print("可能原因:")
        print("1. PDF格式不支持")
        print("2. 字段提取失败")
        print("3. 目录路径错误")


# 使用示例
if __name__ == "__main__":
    pdf_folder = r"C:\Users\26765\Desktop\sxpdf"
    output_file = "报关单数据汇总.xlsx"

    # 创建输出目录（如果不存在）
    # os.makedirs(os.path.dirname(output_file), exist_ok=True)

    process_pdfs(pdf_folder, output_file)
