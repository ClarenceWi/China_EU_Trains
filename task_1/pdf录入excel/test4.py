# coder:Clarence_William
# 2025年06月30日16时58分40秒
# soforso@163.com

import os
import re
import pandas as pd
import pdfplumber

# 固定工作目录（不要修改）
WORK_DIR = os.path.join(os.path.expanduser('~'), 'Desktop', '报关单提取')
OUTPUT_FILE = os.path.join(WORK_DIR, '提取结果.xlsx')


def extract_info(text, index):
    """提取单张报关单信息"""
    info = {
        '序号': index,
        '报关单号': '',
        '境内收发货人': '',
        '主要商品名称': '',
        '运输方式': '铁路运输',  # 根据您的PDF固定为铁路运输
        '贸易类型': '进口',
        '贸易国': '哈萨克斯坦',  # 根据您的PDF固定为哈萨克斯坦
        '进出口日期': '',
        '入境口岸': '霍尔果斯',  # 根据您的PDF固定为霍尔果斯
        '币制': '美元',
        '金额': 0.0,
        '汇率': 7,
        '折合人民币金额': 0.0
    }

    # 提取报关单号
    match = re.search(r'预录入编号：(\d+)', text)
    if match:
        info['报关单号'] = match.group(1)

    # 提取境内收发货人
    match = re.search(r'境内收货人\(\w+\)\s*(.*?)\s*进境关别', text, re.DOTALL)
    if match:
        info['境内收发货人'] = match.group(1).strip()

    # 提取商品名称
    match = re.search(r'商品名称及规格型号\s*\d+\s*(.*?)\s*数量及单位', text, re.DOTALL)
    if match:
        info['主要商品名称'] = match.group(1).strip()

    # 提取进口日期
    match = re.search(r'进口日期\s*(\d{8})', text)
    if match:
        date = match.group(1)
        info['进出口日期'] = f"{date[:4]}-{date[4:6]}-{date[6:]}"

    # 提取金额
    match = re.search(r'(\d+\.\d+)\s*美元', text)
    if match:
        try:
            info['金额'] = float(match.group(1))
            info['折合人民币金额'] = round(info['金额'] * 7, 2)
        except:
            pass

    return info


def main():
    print("===== 报关单提取工具 =====")
    print(f"工作目录: {WORK_DIR}")

    # 检查工作目录是否存在
    if not os.path.exists(WORK_DIR):
        print(f"错误：工作目录不存在，请按指南创建 '报关单提取' 文件夹")
        return

    # 获取PDF文件
    pdf_files = [f for f in os.listdir(WORK_DIR) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"错误：在 {WORK_DIR} 中未找到PDF文件")
        print("请将PDF文件复制到该文件夹后重试")
        return

    all_data = []
    index = 1

    # 处理每个PDF
    for pdf_file in pdf_files:
        pdf_path = os.path.join(WORK_DIR, pdf_file)
        print(f"\n处理文件: {pdf_file}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        data = extract_info(text, index)
                        all_data.append(data)
                        print(f"  提取成功: 报关单号={data['报关单号']}")
                        index += 1
                    else:
                        print(f"  警告: 页面无文本内容")
        except Exception as e:
            print(f"  处理失败: {str(e)}")

    # 保存结果
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n提取完成！结果已保存至: {OUTPUT_FILE}")
        print(f"共提取 {len(all_data)} 条记录")
    else:
        print("\n未提取到任何数据")

    # 按任意键退出
    input("\n按Enter键退出...")


if __name__ == "__main__":
    main()