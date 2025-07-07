import os
import re
import pandas as pd
import pdfplumber

def extract_declaration_info(text, index):
    """根据用户提供的最新规则提取报关单信息"""
    info = {
        '序号': index,
        '报关单号': '',
        '境内收发货人': '',
        '主要商品名称': '',
        '运输方式': '铁路运输',
        '贸易类型': '进口',  # 固定值
        '贸易国': '',
        '进出口日期': '',
        '入境口岸': '',
        '币制': '美元',     # 固定值
        '金额': 0.0,
        '汇率': 7.1761,          # 固定值
        '折合人民币金额': 0.0
    }
    
    # 1. 提取报关单号（预录入编号）
    match = re.search(r'预录入编号：(\d+)', text)
    if match:
        info['报关单号'] = match.group(1)
    
    # 2. 提取境内收发货人、入境口岸、进出口日期
    # 匹配模式："境内收货人(...) 进境关别 (...) 进口日期 申报日期 备案号"后面的字段
    # 采用非贪婪匹配和明确的字段分隔
    pattern = r'境内收货人\(\w+\)\s+进境关别\s+\(\w+\)\s+进口日期\s+申报日期\s+备案号\s+(.*?)\s+境外发货人'
    match = re.search(pattern, text, re.DOTALL)
    if match:
        # 提取到的内容按空格分割（处理多个空格的情况）
        fields = re.split(r'\s+', match.group(1).strip())
        # 确保有足够的字段
        if len(fields) >= 3:
            info['境内收发货人'] = fields[0]
            info['入境口岸'] = fields[1]
            # 处理日期格式 YYYYMMDD -> YYYY-MM-DD
            date_str = fields[2]
            if len(date_str) == 8:
                info['进出口日期'] = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"

    # 4. 提取贸易国（启运国）
    # 优化后的匹配模式，考虑标题行和内容行的结构
    pattern = r'合同协议号\s+贸易国（地区）\(\w+\)\s+启运国（地区）\(\w+\)\s+经停港\(\w+\)\s+入境口岸\(\w+\)\s*\n\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)'
    match = re.search(pattern, text)

    if match:
        # 提取内容行的各个字段
        # group(1): 合同协议号
        # group(2): 贸易国（地区）内容
        # group(3): 启运国（地区）内容
        # group(4): 经停港内容
        # group(5): 入境口岸内容
        info['贸易国'] = match.group(3).strip()
    else:
        # 备选方案：直接匹配启运国字段
        alt_match = re.search(r'启运国（地区）\(\w+\)\s*(\S+)\s*经停港', text)
        if alt_match:
            info['贸易国'] = alt_match.group(1).strip()
        else:
            info['贸易国'] = "未提取到"

    # 5. 提取主要商品名称和金额
    # 优化后的正则表达式，匹配商品名称和金额
    pattern = r'项号\s+商品编号\s+商品名称及规格型号\s+数量及单位\s+单价/总价/币制.*?\d+\s+\d+\s*([^\d]+?)\s*\d+.*?(\d{4,6}(?:,\d{3})*\.\d{2}|\d{1,3}(?:,\d{3})+\.\d{2})'
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if match:
        # 提取商品名称（清理前后空格）
        product_name = match.group(1).strip()
        # 提取金额
        amount_str = match.group(2).strip()

        info['主要商品名称'] = product_name
        info['金额'] = amount_str
    else:
        # 如果主模式失败，尝试备选方案
        alt_pattern = r'项号\s+\d+\s+(\d+)\s*([^\d]+?)\s*\d+.*?(\d+\.\d{2})\s*\(\w{3}\)\s*\(\w{3}\)'
        alt_match = re.search(alt_pattern, text, re.DOTALL)
        if alt_match:
            info['主要商品名称'] = alt_match.group(2).strip()
            info['金额'] = alt_match.group(3).strip()
        else:
            info['主要商品名称'] = "未提取到"
            info['金额'] = "0.00"
    
    return info

def process_pdf_file(pdf_path, serial_number):
    """处理单个PDF文件，返回提取的数据列表"""
    data_list = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"处理文件: {os.path.basename(pdf_path)}，共{len(pdf.pages)}页")
            
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if not text:
                    print(f"警告: 第{page_num}页无文本内容，已跳过")
                    continue
                
                # 提取单页报关单信息
                info = extract_declaration_info(text, serial_number)
                serial_number += 1
                
                # 验证关键信息是否提取成功
                if not info['报关单号']:
                    print(f"警告: 第{page_num}页未提取到报关单号，已跳过")
                    continue
                
                # 打印提取结果供验证
                print(f"成功提取第{page_num}页信息:")
                print(f"  报关单号: {info['报关单号']}")
                print(f"  境内收发货人: {info['境内收发货人']}")
                print(f"  贸易国: {info['贸易国']}")
                print(f"  入境口岸: {info['入境口岸']}")
                print(f"  进出口日期: {info['进出口日期']}")
                print(f"  主要商品名称: {info['主要商品名称']}")
                print(f"  金额: {info['金额']}")
                
                data_list.append(info)
                
    except Exception as e:
        print(f"处理PDF文件时出错: {str(e)}")
    
    return data_list, serial_number

def get_valid_directory():
    """交互式获取有效的PDF目录"""
    print("默认路径为当前脚本所在目录，请将文件复制于此！")
    while True:
        pdf_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 检查路径是否存在
        if os.path.exists(pdf_dir):
            # 检查是否是目录
            if os.path.isdir(pdf_dir):
                return pdf_dir
            else:
                print(f"错误: {pdf_dir} 不是一个有效的目录")
        else:
            print(f"错误: 目录 {pdf_dir} 不存在")
            # 询问用户是否要创建该目录
            create_dir = input("是否要创建这个目录? (y/n): ").strip().lower()
            if create_dir == 'y':
                try:
                    os.makedirs(pdf_dir)
                    print(f"已成功创建目录: {pdf_dir}")
                    print("请将PDF文件放入该目录后按Enter继续...")
                    input()
                    return pdf_dir
                except Exception as e:
                    print(f"创建目录失败: {str(e)}")

def main():
    print("===== 海关报关单信息精准提取工具 =====")
    print("(根据最新提取规则优化)")
    
    # 获取用户选择的PDF目录
    pdf_dir = get_valid_directory()
    
    # 获取输出Excel路径
    default_output = os.path.join(os.getcwd(), "2025年国际贸易报关单台账-投资发展公司.xlsx")
    print(f"\n请输入输出Excel文件路径 (默认: {default_output}):")
    output_excel = input().strip()
    if not output_excel:
        output_excel = default_output
    
    # 获取所有PDF文件
    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"错误: 在目录 {pdf_dir} 中未找到PDF文件")
        print("请将PDF文件放入该目录后重新运行脚本")
        return
    
    all_data = []
    serial_number = 1  # 序号从1开始
    
    # 处理每个PDF文件
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_dir, pdf_file)
        page_data, serial_number = process_pdf_file(pdf_path, serial_number)
        all_data.extend(page_data)
    
    # 保存结果到Excel
    if all_data:
        # 定义列顺序
        columns_order = ['序号', '报关单号', '境内收发货人', '主要商品名称', '运输方式', 
                         '贸易类型', '贸易国', '进出口日期', '入境口岸', '币制', 
                         '金额', '汇率', '折合人民币金额']
        
        df = pd.DataFrame(all_data)

        while True:
            try:
                # 提示用户输入汇率
                exchange_rate = float(input("请输入汇率（例如7.0）："))
                # 将汇率应用到所有行
                df['汇率'] = exchange_rate
                break  # 输入有效，退出循环
            except ValueError:
                print("输入无效，请输入数字（例如7.0）")

        # 添加公式计算折合人民币金额 = 金额 * 汇率
        for idx, row in df.iterrows():
            # 获取金额和汇率的值
            amount = row['金额']
            rate = row['汇率']

            # 计算折合人民币金额
            # 注意：这里使用Excel公式字符串，而不是计算结果
            df['折合人民币金额'] = df['折合人民币金额'].astype(str)
            df.at[idx, '折合人民币金额'] = f'={amount}*{rate}'

        # 重新排列列
        df = df.reindex(columns=columns_order)

        # 保存到Excel
        df.to_excel(output_excel, index=False, engine='openpyxl')
        # df = df.reindex(columns=columns_order)
        # df.to_excel(output_excel, index=False, engine='openpyxl')
        
        print(f"\n===== 处理完成 =====")
        print(f"共处理 {len(pdf_files)} 个PDF文件，提取 {len(all_data)} 条报关单记录")
        print(f"结果已保存至: {output_excel}")
    else:
        print("\n===== 处理完成 =====")
        print("未提取到任何有效数据")

if __name__ == "__main__":
    main()