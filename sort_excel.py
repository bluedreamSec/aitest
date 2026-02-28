#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
对Excel表格按照第二列、第三列内容进行归类排序，将相同的放在一起，输出新的Excel文件。
支持命令行参数指定输入输出路径。
用法:
    python sort_excel.py -i input.xlsx -o output.xlsx
或
    python sort_excel.py --input input.xlsx --output output.xlsx
"""

import pandas as pd
import sys
import os
import argparse

def sort_excel(input_path, output_path):
    """
    读取Excel，按第二列(level1)和第三列(level2)排序，写入新文件。
    """
    try:
        df = pd.read_excel(input_path)
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        sys.exit(1)
    
    # 检查列名
    expected_cols = ['question', 'level1', 'level2']
    if not all(col in df.columns for col in expected_cols):
        print(f"警告: 列名不匹配，实际列: {df.columns.tolist()}")
        # 尝试按位置指定列
        if df.shape[1] >= 3:
            df.columns = ['question', 'level1', 'level2'] + list(df.columns[3:])
        else:
            print("错误: 文件列数不足3列")
            sys.exit(1)
    
    # 排序：先按level1，再按level2，保持原始顺序（稳定排序）
    # 使用sort_values，缺失值放在最后
    df_sorted = df.sort_values(by=['level1', 'level2'], na_position='last')
    
    # 写入新文件
    try:
        df_sorted.to_excel(output_path, index=False)
        print(f"已排序并写入 {output_path}")
        print(f"总行数: {len(df_sorted)}")
        print(f"排序后前几行:")
        print(df_sorted.head())
    except Exception as e:
        print(f"写入Excel文件失败: {e}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description='对Excel表格按照第二列、第三列进行归类排序')
    parser.add_argument('-i', '--input', default=r'D:\002-安全文件\013-AI及安全\demoinfo3.xlsx',
                        help='输入Excel文件路径（默认: demoinfo3.xlsx）')
    parser.add_argument('-o', '--output', default=r'D:\002-安全文件\013-AI及安全\demoinfo4.xlsx',
                        help='输出Excel文件路径（默认: demoinfo4.xlsx）')
    args = parser.parse_args()
    
    input_path = args.input
    output_path = args.output
    
    # 检查输入文件是否存在
    if not os.path.exists(input_path):
        print(f"错误: 输入文件不存在: {input_path}")
        sys.exit(1)
    
    sort_excel(input_path, output_path)

if __name__ == "__main__":
    main()