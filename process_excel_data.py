#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Process Excel data to extract specific columns and remove duplicates based on 款号 and 产品名称.
"""

import os
import sys
import argparse
import pandas as pd
from pathlib import Path

def process_excel_data(input_file, output_file):
    """
    Process Excel data to extract specific columns and remove duplicates.
    
    Args:
        input_file (str): Path to the input Excel file
        output_file (str): Path to save the processed Excel file
    
    Returns:
        bool: True if processing was successful, False otherwise
    """
    try:
        # Check if file exists
        if not os.path.exists(input_file):
            print(f"错误: 找不到Excel文件: {input_file}")
            return False
            
        # Print file size info
        file_size_mb = os.path.getsize(input_file) / (1024 * 1024)
        print(f"Excel文件大小: {file_size_mb:.2f} MB")
        
        # Read Excel file
        print(f"正在读取Excel文件: {input_file}")
        df = pd.read_excel(input_file)
        
        # Display original dataframe info
        print(f"原始数据表格尺寸: {df.shape[0]}行 x {df.shape[1]}列")
        print(f"原始数据列名: {', '.join(df.columns)}")
        
        # 检查并清理列名中的空格和不可见字符
        df.columns = [col.strip() for col in df.columns]
        
        # 显示清理后的列名
        print(f"清理后的列名: {', '.join(df.columns)}")
        
        # Extract only the required columns
        required_columns = ['款号', '产品名称', '品目']
        
        # 尝试查找列名，即使它们有轻微的变化
        actual_columns = []
        missing_columns = []
        
        for req_col in required_columns:
            # 检查精确匹配
            if req_col in df.columns:
                actual_columns.append(req_col)
            else:
                # 检查包含关系
                matching_cols = [col for col in df.columns if req_col in col]
                if matching_cols:
                    print(f"找到类似列名: '{req_col}' -> '{matching_cols[0]}'")
                    actual_columns.append(matching_cols[0])
                else:
                    missing_columns.append(req_col)
        
        if missing_columns:
            print(f"错误: Excel文件中缺少以下列: {', '.join(missing_columns)}")
            return False
        
        # 重新设置列名映射
        column_mapping = dict(zip(actual_columns, required_columns))
        
        try:
            # 提取所需列
            df_extract = df[actual_columns].copy()
            # 重命名列以确保一致性
            df_extract.rename(columns=column_mapping, inplace=True)
            print(f"提取了以下列: {', '.join(required_columns)}")
        except KeyError as e:
            print(f"提取列时出错: {str(e)}")
            print("请确保Excel文件包含以下列: '款号', '产品名称', '品目'")
            return False
        
        # Remove duplicates based on '款号' and '产品名称'
        df_deduplicated = df_extract.drop_duplicates(subset=['款号', '产品名称'])
        
        # Print deduplication results
        duplicate_count = df_extract.shape[0] - df_deduplicated.shape[0]
        print(f"已删除 {duplicate_count} 行重复记录")
        print(f"处理后数据尺寸: {df_deduplicated.shape[0]}行 x {df_deduplicated.shape[1]}列")
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        # Save to a new Excel file
        print(f"正在保存处理后的数据到: {output_file}")
        df_deduplicated.to_excel(output_file, index=False)
        
        print(f"成功处理Excel数据并保存到: {output_file}")
        return True
        
    except Exception as e:
        print(f"处理Excel数据时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 构建数据目录路径
    data_dir = os.path.join(script_dir, "data")
    # 默认输出文件路径
    default_output = os.path.join(data_dir, "processed_data.xlsx")
    
    parser = argparse.ArgumentParser(description="处理Excel数据并提取特定列")
    parser.add_argument("input_file", help="输入Excel文件的路径")
    parser.add_argument("--output", "-o", help=f"输出Excel文件的路径 (默认: {default_output})",
                        default=default_output)
    
    args = parser.parse_args()
    
    print(f"Python版本: {sys.version}")
    print(f"当前工作目录: {os.getcwd()}")
    
    # Convert relative paths to absolute paths
    input_path = os.path.abspath(args.input_file)
    output_path = os.path.abspath(args.output)
    
    print(f"输入文件路径: {input_path}")
    print(f"输出文件路径: {output_path}")
    
    # Process the Excel data
    process_excel_data(input_path, output_path)

if __name__ == "__main__":
    main()