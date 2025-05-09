#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Extract embedded images from an Excel file and save them to a designated folder.
Enhanced version with better debugging and alternative extraction methods.
"""

import os
import sys
import argparse
import zipfile
import io
import traceback
from pathlib import Path

# 检查是否安装了PIL库
try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    print("警告: PIL/Pillow库未安装，某些图像验证功能将被禁用")


def extract_images_using_zipfile(excel_path, output_dir):
    """
    Extract images from Excel file by treating it as a zip archive.
    Excel files (.xlsx) are actually zip archives with XML and binary files inside.
    
    Args:
        excel_path (str): Path to the Excel file
        output_dir (str): Directory to save extracted images
    
    Returns:
        int: Number of images extracted
    """
    # 检查文件是否存在
    if not os.path.exists(excel_path):
        print(f"错误: 找不到Excel文件: {excel_path}")
        return 0
        
    # 打印文件大小信息
    file_size_mb = os.path.getsize(excel_path) / (1024 * 1024)
    print(f"Excel文件大小: {file_size_mb:.2f} MB")

    # 创建输出目录（如果不存在）
    try:
        os.makedirs(output_dir, exist_ok=True)
        print(f"输出目录已创建/确认: {output_dir}")
    except Exception as e:
        print(f"创建输出目录时出错: {str(e)}")
        return 0
    
    print(f"正在打开Excel文件作为zip压缩包: {excel_path}")
    
    image_count = 0
    
    try:
        # 打开Excel文件作为ZIP压缩包
        with zipfile.ZipFile(excel_path, 'r') as zip_ref:
            # 列出压缩包中的所有文件
            all_files = zip_ref.namelist()
            print(f"Excel压缩包中的文件总数: {len(all_files)}")
            print(f"压缩包中的前20个文件:")
            for i, file in enumerate(all_files):
                if i < 20:  # 显示前20个文件以进行调试
                    print(f"  - {file}")
            
            # 搜索媒体文件
            media_files = [f for f in all_files if 'media' in f.lower()]
            print(f"找到 {len(media_files)} 个可能的媒体文件")
            
            # 首先检查标准位置的图像文件
            print("第一遍: 检查标准媒体文件夹...")
            for file in all_files:
                # 检查图像文件（Excel中常见的图像文件扩展名）
                if file.startswith('xl/media/') or '/ppt/media/' in file or '/word/media/' in file:
                    try:
                        image_count += 1
                        image_filename = os.path.basename(file)
                        # 获取图像数据
                        image_data = zip_ref.read(file)
                        
                        # 将图像保存到输出目录
                        output_path = os.path.join(output_dir, f"{image_count}_{image_filename}")
                        
                        with open(output_path, 'wb') as f:
                            f.write(image_data)
                        
                        print(f"已保存图像 {image_count}: {output_path}")
                    except Exception as e:
                        print(f"提取图像 {file} 时出错: {str(e)}")
            
            if image_count == 0:
                print("在标准媒体文件夹中未找到图像。正在搜索其他潜在图像文件...")
                
                # 第二遍查找任何类似图像的文件
                print("第二遍: 检查图像扩展名...")
                for file in all_files:
                    if any(file.endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.emf', '.wmf']):
                        try:
                            image_count += 1
                            image_filename = os.path.basename(file)
                            # 获取图像数据
                            image_data = zip_ref.read(file)
                            
                            # 将图像保存到输出目录
                            output_path = os.path.join(output_dir, f"{image_count}_{image_filename}")
                            
                            with open(output_path, 'wb') as f:
                                f.write(image_data)
                            
                            print(f"已保存图像 {image_count}: {output_path}")
                        except Exception as e:
                            print(f"提取图像 {file} 时出错: {str(e)}")
                
                # 第三遍尝试检测图像二进制数据
                if HAS_PIL:
                    print("第三遍: 使用PIL尝试检测图像内容...")
                    for file in all_files:
                        if 'drawings' in file.lower() or 'image' in file.lower() or 'media' in file.lower():
                            try:
                                image_data = zip_ref.read(file)
                                # 尝试打开为图像以验证
                                img = Image.open(io.BytesIO(image_data))
                                img.verify()  # 验证它是有效的图像
                                
                                image_count += 1
                                output_path = os.path.join(output_dir, f"{image_count}_detected_image.{img.format.lower() if img.format else 'png'}")
                                
                                with open(output_path, 'wb') as f:
                                    f.write(image_data)
                                
                                print(f"已保存检测到的图像 {image_count}: {output_path}")
                            except Exception as e:
                                # 不是有效的图像，跳过
                                pass

            # 确认打印
            if image_count > 0:
                print(f"总共提取了 {image_count} 张图像，保存到: {output_dir}")
            else:
                print(f"在Excel文件中未找到任何图像。")
                # 打印出找到的文件类型以供参考
                extensions = set(os.path.splitext(f)[1] for f in all_files if '.' in f)
                print(f"Excel压缩包中的文件扩展名: {', '.join(extensions)}")
    
    except zipfile.BadZipFile:
        print(f"错误: 文件 {excel_path} 不是有效的zip文件或Excel文件。")
        return 0
    except Exception as e:
        print(f"提取图像时出现错误: {str(e)}")
        print("详细错误信息:")
        traceback.print_exc()
        return 0
    
    print(f"从 {excel_path} 提取了 {image_count} 张图像")
    
    # 确认输出目录中的文件
    if os.path.exists(output_dir):
        files_in_output = os.listdir(output_dir)
        print(f"输出目录 {output_dir} 中有 {len(files_in_output)} 个文件")
        if len(files_in_output) > 0:
            print(f"前几个文件: {', '.join(files_in_output[:5])}")
    
    return image_count


def main():
    parser = argparse.ArgumentParser(description="从Excel文件中提取图像")
    parser.add_argument("excel_file", help="Excel文件的路径")
    parser.add_argument("--output", "-o", default="extracted_images", 
                        help="保存提取的图像的目录 (默认: ./extracted_images)")
    
    args = parser.parse_args()
    
    print(f"Python版本: {sys.version}")
    print(f"当前工作目录: {os.getcwd()}")
    
    # 转换相对路径为绝对路径
    excel_path = os.path.abspath(args.excel_file)
    output_dir = os.path.abspath(args.output)
    
    print(f"Excel文件路径: {excel_path}")
    print(f"输出目录: {output_dir}")
    
    if not os.path.exists(excel_path):
        print(f"错误: 找不到Excel文件: {excel_path}")
        return
    
    # 使用替代方法提取图像
    extract_images_using_zipfile(excel_path, output_dir)


if __name__ == "__main__":
    main()