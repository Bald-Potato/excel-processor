import os
import pandas as pd
from pathlib import Path
import re
import math

def _find_starttime_columns(df):
    """查找包含'starttime'的列"""
    starttime_columns = []
    for col in df.columns:
        if isinstance(col, str) and 'starttime' in col.lower():
            starttime_columns.append(col)
    return starttime_columns

def process_excel_file(file_path, output_folder, check_divisible=False, convert_time=False, custom_suffixes=None):
    """处理单个Excel文件"""
    try:
        print(f"正在处理文件: {file_path.name}")
        
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 查找包含'starttime'的列
        starttime_columns = _find_starttime_columns(df)
        
        if not starttime_columns:
            print(f"❌ 文件 {file_path.name} 中未找到包含'starttime'的列，跳过处理。")
            return
        
        print(f"✅ 找到以下starttime列: {', '.join(starttime_columns)}")
        
        # 判断是否有非40整数倍的值
        has_invalid_values = False
        
        # 处理每个starttime列
        for col in starttime_columns:
            # 如果需要检查是否为40的整数倍
            if check_divisible:
                # 添加一个新列来标记是否为40的整数倍
                is_divisible_col = f"{col}_is_divisible_by_40"
                df[is_divisible_col] = df[col].apply(lambda x: 
                                                    True if pd.notna(x) and x % 40 == 0 
                                                    else False if pd.notna(x) 
                                                    else None)
                
                # 检查是否有非40整数倍的值
                invalid_values = df[~df[is_divisible_col] & df[is_divisible_col].notna()]
                if not invalid_values.empty:
                    has_invalid_values = True
                    print(f"⚠️  列 {col} 中存在非40整数倍的值!")
            
            # 如果需要转换时间格式
            if convert_time:
                time_col = f"{col}_time_format"
                df[time_col] = df[col].apply(lambda x: 
                                            _convert_to_time_format(x) if pd.notna(x) 
                                            else None)
        
        # 确定输出文件名
        file_stem = file_path.stem
        file_suffix = file_path.suffix
        
        # 如果有自定义后缀，使用自定义后缀
        if custom_suffixes and len(custom_suffixes) > 0:
            # 获取文件在当前文件夹中的索引
            file_index = 0
            # 如果索引超出自定义后缀列表长度，使用最后一个后缀加上索引
            if file_index < len(custom_suffixes):
                custom_suffix = custom_suffixes[file_index]
            else:
                last_suffix = custom_suffixes[-1] if custom_suffixes else "已处理"
                custom_suffix = f"{last_suffix}{file_index - len(custom_suffixes) + 1}"
            
            if has_invalid_values and check_divisible:
                output_file_name = f"FALSE_{file_stem}_{custom_suffix}{file_suffix}"
            else:
                output_file_name = f"{file_stem}_{custom_suffix}{file_suffix}"
        else:
            # 使用默认后缀
            if has_invalid_values and check_divisible:
                output_file_name = f"FALSE_{file_stem}_已处理{file_suffix}"
            else:
                output_file_name = f"{file_stem}_已处理{file_suffix}"
        
        # 保存处理后的文件
        output_path = output_folder / output_file_name
        df.to_excel(output_path, index=False)
        print(f"✅ 文件已保存: {output_path.name}")
        
    except Exception as e:
        print(f"❌ 处理文件 {file_path.name} 时出错: {str(e)}")

def _convert_to_time_format(ms):
    """将毫秒转换为时:分:秒:毫秒格式"""
    try:
        # 确保输入是数字
        ms = float(ms)
        
        # 计算时、分、秒和毫秒
        total_seconds = ms / 1000
        hours = math.floor(total_seconds / 3600)
        minutes = math.floor((total_seconds % 3600) / 60)
        seconds = math.floor(total_seconds % 60)
        milliseconds = math.floor(ms % 1000)
        
        # 格式化为时:分:秒:毫秒
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}:{milliseconds:03d}"
    except (ValueError, TypeError):
        return "无效时间"

def process_folder(root_folder, check_divisible=False, convert_time=False, custom_suffixes=None):
    """处理文件夹中的所有Excel文件"""
    root_path = Path(root_folder)
    print(f"开始处理文件夹: {root_path}")
    
    # 计数器
    total_files = 0
    processed_files = 0
    
    # 遍历文件夹中的所有Excel文件
    for file_path in root_path.glob("**/*.xlsx"):
        # 跳过临时文件
        if file_path.name.startswith("~$") or "/~$" in str(file_path):
            continue
            
        total_files += 1
        
        # 为每个文件创建输出文件夹
        output_folder = file_path.parent / "_processed"
        output_folder.mkdir(exist_ok=True)
        
        # 处理文件
        process_excel_file(file_path, output_folder, check_divisible, convert_time, custom_suffixes)
        processed_files += 1
    
    if total_files == 0:
        print("⚠️  未找到任何Excel文件。")
    else:
        print(f"✅ 处理完成! 共处理了 {processed_files} 个文件，跳过了 {total_files - processed_files} 个文件。")
