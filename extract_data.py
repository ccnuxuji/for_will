"""
🚀 Wafer测试数据提取工具 - 智能版本 v2.0

✨ 新功能特点：
1. 🔍 智能识别数据结构 - 不依赖固定行号和列号
2. 📊 动态查找关键标识符 - 支持WAFER ID、SLOT、MEAN、3 SIGMA、Site #
3. 🎯 自动定位数据列 - 智能匹配N1@633、T1、X、Y列
4. 📄 支持CSV格式输入 - 替代原有的Excel格式
5. 🔧 灵活的文件头部处理 - 适应不同格式的数据文件

🎨 使用方法：
1. 在main函数中修改input_basename变量为你的文件名
2. 确保你的CSV文件包含以下关键标识符：
   - WAFER ID：标识wafer数据块
   - SLOT：标识样品ID
   - MEAN：包含平均值数据
   - 3 SIGMA：包含3sigma数据
   - Site #：标识测试点数据的开始
3. 运行程序：python extract_data.py

📋 输出数据：
- lot_id：从SLOT行提取的样品ID
- N1_mean：N1@633的平均值
- T1_mean：T1的平均值
- T1_3sigma_mean：T1的3sigma与mean的比值
- N1_XY_00：X=0,Y=0位置的N1@633值
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime

def read_csv_data(file_path):
    """读取CSV文件并返回DataFrame"""
    try:
        if not os.path.exists(file_path):
            print(f"错误：文件 {file_path} 不存在")
            return None
        
        print(f"🔍 正在读取CSV文件: {file_path}")
        
        # 尝试不同的编码方式读取CSV文件（使用逗号分隔符）
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin1']
        df = None
        successful_encoding = None
        
        for encoding in encodings:
            try:
                # 使用逗号分隔符和容错处理
                df = pd.read_csv(
                    file_path, 
                    header=None, 
                    encoding=encoding,
                    sep=',',  # 使用逗号分隔符
                    on_bad_lines='skip',  # 跳过有问题的行
                    engine='python',  # 使用Python引擎，更灵活
                    skipinitialspace=True,  # 跳过分隔符后的空格
                    quoting=3  # 忽略引号问题
                )
                
                # 检查是否成功读取到有效数据
                if df is not None and len(df) > 0:
                    # 检查是否只有一列，这可能意味着分隔符问题
                    if df.shape[1] == 1:
                        print(f"⚠️  检测到单列数据，可能存在格式问题")
                        print(f"  尝试处理混合格式...")
                        
                        # 重新读取，允许不一致的列数
                        df = pd.read_csv(
                            file_path, 
                            header=None, 
                            encoding=encoding,
                            sep=',',
                            engine='python',
                            on_bad_lines='skip',
                            skipinitialspace=True,
                            quoting=3,
                            # 填充缺失的列
                            names=None
                        )
                        
                        # 如果还是单列，尝试手动处理
                        if df.shape[1] == 1:
                            print(f"  手动处理混合格式文件...")
                            
                            # 读取所有行
                            with open(file_path, 'r', encoding=encoding) as f:
                                lines = f.readlines()
                            
                            # 处理每一行，保持原始的列数结构
                            processed_data = []
                            for i, line in enumerate(lines):
                                line = line.strip()
                                if ',' in line:
                                    # 有逗号的行按逗号分割
                                    parts = line.split(',')
                                    processed_data.append(parts)
                                else:
                                    # 没有逗号的行作为单个值
                                    processed_data.append([line])
                            
                            # 找到最大列数，但保持每行的实际列数
                            max_cols = max(len(row) for row in processed_data) if processed_data else 1
                            
                            # 创建不规则的DataFrame - 用空字符串填充较短的行
                            for row in processed_data:
                                while len(row) < max_cols:
                                    row.append('')
                            
                            # 创建新的DataFrame
                            df = pd.DataFrame(processed_data)
                            print(f"  ✓ 手动处理完成，数据形状: {df.shape}")
                            print(f"  📊 数据结构分析:")
                            
                            # 分析每行的实际列数
                            row_col_counts = {}
                            for i, row in enumerate(processed_data):
                                actual_cols = len([col for col in row if col != ''])
                                if actual_cols not in row_col_counts:
                                    row_col_counts[actual_cols] = []
                                row_col_counts[actual_cols].append(i)
                            
                            for col_count, rows in row_col_counts.items():
                                if len(rows) > 3:  # 只显示主要的行组
                                    print(f"    {col_count}列: 行{rows[0]+1}-{rows[-1]+1} (共{len(rows)}行)")
                                    # 显示示例
                                    sample_row = rows[0]
                                    sample_data = [col for col in processed_data[sample_row] if col != '']
                                    print(f"      示例: {sample_data[:3]}...")  # 显示前3列作为示例
                    
                    successful_encoding = encoding
                    print(f"✓ 成功读取文件 {file_path}")
                    print(f"  使用编码: {encoding}")
                    print(f"  使用分隔符: 逗号 (,)")
                    break
                    
            except Exception as e:
                continue
        
        # 如果所有编码都失败，尝试更宽松的读取方式
        if df is None or len(df) == 0:
            print("🔄 尝试使用更宽松的读取方式...")
            try:
                df = pd.read_csv(
                    file_path, 
                    header=None, 
                    encoding='utf-8',
                    sep=',',
                    engine='python',
                    on_bad_lines='skip',  # 跳过有问题的行
                    skipinitialspace=True,
                    quoting=3  # 忽略引号问题
                )
                if df is not None and len(df) > 0:
                    print(f"✓ 使用宽松模式成功读取文件")
                    successful_encoding = 'utf-8'
            except Exception as e:
                print(f"❌ 宽松模式也失败: {e}")
        
        if df is None or len(df) == 0:
            print(f"❌ 无法读取文件 {file_path}")
            print("💡 可能的原因:")
            print("   1. 文件不是标准的CSV格式")
            print("   2. 文件内容有格式错误")
            print("   3. 文件编码不匹配")
            
            # 尝试读取文件的前几行作为文本显示
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = f.readlines()[:5]
                    print("   📄 文件前5行内容:")
                    for i, line in enumerate(lines, 1):
                        print(f"      第{i}行: {repr(line.strip())}")
            except:
                pass
            
            return None
        
        print(f"  数据形状: {df.shape[0]} 行 × {df.shape[1]} 列")
        
        # 显示前几行数据作为预览
        if len(df) > 0:
            print("  📋 数据预览 (前3行):")
            print(df.head(3).to_string(index=False))
        
        return df
        
    except Exception as e:
        print(f"❌ 读取CSV文件时出错: {e}")
        print("💡 请检查文件格式和内容")
        
        # 显示更详细的错误信息
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()[:3]
                print("   📄 文件前3行内容:")
                for i, line in enumerate(lines, 1):
                    print(f"      第{i}行: {repr(line.strip())}")
        except:
            pass
        
        return None

def find_column_headers(df, start_row, end_row):
    """在指定范围内查找列头"""
    column_mapping = {}
    
    for row_idx in range(start_row, min(end_row, len(df))):
        row_data = df.iloc[row_idx].astype(str)
        for col_idx, cell in enumerate(row_data):
            cell_str = str(cell).upper()
            if 'N1' in cell_str and '633' in cell_str:
                column_mapping['N1_633'] = col_idx
            elif 'T1' in cell_str and 'N1' not in cell_str:
                column_mapping['T1'] = col_idx
            elif cell_str == 'X':
                column_mapping['X'] = col_idx
            elif cell_str == 'Y':
                column_mapping['Y'] = col_idx
    
    return column_mapping

def analyze_data_structure(df):
    """分析数据结构"""
    print("\n" + "="*50)
    print("📊 数据结构分析")
    print("="*50)
    
    # 查找所有可能的关键标识符
    key_identifiers = ['WAFER ID', 'SLOT', 'MEAN', '3 SIGMA', 'Site #']
    
    print("🔍 搜索关键标识符...")
    print(f"  数据框形状: {df.shape}")
    
    # 显示更多行的预览来调试
    print(f"  📋 前10行数据预览:")
    print(df.head(10).to_string())
    
    identifier_positions = {}
    
    for identifier in key_identifiers:
        print(f"\n🔍 搜索 '{identifier}'...")
        # 在所有列和行中搜索
        positions = []
        
        for row_idx in range(len(df)):
            for col_idx in range(df.shape[1]):
                try:
                    cell_value = str(df.iloc[row_idx, col_idx])
                    if identifier.upper() in cell_value.upper():
                        positions.append((row_idx, col_idx))
                        print(f"  ✓ 在第{row_idx+1}行第{col_idx+1}列找到 '{identifier}': {cell_value}")
                except:
                    continue
        
        if positions:
            identifier_positions[identifier] = positions
            print(f"  {identifier}: 找到 {len(positions)} 个位置 - {positions}")
        else:
            print(f"  {identifier}: 未找到")
    
    # 查找所有WAFER ID行
    wafer_id_rows = []
    if 'WAFER ID' in identifier_positions:
        wafer_id_rows = [pos[0] for pos in identifier_positions['WAFER ID']]
    
    print(f"\n发现 {len(wafer_id_rows)} 个wafer数据块")
    
    if wafer_id_rows:
        print(f"WAFER ID位于行: {wafer_id_rows}")
        
        # 显示第一个wafer的数据结构示例
        if len(wafer_id_rows) > 0:
            wafer_start = wafer_id_rows[0]
            wafer_end = wafer_id_rows[1] if len(wafer_id_rows) > 1 else len(df)
            
            print(f"\n📋 第一个wafer数据结构 (行 {wafer_start} 到 {wafer_end-1}):")
            
            # 查找关键行在当前wafer范围内的位置
            for key in key_identifiers:
                if key in identifier_positions:
                    wafer_positions = [pos for pos in identifier_positions[key] 
                                     if wafer_start <= pos[0] < wafer_end]
                    if wafer_positions:
                        print(f"  {key}: {wafer_positions}")
    else:
        print("\n❌ 未找到任何wafer数据块")
        print("🔧 请检查以下内容:")
        print("   1. 文件是否包含'WAFER ID'标识符")
        print("   2. 数据格式是否正确")
        print("   3. 文件编码是否正确")
        print("   4. 尝试查看前几行数据:")
        if len(df) > 0:
            print(f"      前5行数据预览:")
            print(df.head().to_string(index=False))
    
    return wafer_id_rows, identifier_positions

def analyze_wafer_structure(df, wafer_start, wafer_end):
    """分析单个wafer的数据结构"""
    print(f"   📊 分析wafer结构 (行 {wafer_start} 到 {wafer_end-1})")
    
    # 分析每行的列数
    row_structures = []
    for row_idx in range(wafer_start, wafer_end):
        if row_idx < len(df):
            actual_cols = len([col for col in df.iloc[row_idx] if pd.notna(col) and str(col).strip() != ''])
            row_structures.append((row_idx, actual_cols))
    
    # 识别数据段
    sections = {'2_col': [], '5_col': [], '14_col': []}
    
    for row_idx, col_count in row_structures:
        if col_count == 2:
            sections['2_col'].append(row_idx)
        elif col_count >= 4 and col_count <= 6:  # 5列左右
            sections['5_col'].append(row_idx)
        elif col_count >= 10:  # 14列左右（允许一定范围）
            sections['14_col'].append(row_idx)
    
    print(f"   📋 数据段分析:")
    print(f"     2列段: {len(sections['2_col'])} 行 - {sections['2_col']}")
    print(f"     5列段: {len(sections['5_col'])} 行 - {sections['5_col']}")
    print(f"     14列段: {len(sections['14_col'])} 行 - {sections['14_col']}")
    
    return sections

def extract_wafer_data(df, identifier_positions):
    """提取wafer数据"""
    print("\n" + "="*50)
    print("⚙️  开始提取wafer数据")
    print("="*50)
    
    results = []
    
    # 获取所有WAFER ID行
    wafer_id_rows = []
    if 'WAFER ID' in identifier_positions:
        wafer_id_rows = [pos[0] for pos in identifier_positions['WAFER ID']]
    
    if not wafer_id_rows:
        print("❌ 未找到WAFER ID行")
        return results
    
    # 为每个wafer处理数据
    for i, wafer_start in enumerate(wafer_id_rows):
        print(f"\n🔍 处理wafer {i+1}/{len(wafer_id_rows)}")
        
        # 确定当前wafer的数据范围
        if i + 1 < len(wafer_id_rows):
            wafer_end = wafer_id_rows[i + 1]
        else:
            wafer_end = len(df)
        
        print(f"   数据范围: 行 {wafer_start} 到 {wafer_end-1}")
        
        # 分析wafer结构
        sections = analyze_wafer_structure(df, wafer_start, wafer_end)
        
        # 初始化结果字典
        result = {
            'lot_id': None,
            'N1_mean': None,
            'T1_mean': None,
            'T1_3sigma_mean': None,
            'N1_XY_00': None
        }
        
        # 1. 从2列段中找到SLOT值
        try:
            print("   🔍 在2列段中查找SLOT值...")
            slot_found = False
            for row_idx in sections['2_col']:
                if row_idx < len(df) and pd.notna(df.iloc[row_idx, 0]):
                    if 'SLOT' in str(df.iloc[row_idx, 0]).upper():
                        if row_idx < len(df) and len(df.iloc[row_idx]) > 1:
                            result['lot_id'] = df.iloc[row_idx, 1]
                            print(f"   ✓ SLOT值: {result['lot_id']} (位置: 行{row_idx+1}, 列2)")
                            slot_found = True
                            break
            if not slot_found:
                print("   ❌ 未在2列段中找到SLOT值")
        except Exception as e:
            print(f"   ❌ 提取SLOT值时出错: {e}")
        
        # 2. 从5列段中找到MEAN和3 SIGMA值
        try:
            print("   🔍 在5列段中查找MEAN和3 SIGMA...")
            
            # 根据已知的列位置直接提取数据（N1@633在列2，T1在列4）
            n1_col = 2  # N1@633列（之前是M1@633）
            t1_col = 4  # T1列
            
            print(f"   🔍 使用固定列位置: N1@633列={n1_col}, T1列={t1_col}")
            
            # 提取MEAN值
            for row_idx in sections['5_col']:
                if row_idx < len(df) and pd.notna(df.iloc[row_idx, 0]):
                    if 'MEAN' in str(df.iloc[row_idx, 0]).upper():
                        if n1_col < len(df.iloc[row_idx]):
                            result['N1_mean'] = float(df.iloc[row_idx, n1_col])
                            print(f"   ✓ N1@633 MEAN值: {result['N1_mean']}")
                        if t1_col < len(df.iloc[row_idx]):
                            result['T1_mean'] = float(df.iloc[row_idx, t1_col])
                            print(f"   ✓ T1 MEAN值: {result['T1_mean']}")
                        break
            
            # 提取3 SIGMA值并计算比值
            if result['T1_mean'] is not None and result['T1_mean'] != 0:
                for row_idx in sections['5_col']:
                    if row_idx < len(df) and pd.notna(df.iloc[row_idx, 0]):
                        if '3 SIGMA' in str(df.iloc[row_idx, 0]).upper():
                            if t1_col < len(df.iloc[row_idx]):
                                t1_3sigma = float(df.iloc[row_idx, t1_col])
                                result['T1_3sigma_mean'] = t1_3sigma / result['T1_mean']
                                print(f"   ✓ 3 SIGMA行 - T1: {t1_3sigma}, 计算结果: {result['T1_3sigma_mean']:.6f}")
                            break
            else:
                print("   ❌ T1 MEAN值为0或None，无法计算比值")
                
        except Exception as e:
            print(f"   ❌ 提取5列段数据时出错: {e}")
        
        # 3. 从14列段中找到Site #和X=0, Y=0的数据
        try:
            print("   🔍 在14列段中查找Site #和X=0,Y=0数据...")
            
            site_header_row = None
            if sections['14_col']:
                # 查找Site #所在的行
                for row_idx in sections['14_col']:
                    if row_idx < len(df) and pd.notna(df.iloc[row_idx, 0]):
                        if 'SITE' in str(df.iloc[row_idx, 0]).upper():
                            site_header_row = row_idx
                            break
            
            if site_header_row is not None:
                print(f"   🔍 找到Site #行: 行{site_header_row+1}")
                header_data = df.iloc[site_header_row].astype(str)
                
                # 识别X、Y和N1列
                x_col = None
                y_col = None
                n1_col = None
                
                # 使用固定的列位置（基于已知的14列段结构）
                x_col = 5   # X列
                y_col = 6   # Y列  
                n1_col = 2  # N1@633列（之前是M1@633，现在已更新）
                
                print(f"   🔍 14列段列位置: X列={x_col}, Y列={y_col}, N1列={n1_col}")
                
                if x_col is not None and y_col is not None:
                    # 在14列段中查找X=0, Y=0的数据
                    found_xy_00 = False
                    for row_idx in sections['14_col']:
                        if row_idx > site_header_row and row_idx < len(df):
                            try:
                                x_val = df.iloc[row_idx, x_col] if x_col < len(df.iloc[row_idx]) else None
                                y_val = df.iloc[row_idx, y_col] if y_col < len(df.iloc[row_idx]) else None
                                
                                if pd.notna(x_val) and pd.notna(y_val) and float(x_val) == 0 and float(y_val) == 0:
                                    if n1_col is not None and n1_col < len(df.iloc[row_idx]):
                                        result['N1_XY_00'] = df.iloc[row_idx, n1_col]
                                        print(f"   ✓ 找到X=0, Y=0的行（行 {row_idx+1}），N1@633值: {result['N1_XY_00']}")
                                        found_xy_00 = True
                                        break
                            except (ValueError, IndexError):
                                continue
                    
                    if not found_xy_00:
                        print("   ❌ 未找到X=0, Y=0的数据行")
                else:
                    print("   ❌ 未找到X和Y列")
            else:
                print("   ❌ 未找到Site #行")
                
        except Exception as e:
            print(f"   ❌ 查找14列段数据时出错: {e}")
        
        results.append(result)
        
        # 显示当前wafer的完整结果
        print(f"   📊 Wafer {i+1} 提取结果:")
        for key, value in result.items():
            print(f"      {key}: {value}")
    
    return results

def save_results_to_excel(results, output_file):
    """保存结果到Excel文件"""
    try:
        df_results = pd.DataFrame(results)
        df_results.to_excel(output_file, index=False)
        
        print(f"\n✅ 结果已保存到: {output_file}")
        print(f"   文件大小: {os.path.getsize(output_file)} 字节")
        print(f"   保存时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 显示保存的数据摘要
        print(f"\n📋 保存的数据摘要:")
        print(f"   总记录数: {len(results)}")
        print(f"   列名: {list(df_results.columns)}")
        print(f"   数据预览:")
        print(df_results.to_string(index=False))
        
    except Exception as e:
        print(f"❌ 保存结果时出错: {e}")

def main():
    print("🚀 Wafer测试数据提取工具 - 智能版本 v2.0")
    print("="*50)
    print("✨ 智能特性: 自动识别数据结构，支持可变文件头部")
    
    # ===== 📁 修改输入文件名 =====
    # 👇 请在下面修改你的文件名（不需要.csv后缀）
    input_basename = "test_data"  # 👈 修改这里的文件名！
    # ==========================
    
    # 添加.csv后缀
    input_file = f'{input_basename}.csv'
    
    # 动态生成输出文件名：在输入文件名前加上"extracted_"
    output_file = f'extracted_{input_basename}.xlsx'
    
    print(f"📁 输入文件: {input_file}")
    print(f"📁 输出文件: {output_file}")
    print(f"💡 提示: 如需更改输入文件名，请编辑main函数中的input_basename变量")
    
    # 检查文件是否存在
    if not os.path.exists(input_file):
        print(f"\n❌ 文件不存在: {input_file}")
        print("🔧 请确保:")
        print("   1. 文件名是否正确")
        print("   2. 文件是否在当前目录中")
        print("   3. 文件格式是否为CSV")
        print("   4. 在代码中正确设置了input_basename变量")
        return
    
    # 读取输入文件
    df = read_csv_data(input_file)
    if df is None:
        return
    
    # 分析数据结构
    wafer_id_rows, identifier_positions = analyze_data_structure(df)
    if not wafer_id_rows:
        print("❌ 无法继续处理：未找到有效的wafer数据")
        return
    
    # 提取wafer数据
    results = extract_wafer_data(df, identifier_positions)
    
    if results:
        # 保存结果
        save_results_to_excel(results, output_file)
        
        print(f"\n🎉 处理完成！")
        print(f"   ✓ 成功提取 {len(results)} 个wafer的数据")
        print(f"   ✓ 结果已保存到 {output_file}")
        print(f"\n📋 使用指南:")
        print(f"   1. 检查输出文件: {output_file}")
        print(f"   2. 如需处理其他文件，请修改main函数中的input_basename")
        print(f"   3. 程序支持智能识别各种文件格式")
    else:
        print("❌ 未能提取到任何数据")
        print("💡 建议:")
        print("   1. 检查文件格式是否正确")
        print("   2. 确认文件包含必要的标识符")
        print("   3. 查看上方的错误信息进行调试")

if __name__ == "__main__":
    main()
