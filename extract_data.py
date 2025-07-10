"""
🚀 Wafer测试数据提取工具 - 智能版本 v2.0

✨ 新功能特点：
1. 🔍 智能识别数据结构 - 不依赖固定行号和列号
2. 📊 动态查找关键标识符 - 支持WAFER ID、SLOT、MEAN、3SIGMA、Site #
3. 🎯 自动定位数据列 - 智能匹配N1@633、T1、X、Y列
4. 📄 支持CSV格式输入 - 替代原有的Excel格式
5. 🔧 灵活的文件头部处理 - 适应不同格式的数据文件

🎨 使用方法：
1. 在main函数中修改input_basename变量为你的文件名
2. 确保你的CSV文件包含以下关键标识符：
   - WAFER ID：标识wafer数据块
   - SLOT：标识样品ID
   - MEAN：包含平均值数据
   - 3SIGMA：包含3sigma数据
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
        
        # 尝试不同的编码方式读取CSV文件
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin1']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, header=None, encoding=encoding)
                print(f"✓ 成功读取文件 {file_path} (编码: {encoding})")
                break
            except UnicodeDecodeError:
                continue
        
        if df is None:
            print(f"❌ 无法读取文件 {file_path}，尝试了多种编码方式")
            return None
        
        print(f"  数据形状: {df.shape[0]} 行 × {df.shape[1]} 列")
        return df
    except Exception as e:
        print(f"❌ 读取CSV文件时出错: {e}")
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
    key_identifiers = ['WAFER ID', 'SLOT', 'MEAN', '3SIGMA', 'Site #']
    
    print("🔍 搜索关键标识符...")
    identifier_positions = {}
    
    for identifier in key_identifiers:
        # 在所有列中搜索，不仅限于第一列
        positions = []
        for col in range(df.shape[1]):
            if col < df.shape[1]:
                try:
                    mask = df.iloc[:, col].astype(str).str.contains(identifier, case=False, na=False)
                    rows = df[mask].index.tolist()
                    if rows:
                        positions.extend([(row, col) for row in rows])
                except:
                    continue
        
        if positions:
            identifier_positions[identifier] = positions
            print(f"  {identifier}: 找到 {len(positions)} 个位置 - {positions}")
    
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
        
        # 提取当前wafer的数据区域
        wafer_data = df.iloc[wafer_start:wafer_end]
        
        # 初始化结果字典
        result = {
            'lot_id': None,
            'N1_mean': None,
            'T1_mean': None,
            'T1_3sigma_mean': None,
            'N1_XY_00': None
        }
        
        # 1. 找到SLOT值
        try:
            slot_positions = [pos for pos in identifier_positions.get('SLOT', []) 
                            if wafer_start <= pos[0] < wafer_end]
            if slot_positions:
                slot_row, slot_col = slot_positions[0]
                # SLOT值通常在标识符的下一列
                if slot_col + 1 < df.shape[1]:
                    result['lot_id'] = df.iloc[slot_row, slot_col + 1]
                    print(f"   ✓ SLOT值: {result['lot_id']} (位置: 行{slot_row}, 列{slot_col+1})")
                else:
                    print("   ❌ SLOT值列超出范围")
            else:
                print("   ❌ 未找到SLOT行")
        except Exception as e:
            print(f"   ❌ 提取SLOT值时出错: {e}")
        
        # 2. 找到MEAN行，智能识别N1@633和T1列
        try:
            mean_positions = [pos for pos in identifier_positions.get('MEAN', []) 
                            if wafer_start <= pos[0] < wafer_end]
            if mean_positions:
                mean_row, mean_col = mean_positions[0]
                print(f"   🔍 MEAN行位置: 行{mean_row}, 列{mean_col}")
                
                # 智能查找N1@633和T1列
                header_row = None
                for search_row in range(max(0, mean_row - 10), mean_row):
                    row_data = df.iloc[search_row].astype(str)
                    if any('N1' in str(cell) and '633' in str(cell) for cell in row_data):
                        header_row = search_row
                        break
                
                if header_row is not None:
                    header_data = df.iloc[header_row].astype(str)
                    n1_col = None
                    t1_col = None
                    
                    for col_idx, cell in enumerate(header_data):
                        if 'N1' in str(cell) and '633' in str(cell):
                            n1_col = col_idx
                        elif 'T1' in str(cell) and 'N1' not in str(cell):
                            t1_col = col_idx
                    
                    if n1_col is not None:
                        result['N1_mean'] = df.iloc[mean_row, n1_col]
                        print(f"   ✓ N1@633 MEAN值: {result['N1_mean']} (列{n1_col})")
                    if t1_col is not None:
                        result['T1_mean'] = df.iloc[mean_row, t1_col]
                        print(f"   ✓ T1 MEAN值: {result['T1_mean']} (列{t1_col})")
                else:
                    print("   ❌ 未找到数据列头")
            else:
                print("   ❌ 未找到MEAN行")
        except Exception as e:
            print(f"   ❌ 提取MEAN值时出错: {e}")
        
        # 3. 找到3SIGMA行，计算T1的3sigma/mean比值
        try:
            sigma_positions = [pos for pos in identifier_positions.get('3SIGMA', []) 
                             if wafer_start <= pos[0] < wafer_end]
            if sigma_positions and result['T1_mean'] is not None:
                sigma_row, sigma_col = sigma_positions[0]
                
                # 使用与MEAN行相同的T1列位置
                mean_positions = [pos for pos in identifier_positions.get('MEAN', []) 
                                if wafer_start <= pos[0] < wafer_end]
                if mean_positions:
                    mean_row, mean_col = mean_positions[0]
                    
                    # 找到T1列
                    header_row = None
                    for search_row in range(max(0, mean_row - 10), mean_row):
                        row_data = df.iloc[search_row].astype(str)
                        if any('T1' in str(cell) for cell in row_data):
                            header_row = search_row
                            break
                    
                    if header_row is not None:
                        header_data = df.iloc[header_row].astype(str)
                        t1_col = None
                        for col_idx, cell in enumerate(header_data):
                            if 'T1' in str(cell) and 'N1' not in str(cell):
                                t1_col = col_idx
                                break
                        
                        if t1_col is not None:
                            t1_3sigma = df.iloc[sigma_row, t1_col]
                            if pd.notna(t1_3sigma) and pd.notna(result['T1_mean']) and result['T1_mean'] != 0:
                                result['T1_3sigma_mean'] = t1_3sigma / result['T1_mean']
                                print(f"   ✓ 3SIGMA行 - T1: {t1_3sigma}, 计算结果: {result['T1_3sigma_mean']:.6f}")
                            else:
                                print("   ❌ 3SIGMA或MEAN的T1值无效")
                        else:
                            print("   ❌ 未找到T1列")
                    else:
                        print("   ❌ 未找到数据列头")
            else:
                print("   ❌ 未找到3SIGMA行或MEAN的T1值")
        except Exception as e:
            print(f"   ❌ 计算T1_3sigma_mean时出错: {e}")
        
        # 4. 从"Site #"开始往下，智能查找X=0, Y=0的数据
        try:
            site_positions = [pos for pos in identifier_positions.get('Site #', []) 
                            if wafer_start <= pos[0] < wafer_end]
            if site_positions:
                site_row, site_col = site_positions[0]
                print(f"   🔍 从Site #行（行 {site_row}）开始查找X=0, Y=0的数据")
                
                # 智能查找X和Y列
                header_data = df.iloc[site_row].astype(str)
                x_col = None
                y_col = None
                n1_col = None
                
                for col_idx, cell in enumerate(header_data):
                    cell_str = str(cell).upper()
                    if 'X' == cell_str:
                        x_col = col_idx
                    elif 'Y' == cell_str:
                        y_col = col_idx
                    elif 'N1' in cell_str and '633' in cell_str:
                        n1_col = col_idx
                
                if x_col is not None and y_col is not None:
                    print(f"   🔍 X列: {x_col}, Y列: {y_col}, N1列: {n1_col}")
                    
                    found_xy_00 = False
                    for row_idx in range(site_row + 1, wafer_end):
                        if row_idx < len(df):
                            x_val = df.iloc[row_idx, x_col]
                            y_val = df.iloc[row_idx, y_col]
                            if pd.notna(x_val) and pd.notna(y_val) and x_val == 0 and y_val == 0:
                                if n1_col is not None:
                                    result['N1_XY_00'] = df.iloc[row_idx, n1_col]
                                    print(f"   ✓ 找到X=0, Y=0的行（行 {row_idx}），N1@633值: {result['N1_XY_00']}")
                                    found_xy_00 = True
                                    break
                    
                    if not found_xy_00:
                        print("   ❌ 未找到X=0, Y=0的数据行")
                else:
                    print("   ❌ 未找到X和Y列")
            else:
                print("   ❌ 未找到Site #行")
        except Exception as e:
            print(f"   ❌ 查找X=0, Y=0数据时出错: {e}")
        
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
