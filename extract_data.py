import pandas as pd
import numpy as np
import os
import sys
from datetime import datetime

def read_excel_data(file_path):
    """读取Excel文件并返回DataFrame"""
    try:
        if not os.path.exists(file_path):
            print(f"错误：文件 {file_path} 不存在")
            return None
        
        df = pd.read_excel(file_path, header=None)
        print(f"✓ 成功读取文件 {file_path}")
        print(f"  数据形状: {df.shape[0]} 行 × {df.shape[1]} 列")
        return df
    except Exception as e:
        print(f"❌ 读取Excel文件时出错: {e}")
        return None

def analyze_data_structure(df):
    """分析数据结构"""
    print("\n" + "="*50)
    print("📊 数据结构分析")
    print("="*50)
    
    # 查找所有WAFER ID行
    wafer_id_rows = df[df.iloc[:, 0] == 'WAFER ID'].index.tolist()
    print(f"发现 {len(wafer_id_rows)} 个wafer数据块")
    
    if wafer_id_rows:
        print(f"WAFER ID位于行: {wafer_id_rows}")
        
        # 显示第一个wafer的数据结构示例
        if len(wafer_id_rows) > 0:
            wafer_start = wafer_id_rows[0]
            wafer_end = wafer_id_rows[1] if len(wafer_id_rows) > 1 else len(df)
            
            print(f"\n📋 第一个wafer数据结构 (行 {wafer_start} 到 {wafer_end-1}):")
            
            # 查找关键行
            key_rows = ['SLOT', 'MEAN', '3SIGMA', 'Site #']
            for key in key_rows:
                rows = df[df.iloc[:, 0] == key].index.tolist()
                if rows:
                    row_idx = rows[0]
                    if wafer_start <= row_idx < wafer_end:
                        print(f"  {key}: 行 {row_idx}")
    
    return wafer_id_rows

def extract_wafer_data(df):
    """提取wafer数据"""
    print("\n" + "="*50)
    print("⚙️  开始提取wafer数据")
    print("="*50)
    
    results = []
    
    # 查找所有WAFER ID行
    wafer_id_rows = df[df.iloc[:, 0] == 'WAFER ID'].index.tolist()
    
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
            slot_rows = wafer_data[wafer_data.iloc[:, 0] == 'SLOT'].index.tolist()
            if slot_rows:
                slot_row = slot_rows[0]
                result['lot_id'] = df.iloc[slot_row, 1]
                print(f"   ✓ SLOT值: {result['lot_id']}")
            else:
                print("   ❌ 未找到SLOT行")
        except Exception as e:
            print(f"   ❌ 提取SLOT值时出错: {e}")
        
        # 2. 找到MEAN行，获取C列（N1@633）和E列（T1）的值
        try:
            mean_rows = wafer_data[wafer_data.iloc[:, 0] == 'MEAN'].index.tolist()
            if mean_rows:
                mean_row = mean_rows[0]
                result['N1_mean'] = df.iloc[mean_row, 2]  # C列（索引2）对应N1@633
                result['T1_mean'] = df.iloc[mean_row, 4]  # E列（索引4）对应T1
                print(f"   ✓ MEAN行 - N1@633: {result['N1_mean']}, T1: {result['T1_mean']}")
            else:
                print("   ❌ 未找到MEAN行")
        except Exception as e:
            print(f"   ❌ 提取MEAN值时出错: {e}")
        
        # 3. 找到3SIGMA行，计算E列（T1）的值除以MEAN行的E列（T1）的值
        try:
            sigma_rows = wafer_data[wafer_data.iloc[:, 0] == '3SIGMA'].index.tolist()
            if sigma_rows and result['T1_mean'] is not None:
                sigma_row = sigma_rows[0]
                t1_3sigma = df.iloc[sigma_row, 4]  # E列（索引4）对应T1
                if pd.notna(t1_3sigma) and pd.notna(result['T1_mean']) and result['T1_mean'] != 0:
                    result['T1_3sigma_mean'] = t1_3sigma / result['T1_mean']
                    print(f"   ✓ 3SIGMA行 - T1: {t1_3sigma}, 计算结果: {result['T1_3sigma_mean']:.6f}")
                else:
                    print("   ❌ 3SIGMA或MEAN的T1值无效")
            else:
                print("   ❌ 未找到3SIGMA行或MEAN的T1值")
        except Exception as e:
            print(f"   ❌ 计算T1_3sigma_mean时出错: {e}")
        
        # 4. 从"Site #"开始往下，找到F列（X）和G列（Y）都为0的那一行
        try:
            site_rows = wafer_data[wafer_data.iloc[:, 0] == 'Site #'].index.tolist()
            if site_rows:
                site_start = site_rows[0]
                print(f"   🔍 从Site #行（行 {site_start}）开始查找X=0, Y=0的数据")
                
                found_xy_00 = False
                for row_idx in range(site_start + 1, wafer_end):
                    if row_idx < len(df):
                        x_val = df.iloc[row_idx, 5]  # F列（索引5）对应X
                        y_val = df.iloc[row_idx, 6]  # G列（索引6）对应Y
                        if pd.notna(x_val) and pd.notna(y_val) and x_val == 0 and y_val == 0:
                            result['N1_XY_00'] = df.iloc[row_idx, 2]  # C列（索引2）对应N1@633
                            print(f"   ✓ 找到X=0, Y=0的行（行 {row_idx}），N1@633值: {result['N1_XY_00']}")
                            found_xy_00 = True
                            break
                
                if not found_xy_00:
                    print("   ❌ 未找到X=0, Y=0的数据行")
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
    print("🚀 Wafer测试数据提取工具")
    print("="*50)
    
    # 检查命令行参数
    if len(sys.argv) < 2:
        print("❌ 请提供输入文件名（不需要.xlsx后缀）")
        print("使用方法: python extract_data.py <文件名>")
        print("例如: python extract_data.py test_data")
        return
    
    # 获取输入文件名并添加.xlsx后缀
    input_basename = sys.argv[1]
    input_file = f'{input_basename}.xlsx'
    
    # 动态生成输出文件名：在输入文件名前加上"extracted_"
    output_file = f'extracted_{input_basename}.xlsx'
    
    print(f"📁 输入文件: {input_file}")
    print(f"📁 输出文件: {output_file}")
    
    # 读取输入文件
    df = read_excel_data(input_file)
    if df is None:
        return
    
    # 分析数据结构
    wafer_id_rows = analyze_data_structure(df)
    if not wafer_id_rows:
        print("❌ 无法继续处理：未找到有效的wafer数据")
        return
    
    # 提取wafer数据
    results = extract_wafer_data(df)
    
    if results:
        # 保存结果
        save_results_to_excel(results, output_file)
        
        print(f"\n🎉 处理完成！")
        print(f"   ✓ 成功提取 {len(results)} 个wafer的数据")
        print(f"   ✓ 结果已保存到 {output_file}")
    else:
        print("❌ 未能提取到任何数据")

if __name__ == "__main__":
    main()
