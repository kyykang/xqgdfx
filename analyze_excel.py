#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
需求工单数据分析脚本
用于读取Excel文件并分析数据结构
"""

import pandas as pd
import json

def analyze_excel_data(file_path):
    """
    分析Excel文件的数据结构
    
    Args:
        file_path (str): Excel文件路径
    
    Returns:
        dict: 包含数据分析结果的字典
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 基本信息
        analysis_result = {
            "总行数": len(df),
            "总列数": len(df.columns),
            "列名": list(df.columns),
            "数据类型": {col: str(dtype) for col, dtype in df.dtypes.items()},
            "前5行数据": df.head().to_dict('records'),
            "缺失值统计": {col: int(count) for col, count in df.isnull().sum().items()}
        }
        
        # 如果有部门相关列，进行部门统计
        dept_columns = [col for col in df.columns if '部门' in col or 'department' in col.lower()]
        if dept_columns:
            analysis_result["部门统计"] = {}
            for col in dept_columns:
                analysis_result["部门统计"][col] = df[col].value_counts().to_dict()
        
        # 如果有年度相关列，进行年度统计
        year_columns = [col for col in df.columns if '年' in col or 'year' in col.lower() or '时间' in col or 'date' in col.lower()]
        if year_columns:
            analysis_result["年度统计"] = {}
            for col in year_columns:
                try:
                    # 尝试提取年份信息
                    if df[col].dtype == 'object':
                        # 如果是字符串类型，尝试转换为日期
                        dates = pd.to_datetime(df[col], errors='coerce')
                        years = dates.dt.year.value_counts().to_dict()
                        analysis_result["年度统计"][col] = years
                    else:
                        analysis_result["年度统计"][col] = df[col].value_counts().to_dict()
                except:
                    analysis_result["年度统计"][col] = "无法解析"
        
        # 查找可能的系统相关列（OA系统、营销平台、u8c）
        system_keywords = ['OA', 'oa', '营销', '平台', 'u8c', 'U8C', '系统']
        system_columns = []
        for col in df.columns:
            for keyword in system_keywords:
                if keyword in col:
                    system_columns.append(col)
                    break
        
        if system_columns:
            analysis_result["系统相关列"] = {}
            for col in system_columns:
                analysis_result["系统相关列"][col] = df[col].value_counts().to_dict()
        
        return analysis_result
        
    except Exception as e:
        return {"错误": str(e)}

if __name__ == "__main__":
    # Excel文件路径
    excel_file = "/Users/kangyiyuan/Desktop/AI编程项目/需求工单分析2/需求工单统计表.xlsx"
    
    # 分析数据
    result = analyze_excel_data(excel_file)
    
    # 输出结果
    print("=== 需求工单数据分析结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    
    # 保存分析结果到文件
    with open("data_analysis_result.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print("\n分析结果已保存到 data_analysis_result.json 文件")