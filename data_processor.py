#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
需求工单数据处理器
用于处理Excel数据并生成网站所需的JSON数据
"""

import pandas as pd
import json
from datetime import datetime
from collections import Counter

def process_ticket_data(excel_file):
    """
    处理需求工单数据
    
    Args:
        excel_file (str): Excel文件路径
    
    Returns:
        dict: 处理后的数据
    """
    # 读取Excel文件
    df = pd.read_excel(excel_file)
    
    # 跳过前两行（标题行和表头说明），从第三行开始是真实数据
    df = df.iloc[2:].reset_index(drop=True)
    
    # 重新设置列名
    df.columns = [
        '流水号', '申请人', '所在部门', '创建日期', '工单类型', 
        '工单类型子类型', 'OA系统', '营销平台', 'U8C', 
        '需求内容', '审核状态', '流程状态'
    ]
    
    # 清理数据：移除空行
    df = df.dropna(subset=['流水号'])
    
    # 处理日期列
    df['创建日期'] = pd.to_datetime(df['创建日期'], errors='coerce')
    df['年份'] = df['创建日期'].dt.year
    
    # 1. 按部门统计工单数量（Top 10）
    dept_counts = df['所在部门'].value_counts().head(10)
    dept_top10 = {
        'labels': dept_counts.index.tolist(),
        'data': dept_counts.values.tolist()
    }
    
    # 2. 统计各系统勾选情况
    oa_count = len(df[df['OA系统'] == '勾选'])
    marketing_count = len(df[df['营销平台'] == '勾选'])
    u8c_count = len(df[df['U8C'] == '勾选'])
    
    system_stats = {
        'OA系统': oa_count,
        '营销平台': marketing_count,
        'U8C': u8c_count
    }
    
    # 3. 按年度统计工单数量
    year_counts = df['年份'].value_counts().sort_index()
    year_stats = {
        'labels': [str(int(year)) for year in year_counts.index if pd.notna(year)],
        'data': [int(count) for year, count in year_counts.items() if pd.notna(year)]
    }
    
    # 4. 工单类型统计
    type_counts = df['工单类型'].value_counts()
    type_stats = {
        'labels': type_counts.index.tolist(),
        'data': type_counts.values.tolist()
    }
    
    # 5. 工单状态统计
    status_counts = df['流程状态'].value_counts()
    status_stats = {
        'labels': status_counts.index.tolist(),
        'data': status_counts.values.tolist()
    }
    
    # 6. 月度趋势分析
    df['年月'] = df['创建日期'].dt.to_period('M')
    monthly_counts = df['年月'].value_counts().sort_index()
    monthly_stats = {
        'labels': [str(period) for period in monthly_counts.index if pd.notna(period)],
        'data': [int(count) for period, count in monthly_counts.items() if pd.notna(period)]
    }
    
    # 汇总所有统计数据
    result = {
        'summary': {
            'total_tickets': len(df),
            'total_departments': df['所在部门'].nunique(),
            'date_range': {
                'start': df['创建日期'].min().strftime('%Y-%m-%d') if pd.notna(df['创建日期'].min()) else 'N/A',
                'end': df['创建日期'].max().strftime('%Y-%m-%d') if pd.notna(df['创建日期'].max()) else 'N/A'
            }
        },
        'dept_top10': dept_top10,
        'system_stats': system_stats,
        'year_stats': year_stats,
        'type_stats': type_stats,
        'status_stats': status_stats,
        'monthly_stats': monthly_stats
    }
    
    return result

def save_data_for_web(data, output_file):
    """
    保存数据为网站可用的JSON格式
    
    Args:
        data (dict): 处理后的数据
        output_file (str): 输出文件路径
    """
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"数据已保存到 {output_file}")

if __name__ == "__main__":
    # 处理数据
    excel_file = "/Users/kangyiyuan/Desktop/AI编程项目/需求工单分析2/需求工单统计表.xlsx"
    
    print("正在处理需求工单数据...")
    processed_data = process_ticket_data(excel_file)
    
    # 保存为JSON文件供网站使用
    save_data_for_web(processed_data, "ticket_data.json")
    
    # 打印基本统计信息
    print("\n=== 数据处理完成 ===")
    print(f"总工单数: {processed_data['summary']['total_tickets']}")
    print(f"涉及部门数: {processed_data['summary']['total_departments']}")
    print(f"数据时间范围: {processed_data['summary']['date_range']['start']} 至 {processed_data['summary']['date_range']['end']}")
    print(f"\n各系统勾选统计:")
    for system, count in processed_data['system_stats'].items():
        print(f"  {system}: {count} 个")