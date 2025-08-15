#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
需求工单数据处理器
用于处理Excel数据并生成网站所需的JSON数据
"""

import pandas as pd
import json
from datetime import datetime
import re
from collections import Counter

def clean_department_name(dept_name):
    """
    清理部门名称，去除括号及其内容
    例如："项目管理三部(哈尔滨安天系统安全技术有限公司)" -> "项目管理三部"
    """
    if pd.isna(dept_name) or dept_name == '':
        return dept_name
    
    # 使用正则表达式去除括号及其内容
    cleaned_name = re.sub(r'\([^)]*\)', '', str(dept_name))
    return cleaned_name.strip()

def load_organization_structure():
    """
    读取启用组织.xlsx文件，构建部门层级映射关系
    返回：部门编码到部门名称的映射，以及部门到一级部门的映射
    """
    try:
        # 读取组织结构文件
        org_df = pd.read_excel('启用组织.xlsx')
        
        # 创建编码到名称的映射
        code_to_name = {}
        # 创建部门名称到一级部门的映射
        dept_to_top_level = {}
        
        # 第一步：建立完整的编码到名称的映射
        for index, row in org_df.iterrows():
            if index == 0:  # 跳过标题行
                continue
                
            dept_name = str(row['Name']).strip()
            dept_code = str(row['Code']).strip()
            
            # 建立编码到名称的映射
            if dept_code and dept_name:
                code_to_name[dept_code] = dept_name
        
        # 第二步：建立部门到一级部门的映射
        for index, row in org_df.iterrows():
            if index == 0:  # 跳过标题行
                continue
                
            dept_name = str(row['Name']).strip()
            dept_code = str(row['Code']).strip()
            parent_code = str(row['PDepartmentCode']).strip() if pd.notna(row['PDepartmentCode']) else None
            
            # 如果没有上级部门编码，说明是一级部门
            if not parent_code or parent_code == 'nan':
                dept_to_top_level[dept_name] = dept_name
            else:
                # 递归查找一级部门
                top_level_dept = find_top_level_department(dept_name, org_df, code_to_name)
                dept_to_top_level[dept_name] = top_level_dept
        
        print(f"成功加载组织结构，共{len(dept_to_top_level)}个部门")
        return code_to_name, dept_to_top_level
        
    except Exception as e:
        print(f"加载组织结构文件失败: {e}")
        return {}, {}

def find_top_level_department(dept_name, org_df, code_to_name):
    """
    递归查找部门的一级部门
    """
    # 查找当前部门的编码
    current_code = None
    for index, row in org_df.iterrows():
        if str(row['Name']).strip() == dept_name:
            current_code = str(row['Code']).strip()
            parent_code = str(row['PDepartmentCode']).strip() if pd.notna(row['PDepartmentCode']) else None
            
            # 如果没有上级部门，说明是一级部门
            if not parent_code or parent_code == 'nan':
                return dept_name
            
            # 如果有上级部门，继续向上查找
            parent_name = code_to_name.get(parent_code, None)
            if parent_name:
                return find_top_level_department(parent_name, org_df, code_to_name)
            else:
                return dept_name
    
    # 如果找不到，返回原名称
    return dept_name

def map_to_top_level_department(dept_name, dept_to_top_level):
    """
    将部门名称映射到一级部门
    """
    if not dept_name or pd.isna(dept_name):
        return dept_name
    
    # 先清理部门名称
    cleaned_name = clean_department_name(dept_name)
    
    # 查找映射
    return dept_to_top_level.get(cleaned_name, cleaned_name)

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
    
    # 加载组织结构映射
    code_to_name, dept_to_top_level = load_organization_structure()
    
    # 清理部门名称，去除括号内容
    df['所在部门'] = df['所在部门'].apply(clean_department_name)
    
    # 映射到一级部门
    df['一级部门'] = df['所在部门'].apply(lambda x: map_to_top_level_department(x, dept_to_top_level))
    
    print(f"部门映射完成，共映射到 {len(df['一级部门'].unique())} 个一级部门")
    
    # 处理日期列
    df['创建日期'] = pd.to_datetime(df['创建日期'], errors='coerce')
    df['年份'] = df['创建日期'].dt.year
    
    # 1. 按部门统计工单数量（Top 10）
    dept_counts = df['所在部门'].value_counts().head(10)
    dept_top10 = {
        'labels': dept_counts.index.tolist(),
        'data': dept_counts.values.tolist()
    }
    
    # 1.0 按一级部门统计工单数量（完整列表，用于模态框显示）
    dept_counts_all = df['一级部门'].value_counts()
    dept_all = {
        'labels': dept_counts_all.index.tolist(),
        'data': dept_counts_all.values.tolist()
    }
    
    # 1.0.1 按原始部门统计工单数量（完整列表，用于原始部门显示）
    original_dept_counts_all = df['所在部门'].value_counts()
    original_dept_all = {
        'labels': original_dept_counts_all.index.tolist(),
        'data': original_dept_counts_all.values.tolist()
    }
    
    # 保留原始部门统计的前10名（用于兼容性）
    dept_counts_top10 = df['一级部门'].value_counts().head(10)
    dept_top10 = {
        'labels': dept_counts_top10.index.tolist(),
        'data': dept_counts_top10.values.tolist()
    }
    
    # 1.1 按年度分组的一级部门统计
    dept_by_year = {}
    dept_by_year_all = {}  # 完整的年度部门统计
    original_dept_by_year_all = {}  # 完整的年度原始部门统计
    for year in df['年份'].dropna().unique():
        year_data = df[df['年份'] == year]
        year_dept_counts = year_data['一级部门'].value_counts().head(10)
        dept_by_year[str(int(year))] = {
            'labels': year_dept_counts.index.tolist(),
            'data': year_dept_counts.values.tolist()
        }
        
        # 完整的年度部门统计（不限制数量）
        year_dept_counts_all = year_data['一级部门'].value_counts()
        dept_by_year_all[str(int(year))] = {
            'labels': year_dept_counts_all.index.tolist(),
            'data': year_dept_counts_all.values.tolist()
        }
        
        # 完整的年度原始部门统计（不限制数量）
        year_original_dept_counts_all = year_data['所在部门'].value_counts()
        original_dept_by_year_all[str(int(year))] = {
            'labels': year_original_dept_counts_all.index.tolist(),
            'data': year_original_dept_counts_all.values.tolist()
        }
    
    # 1.2 排除草稿的一级部门统计
    df_no_draft = df[df['审核状态'] != '草稿']
    dept_counts_no_draft = df_no_draft['一级部门'].value_counts().head(10)
    dept_top10_no_draft = {
        'labels': dept_counts_no_draft.index.tolist(),
        'data': dept_counts_no_draft.values.tolist()
    }
    
    # 1.2.1 排除草稿的一级部门统计（完整列表）
    dept_counts_all_no_draft = df_no_draft['一级部门'].value_counts()
    dept_all_no_draft = {
        'labels': dept_counts_all_no_draft.index.tolist(),
        'data': dept_counts_all_no_draft.values.tolist()
    }
    
    # 1.2.2 排除草稿的原始部门统计（完整列表）
    original_dept_counts_all_no_draft = df_no_draft['所在部门'].value_counts()
    original_dept_all_no_draft = {
        'labels': original_dept_counts_all_no_draft.index.tolist(),
        'data': original_dept_counts_all_no_draft.values.tolist()
    }
    
    # 1.3 按年度分组的一级部门统计（排除草稿）
    dept_by_year_no_draft = {}
    dept_by_year_all_no_draft = {}  # 完整的年度部门统计（排除草稿）
    original_dept_by_year_all_no_draft = {}  # 完整的年度原始部门统计（排除草稿）
    for year in df_no_draft['年份'].dropna().unique():
        year_data = df_no_draft[df_no_draft['年份'] == year]
        year_dept_counts = year_data['一级部门'].value_counts().head(10)
        dept_by_year_no_draft[str(int(year))] = {
            'labels': year_dept_counts.index.tolist(),
            'data': year_dept_counts.values.tolist()
        }
        
        # 完整的年度部门统计（排除草稿，不限制数量）
        year_dept_counts_all = year_data['一级部门'].value_counts()
        dept_by_year_all_no_draft[str(int(year))] = {
            'labels': year_dept_counts_all.index.tolist(),
            'data': year_dept_counts_all.values.tolist()
        }
        
        # 完整的年度原始部门统计（排除草稿，不限制数量）
        year_original_dept_counts_all = year_data['所在部门'].value_counts()
        original_dept_by_year_all_no_draft[str(int(year))] = {
            'labels': year_original_dept_counts_all.index.tolist(),
            'data': year_original_dept_counts_all.values.tolist()
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
    
    # 2.1 按年度分组的系统统计
    system_by_year = {}
    for year in df['年份'].dropna().unique():
        year_data = df[df['年份'] == year]
        oa_year = len(year_data[year_data['OA系统'] == '勾选'])
        marketing_year = len(year_data[year_data['营销平台'] == '勾选'])
        u8c_year = len(year_data[year_data['U8C'] == '勾选'])
        system_by_year[str(int(year))] = {
            'OA系统': oa_year,
            '营销平台': marketing_year,
            'U8C': u8c_year
        }
    
    # 2.2 系统统计（排除草稿）
    oa_count_no_draft = len(df_no_draft[df_no_draft['OA系统'] == '勾选'])
    marketing_count_no_draft = len(df_no_draft[df_no_draft['营销平台'] == '勾选'])
    u8c_count_no_draft = len(df_no_draft[df_no_draft['U8C'] == '勾选'])
    
    system_stats_no_draft = {
        'OA系统': oa_count_no_draft,
        '营销平台': marketing_count_no_draft,
        'U8C': u8c_count_no_draft
    }
    
    # 2.3 按年度分组的系统统计（排除草稿）
    system_by_year_no_draft = {}
    for year in df_no_draft['年份'].dropna().unique():
        year_data = df_no_draft[df_no_draft['年份'] == year]
        oa_year = len(year_data[year_data['OA系统'] == '勾选'])
        marketing_year = len(year_data[year_data['营销平台'] == '勾选'])
        u8c_year = len(year_data[year_data['U8C'] == '勾选'])
        system_by_year_no_draft[str(int(year))] = {
            'OA系统': oa_year,
            '营销平台': marketing_year,
            'U8C': u8c_year
        }
    
    # 3. 按年度统计工单数量
    year_counts = df['年份'].value_counts().sort_index()
    year_stats = {
        'labels': [str(int(year)) for year in year_counts.index if pd.notna(year)],
        'data': [int(count) for year, count in year_counts.items() if pd.notna(year)]
    }
    
    # 3.1 按年度统计工单数量（排除草稿）
    year_counts_no_draft = df_no_draft['年份'].value_counts().sort_index()
    year_stats_no_draft = {
        'labels': [str(int(year)) for year in year_counts_no_draft.index if pd.notna(year)],
        'data': [int(count) for year, count in year_counts_no_draft.items() if pd.notna(year)]
    }
    
    # 4. 工单类型统计
    type_counts = df['工单类型'].value_counts()
    type_stats = {
        'labels': type_counts.index.tolist(),
        'data': type_counts.values.tolist()
    }
    
    # 4.1 按年度分组的工单类型统计
    type_by_year = {}
    for year in df['年份'].dropna().unique():
        year_data = df[df['年份'] == year]
        year_type_counts = year_data['工单类型'].value_counts()
        type_by_year[str(int(year))] = {
            'labels': year_type_counts.index.tolist(),
            'data': year_type_counts.values.tolist()
        }
    
    # 4.2 工单类型统计（排除草稿）
    type_counts_no_draft = df_no_draft['工单类型'].value_counts()
    type_stats_no_draft = {
        'labels': type_counts_no_draft.index.tolist(),
        'data': type_counts_no_draft.values.tolist()
    }
    
    # 4.3 按年度分组的工单类型统计（排除草稿）
    type_by_year_no_draft = {}
    for year in df_no_draft['年份'].dropna().unique():
        year_data = df_no_draft[df_no_draft['年份'] == year]
        year_type_counts = year_data['工单类型'].value_counts()
        type_by_year_no_draft[str(int(year))] = {
            'labels': year_type_counts.index.tolist(),
            'data': year_type_counts.values.tolist()
        }
    
    # 5. 工单状态统计
    status_counts = df['流程状态'].value_counts()
    status_stats = {
        'labels': status_counts.index.tolist(),
        'data': status_counts.values.tolist()
    }
    
    # 5.1 按年度分组的工单状态统计
    status_by_year = {}
    for year in df['年份'].dropna().unique():
        year_data = df[df['年份'] == year]
        year_status_counts = year_data['流程状态'].value_counts()
        status_by_year[str(int(year))] = {
            'labels': year_status_counts.index.tolist(),
            'data': year_status_counts.values.tolist()
        }
    
    # 5.2 审核状态统计
    audit_counts = df['审核状态'].value_counts()
    audit_stats = {
        'labels': audit_counts.index.tolist(),
        'data': audit_counts.values.tolist()
    }
    
    # 5.3 按年度分组的审核状态统计
    audit_by_year = {}
    for year in df['年份'].dropna().unique():
        year_data = df[df['年份'] == year]
        year_audit_counts = year_data['审核状态'].value_counts()
        audit_by_year[str(int(year))] = {
            'labels': year_audit_counts.index.tolist(),
            'data': year_audit_counts.values.tolist()
        }
    
    # 6. 月度趋势分析
    df['年月'] = df['创建日期'].dt.to_period('M')
    monthly_counts = df['年月'].value_counts().sort_index()
    monthly_stats = {
        'labels': [str(period) for period in monthly_counts.index if pd.notna(period)],
        'data': [int(count) for period, count in monthly_counts.items() if pd.notna(period)]
    }
    
    # 6.1 按年度分组的月度趋势分析
    monthly_by_year = {}
    for year in df['年份'].dropna().unique():
        year_data = df[df['年份'] == year]
        year_data['年月'] = year_data['创建日期'].dt.to_period('M')
        year_monthly_counts = year_data['年月'].value_counts().sort_index()
        monthly_by_year[str(int(year))] = {
            'labels': [str(period) for period in year_monthly_counts.index if pd.notna(period)],
            'data': [int(count) for period, count in year_monthly_counts.items() if pd.notna(period)]
        }
    
    # 6.2 月度趋势分析（排除草稿）
    df_no_draft['年月'] = df_no_draft['创建日期'].dt.to_period('M')
    monthly_counts_no_draft = df_no_draft['年月'].value_counts().sort_index()
    monthly_stats_no_draft = {
        'labels': [str(period) for period in monthly_counts_no_draft.index if pd.notna(period)],
        'data': [int(count) for period, count in monthly_counts_no_draft.items() if pd.notna(period)]
    }
    
    # 6.3 按年度分组的月度趋势分析（排除草稿）
    monthly_by_year_no_draft = {}
    for year in df_no_draft['年份'].dropna().unique():
        year_data = df_no_draft[df_no_draft['年份'] == year]
        year_data['年月'] = year_data['创建日期'].dt.to_period('M')
        year_monthly_counts = year_data['年月'].value_counts().sort_index()
        monthly_by_year_no_draft[str(int(year))] = {
            'labels': [str(period) for period in year_monthly_counts.index if pd.notna(period)],
            'data': [int(count) for period, count in year_monthly_counts.items() if pd.notna(period)]
        }
    
    # 提取未结束工单的详细信息
    unfinished_tickets = df[df['流程状态'] == '未结束'][['流水号', '需求内容', '申请人', '所在部门', '创建日期', '工单类型', '审核状态']].copy()
    unfinished_tickets['创建日期'] = unfinished_tickets['创建日期'].dt.strftime('%Y-%m-%d')
    unfinished_tickets_list = unfinished_tickets.to_dict('records')
    
    # 按年度分组的未结束工单详细信息
    unfinished_by_year = {}
    for year in df['年份'].dropna().unique():
        year_unfinished = df[(df['年份'] == year) & (df['流程状态'] == '未结束')][['流水号', '需求内容', '申请人', '所在部门', '创建日期', '工单类型', '审核状态']].copy()
        year_unfinished['创建日期'] = year_unfinished['创建日期'].dt.strftime('%Y-%m-%d')
        unfinished_by_year[str(int(year))] = year_unfinished.to_dict('records')
    
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
        'dept_all': dept_all,  # 完整的部门统计
        'original_dept_all': original_dept_all,  # 完整的原始部门统计
        'dept_by_year': dept_by_year,
        'dept_by_year_all': dept_by_year_all,  # 完整的年度部门统计
        'original_dept_by_year_all': original_dept_by_year_all,  # 完整的年度原始部门统计
        'dept_top10_no_draft': dept_top10_no_draft,
        'dept_all_no_draft': dept_all_no_draft,  # 完整的部门统计（排除草稿）
        'original_dept_all_no_draft': original_dept_all_no_draft,  # 完整的原始部门统计（排除草稿）
        'dept_by_year_no_draft': dept_by_year_no_draft,
        'dept_by_year_all_no_draft': dept_by_year_all_no_draft,  # 完整的年度部门统计（排除草稿）
        'original_dept_by_year_all_no_draft': original_dept_by_year_all_no_draft,  # 完整的年度原始部门统计（排除草稿）
        'system_stats': system_stats,
        'system_by_year': system_by_year,
        'system_stats_no_draft': system_stats_no_draft,
        'system_by_year_no_draft': system_by_year_no_draft,
        'year_stats': year_stats,
        'year_stats_no_draft': year_stats_no_draft,
        'type_stats': type_stats,
        'type_by_year': type_by_year,
        'type_stats_no_draft': type_stats_no_draft,
        'type_by_year_no_draft': type_by_year_no_draft,
        'status_stats': status_stats,
        'status_by_year': status_by_year,
        'audit_stats': audit_stats,
        'audit_by_year': audit_by_year,
        'monthly_stats': monthly_stats,
        'monthly_by_year': monthly_by_year,
        'monthly_stats_no_draft': monthly_stats_no_draft,
        'monthly_by_year_no_draft': monthly_by_year_no_draft,
        'unfinished_tickets': unfinished_tickets_list,
        'unfinished_by_year': unfinished_by_year
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