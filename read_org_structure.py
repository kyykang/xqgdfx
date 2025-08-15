#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
读取启用组织.xlsx文件，分析部门层级结构
用于理解部门编码和上级部门关系
"""

import pandas as pd
import json

def read_organization_structure():
    """
    读取启用组织.xlsx文件，分析部门结构
    """
    try:
        # 读取Excel文件
        df = pd.read_excel('启用组织.xlsx')
        
        print("=== 启用组织.xlsx 文件结构 ===")
        print(f"总行数: {len(df)}")
        print(f"列名: {list(df.columns)}")
        print("\n=== 前10行数据 ===")
        print(df.head(10))
        
        print("\n=== 数据类型 ===")
        print(df.dtypes)
        
        print("\n=== 数据统计 ===")
        print(df.describe())
        
        # 如果有部门相关的列，显示唯一值
        for col in df.columns:
            if any(keyword in str(col).lower() for keyword in ['部门', 'dept', '编码', 'code', '名称', 'name']):
                print(f"\n=== {col} 列的唯一值数量: {df[col].nunique()} ===")
                if df[col].nunique() < 50:  # 如果唯一值不太多，显示所有
                    print(df[col].unique())
                else:  # 否则只显示前20个
                    print("前20个值:")
                    print(df[col].unique()[:20])
        
        return df
        
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None

if __name__ == "__main__":
    df = read_organization_structure()