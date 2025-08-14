#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

# 创建测试Excel文件
def create_test_excel():
    # 生成测试数据
    departments = ['技术部', '市场部', '人事部', '财务部', '运营部', '客服部', '产品部', '设计部']
    ticket_types = ['功能需求', '系统优化', '数据分析', '界面改进', '流程优化', '技术支持']
    statuses = ['已完成', '进行中', '待处理', '已取消']
    systems = ['OA系统', '营销平台', 'U8C系统']
    
    # 生成随机日期（2023-2024年）
    start_date = datetime(2023, 1, 1)
    end_date = datetime(2024, 12, 31)
    
    data = []
    for i in range(100):  # 生成100条测试数据
        # 随机日期
        random_days = random.randint(0, (end_date - start_date).days)
        ticket_date = start_date + timedelta(days=random_days)
        
        data.append({
            '工单编号': f'T{2023}{i+1:04d}',
            '提交日期': ticket_date.strftime('%Y-%m-%d'),
            '部门': random.choice(departments),
            '工单类型': random.choice(ticket_types),
            '工单状态': random.choice(statuses),
            '系统': random.choice(systems),
            '描述': f'测试工单{i+1}的详细描述',
            '处理人': f'处理人{random.randint(1, 10)}',
            '优先级': random.choice(['高', '中', '低'])
        })
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存为Excel文件
    test_file = '测试工单数据.xlsx'
    df.to_excel(test_file, index=False, sheet_name='工单数据')
    
    print(f"测试Excel文件已创建: {test_file}")
    print(f"包含 {len(df)} 条测试数据")
    print("\n数据预览:")
    print(df.head())
    
    return test_file

if __name__ == '__main__':
    create_test_excel()