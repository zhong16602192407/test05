#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
比较CSV文件和Excel文件中的企业名称和手机号（优化版）
"""

import pandas as pd
import re
from collections import defaultdict
from difflib import SequenceMatcher

def normalize_phone(phone):
    """标准化电话号码"""
    if pd.isna(phone):
        return None
    phone_str = str(phone).strip()
    # 移除空格、破折号等
    phone_str = re.sub(r'[\s\-\(\)]', '', phone_str)
    # 移除前导的+966或0
    phone_str = re.sub(r'^(\+966|966|0+)', '', phone_str)
    if phone_str and len(phone_str) >= 7:
        return phone_str
    return None

def normalize_company_name(name):
    """标准化公司名称"""
    if pd.isna(name):
        return None
    name_str = str(name).strip()
    # 移除常见的公司后缀词
    name_str = re.sub(r'(شركة|مصنع|فرع|المحدودة|للصناعة|للتجارة|التجارية)', '', name_str)
    name_str = name_str.strip()
    return name_str if name_str else None

def similar(a, b):
    """计算字符串相似度"""
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

# 读取CSV文件
print("正在读取CSV文件...")
companysa_df = pd.read_csv('companysa_companies.csv')
eyeofriyadh_df = pd.read_csv('eyeofriyadh_contacts.csv')
findsaudi_df = pd.read_csv('findsaudi_companies.csv')

# 读取Excel文件
print("正在读取Excel文件...")
excel_df = pd.read_excel('未爬取的沙特企业.xlsx')

print(f"\n文件信息:")
print(f"companysa_companies.csv: {len(companysa_df)} 条记录")
print(f"eyeofriyadh_contacts.csv: {len(eyeofriyadh_df)} 条记录")
print(f"findsaudi_companies.csv: {len(findsaudi_df)} 条记录")
print(f"未爬取的沙特企业.xlsx: {len(excel_df)} 条记录")

# 合并所有CSV数据到字典中以便快速查找
print("\n正在建立索引...")
csv_companies_by_name = defaultdict(list)
csv_companies_by_phone = defaultdict(list)

def add_to_index(source, company_name, phone, row_data):
    """添加到索引"""
    if not company_name or str(company_name) == 'nan':
        return

    company_data = {
        'source': source,
        'company_name': company_name,
        'normalized_name': normalize_company_name(company_name),
        'phone': phone,
        'raw_data': row_data
    }

    # 按名称索引
    csv_companies_by_name[company_name].append(company_data)
    normalized = normalize_company_name(company_name)
    if normalized:
        csv_companies_by_name[normalized].append(company_data)

    # 按电话索引
    if phone:
        csv_companies_by_phone[phone].append(company_data)

# 从companysa添加
for idx, row in companysa_df.iterrows():
    company_name = row.get('company_name', '')
    phone = normalize_phone(row.get('phone_number', ''))
    add_to_index('companysa', company_name, phone, row.to_dict())

# 从eyeofriyadh添加
for idx, row in eyeofriyadh_df.iterrows():
    company_name = row.get('name', '')
    phone = normalize_phone(row.get('phone', ''))
    add_to_index('eyeofriyadh', company_name, phone, row.to_dict())

# 从findsaudi添加
for idx, row in findsaudi_df.iterrows():
    company_name = row.get('company_name', '')
    phone = normalize_phone(row.get('phone_number', ''))
    add_to_index('findsaudi', company_name, phone, row.to_dict())

print(f"索引建立完成，共 {len(csv_companies_by_name)} 个不同的公司名称")

# 进行匹配
matches = []
excel_name_col = '企业名称'

print("\n开始匹配（使用优化算法）...")
match_count = 0
total = len(excel_df)

for excel_idx, excel_row in excel_df.iterrows():
    if excel_idx % 1000 == 0:
        print(f"进度: {excel_idx}/{total} ({excel_idx*100//total}%)")

    excel_company = str(excel_row[excel_name_col]).strip()

    if not excel_company or excel_company == 'nan':
        continue

    excel_company_normalized = normalize_company_name(excel_company)

    # 先尝试精确匹配
    matched = set()

    # 1. 完全匹配
    if excel_company in csv_companies_by_name:
        for csv_data in csv_companies_by_name[excel_company]:
            key = (csv_data['source'], csv_data['company_name'])
            if key not in matched:
                matched.add(key)
                matches.append({
                    'Excel公司名称': excel_company,
                    'CSV来源': csv_data['source'],
                    'CSV公司名称': csv_data['company_name'],
                    'CSV电话': csv_data['phone'],
                    '名称相似度': 1.0,
                    '匹配类型': '完全匹配'
                })
                match_count += 1

    # 2. 标准化后匹配
    if excel_company_normalized and excel_company_normalized in csv_companies_by_name:
        for csv_data in csv_companies_by_name[excel_company_normalized]:
            key = (csv_data['source'], csv_data['company_name'])
            if key not in matched:
                matched.add(key)
                matches.append({
                    'Excel公司名称': excel_company,
                    'CSV来源': csv_data['source'],
                    'CSV公司名称': csv_data['company_name'],
                    'CSV电话': csv_data['phone'],
                    '名称相似度': 0.95,
                    '匹配类型': '标准化匹配'
                })
                match_count += 1

print(f"\n匹配完成!")

# 创建结果DataFrame
matches_df = pd.DataFrame(matches)

if len(matches_df) > 0:
    print(f"\n找到 {len(matches_df)} 条匹配记录")

    # 按来源统计
    print("\n按CSV来源统计:")
    print(matches_df['CSV来源'].value_counts())

    # 按匹配类型统计
    print("\n按匹配类型统计:")
    print(matches_df['匹配类型'].value_counts())

    # 保存结果
    output_file = '企业匹配结果.xlsx'
    matches_df.to_excel(output_file, index=False)
    print(f"\n结果已保存到: {output_file}")

    # 显示前30条匹配结果
    print("\n前30条匹配结果:")
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 50)
    print(matches_df.head(30))

    # 显示带电话的记录
    has_phone = matches_df[matches_df['CSV电话'].notna()]
    print(f"\n有电话号码的匹配记录: {len(has_phone)} 条")
    if len(has_phone) > 0:
        print("\n带电话号码的匹配示例:")
        print(has_phone.head(20))
else:
    print("\n未找到匹配记录")

print("\n分析完成!")
