#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
比较CSV文件和Excel文件中的企业名称和手机号
"""

import pandas as pd
import re
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

# 打印Excel文件的列名
print(f"\nExcel文件列名: {excel_df.columns.tolist()}")

# 打印CSV文件的列名
print(f"companysa列名: {companysa_df.columns.tolist()}")
print(f"eyeofriyadh列名: {eyeofriyadh_df.columns.tolist()}")
print(f"findsaudi列名: {findsaudi_df.columns.tolist()}")

# 尝试识别公司名称列和电话列
excel_name_col = None
excel_phone_col = None

for col in excel_df.columns:
    col_lower = str(col).lower()
    if '名称' in col or 'name' in col_lower or '公司' in col or 'company' in col_lower:
        excel_name_col = col
    if '电话' in col or 'phone' in col_lower or '手机' in col or 'mobile' in col_lower or '联系' in col:
        excel_phone_col = col

print(f"\nExcel文件识别的列:")
print(f"公司名称列: {excel_name_col}")
print(f"电话列: {excel_phone_col}")

if not excel_name_col:
    print("\n警告: 未能自动识别Excel文件的公司名称列，使用第一列")
    excel_name_col = excel_df.columns[0]

# 合并所有CSV数据
all_csv_companies = []

# 从companysa添加
for idx, row in companysa_df.iterrows():
    company_name = row.get('company_name', '')
    phone = normalize_phone(row.get('phone_number', ''))
    if company_name:
        all_csv_companies.append({
            'source': 'companysa',
            'company_name': company_name,
            'normalized_name': normalize_company_name(company_name),
            'phone': phone,
            'raw_data': row.to_dict()
        })

# 从eyeofriyadh添加
for idx, row in eyeofriyadh_df.iterrows():
    # 尝试找到公司名称列
    company_name = row.get('company_name', row.get('name', row.get('Company Name', '')))
    phone = normalize_phone(row.get('phone', row.get('mobile', row.get('Phone', ''))))
    if company_name:
        all_csv_companies.append({
            'source': 'eyeofriyadh',
            'company_name': company_name,
            'normalized_name': normalize_company_name(company_name),
            'phone': phone,
            'raw_data': row.to_dict()
        })

# 从findsaudi添加
for idx, row in findsaudi_df.iterrows():
    company_name = row.get('company_name', row.get('name', row.get('Company Name', '')))
    phone = normalize_phone(row.get('phone', row.get('mobile', row.get('Phone', ''))))
    if company_name:
        all_csv_companies.append({
            'source': 'findsaudi',
            'company_name': company_name,
            'normalized_name': normalize_company_name(company_name),
            'phone': phone,
            'raw_data': row.to_dict()
        })

print(f"\n总计CSV记录数: {len(all_csv_companies)}")

# 进行匹配
matches = []
similarity_threshold = 0.6  # 相似度阈值

print("\n开始匹配...")
for excel_idx, excel_row in excel_df.iterrows():
    excel_company = str(excel_row[excel_name_col]).strip()
    excel_company_normalized = normalize_company_name(excel_company)
    excel_phone = normalize_phone(excel_row.get(excel_phone_col, '')) if excel_phone_col else None

    if not excel_company or excel_company == 'nan':
        continue

    # 寻找匹配
    for csv_company in all_csv_companies:
        csv_name = csv_company['company_name']
        csv_normalized = csv_company['normalized_name']
        csv_phone = csv_company['phone']

        # 完全匹配
        exact_match = (excel_company == csv_name)

        # 计算相似度
        similarity = similar(excel_company, csv_name)
        normalized_similarity = similar(excel_company_normalized, csv_normalized) if excel_company_normalized and csv_normalized else 0

        max_similarity = max(similarity, normalized_similarity)

        # 电话匹配
        phone_match = False
        if excel_phone and csv_phone:
            phone_match = (excel_phone == csv_phone)

        # 如果名称高度相似或电话匹配，则认为是同一家公司
        if exact_match or max_similarity >= similarity_threshold or phone_match:
            matches.append({
                'Excel公司名称': excel_company,
                'Excel电话': excel_phone,
                'CSV来源': csv_company['source'],
                'CSV公司名称': csv_name,
                'CSV电话': csv_phone,
                '名称相似度': round(max_similarity, 2),
                '电话匹配': '是' if phone_match else '否',
                '匹配类型': '完全匹配' if exact_match else ('电话匹配' if phone_match else '相似匹配')
            })

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

    # 电话匹配统计
    phone_matched = matches_df[matches_df['电话匹配'] == '是']
    print(f"\n电话号码匹配的记录: {len(phone_matched)} 条")

    # 保存结果
    output_file = '企业匹配结果.xlsx'
    matches_df.to_excel(output_file, index=False)
    print(f"\n结果已保存到: {output_file}")

    # 显示前20条匹配结果
    print("\n前20条匹配结果:")
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 50)
    print(matches_df.head(20))

    # 显示有电话匹配的记录
    if len(phone_matched) > 0:
        print("\n有电话匹配的记录示例:")
        print(phone_matched.head(10))
else:
    print("\n未找到匹配记录")

print("\n分析完成!")
