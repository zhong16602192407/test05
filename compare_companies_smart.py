#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能比较CSV文件和Excel文件中的企业名称和手机号
"""

import pandas as pd
import re
from collections import defaultdict
from difflib import SequenceMatcher

def clean_company_name(name):
    """清理企业名称，去除编号、括号等"""
    if pd.isna(name):
        return None
    name_str = str(name).strip()

    # 去除括号及其内容（包括编号）
    name_str = re.sub(r'\([^)]*\)', '', name_str)
    name_str = re.sub(r'\[[^\]]*\]', '', name_str)

    # 去除常见的公司后缀词
    name_str = re.sub(r'\b(LTD|LIMITED|LLC|INC|CORP|CO|COMPANY|EST|ESTABLISHMENT)\b', '', name_str, flags=re.IGNORECASE)
    name_str = re.sub(r'(شركة|مصنع|فرع|المحدودة|للصناعة|للتجارة|التجارية)', '', name_str)

    # 去除多余的空格
    name_str = re.sub(r'\s+', ' ', name_str)
    name_str = name_str.strip()

    return name_str.upper() if name_str else None

def get_name_keywords(name):
    """提取企业名称中的关键词"""
    if not name:
        return set()
    # 分词并过滤掉太短的词
    words = re.findall(r'\b\w{3,}\b', name.upper())
    return set(words)

def normalize_phone(phone):
    """标准化电话号码"""
    if pd.isna(phone):
        return None
    phone_str = str(phone).strip()
    phone_str = re.sub(r'[\s\-\(\)]', '', phone_str)
    phone_str = re.sub(r'^(\+966|966|0+)', '', phone_str)
    if phone_str and len(phone_str) >= 7:
        return phone_str
    return None

def similarity_score(name1, name2):
    """计算两个名称的相似度"""
    if not name1 or not name2:
        return 0.0

    # 基本相似度
    basic_sim = SequenceMatcher(None, name1, name2).ratio()

    # 关键词匹配度
    keywords1 = get_name_keywords(name1)
    keywords2 = get_name_keywords(name2)

    if keywords1 and keywords2:
        common_keywords = keywords1 & keywords2
        keyword_sim = len(common_keywords) / max(len(keywords1), len(keywords2))
    else:
        keyword_sim = 0.0

    # 综合得分
    return max(basic_sim, keyword_sim)

# 读取数据
print("正在读取文件...")
companysa_df = pd.read_csv('companysa_companies.csv')
eyeofriyadh_df = pd.read_csv('eyeofriyadh_contacts.csv')
findsaudi_df = pd.read_csv('findsaudi_companies.csv')
excel_df = pd.read_excel('未爬取的沙特企业.xlsx')

print(f"\n文件信息:")
print(f"companysa: {len(companysa_df)} 条")
print(f"eyeofriyadh: {len(eyeofriyadh_df)} 条")
print(f"findsaudi: {len(findsaudi_df)} 条")
print(f"Excel: {len(excel_df)} 条")

# 建立CSV企业索引
print("\n正在建立索引...")
csv_companies = []

# 从companysa添加
for idx, row in companysa_df.iterrows():
    name = row.get('company_name', '')
    if name and str(name) != 'nan':
        csv_companies.append({
            'source': 'companysa',
            'original_name': name,
            'cleaned_name': clean_company_name(name),
            'phone': normalize_phone(row.get('phone_number', '')),
            'keywords': get_name_keywords(clean_company_name(name))
        })

# 从eyeofriyadh添加
for idx, row in eyeofriyadh_df.iterrows():
    name = row.get('name', '')
    if name and str(name) != 'nan':
        csv_companies.append({
            'source': 'eyeofriyadh',
            'original_name': name,
            'cleaned_name': clean_company_name(name),
            'phone': normalize_phone(row.get('phone', '')),
            'keywords': get_name_keywords(clean_company_name(name))
        })

# 从findsaudi添加
for idx, row in findsaudi_df.iterrows():
    name = row.get('company_name', '')
    if name and str(name) != 'nan':
        csv_companies.append({
            'source': 'findsaudi',
            'original_name': name,
            'cleaned_name': clean_company_name(name),
            'phone': normalize_phone(row.get('phone_number', '')),
            'keywords': get_name_keywords(clean_company_name(name))
        })

print(f"索引建立完成，共 {len(csv_companies)} 条CSV记录")

# 按关键词建立索引以加速搜索
keyword_index = defaultdict(list)
for idx, company in enumerate(csv_companies):
    for keyword in company['keywords']:
        keyword_index[keyword].append(idx)

print(f"关键词索引包含 {len(keyword_index)} 个不同的关键词")

# 进行匹配
print("\n开始智能匹配...")
matches = []
threshold = 0.5  # 相似度阈值

for excel_idx, excel_row in excel_df.iterrows():
    if excel_idx % 1000 == 0:
        print(f"进度: {excel_idx}/{len(excel_df)} ({excel_idx*100//len(excel_df)}%)")

    excel_name = str(excel_row['企业名称']).strip()
    if not excel_name or excel_name == 'nan':
        continue

    excel_cleaned = clean_company_name(excel_name)
    if not excel_cleaned:
        continue

    excel_keywords = get_name_keywords(excel_cleaned)

    # 找出可能的候选
    candidates = set()
    for keyword in excel_keywords:
        if keyword in keyword_index:
            candidates.update(keyword_index[keyword])

    # 如果没有关键词匹配，跳过（避免不必要的比较）
    if not candidates:
        continue

    # 对候选进行详细比较
    best_matches = []
    for csv_idx in candidates:
        csv_company = csv_companies[csv_idx]
        csv_cleaned = csv_company['cleaned_name']

        if not csv_cleaned:
            continue

        # 计算相似度
        sim = similarity_score(excel_cleaned, csv_cleaned)

        if sim >= threshold:
            best_matches.append({
                'csv_idx': csv_idx,
                'similarity': sim
            })

    # 添加最佳匹配
    if best_matches:
        # 按相似度排序，只保留最好的匹配
        best_matches.sort(key=lambda x: x['similarity'], reverse=True)

        for match in best_matches[:3]:  # 最多保留前3个匹配
            csv_company = csv_companies[match['csv_idx']]
            matches.append({
                'Excel公司名称': excel_name,
                'Excel清理后': excel_cleaned,
                'CSV来源': csv_company['source'],
                'CSV公司名称': csv_company['original_name'],
                'CSV清理后': csv_company['cleaned_name'],
                'CSV电话': csv_company['phone'],
                '相似度': round(match['similarity'], 3)
            })

print(f"\n匹配完成!")

# 生成结果
matches_df = pd.DataFrame(matches)

if len(matches_df) > 0:
    print(f"\n找到 {len(matches_df)} 条匹配记录")

    # 统计
    print("\n按CSV来源统计:")
    print(matches_df['CSV来源'].value_counts())

    print("\n按相似度区间统计:")
    print(pd.cut(matches_df['相似度'], bins=[0, 0.6, 0.7, 0.8, 0.9, 1.0]).value_counts().sort_index())

    # 有电话的记录
    has_phone = matches_df[matches_df['CSV电话'].notna()]
    print(f"\n有电话号码的匹配: {len(has_phone)} 条")

    # 保存结果
    output_file = '企业匹配结果_智能版.xlsx'
    matches_df.to_excel(output_file, index=False)
    print(f"\n结果已保存到: {output_file}")

    # 显示示例
    print("\n相似度最高的20条匹配:")
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 40)
    top_matches = matches_df.nlargest(20, '相似度')
    print(top_matches)

    if len(has_phone) > 0:
        print("\n有电话号码的匹配示例（前20条）:")
        print(has_phone.nlargest(20, '相似度')[['Excel公司名称', 'CSV来源', 'CSV公司名称', 'CSV电话', '相似度']])
else:
    print("\n未找到匹配记录")

print("\n分析完成!")
