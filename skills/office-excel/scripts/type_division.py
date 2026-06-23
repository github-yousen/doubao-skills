#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用文本内容分类统计脚本
"""

import pandas as pd
import re
from collections import Counter

# 读取数据
df = pd.read_excel('/path/to/input_file.xlsx')

# 定义分类关键词映射
# 基于对文本内容的分析，定义分类规则

def extract_specific_issues(content, case_type):
    """
    从文本内容中提取分类事项
    返回分类事项列表（一条文本可能包含多个事项）
    """
    issues = []
    content = str(content)
    
    # ============ 类别1相关问题 ============
    if any(kw in content for kw in ['关键词1', '关键词2', '关键词3', '关键词4', '关键词5', '关键词6']):
        issues.append('类别1-子类别1')
    elif any(kw in content for kw in ['关键词7', '关键词8', '关键词9', '关键词10', '关键词11']):
        issues.append('类别1-子类别2')
    elif any(kw in content for kw in ['关键词12', '关键词13', '关键词14']):
        if '关键词1' not in content and '关键词2' not in content:
            issues.append('类别1-子类别3')
    elif '关键词15' in content and not issues:
        # 其他类别1相关
        if '关键词16' in content:
            issues.append('类别1-子类别4')
        elif '关键词17' in content or '关键词18' in content:
            issues.append('类别1-子类别5')
        else:
            issues.append('类别1-子类别1')
    
    # ============ 类别2相关问题 ============
    if any(kw in content for kw in ['关键词19', '关键词20']):
        if any(kw in content for kw in ['关键词21', '关键词22', '关键词23', '关键词24', '关键词25']):
            issues.append('类别2-子类别1')
        elif '关键词26' in content:
            issues.append('类别2-子类别2')
        else:
            issues.append('类别2-子类别3')
    
    # ============ 类别3相关问题 ============
    if any(kw in content for kw in ['关键词27', '关键词28', '关键词29']):
        if any(kw in content for kw in ['关键词21', '关键词22', '关键词23', '关键词24']):
            issues.append('类别3-子类别1')
        elif '关键词30' in content or '关键词31' not in content and '关键词32' in content:
            issues.append('类别3-子类别2')
        else:
            issues.append('类别3-子类别3')
    
    # ============ 类别4相关问题 ============
    if '关键词33' in content:
        if any(kw in content for kw in ['关键词21', '关键词22', '关键词23']):
            issues.append('类别4-子类别1')
        elif '关键词34' in content:
            issues.append('类别4-子类别2')
        elif '关键词35' in content or '关键词36' in content:
            issues.append('类别4-子类别3')
        else:
            issues.append('类别4-子类别4')
    
    # ============ 类别5相关问题 ============
    if any(kw in content for kw in ['关键词37', '关键词38', '关键词39']):
        if any(kw in content for kw in ['关键词21', '关键词22', '关键词23', '关键词24']):
            issues.append('类别5-子类别1')
        else:
            issues.append('类别5-子类别2')
    
    # ============ 类别6相关问题 ============
    if any(kw in content for kw in ['关键词40', '关键词41', '关键词42']):
        issues.append('类别6-子类别1')
    
    # ============ 类别7相关问题 ============
    if any(kw in content for kw in ['关键词43', '关键词44', '关键词45']):
        if any(kw in content for kw in ['关键词46', '关键词47', '关键词22', '关键词23', '关键词24']):
            issues.append('类别7-子类别1')
        else:
            issues.append('类别7-子类别2')
    
    # ============ 类别8相关问题 ============
    if any(kw in content for kw in ['关键词48', '关键词49', '关键词50']):
        if '关键词51' in content or '关键词52' in content:
            issues.append('类别8-子类别1')
        else:
            issues.append('类别8-子类别2')
    
    # ============ 类别9相关问题 ============
    if any(kw in content for kw in ['关键词53', '关键词54', '关键词55']):
        if '关键词56' in content or '关键词57' in content or '关键词58' in content:
            issues.append('类别9-子类别1')
        elif '关键词59' in content:
            if '关键词37' not in str(issues):
                issues.append('类别9-子类别2')
    
    # ============ 类别10相关问题 ============
    if any(kw in content for kw in ['关键词60', '关键词61', '关键词62']):
        issues.append('类别10-子类别1')
    if any(kw in content for kw in ['关键词63', '关键词64']):
        if '关键词65' in content or '关键词66' in content or '关键词67' in content:
            issues.append('类别10-子类别2')
    
    # ============ 类别11相关问题 ============
    if any(kw in content for kw in ['关键词68', '关键词69', '关键词70']):
        issues.append('类别11-子类别1')
    
    # ============ 类别12相关问题 ============
    if any(kw in content for kw in ['关键词71', '关键词72', '关键词73']):
        if '关键词74' in content:
            issues.append('类别12-子类别1')
        else:
            issues.append('类别12-子类别2')
    
    # ============ 类别13相关问题 ============
    if '关键词75' in content:
        if '关键词76' in content or '关键词77' in content:
            issues.append('类别13-子类别1')
        elif '关键词78' in content or '关键词79' in content:
            issues.append('类别13-子类别2')
    
    # ============ 类别14相关问题 ============
    if any(kw in content for kw in ['关键词80', '关键词81']):
        issues.append('类别14-子类别1')
    
    # ============ 类别15相关问题 ============
    if any(kw in content for kw in ['关键词82', '关键词83', '关键词84']):
        issues.append('类别15-子类别1')
    
    # ============ 类别16相关问题 ============
    if any(kw in content for kw in ['关键词85', '关键词86', '关键词87']):
        issues.append('类别16-子类别1')
    
    # ============ 类别17相关问题 ============
    if '关键词88' in content and '关键词15' not in content:
        issues.append('类别17-子类别1')
    
    # ============ 类别18相关问题 ============
    if any(kw in content for kw in ['关键词89', '关键词90', '关键词91']):
        issues.append('类别18-子类别1')
    
    # ============ 类别19相关问题 ============
    if any(kw in content for kw in ['关键词92', '关键词93', '关键词94']):
        issues.append('类别19-子类别1')
    
    # ============ 类别20相关问题 ============
    if any(kw in content for kw in ['关键词95', '关键词96', '关键词97', '关键词98']):
        issues.append('类别20-子类别1')
    
    # ============ 类别21相关问题 ============
    if '关键词99' in content and '关键词15' not in content:
        issues.append('类别21-子类别1')
    
    # ============ 类别22相关问题 ============
    if '关键词100' in content and not issues:
        if '关键词15' in content:
            issues.append('类别22-子类别1')
        elif '关键词33' in content:
            issues.append('类别22-子类别2')
        else:
            issues.append('类别22-子类别3')
    
    # ============ 类别23相关问题 ============
    if any(kw in content for kw in ['关键词101', '关键词102', '关键词103']):
        issues.append('类别23-子类别1')
    
    # 如果没有识别出具体事项，根据案件类型给出默认分类
    if not issues:
        if case_type == '类型1':
            issues.append('类别24-子类别1')
        elif case_type == '类型2':
            issues.append('类别24-子类别2')
        elif case_type == '类型3':
            issues.append('类别24-子类别3')
        elif case_type == '类型4':
            issues.append('类别24-子类别4')
        elif case_type == '类型5':
            issues.append('类别24-子类别5')
        elif case_type == '类型6':
            issues.append('类别24-子类别6')
        elif case_type == '类型7':
            issues.append('类别24-子类别7')
        elif case_type == '类型8':
            issues.append('类别24-子类别8')
        else:
            issues.append('类别25')
    
    return issues


# 提取所有文本的具体事项
all_issues = []
issue_details = []

for idx, row in df.iterrows():
    content = row['文本内容']
    case_type = row['类型']
    issues = extract_specific_issues(content, case_type)
    
    for issue in issues:
        all_issues.append(issue)
        issue_details.append({
            '序号': row['序号'],
            '类型': case_type,
            '具体事项': issue,
            '文本内容摘要': str(content)[:100] + '...' if len(str(content)) > 100 else content
        })

# 统计具体事项
issue_counts = Counter(all_issues)

# 输出统计结果
print("=" * 60)
print("文本内容分类统计")
print("=" * 60)
print(f"\n总文本记录数: {len(df)}")
print(f"识别出的具体事项总数: {len(all_issues)}")
print(f"具体事项类别数: {len(issue_counts)}")

print("\n" + "-" * 60)
print("具体事项分类统计（按数量降序排列）")
print("-" * 60)

for issue, count in issue_counts.most_common():
    print(f"{issue}: {count}条")

# 创建详细的DataFrame
details_df = pd.DataFrame(issue_details)

# 保存结果
details_df.to_excel('/path/to/output_details.xlsx', index=False)

# 创建统计汇总表
summary_data = []
for issue, count in issue_counts.most_common():
    summary_data.append({
        '具体事项': issue,
        '数量': count,
        '占比': f'{count/len(all_issues)*100:.1f}%'
    })

summary_df = pd.DataFrame(summary_data)
summary_df.to_excel('/path/to/output_summary.xlsx', index=False)

print("\n" + "=" * 60)
print("结果已保存到:")
print("1. /path/to/output_details.xlsx")
print("2. /path/to/output_summary.xlsx")
print("=" * 60)