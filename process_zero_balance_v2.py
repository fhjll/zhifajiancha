# -*- coding: utf-8 -*-
"""
零余额账户延迟清算积数检测 — 独立版（v2）

功能：
  分析零余额账户中正向非清算交易（退款）与负向清算交易（垫款户扣款）的匹配情况，
  找出未在当天或第二个工作日内完成清算的交易，计算"积数"（金额 × 延迟天数）。

这与 process_zero_balance.py 中的 detect_delayed_settlement() 功能相同，
但作为独立脚本保留，可直接运行：
  python process_zero_balance_v2.py

数据流：
  零余额账户.xlsx（多 Sheet，每个 Sheet 对应一个账号）
    └─→ 读取每个 Sheet，提取正负交易
    └─→ 正向非清算交易（退款入账）→ 等待匹配
    └─→ 负向清算交易（向垫款户付款）→ 用于匹配
    └─→ 逐条正向交易，在负向交易中找同金额匹配
    └─→ 未在1个工作日内匹配的 → 计算积数
    └─→ 结果输出到 零余额账户延迟清算积数.xlsx
"""

import pandas as pd
from datetime import datetime, timedelta
import numpy as np


def parse_date(d):
    """将交易日期转为 datetime 对象，支持多种格式"""
    if pd.isna(d):
        return None
    if isinstance(d, datetime):
        return d
    if isinstance(d, (int, float)):
        s = str(int(d))
        return datetime.strptime(s, '%Y%m%d')
    if isinstance(d, str):
        s = d.replace('-', '').replace('/', '')
        return datetime.strptime(s, '%Y%m%d')
    return None


def is_weekend(dt):
    """判断是否为周末（周六=5, 周日=6）"""
    return dt.weekday() >= 5


def next_workday(dt):
    """获取下一个工作日（跳过周六、周日）"""
    nxt = dt + timedelta(days=1)
    while is_weekend(nxt):
        nxt += timedelta(days=1)
    return nxt


def is_within_same_or_next_workday(date1, date2):
    """
    判断 date2 是否是 date1 的当天或第二个工作日。

    规则：当天清算 → 正常；次日清算（跳过周末）→ 正常；
    超过次日（且非周末）→ 延迟。
    """
    if date1 == date2:
        return True
    if next_workday(date1) == date2:
        return True
    return False


# =========================
# 主流程
# =========================

# 1. 读取零余额账户Excel（所有 Sheet）
xls = pd.ExcelFile('零余额账户.xlsx')
sheet_names = xls.sheet_names

results = []

for sheet_name in sheet_names:
    # --- 1a. 读取 Sheet 数据 ---
    # 零余额账户 Excel 标准格式：表头在第 9 行（header=9）
    # 列约定：交易日期、交易金额、对方户名、借贷标志
    df = pd.read_excel(xls, sheet_name=sheet_name, header=9)
    df = df.dropna(subset=['交易日期'])
    if len(df) == 0:
        continue

    # --- 1b. 解析日期 ---
    df['日期对象'] = df['交易日期'].apply(parse_date)
    df = df.dropna(subset=['日期对象'])

    # --- 1c. 分离正向交易和负向交易 ---
    # 正向非清算交易：金额>0，对方户名≠垫款户 → 代表"退款入账"
    pos_mask = (df['交易金额'] > 0) & (df['对方户名'] != '集中支付零余额清算待转')
    pos_df = df[pos_mask].copy()

    # 负向清算交易：金额<0，对方户名=垫款户 → 代表"向垫款户划转资金完成清算"
    neg_mask = (df['交易金额'] < 0) & (df['对方户名'] == '集中支付零余额清算待转')
    neg_df = df[neg_mask].copy()

    if len(pos_df) == 0:
        continue

    # --- 1d. 将负向交易按金额索引，并标记是否已被使用 ---
    # 一条负向交易只能匹配一条正向交易（1:1 金额精确匹配）
    neg_by_amount = {}
    for idx, row in neg_df.iterrows():
        amt = abs(row['交易金额'])
        if amt not in neg_by_amount:
            neg_by_amount[amt] = []
        neg_by_amount[amt].append({
            'index': idx,
            'row': row,
            'used': False,  # 是否已被某条正向交易匹配
        })

    # 每个金额组内的交易按日期排序
    for amt in neg_by_amount:
        neg_by_amount[amt].sort(key=lambda r: r['row']['日期对象'])

    # --- 1e. 遍历正向交易，按日期排序后逐条匹配 ---
    pos_list = []
    for _, pos_row in pos_df.iterrows():
        pos_list.append(pos_row)
    pos_list.sort(key=lambda r: r['日期对象'])

    for pos_row in pos_list:
        pos_date = pos_row['日期对象']
        pos_amt = pos_row['交易金额']

        # 根据金额索引查找可匹配的负向交易
        matched_negs = neg_by_amount.get(pos_amt, [])

        if not matched_negs:
            # 无同金额负向交易 → 跳过（无法匹配）
            continue

        # 检查当天或第二个工作日内是否有未使用的匹配
        immediate_match = None      # 及时匹配（当天/次日）
        delayed_matches = []        # 延迟匹配（更晚日期）

        for neg_item in matched_negs:
            if neg_item['used']:
                continue  # 已被其他正向交易使用
            neg_date = neg_item['row']['日期对象']
            if is_within_same_or_next_workday(pos_date, neg_date):
                immediate_match = neg_item
                break  # 找到最早及时匹配，停止搜索
            elif neg_date > pos_date:
                delayed_matches.append(neg_item)

        # --- 1f. 判断匹配结果 ---
        if immediate_match is not None:
            # 当天或第二个工作日有匹配 → 正常清算，跳过
            immediate_match['used'] = True
            continue

        # 没有及时匹配 → 记录为延迟清算
        final_neg = None
        for neg_item in delayed_matches:
            if not neg_item['used']:
                final_neg = neg_item
                break  # 取最早的未使用延迟匹配

        if final_neg:
            final_neg['used'] = True
            final_neg_date = final_neg['row']['日期对象']
            days_diff = (final_neg_date - pos_date).days
            jishu = pos_amt * days_diff  # 积数 = 金额 × 延迟天数

            results.append({
                'sheet名称': sheet_name,
                '交易日期': pos_date.strftime('%Y-%m-%d'),
                '交易金额': pos_amt,
                '清算日期': final_neg_date.strftime('%Y-%m-%d'),
                '日期间隔': days_diff,
                '积数': jishu,
                '对方户名': pos_row['对方户名'],
                '交易描述': pos_row['交易描述']
            })

# =========================
# 结果输出
# =========================

# 按积数从大到小排序（积数越大，违规越严重）
results_df = pd.DataFrame(results)
if len(results_df) > 0:
    results_df = results_df.sort_values('积数', ascending=False)
    results_df.to_excel('零余额账户延迟清算积数.xlsx', index=False)
    print(f"共找到 {len(results_df)} 条延迟清算记录")
    print(f"结果已保存到: 零余额账户延迟清算积数.xlsx")
    print(f"\n前15条记录:")
    print(results_df[['sheet名称', '交易日期', '交易金额', '清算日期',
                      '日期间隔', '积数']].head(15).to_string(index=False))
else:
    print("没有找到延迟清算记录")
