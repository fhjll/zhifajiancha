"""
零余额账户违规检测模块

提供两类检测功能：
  1. detect_fund_matching  — 资金匹配检测：从待查记录和零余额账户中匹配退款记录，计算各环节日期
  2. detect_delayed_settlement — 延迟清算积数检测：分析零余额账户中延迟结算的交易

输入：流水账户报表（Excel/CSV文件）
输出：检测结果的DataFrame
"""

import os
import json
import bisect
import pandas as pd
from datetime import datetime, timedelta
from typing import Optional
import numpy as np
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

# =========================
# 列名映射配置（从 JSON 文件加载，方便用户自定义）
# =========================

_CONFIG_CACHE = None

def _load_column_config():
    """加载 column_mapping.json 列名映射配置。

    查找顺序：
      1. 项目根目录下的 column_mapping.json
      2. 若文件不存在或格式错误，使用硬编码默认值

    返回 dict，包含 csv_column_aliases / amount_in_fen / preprocess_column_groups。
    """
    global _CONFIG_CACHE
    if _CONFIG_CACHE is not None:
        return _CONFIG_CACHE

    # 硬编码默认值（JSON 文件缺失时的回退）
    defaults = {
        "csv_column_aliases": {
            '交易日期': '交易日期',
            '入帐日期': '入帐日期',
            '入账日期': '入账日期',
            '日期': '交易日期',
            '交易金额': '交易金额',
            '发生额': '发生额',
            '发生额（分）': '发生额',
            '借贷标志': '借贷标志',
            '借贷标志（1-借 2-贷）': '借贷标志',
            '对方户名': '对方户名',
            '账号': '账号',
            '帐号': '账号',
            '对方账号': '对方账号',
            '对方帐号': '对方账号',
            '余额': '余额',
            '账户余额': '账户余额',
            '摘要': '摘要',
            '摘要信息': '摘要',
            '交易描述': '交易描述',
        },
        "amount_in_fen_columns": ["发生额（分）"],
        "preprocess_column_groups": {
            "account": ["账户", "账号", "帐号"],
            "date": ["交易日期", "入帐日期", "入账日期", "日期"],
            "time": ["交易时间", "时间", "交易时刻"],
            "balance": ["余额", "账户余额"],
            "amount": ["交易金额", "发生额", "发生额（分）", "金额"],
        },
    }

    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'column_mapping.json')
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            # 合并：以 JSON 为准，缺失的键用默认值补充
            for key in defaults:
                if key not in loaded:
                    loaded[key] = defaults[key]
            _CONFIG_CACHE = loaded
            return _CONFIG_CACHE
        except Exception:
            print(f"  [提示] 读取列名映射文件失败，使用默认配置: {config_path}")
    _CONFIG_CACHE = defaults
    return _CONFIG_CACHE


def _get_csv_aliases():
    """获取 CSV 列名映射字典（原始列名 → 标准列名），过滤掉以 _ 开头的注释键"""
    raw = _load_column_config()["csv_column_aliases"]
    return {k: v for k, v in raw.items() if not k.startswith('_')}


def _get_fen_columns():
    """获取金额单位为"分"的列名集合。

    支持两种 JSON 格式：
      - 扁平列表：["发生额（分）"]
      - 带注释的对象：{"__comment__": "...", "columns": ["发生额（分）"]}
    """
    raw = _load_column_config()["amount_in_fen_columns"]
    if isinstance(raw, dict) and "columns" in raw:
        return set(raw["columns"])
    if isinstance(raw, list):
        return set(raw)
    return set()


def _get_preprocess_groups():
    """获取预处理模块的列名分组配置（过滤掉以 _ 开头的注释键）"""
    raw = _load_column_config()["preprocess_column_groups"]
    return {k: v for k, v in raw.items() if not k.startswith('_')}


# =========================
# 日期工具
# =========================

def parse_date(d):
    """将各种格式的日期转换为datetime对象"""
    if pd.isna(d):
        return None
    if isinstance(d, datetime):
        return d
    if isinstance(d, (int, float)):
        s = str(int(d))
        return datetime.strptime(s, '%Y%m%d')
    if isinstance(d, str):
        s_raw = d.strip()
        # 标准 8 位数字日期 (如 20250123, 2025-01-23, 2025/01/23)
        s = s_raw.replace('-', '').replace('/', '')
        if s.isdigit() and len(s) == 8:
            return datetime.strptime(s, '%Y%m%d')
        # 处理单月/单日格式 (如 2025/7/11 → 2025-07-11, 2025/7/1 → 2025-07-01)
        for sep in ('/', '-'):
            if sep in s_raw:
                parts = s_raw.split(sep)
                if len(parts) == 3:
                    try:
                        return datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                    except (ValueError, TypeError):
                        pass
    return None



# =========================
# CSV 文件支持
# =========================

# 列名映射：从 column_mapping.json 加载，可通过修改 JSON 文件自定义
# CSV_COLUMN_ALIASES — 原始列名 → 标准列名（调用 _get_csv_aliases() 获取）
# AMOUNT_IN_FEN_COLUMNS — 金额单位为"分"的列名（调用 _get_fen_columns() 获取）

# 标准列名列表（所有检测模块应统一使用的列名）
STANDARD_COLUMNS = {
    '交易日期', '入帐日期', '交易金额', '发生额',
    '借贷标志', '对方户名', '账号', '对方账号',
    '余额', '账户余额', '摘要',
}


def _normalize_csv_columns(df, convert_fen=False):
    """
    统一CSV文件的列名，处理各种银行流水格式的差异。

    1. 将别名列名映射为标准列名
    2. 将"发生额（分）" ÷100 转换为元
    3. 自动检测"非税专户"格式（含地区号+账号列），金额÷100
    4. 确保必要列存在（对方户名、账号等）

    参数
    ----------
    df : pd.DataFrame
        从CSV读取的原始DataFrame
    convert_fen : bool
        是否将"分"为单位的金额自动转为"元"（默认 True）。
        设为 False 时，源数据中的金额原样保留（1000 即为一千元）。

    返回
    -------
    pd.DataFrame
        列名统一后的DataFrame
    """
    df = df.copy()
    col_map = {}
    amount_is_fen = False

    csv_aliases = _get_csv_aliases()
    fen_columns = _get_fen_columns()
    for actual_name in df.columns:
        name_clean = str(actual_name).strip()
        if name_clean in csv_aliases:
            std_name = csv_aliases[name_clean]
            col_map[actual_name] = std_name
            if name_clean in fen_columns:
                amount_is_fen = True
        else:
            col_map[actual_name] = actual_name  # keep as-is

    df = df.rename(columns=col_map)

    # 处理金额为"分"的情况：÷100 转换为元（仅当列名明确标注"分"时）
    if amount_is_fen and '发生额' in df.columns:
        df['发生额'] = pd.to_numeric(df['发生额'], errors='coerce') / 100.0

    # 自动检测"非税专户"银行格式：含有"地区号"和"帐号"/"账号"列，金额以"分"为单位
    # 注意：此时列名已经过 rename（"帐号"→"账号"），需检查两种形式
    if convert_fen:
        _cols_str = list(map(str, df.columns))
        _is_feishui_format = (
            '地区号' in _cols_str and
            ('帐号' in _cols_str or '账号' in _cols_str)
        )
        if _is_feishui_format and '发生额' in df.columns and not amount_is_fen:
            df['发生额'] = pd.to_numeric(df['发生额'], errors='coerce') / 100.0

    # 对方户名回退：如果没有对方户名但有对方行名，用对方行名替代
    if '对方户名' not in df.columns and '对方行名' in df.columns:
        df['对方户名'] = df['对方行名']
    elif '对方户名' not in df.columns:
        df['对方户名'] = ''

    # 账号列统一
    if '账号' not in df.columns and '帐号' in df.columns:
        df['账号'] = df['帐号']

    # 确保账号列不出现科学计数法（如 1.82E+16）
    if '账号' in df.columns:
        # 先尝试保留原始字符串形式
        df['账号'] = df['账号'].apply(
            lambda x: str(int(float(x))) if pd.notna(x) and isinstance(x, (int, float)) and not isinstance(x, str)
            else str(x).strip() if pd.notna(x)
            else ''
        )

    # 确保"发生额"可作为"交易金额"使用（在统一处理时做别名）
    if '交易金额' not in df.columns and '发生额' in df.columns:
        df['交易金额'] = df['发生额']

    # 清理所有字符串值的前后空白（如制表符前缀）
    for col in df.columns:
        if df[col].dtype == object:  # 只处理字符串列
            df[col] = df[col].apply(
                lambda x: str(x).strip() if pd.notna(x) and isinstance(x, str) else x
            )

    return df


def _is_csv_file(filepath):
    """检查文件是否为CSV/TXT格式"""
    ext = os.path.splitext(filepath)[1].lower()
    return ext in ('.csv', '.txt')


def _read_csv_raw(filepath):
    """
    读取CSV文件，自动检测编码，处理"所有数据在第一列"的情况。

    银行导出的 CSV 有时会被 pandas 误读为单列，
    此处检测到单列后尝试用逗号重新分割。
    """
    # 尝试多种编码
    df = None
    for encoding in ['utf-8', 'gbk', 'gb2312', 'utf-16']:
        try:
            df = pd.read_csv(filepath, header=None, encoding=encoding)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if df is None:
        # fallback：用二进制模式读取并猜测编码
        import codecs
        with open(filepath, 'rb') as f:
            raw = f.read(8192)
        try:
            enc = codecs.lookup('utf-8').name
            df = pd.read_csv(filepath, header=None, encoding=enc)
        except Exception:
            try:
                enc = 'gbk'
                df = pd.read_csv(filepath, header=None, encoding=enc)
            except Exception:
                raise ValueError(f"无法识别文件编码: {filepath}")

    # ---- 处理「所有数据挤在第一列」的情况 ----
    if df.shape[1] <= 1:
        # 依次尝试多种分隔符：逗号、Tab、分号
        best_df = None
        best_cols = 0
        for sep in [',', '\t', ';', '|']:
            split_df = df.iloc[:, 0].str.split(sep, expand=True)
            if split_df.shape[1] > 1:
                col_counts = split_df.apply(lambda r: r.notna().sum(), axis=1)
                median_cols = col_counts.median()
                if median_cols >= 4 and median_cols > best_cols:  # 至少4列、列数最多者获胜
                    best_df = split_df
                    best_cols = median_cols
        if best_df is not None:
            df = best_df

    return df


def _detect_header_row_from_df(df_raw, max_rows=15):
    """
    从原始 DataFrame（无表头）中检测表头行。
    与 _detect_header_row 逻辑相同，但直接操作 DataFrame。
    """
    best_row = 0
    best_score = -1.0

    for h in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[h]
        cols = [str(c).strip() for c in row if pd.notna(c)]
        if not cols:
            continue

        unique_ratio = len(set(cols)) / len(cols)

        score = 0
        for c in cols:
            if '入帐' in c or ('日期' in c and '交易' in c):
                score += 3
            elif '日期' in c or '借贷标志' in c:
                score += 3
            elif '发生额' in c or ('金额' in c and '交易' in c):
                score += 3
            elif '金额' in c:
                score += 2
            elif '户名' in c or '行名' in c:
                score += 2
            elif '余额' in c:
                score += 2
            elif '摘要' in c or '描述' in c:
                score += 1
            elif '账号' in c or '帐号' in c:
                score += 1
            elif '对方' in c:
                score += 1
            elif '记录号' in c:
                score += 1

        score = score * unique_ratio

        if score > best_score:
            best_score = score
            best_row = h

    return best_row


def _load_dataframe(filepath, header_row=None):
    """
    通用加载函数：支持 Excel 和 CSV，返回 (检测到的表头行号, DataFrame)。

    参数
    ----------
    filepath : str
        文件路径
    header_row : int or None
        指定表头行；None 表示自动检测
    """
    if _is_csv_file(filepath):
        df_raw = _read_csv_raw(filepath)
        if header_row is None:
            header_row = _detect_header_row_from_df(df_raw)
        # 以 header_row 为表头构建 DataFrame
        df = df_raw.iloc[header_row:].reset_index(drop=True)
        df.columns = [
            str(c).strip() if pd.notna(c) else f'col_{i}'
            for i, c in enumerate(df.iloc[0])
        ]
        df = df.iloc[1:].reset_index(drop=True)
        # 数值列转换（逐列尝试，不抛出异常）
        # 注意：跳过疑似账号/卡号的列，避免科学计数法导致精度丢失
        ACCOUNT_LIKE_COLS = {'账号', '帐号', '卡号', '对方账号', '对方帐号', '对方行号'}
        for col in df.columns:
            col_clean = str(col).strip()
            if col_clean in ACCOUNT_LIKE_COLS:
                continue
            # 也跳过列名中包含"账号"的变体
            if any(kw in col_clean for kw in ('账号', '帐号', '卡号', '行号')):
                continue
            try:
                converted = pd.to_numeric(df[col], errors='coerce')
                if converted.notna().sum() > len(df) * 0.5:  # 超过50%可转为数值才保留
                    df[col] = converted
            except Exception:
                pass
        return header_row, df
    else:
        # Excel 文件：用原有的 _detect_header_row
        xls = pd.ExcelFile(filepath)
        sheet_name = xls.sheet_names[0]
        if header_row is None:
            header_row = _detect_header_row(xls, sheet_name)
        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
        return header_row, df


def _detect_header_row(xls, sheet_name, max_rows=15):
    """自动检测表头所在行。扫描前 max_rows 行，按列名匹配度评分。"""
    best_row = 0
    best_score = -1.0

    for h in range(max_rows):
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=h, nrows=0)
            cols = [str(c).strip() for c in df.columns]
            if not cols:
                continue

            # 去重比率：重复列名说明是元数据行
            unique_ratio = len(set(cols)) / len(cols)

            # 根据关键词评分
            score = 0
            for c in cols:
                if '入帐' in c or ('日期' in c and '交易' in c):
                    score += 3
                elif '日期' in c or '借贷标志' in c:
                    score += 3
                elif '发生额' in c or ('金额' in c and '交易' in c):
                    score += 3
                elif '金额' in c:
                    score += 2
                elif '户名' in c or '行名' in c:
                    score += 2
                elif '余额' in c:
                    score += 2
                elif '摘要' in c or '描述' in c:
                    score += 1
                elif '账号' in c or '帐号' in c:
                    score += 1
                elif '对方' in c:
                    score += 1
                elif '记录号' in c:
                    score += 1

            score = score * unique_ratio

            if score > best_score:
                best_score = score
                best_row = h
        except Exception:
            continue

    return best_row

def is_weekend(dt):
    """判断是否为周末"""
    return dt.weekday() >= 5


def next_workday(dt):
    """获取下一个工作日"""
    nxt = dt + timedelta(days=1)
    while is_weekend(nxt):
        nxt += timedelta(days=1)
    return nxt


def is_within_same_or_next_workday(date1, date2):
    """判断date2是否是date1的当天或第二个工作日"""
    if date1 == date2:
        return True
    if next_workday(date1) == date2:
        return True
    return False


def add_working_days(dt, n):
    """
    往后加 n 个工作日（跳过周末）。

    >>> add_working_days(datetime(2026, 4, 28), 2)  # 周二 → 周四
    datetime(2026, 4, 30)
    >>> add_working_days(datetime(2026, 4, 30), 1)  # 周四 → 周五
    datetime(2026, 5, 1)
    """
    result = dt
    for _ in range(n):
        result += timedelta(days=1)
        while is_weekend(result):
            result += timedelta(days=1)
    return result


def working_days_between(d1, d2):
    """
    计算 d1 到 d2 之间的工作日数（不含 d1，含 d2）。

    >>> working_days_between(datetime(2026, 4, 28), datetime(2026, 4, 29))  # 周二→周三
    1
    >>> working_days_between(datetime(2026, 4, 24), datetime(2026, 4, 27))  # 周五→周一
    1
    """
    if d1 is None or d2 is None or d1 >= d2:
        return 0
    count = 0
    current = d1
    while current < d2:
        current += timedelta(days=1)
        if not is_weekend(current):
            count += 1
    return count


def is_within_working_days(date_from, date_to, max_days):
    """
    判断 date_to 是否在 date_from 之后的 max_days 个工作日内。
    """
    if date_from is None or date_to is None:
        return False
    if date_to < date_from:
        return False
    return working_days_between(date_from, date_to) <= max_days


# =========================
# 数据加载
# =========================


def _normalize_amount(row):
    """
    根据借贷标志或正负值确定交易金额。
    
    银行流水有两种金额表示方式：
    - 正负数表示：收入为正，支出为负（标准零余额格式）
    - 借贷标志 + 正数：借贷标志=2(贷/收入)为正，借贷标志=1(借/支出)为负（省垫款户格式）
    """
    raw = row.get('交易金额') or row.get('发生额')
    amount = pd.to_numeric(raw, errors='coerce')
    if pd.isna(amount):
        return 0.0
    
    debit_flag = row.get('借贷标志')
    if pd.notna(debit_flag):
        try:
            flag = int(debit_flag)
            if flag == 2:   # 贷 / 收入
                return abs(amount)
            elif flag == 1: # 借 / 支出
                return -abs(amount)
        except (ValueError, TypeError):
            # 数值转换失败，尝试字符串模式匹配（银行可能有 "借"/"贷" 等文字值）
            flag_str = str(debit_flag).strip()
            if '借' in flag_str:
                return -abs(amount)
            elif '贷' in flag_str:
                return abs(amount)
    
    # 没有借贷标志，直接用原值
    return amount


def _normalize_amount_vectorized(df):
    """
    向量化版本：一次性对整个 DataFrame 的金额进行借贷标志规范化。

    替代逐行 .apply(_normalize_amount, axis=1)，性能提升 50-100 倍。
    对 10 万行数据：原版 ~2 秒 → 向量化 ~0.02 秒。
    """
    if '交易金额' in df.columns:
        amounts = pd.to_numeric(df['交易金额'], errors='coerce').fillna(0)
    elif '发生额' in df.columns:
        amounts = pd.to_numeric(df['发生额'], errors='coerce').fillna(0)
    else:
        df['交易金额'] = 0.0
        return df

    if '借贷标志' in df.columns:
        # 先尝试数值转换（1=借/支出, 2=贷/收入）
        flags_num = pd.to_numeric(df['借贷标志'], errors='coerce')
        # 对无法转数值的，尝试字符串模式匹配（银行可能有 "借"/"贷" 等文字值）
        flags_str = df['借贷标志'].astype(str).str.strip()
        # 最终标志：数字优先，字符串回退
        debit_mask = (flags_num == 1) | (
            flags_num.isna() & flags_str.str.contains('借', na=False)
        )
        credit_mask = (flags_num == 2) | (
            flags_num.isna() & flags_str.str.contains('贷', na=False) & ~debit_mask
        )
        result = amounts.copy()
        result[credit_mask] = result[credit_mask].abs()
        result[debit_mask] = -result[debit_mask].abs()
        df['交易金额'] = result
    else:
        df['交易金额'] = amounts

    return df


# ============================================================
# 修改 _load_zero_balance 以支持多种银行流水格式
# ============================================================
def _load_zero_balance(xls, sheet_name):
    """
    读取零余额账户sheet——支持多种银行流水格式。

    自动检测格式：
    - 标准零余额格式：header=9，列含 交易日期
    - 省垫款户格式：header=0，列含 入帐日期/入账日期、借贷标志、发生额

    日期列名兼容（按优先级）：
      交易日期 → 入帐日期 → 入账日期
    """
    # 自动检测表头行
    header_row = _detect_header_row(xls, sheet_name)
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

    # 标准格式：含 交易日期 列
    if '交易日期' in df.columns:
        df = df.dropna(subset=['交易日期'])
        if len(df) > 0:
            df['日期对象'] = df['交易日期'].apply(parse_date)
            df = df.dropna(subset=['日期对象'])
            df['交易金额'] = pd.to_numeric(df['交易金额'], errors='coerce')
            # 如有借贷标志，根据借贷方向归一化金额正负号
            # 借=支出(负), 贷=收入(正)，避免"借"记录被误判为来账
            if '借贷标志' in df.columns:
                df = _normalize_amount_vectorized(df)
            if len(df) > 0:
                return df

    # 省垫款户格式：含 入帐日期/入账日期、借贷标志 列
    # 优先检查"入帐日期"，再回退检查"入账日期"
    date_col = None
    for candidate in ['入帐日期', '入账日期']:
        if candidate in df.columns:
            date_col = candidate
            break
    if date_col is not None:
        df = df.dropna(subset=[date_col])
        if len(df) > 0:
            df['日期对象'] = df[date_col].apply(parse_date)
            df = df.dropna(subset=['日期对象'])
            df = _normalize_amount_vectorized(df)
            if '对方户名' not in df.columns:
                df['对方户名'] = df.get('对方行名', '')
            return df

    return pd.DataFrame()



def _load_qing_suan(filepath):
    """
    读取集中支付零余额清算待转（支持 Excel 和 CSV）。

    自动检测表头行，不再硬编码 header=5。
    """
    if _is_csv_file(filepath):
        df_raw = _read_csv_raw(filepath)
        header_row = _detect_header_row_from_df(df_raw)
        df = df_raw.iloc[header_row:].reset_index(drop=True)
        df.columns = [
            str(c).strip() if pd.notna(c) else f'col_{i}'
            for i, c in enumerate(df.iloc[0])
        ]
        df = df.iloc[1:].reset_index(drop=True)
        # 归一化列名
        df = _normalize_csv_columns(df)
    else:
        xls = pd.ExcelFile(filepath)
        sheet_name = xls.sheet_names[0]
        header_row = _detect_header_row(xls, sheet_name)
        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

    # 检查是否存在交易日期列
    date_col = None
    for candidate in ['交易日期', '入帐日期', '入账日期']:
        if candidate in df.columns:
            date_col = candidate
            break

    if date_col is None:
        raise KeyError(f"文件 {filepath} 中未找到交易日期列（'交易日期'/'入帐日期'），当前列名: {list(df.columns)}")

    df = df.dropna(subset=[date_col])
    df['日期对象'] = df[date_col].apply(parse_date)
    df = df.dropna(subset=['日期对象'])

    # 金额列检测
    amt_col = None
    for candidate in ['交易金额', '发生额']:
        if candidate in df.columns:
            amt_col = candidate
            break
    if amt_col is None:
        raise KeyError(f"文件 {filepath} 中未找到金额列（'交易金额'/'发生额'）")

    df[amt_col] = pd.to_numeric(df[amt_col], errors='coerce')
    df['交易金额'] = df[amt_col]  # 统一列名
    # 根据借贷标志规范化金额符号（如有借贷标志列）
    if '借贷标志' in df.columns:
        df['交易金额'] = df.apply(_normalize_amount, axis=1)

    # 余额列（兼容"账户余额"和"余额"两种列名）
    balance_col = None
    for c in ['账户余额', '余额']:
        if c in df.columns:
            balance_col = c
            break
    if balance_col:
        df[balance_col] = pd.to_numeric(df[balance_col], errors='coerce')
        # 统一为"账户余额"
        if balance_col != '账户余额':
            df['账户余额'] = df[balance_col]

    return df


# =========================
# 资金匹配检测（原main.py逻辑）
# =========================

def _calc_tui_guo_ku(qing_df, account, tui_zhuan_date):
    """
    计算退至国库日期

    校验逻辑：在收到零余额账户转至待转户的资金后（即退至待转户日期当天），
    检查当天待转户（集中支付零余额清算待转）的余额是否清零，
    或当天是否有一笔转至"往账待转"的资金且金额 >= 收到的资金总和。
    若满足任一条件，则退至国库日期为该日期；否则为空。
    """
    if pd.isna(tui_zhuan_date):
        return None
    day_df = qing_df[qing_df['日期对象'] == tui_zhuan_date]
    if len(day_df) == 0:
        return None

    # 当天来自该零余额账号的正数交易（收入）总和
    income_mask = (day_df['对方账号'].astype(str).str.strip() == account) & (day_df['交易金额'] > 0)
    income_sum = day_df[income_mask]['交易金额'].sum()

    # 当天转至往账待转的支出总和
    wangzhang_mask = day_df['对方户名'] == '往账待转'
    wangzhang_sum = abs(day_df[wangzhang_mask]['交易金额'].sum())

    # 当天最后一笔交易的余额（按 Excel 行序，越靠后越晚发生）
    last_balance = day_df.iloc[-1]['账户余额']

    if (wangzhang_sum >= income_sum * 0.99 and income_sum > 0) or last_balance == 0:
        return tui_zhuan_date.strftime('%Y-%m-%d')
    return None


def _calc_tui_guo_ku_indexed(qing_by_date, account, tui_zhuan_date):
    """
    计算退至国库日期（O(1) 预索引版）。

    与 _calc_tui_guo_ku 逻辑相同，但接收预分组的 qing_by_date dict
    而非完整 qing_df，将日期查找从 O(M) 降为 O(1)。
    """
    if pd.isna(tui_zhuan_date) or tui_zhuan_date is None:
        return None

    day_df = qing_by_date.get(tui_zhuan_date)
    if day_df is None or len(day_df) == 0:
        return None

    # 当天来自该零余额账号的正数交易（收入）总和
    income_mask = (day_df['对方账号'].astype(str).str.strip() == account) & (day_df['交易金额'] > 0)
    income_sum = day_df[income_mask]['交易金额'].sum()

    # 当天转至往账待转的支出总和
    wangzhang_mask = day_df['对方户名'] == '往账待转'
    wangzhang_sum = abs(day_df[wangzhang_mask]['交易金额'].sum())

    # 当天最后一笔交易的余额
    last_balance = day_df.iloc[-1]['账户余额']

    if (wangzhang_sum >= income_sum * 0.99 and income_sum > 0) or last_balance == 0:
        return tui_zhuan_date.strftime('%Y-%m-%d')
    return None


def _detect_fund_matching_sheet(account, zero_df, sheet_name, advance_acct_name, qing_by_date):
    """
    处理单个 sheet 的资金匹配检测（模块级函数，支持 pickle 序列化供 ProcessPoolExecutor 使用）。

    使用预索引替代全表扫描：O(N log N + P log N)，原版为 O(P×N)。
    """
    results = []

    if len(zero_df) == 0:
        return results

    # 确保对方户名列存在
    if '对方户名' in zero_df.columns:
        zero_df['对方户名'] = zero_df['对方户名'].fillna(zero_df.get('对方行名', ''))
    elif '对方行名' in zero_df.columns:
        zero_df['对方户名'] = zero_df['对方行名']
    else:
        zero_df['对方户名'] = ''

    # ================================================================
    # 阶段1：预建索引 (O(N log N)) —— 只做一次，所有待查记录共享
    # ================================================================

    # 1a. 清算日期候选：对方户名==垫款户 AND 金额>0 的日期（排序去重）
    qs_mask = (
        (zero_df['交易金额'] > 0) &
        (zero_df['对方户名'] == advance_acct_name)
    )
    qs_dates = sorted(zero_df.loc[qs_mask, '日期对象'].unique())

    # 1b. 负向-垫款户交易：按金额分组，每组按日期排序
    neg_mask = (
        (zero_df['交易金额'] < 0) &
        (zero_df['对方户名'] == advance_acct_name)
    )
    neg_subset = zero_df.loc[neg_mask, ['日期对象', '交易金额']]
    neg_by_amount = {}
    for _, neg_row in neg_subset.iterrows():
        amt_key = round(abs(neg_row['交易金额']), 2)
        if amt_key not in neg_by_amount:
            neg_by_amount[amt_key] = []
        neg_by_amount[amt_key].append(neg_row['日期对象'])

    for amt_key in neg_by_amount:
        neg_by_amount[amt_key].sort()

    # ================================================================
    # 阶段2：提取待查记录 (O(N))
    # ================================================================
    dai_cha_df = zero_df[
        (zero_df['交易金额'] > 0) &
        (zero_df['对方户名'] != advance_acct_name)
    ]

    # ================================================================
    # 阶段3：逐条匹配 (O(P log N)) — 每条用二分查找/字典查找
    # ================================================================
    for _, row in dai_cha_df.iterrows():
        refund_date = row['日期对象']
        target_amt = round(abs(row['交易金额']), 2)

        if pd.isna(target_amt) or target_amt == 0:
            continue

        # 3a. 清算日期：二分查找 qs_dates 中 < refund_date 的最大日期
        qing_suan_date = None
        idx = bisect.bisect_left(qs_dates, refund_date) - 1
        if idx >= 0:
            qing_suan_date = qs_dates[idx]

        # 3b. 退至待转户日期：金额索引 → 找 > refund_date 的最早日期
        tui_zhuan_date = None
        amt_dates = neg_by_amount.get(target_amt, [])
        for dt in amt_dates:
            if dt > refund_date:
                tui_zhuan_date = dt
                break

        # 3c. 退至国库日期：预索引 qing_by_date → O(1)
        tui_guo_ku_date = _calc_tui_guo_ku_indexed(
            qing_by_date, account, tui_zhuan_date
        )

        # 提取摘要和对方账号
        summary = ''
        for c in ('摘要', '交易描述', '附言', '用途'):
            v = row.get(c)
            if pd.notna(v) and str(v).strip():
                summary = str(v).strip()
                break
        counterparty_acct = ''
        for c in ('对方账号', '对方帐号'):
            v = row.get(c)
            if pd.notna(v):
                counterparty_acct = f'{int(float(v))}' if isinstance(v, (int, float)) else str(v).strip()
                break

        results.append({
            'sheet名称': sheet_name,
            '账号': account,
            '对方账号': counterparty_acct,
            '对方户名': row.get('对方户名', ''),
            '摘要': summary,
            '清算日期': (
                qing_suan_date.strftime('%Y-%m-%d')
                if qing_suan_date is not None else None
            ),
            '退款日期': refund_date.strftime('%Y-%m-%d'),
            '退至垫款户日期': (
                tui_zhuan_date.strftime('%Y-%m-%d')
                if tui_zhuan_date is not None else None
            ),
            '退至金库日期': tui_guo_ku_date,
            '交易金额': target_amt,
        })

    return results


def detect_fund_matching(
    zero_balance_path='零余额账户.xlsx',
    qing_suan_path: str = '集中支付零余额清算待转.xlsx',
    advance_acct_name: str = '集中支付零余额清算待转',
    max_workers: int = 1,
) -> pd.DataFrame:
    """
    资金匹配检测

    从零余额账户中提取正向非清算交易作为待查记录，自动匹配各环节日期：
      - 退款日期
      - 清算日期
      - 退至待转户日期
      - 退至国库日期

    参数
    ----------
    zero_balance_path : str
        零余额账户 Excel 文件路径
    qing_suan_path : str
        集中支付零余额清算待转 Excel 文件路径
    advance_acct_name : str
        垫款户名称（对方户名），用于识别清算交易；默认 "集中支付零余额清算待转"

    返回
    -------
    pd.DataFrame
        包含资金匹配检测结果的DataFrame
    """
    # 如果垫款户名称为空，使用默认值
    if not advance_acct_name:
        advance_acct_name = '集中支付零余额清算待转'
    # 支持单个文件路径或文件路径列表
    if isinstance(zero_balance_path, str):
        zero_balance_paths = [zero_balance_path]
    else:
        zero_balance_paths = zero_balance_path

    qing_df = _load_qing_suan(qing_suan_path)
    qing_by_date = {}
    for dt, grp in qing_df.groupby('日期对象'):
        qing_by_date[dt] = grp

    # ---- 构建处理任务 (account, zero_df, sheet_name) ----
    tasks = []
    for zb_path in zero_balance_paths:
        if _is_csv_file(zb_path):
            _, zero_df = _load_dataframe(zb_path)
            if len(zero_df) == 0:
                continue
            zero_df = _normalize_csv_columns(zero_df)
            date_col = None
            for c in ['交易日期', '入帐日期', '入账日期']:
                if c in zero_df.columns:
                    date_col = c
                    break
            if date_col is None:
                continue
            zero_df = zero_df.dropna(subset=[date_col])
            if len(zero_df) == 0:
                continue
            zero_df['日期对象'] = zero_df[date_col].apply(parse_date)
            zero_df = zero_df.dropna(subset=['日期对象'])
            if '交易金额' in zero_df.columns:
                zero_df['交易金额'] = pd.to_numeric(zero_df['交易金额'], errors='coerce')
            elif '发生额' in zero_df.columns:
                zero_df['交易金额'] = pd.to_numeric(zero_df['发生额'], errors='coerce')
            zero_df = _normalize_amount_vectorized(zero_df)
            if '账号' in zero_df.columns and len(zero_df) > 0:
                raw_acct = zero_df['账号'].iloc[0]
                account = str(raw_acct).strip() if pd.notna(raw_acct) else ''
            else:
                account = os.path.splitext(os.path.basename(zb_path))[0]
            tasks.append((account, zero_df, os.path.basename(zb_path)))
        else:
            zero_xls = pd.ExcelFile(zb_path)
            for s in zero_xls.sheet_names:
                zero_df = _load_zero_balance(zero_xls, s)
                if len(zero_df) == 0:
                    continue
                account = s.split('(')[0] if '(' in s else s
                if account in ('Sheet1', 'Sheet', 'sheet') and '账号' in zero_df.columns:
                    account = str(zero_df['账号'].iloc[0])
                tasks.append((account, zero_df, s))

    # ---- 处理 (多进程并行处理) ----
    if max_workers > 1 and len(tasks) > 1:
        all_results = []
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for account, zero_df, sheet_name in tasks:
                f = executor.submit(
                    _detect_fund_matching_sheet,
                    account, zero_df, sheet_name,
                    advance_acct_name, qing_by_date,
                )
                futures[f] = sheet_name

            for f in as_completed(futures):
                sheet_name = futures[f]
                try:
                    sheet_results = f.result()
                    all_results.extend(sheet_results)
                except Exception as e:
                    print(f"  [错误] 处理 '{sheet_name}' 失败: {e}")
    else:
        all_results = []
        for account, zero_df, sheet_name in tasks:
            sheet_results = _detect_fund_matching_sheet(
                account, zero_df, sheet_name,
                advance_acct_name, qing_by_date,
            )
            all_results.extend(sheet_results)

    return pd.DataFrame(all_results)


# =========================
# 延迟清算积数检测（原 process_zero_balance_v2.py 逻辑）
# =========================

def detect_delayed_settlement(
    zero_balance_path='零余额账户.xlsx',
    advance_acct_name: str = '集中支付零余额清算待转',
) -> pd.DataFrame:
    """
    延迟清算积数检测

    分析零余额账户中正向非清算交易与负向清算交易的匹配情况，
    找出未在当天或第二个工作日完成清算的交易，计算积数（金额 × 延迟天数）。

    参数
    ----------
    zero_balance_path : str
        零余额账户 Excel 文件路径
    advance_acct_name : str
        垫款户名称（对方户名），用于识别清算交易；默认 "集中支付零余额清算待转"

    返回
    -------
    pd.DataFrame
        包含延迟清算记录的DataFrame（按积数降序排列）
    """
    # 如果垫款户名称为空，使用默认值
    if not advance_acct_name:
        advance_acct_name = '集中支付零余额清算待转'
    if isinstance(zero_balance_path, str):
        zero_balance_paths = [zero_balance_path]
    else:
        zero_balance_paths = zero_balance_path

    results = []

    for zb_path in zero_balance_paths:
        if _is_csv_file(zb_path):
            # CSV 文件：读取为单个 DataFrame，使用列名归一化
            _, df = _load_dataframe(zb_path)
            if len(df) == 0:
                continue
            # 归一化列名（"日期"→"交易日期"、"发生额（分）"→"发生额"/100 等）
            df = _normalize_csv_columns(df)
            # 统一日期列名
            date_col = None
            for c in ['交易日期', '入帐日期', '入账日期']:
                if c in df.columns:
                    date_col = c
                    break
            if date_col is None:
                continue
            df = df.dropna(subset=[date_col])
            if len(df) == 0:
                continue
            df['日期对象'] = df[date_col].apply(parse_date)
            df = df.dropna(subset=['日期对象'])
            # 金额列（_normalize_csv_columns 已处理"发生额（分）→发生额/100"）
            if '交易金额' in df.columns:
                df['交易金额'] = pd.to_numeric(df['交易金额'], errors='coerce')
            elif '发生额' in df.columns:
                df['交易金额'] = pd.to_numeric(df['发生额'], errors='coerce')
            # 应用借贷标志规范化
            df = _normalize_amount_vectorized(df)
            # 对方户名（_normalize_csv_columns 已确保存在）
            # 交易描述
            if '交易描述' not in df.columns:
                df['交易描述'] = df.get('摘要', df.get('摘要信息', ''))
            sheets_iter = [('CSV', df)]
        else:
            xls = pd.ExcelFile(zb_path)
            sheets_iter = [(s, _load_zero_balance(xls, s)) for s in xls.sheet_names]

        for sheet_name, df in sheets_iter:
            if len(df) == 0:
                continue

            # 提取账号
            account = sheet_name
            for c in ('账号', '帐号'):
                if c in df.columns and len(df) > 0:
                    v = df[c].iloc[0]
                    if pd.notna(v):
                        account = f'{int(float(v))}' if isinstance(v, (int, float)) else str(v).strip()
                        break

            # 正向非清算交易：金额>0，对方户名不为"集中支付零余额清算待转"
            pos_mask = (
                (df['交易金额'] > 0) &
                (df['对方户名'] != advance_acct_name)
            )
            pos_df = df[pos_mask].copy()

            # 负向清算交易：金额<0，对方户名为"集中支付零余额清算待转"
            neg_mask = (
                (df['交易金额'] < 0) &
                (df['对方户名'] == advance_acct_name)
            )
            neg_df = df[neg_mask].copy()

            if len(pos_df) == 0:
                continue

            # 将负向交易按金额分组，并标记是否已使用
            neg_by_amount = {}
            for idx, row in neg_df.iterrows():
                amt = abs(row['交易金额'])
                if amt not in neg_by_amount:
                    neg_by_amount[amt] = []
                neg_by_amount[amt].append({
                    'index': idx,
                    'row': row,
                    'used': False,
                })

            # 按日期排序
            for amt in neg_by_amount:
                neg_by_amount[amt].sort(key=lambda r: r['row']['日期对象'])

            # 遍历每条正向非清算交易（按日期排序）
            pos_list = sorted(
                [row for _, row in pos_df.iterrows()],
                key=lambda r: r['日期对象'],
            )

            for pos_row in pos_list:
                pos_date = pos_row['日期对象']
                pos_amt = pos_row['交易金额']

                matched_negs = neg_by_amount.get(pos_amt, [])
                if not matched_negs:
                    continue

                # 检查当天或第二个工作日是否有未使用的匹配
                immediate_match = None
                delayed_matches = []

                for neg_item in matched_negs:
                    if neg_item['used']:
                        continue
                    neg_date = neg_item['row']['日期对象']
                    if is_within_same_or_next_workday(pos_date, neg_date):
                        immediate_match = neg_item
                        break
                    elif neg_date > pos_date:
                        delayed_matches.append(neg_item)

                # 当天或次日有匹配 → 正常，跳过
                if immediate_match is not None:
                    immediate_match['used'] = True
                    continue

                # 无及时匹配，找最终匹配的负向交易（日期最早的未使用延迟匹配）
                final_neg = None
                for neg_item in delayed_matches:
                    if not neg_item['used']:
                        final_neg = neg_item
                        break

                if final_neg:
                    final_neg['used'] = True
                    final_neg_date = final_neg['row']['日期对象']
                    days_diff = (final_neg_date - pos_date).days
                    jishu = pos_amt * days_diff

                    # 提取对方账号
                    counterparty_acct = ''
                    for c in ('对方账号', '对方帐号'):
                        v = pos_row.get(c)
                        if pd.notna(v):
                            counterparty_acct = f'{int(float(v))}' if isinstance(v, (int, float)) else str(v).strip()
                            break

                    results.append({
                        'sheet名称': sheet_name,
                        '账号': account,
                        '对方账号': counterparty_acct,
                        '对方户名': pos_row['对方户名'],
                        '交易描述': pos_row['交易描述'],
                        '退款日期': pos_date.strftime('%Y-%m-%d'),
                        '交易金额': pos_amt,
                        '退回待转户日期': final_neg_date.strftime('%Y-%m-%d'),
                        '日期间隔': days_diff,
                        '积数': jishu,
                    })

    results_df = pd.DataFrame(results)
    if len(results_df) > 0:
        results_df = results_df.sort_values('积数', ascending=False)

    return results_df


# =========================
# 非税核查
# =========================

def detect_non_tax_verification(
    file_path='非税专户.csv',
    designated_account: str = '待报解预算收入',
    days_threshold: int = 2,
) -> pd.DataFrame:
    """
    非税专户核查（v2）：累计滚动匹配 + 余额归零检测。

    核心规则（v2 优化）：
    - FIFO 累计匹配：不再要求 1:1 精确金额，采用先入先出累计匹配
      - 一笔来账可被多笔划转分批转出（拆分划转）
      - 一笔划转可覆盖多笔来账（合并划转）
      - 当天划转一部分，剩余第二天同新来账一起划转
    - 零余额检测：若某笔交易后余额为零，则截至当时的所有待匹配来账均视为已划转
      - 此类记录备注标注"余额归零"
    - 工作日计算：跳过周六周日

    参数
    ----------
    file_path : str
        非税专户流水文件路径（CSV 或 Excel）
    designated_account : str
        指定划转账户的对方户名；默认 "待报解预算收入"（国库经收处）
    days_threshold : int
        工作日阈值（默认 2）。来账超过此工作日未匹配支出 → 标记"可疑"

    返回
    -------
    pd.DataFrame
        每行一笔来账，包含匹配状态：
        - 来源日期 / 来源金额 / 摘要 / 来账对方户名
        - 划转日期 / 划转金额 / 划转对方行名
        - 等待工作日 / 阈值天数 / 备注
        - 状态：已划转 / 延迟划转 / 未划转 / 可疑
        - 末尾附汇总行
    """
    # ================================================================
    # 1. 加载数据
    # ================================================================
    if _is_csv_file(file_path):
        _, df = _load_dataframe(file_path)
        if len(df) == 0:
            return pd.DataFrame()
        df = _normalize_csv_columns(df, convert_fen=False)
        # 日期列
        date_col = None
        for c in ['交易日期', '入帐日期', '入账日期']:
            if c in df.columns:
                date_col = c
                break
        if date_col is None:
            return pd.DataFrame()
        df = df.dropna(subset=[date_col])
        df['日期对象'] = df[date_col].apply(parse_date)
        df = df.dropna(subset=['日期对象'])
        # 金额列（_normalize_csv_columns 已处理 fen→元）
        if '交易金额' in df.columns:
            df['交易金额'] = pd.to_numeric(df['交易金额'], errors='coerce')
        elif '发生额' in df.columns:
            df['交易金额'] = pd.to_numeric(df['发生额'], errors='coerce')
        df = _normalize_amount_vectorized(df)
        # 对方户名
        if '对方户名' not in df.columns or df['对方户名'].isna().all():
            df['对方户名'] = df.get('对方行名', '')
        # 摘要
        if '摘要' not in df.columns:
            df['摘要'] = df.get('附言', '')
        # 对方行名
        if '对方行名' not in df.columns:
            df['对方行名'] = ''
    else:
        xls = pd.ExcelFile(file_path)
        sheet_name = xls.sheet_names[0]
        df = _load_zero_balance(xls, sheet_name)
        if '摘要' not in df.columns:
            df['摘要'] = ''
        if '对方行名' not in df.columns:
            df['对方行名'] = ''

    if len(df) == 0:
        return pd.DataFrame()

    # 提取数据来源标识
    file_label = os.path.splitext(os.path.basename(file_path))[0]
    account_num = ''
    for c in ('账号', '帐号'):
        if c in df.columns and len(df) > 0:
            v = df[c].iloc[0]
            if pd.notna(v):
                # 确保账号不以科学计数法显示
                if isinstance(v, (int, float)):
                    account_num = f'{int(float(v))}'
                else:
                    account_num = str(v).strip()
                break

    # 检测余额列（用于余额归零判断）
    balance_col = None
    for c in ['账户余额', '余额']:
        if c in df.columns:
            balance_col = c
            df[balance_col] = pd.to_numeric(df[balance_col], errors='coerce')
            break

    # ================================================================
    # 2. 数据准备：按时间顺序遍历，构建来账列表与 FIFO 匹配
    # ================================================================
    df = df.sort_values('日期对象').reset_index(drop=True)
    latest_date = df['日期对象'].max()

    # 来账列表（保存元数据，按时间顺序）
    credits = []
    # FIFO 待匹配队列：[{credit_idx, remaining}]
    pending = []
    # 匹配记录：credit_idx -> [{date, amount, bank, match_type}]
    match_records = {}

    for idx, row in df.iterrows():
        amount = row['交易金额']
        if pd.isna(amount) or amount == 0:
            continue

        date = row['日期对象']

        if amount > 0:
            # === 来账（收入） → 加入待匹配池 ===
            cpty_acct = ''
            for c in ('对方账号', '对方帐号'):
                v = row.get(c)
                if pd.notna(v):
                    cpty_acct = f'{int(float(v))}' if isinstance(v, (int, float)) else str(v).strip()
                    break

            ci = len(credits)
            credits.append({
                'date': date,
                'amount': amount,
                'summary': str(row.get('摘要', '') or ''),
                'counterparty': str(row.get('对方户名', '') or ''),
                'counterparty_acct': cpty_acct,
            })
            pending.append({'credit_idx': ci, 'remaining': amount})
            match_records[ci] = []

        elif amount < 0:
            # === 支出 → 若对方为指定账户，则进行 FIFO 累计匹配 ===
            counterparty = str(row.get('对方户名', '')).strip()
            if counterparty == designated_account:
                debit_amt = abs(amount)
                debit_bank = str(row.get('对方行名', '') or '')

                # FIFO：从最早待匹配来账开始消耗
                while debit_amt > 0.005 and pending:
                    oldest = pending[0]
                    match_amt = min(debit_amt, oldest['remaining'])
                    match_records[oldest['credit_idx']].append({
                        'date': date,
                        'amount': match_amt,
                        'bank': debit_bank,
                        'type': 'direct',
                    })
                    oldest['remaining'] = oldest['remaining'] - match_amt
                    debit_amt = debit_amt - match_amt
                    if oldest['remaining'] < 0.005:
                        pending.pop(0)

        # === 零余额检测：当前交易后余额为零 → 待匹配来账全部视为已划转 ===
        if balance_col and pending:
            bal = row.get(balance_col)
            if pd.notna(bal) and float(bal) == 0.0:
                for pc in list(pending):
                    if pc['remaining'] > 0.005:
                        match_records[pc['credit_idx']].append({
                            'date': date,
                            'amount': pc['remaining'],
                            'bank': '',
                            'type': 'zero_balance',
                        })
                        pc['remaining'] = 0.0
                pending.clear()

    # ================================================================
    # 3. 生成结果：根据匹配记录判定每条来账的状态
    # ================================================================
    results = []
    total_in = 0.0
    total_matched = 0.0

    for ci, credit in enumerate(credits):
        credit_date = credit['date']
        credit_amt = credit['amount']
        total_in += credit_amt

        events = match_records.get(ci, [])
        matched_total = sum(e['amount'] for e in events)

        if matched_total > 0.005:
            # 有匹配记录
            total_matched += matched_total
            last_event = events[-1]
            last_match_date = last_event['date']

            # 汇总划转对方行名（去重拼接）
            banks = list(dict.fromkeys(
                e['bank'] for e in events if e['bank']
            ))
            match_bank = ' / '.join(banks) if banks else ''

            # 划转日期取最后一笔匹配的日期
            transfer_date = last_match_date.strftime('%Y-%m-%d')

            wdays = working_days_between(credit_date, last_match_date)

            # 判断匹配类型 → 备注
            direct_count = sum(1 for e in events if e['type'] == 'direct')
            zb_count = sum(1 for e in events if e['type'] == 'zero_balance')
            if zb_count > 0 and direct_count > 0:
                remark = f'分{direct_count}批划转 + 余额归零'
            elif zb_count > 0:
                remark = '余额归零'
            elif direct_count > 1:
                remark = f'分{direct_count}批划转'
            else:
                remark = ''

            # 判断状态
            if matched_total >= credit_amt - 0.005:
                # 完全匹配
                within_threshold = wdays <= days_threshold
                status = '已划转' if within_threshold else '延迟划转'
            else:
                # 部分匹配（仍有剩余未匹配）
                wdays_since = working_days_between(credit_date, latest_date)
                if wdays_since > days_threshold:
                    status = '可疑'
                else:
                    status = '未划转'
                if remark:
                    remark += f'；部分划转(已匹配{matched_total}/总额{credit_amt})'
                else:
                    remark = f'部分划转(已匹配{matched_total}/总额{credit_amt})'

            results.append({
                '文件来源': file_label,
                '账号': account_num,
                '来源日期': credit_date.strftime('%Y-%m-%d'),
                '来源金额': credit_amt,
                '摘要': credit['summary'],
                '对方账号': credit['counterparty_acct'],
                '来账对方户名': credit['counterparty'],
                '划转日期': transfer_date,
                '划转金额': matched_total,
                '划转对方行名': match_bank,
                '等待工作日': wdays,
                '阈值天数': days_threshold,
                '备注': remark,
                '状态': status,
            })
        else:
            # 完全无匹配
            wdays_since = working_days_between(credit_date, latest_date)
            if wdays_since > days_threshold:
                status = '可疑'
            else:
                status = '未划转'

            results.append({
                '文件来源': file_label,
                '账号': account_num,
                '来源日期': credit_date.strftime('%Y-%m-%d'),
                '来源金额': credit_amt,
                '摘要': credit['summary'],
                '对方账号': credit['counterparty_acct'],
                '来账对方户名': credit['counterparty'],
                '划转日期': '',
                '划转金额': 0,
                '划转对方行名': '',
                '等待工作日': wdays_since,
                '阈值天数': days_threshold,
                '备注': '',
                '状态': status,
            })

    # ================================================================
    # 4. 排序 + 汇总行
    # ================================================================
    results_df = pd.DataFrame(results)
    if len(results_df) > 0:
        # 可疑排最前面，然后未划转，然后延迟划转，然后已划转
        status_order = {'可疑': 0, '未划转': 1, '延迟划转': 2, '已划转': 3}
        results_df['_sort'] = results_df['状态'].map(status_order).fillna(4)
        results_df = results_df.sort_values(['_sort', '等待工作日', '来源日期'],
                                             ascending=[True, False, True])
        results_df = results_df.drop(columns=['_sort'])

        # 汇总行
        suspicious_count = int((results_df['状态'] == '可疑').sum())
        unmatched_count = int((results_df['状态'] == '未划转').sum())
        delayed_count = int((results_df['状态'] == '延迟划转').sum())
        matched_count = int((results_df['状态'] == '已划转').sum())
        zb_cleared = int((results_df['备注'].str.contains('余额归零', na=False)).sum())

        total_row = {
            '文件来源': file_label,
            '账号': account_num,
            '来源日期': '--- 合计 ---',
            '来源金额': total_in,
            '摘要': '',
            '对方账号': '',
            '来账对方户名': f'已划转{matched_count} / 延迟{delayed_count} / 未划转{unmatched_count} / 可疑{suspicious_count} / 余额归零{zb_cleared}',
            '划转日期': '',
            '划转金额': total_matched,
            '划转对方行名': '',
            '等待工作日': '',
            '阈值天数': '',
            '备注': '',
            '状态': f'流入{total_in} / 已匹配{total_matched} / 未匹配{total_in-total_matched}',
        }
        results_df = pd.concat([
            results_df,
            pd.DataFrame([total_row])
        ], ignore_index=True)

    return results_df


# =========================
# 清算核查二次确认
# =========================


def _normalize_confirm_account_value(value):
    """归一化账号，避免科学计数法导致的精度丢失。"""
    if pd.isna(value):
        return ''
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            if float(value).is_integer():
                return str(int(value))
        except (ValueError, TypeError, OverflowError):
            pass
        return str(value).strip()
    return str(value).strip()


def _normalize_confirm_text_value(value):
    """归一化户名，忽略名称中的空白差异。"""
    if pd.isna(value):
        return ''
    return ''.join(str(value).split())


def _normalize_confirm_code_value(value):
    """归一化凭证类型，兼容 2302 / 2302.0 / '2302'。"""
    if pd.isna(value):
        return ''
    try:
        num = float(str(value).strip())
        if num.is_integer():
            return str(int(num))
    except (ValueError, TypeError):
        pass
    return str(value).strip()


def _normalize_confirm_numeric_series(series):
    """将 CSV 金额列转换为数值，兼容千分位和人民币符号。"""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors='coerce')
    cleaned = (
        series.astype(str)
        .str.replace(',', '', regex=False)
        .str.replace('，', '', regex=False)
        .str.replace('￥', '', regex=False)
        .str.replace('元', '', regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors='coerce')


def _parse_confirm_date(value):
    """解析二次确认 CSV 的日期，兼容带时间的日期字符串。"""
    parsed = parse_date(value)
    if parsed is not None:
        return parsed
    try:
        ts = pd.to_datetime(value, errors='coerce')
        if pd.notna(ts):
            return ts.to_pydatetime()
    except Exception:
        pass
    return None


def _load_settlement_confirm_csv(filepath):
    """加载二次确认 CSV，并仅保留符合 2302 金额换算规则的记录。"""
    if not _is_csv_file(filepath):
        raise ValueError(f"二次确认数据需为 CSV 文件: {filepath}")

    _, df = _load_dataframe(filepath)
    if len(df) == 0:
        return pd.DataFrame()

    alias_map = {
        '付款人名称': '付款人名称',
        '付款人户名': '付款人名称',
        '付款人姓名': '付款人名称',
        '付款人账号': '付款人账号',
        '付款人帐号': '付款人账号',
        '凭证类型': '凭证类型',
        '票据类型': '凭证类型',
        '支付金额': '支付金额',
        '付款金额': '支付金额',
        '发生额': '发生额',
        '发生金额': '发生额',
        '发生额（分）': '发生额',
        '发生额(分)': '发生额',
        '交易日期': '交易日期',
        '业务日期': '交易日期',
        '发生日期': '交易日期',
        '来源日期': '交易日期',
        '支付日期': '交易日期',
        '日期': '交易日期',
        '入账日期': '交易日期',
        '入帐日期': '交易日期',
    }
    rename_map = {}
    for col in df.columns:
        clean_col = str(col).strip()
        if clean_col.startswith('\ufeff'):
            clean_col = clean_col[1:]
        if clean_col in alias_map:
            rename_map[col] = alias_map[clean_col]
    if rename_map:
        df = df.rename(columns=rename_map)

    required_cols = ['交易日期', '付款人名称', '付款人账号', '凭证类型', '支付金额', '发生额']
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        raise ValueError(
            f"确认文件 {filepath} 缺少必需列: {', '.join(missing_cols)}；"
            f"当前列名: {list(df.columns)}"
        )

    df['日期对象'] = df['交易日期'].apply(_parse_confirm_date)
    df = df.dropna(subset=['日期对象'])
    df['付款人名称'] = df['付款人名称'].map(_normalize_confirm_text_value)
    df['付款人账号'] = df['付款人账号'].map(_normalize_confirm_account_value)
    df['凭证类型'] = df['凭证类型'].map(_normalize_confirm_code_value)
    df['支付金额'] = _normalize_confirm_numeric_series(df['支付金额'])
    df['发生额'] = _normalize_confirm_numeric_series(df['发生额'])

    expected_fen = -(df['支付金额'] * 100).round()
    valid_amount = (
        df['支付金额'].notna()
        & df['发生额'].notna()
        & ((df['发生额'] - expected_fen).abs() < 0.01)
    )
    return df.loc[valid_amount].sort_values('日期对象').reset_index(drop=True)


def confirm_settlement_suspicious(suspicious_df, confirm_file_path):
    """
    对清算核查可疑记录做二次确认。

    以可疑记录来源日期为基准，检查当天或下一工作日内是否存在：
    - 付款人名称 = 来账对方户名
    - 付款人账号 = 对方账号
    - 凭证类型 = 2302
    - 支付金额 × 100 取负数后 = 该条记录的发生额

    找到的记录从可疑结果中剔除；找不到则保留。确认 CSV 中的每条 2302 记录
    最多用于剔除一条可疑记录，避免重复使用同一笔确认数据。
    """
    if suspicious_df is None or len(suspicious_df) == 0:
        return suspicious_df.copy() if suspicious_df is not None else pd.DataFrame()
    if not confirm_file_path:
        return suspicious_df.copy()

    required_cols = ['来源日期', '对方账号', '来账对方户名']
    missing_cols = [c for c in required_cols if c not in suspicious_df.columns]
    if missing_cols:
        raise ValueError(
            f"清算核查结果缺少二次确认所需列: {', '.join(missing_cols)}"
        )

    confirm_df = _load_settlement_confirm_csv(confirm_file_path)
    if len(confirm_df) == 0:
        print(f"[清算核查二次确认] 确认文件中无有效 2302 记录，保留全部 {len(suspicious_df)} 条")
        return suspicious_df.copy()

    by_key = {}
    for idx, row in confirm_df.iterrows():
        key = (row['日期对象'].date(), row['付款人名称'], row['付款人账号'], row['凭证类型'])
        by_key.setdefault(key, []).append(idx)

    matched_confirm = set()
    kept = []
    removed = 0
    for _, row in suspicious_df.iterrows():
        source_date = _parse_confirm_date(row['来源日期'])
        if source_date is None:
            kept.append(row.to_dict())
            continue

        name = _normalize_confirm_text_value(row['来账对方户名'])
        account = _normalize_confirm_account_value(row['对方账号'])
        confirm_dates = {source_date.date(), next_workday(source_date).date()}

        found = False
        for d in confirm_dates:
            for confirm_idx in by_key.get((d, name, account, '2302'), []):
                if confirm_idx not in matched_confirm:
                    matched_confirm.add(confirm_idx)
                    found = True
                    break
            if found:
                break

        if found:
            removed += 1
        else:
            kept.append(row.to_dict())

    result_df = pd.DataFrame(kept, columns=suspicious_df.columns)
    print(
        f"[清算核查二次确认] 可疑记录 {len(suspicious_df)} 条，"
        f"确认剔除 {removed} 条，保留 {len(result_df)} 条"
    )
    return result_df


# =========================
# 清算核查简化版：仅输出可疑记录）
# =========================

def detect_settlement_verification(
    file_path: str,
    designated_account = '待报解预算收入',
    days_threshold: int = 2,
    confirm_file_path: Optional[str] = None,
) -> pd.DataFrame:
    """
    清算核查：按退款笔数核对待查资金与匹配资金。

    校验逻辑：
    1. 待查资金：非指定账户汇入的款项（金额>0）。
    2. 匹配资金：所有汇往指定账户的出账（金额<0）。
    3. 遍历交易，对每笔匹配资金：
       - 从附言/摘要中解析退款笔数 N。
       - 结合入帐日期与入帐时间，找出匹配资金发生时间点之前的待查资金。
       - 若 N >= 剩余笔数：认为这些待查资金全部退回，从待查记录中删除。
       - 若 N < 剩余笔数：按金额匹配确定具体哪 N 笔被退回，退回的从待查记录中删除。
    4. 每笔待查资金超过 days_threshold 个工作日仍未退回的，记入超期集合；
       超期记录不从待查池移除，后续匹配资金仍可匹配，最终仅输出超期集合中的信息。

    参数
    ----------
    file_path : str
        流水文件路径（CSV 或 Excel）
    designated_account : str 或 list[str]
        指定划转账户的对方户名，支持多个；
        默认 "待报解预算收入"
    days_threshold : int
        工作日阈值（默认 2）。待查资金超过此工作日未退回 → 标记"违规"
    confirm_file_path : str, optional
        二次确认 CSV 文件路径。提供时，会按 2302 凭证记录剔除已确认的可疑记录。

    返回
    -------
    pd.DataFrame
        仅包含违规的待查资金记录，按来源日期降序排列。
        若无可疑记录则返回空 DataFrame。
    """
    import re
    import bisect

    # ================================================================
    # 1. 加载数据
    # ================================================================
    if _is_csv_file(file_path):
        _, df = _load_dataframe(file_path)
        if len(df) == 0:
            return pd.DataFrame()
        df = _normalize_csv_columns(df, convert_fen=False)
        date_col = None
        for c in ['交易日期', '入帐日期', '入账日期']:
            if c in df.columns:
                date_col = c
                break
        if date_col is None:
            return pd.DataFrame()
        df = df.dropna(subset=[date_col])
        df['日期对象'] = df[date_col].apply(parse_date)
        df = df.dropna(subset=['日期对象'])
        if '交易金额' in df.columns:
            df['交易金额'] = pd.to_numeric(df['交易金额'], errors='coerce')
        elif '发生额' in df.columns:
            df['交易金额'] = pd.to_numeric(df['发生额'], errors='coerce')
        df = _normalize_amount_vectorized(df)
        if '对方户名' not in df.columns or df['对方户名'].isna().all():
            df['对方户名'] = df.get('对方行名', '')
        if '摘要' not in df.columns:
            df['摘要'] = df.get('附言', '')
    else:
        _, df = _load_dataframe(file_path)
        if len(df) == 0:
            return pd.DataFrame()
        df = _normalize_csv_columns(df, convert_fen=False)
        date_col = None
        for c in ['交易日期', '入帐日期', '入账日期']:
            if c in df.columns:
                date_col = c
                break
        if date_col:
            df = df.dropna(subset=[date_col])
            df['日期对象'] = df[date_col].apply(parse_date)
            df = df.dropna(subset=['日期对象'])
        if '交易金额' in df.columns:
            df['交易金额'] = pd.to_numeric(df['交易金额'], errors='coerce')
        elif '发生额' in df.columns:
            df['交易金额'] = pd.to_numeric(df['发生额'], errors='coerce')
        df = _normalize_amount_vectorized(df)
        if '对方户名' not in df.columns or df['对方户名'].isna().all():
            df['对方户名'] = df.get('对方行名', '')
        if '摘要' not in df.columns:
            df['摘要'] = ''

    if len(df) == 0:
        return pd.DataFrame()

    # 合并日期与时间，生成精确到秒的时间对象（向量化）
    time_col = None
    for c in ['入帐时间', '交易时间', '入账时间', '时间']:
        if c in df.columns:
            time_col = c
            break
    if time_col:
        time_str = (
            df[time_col].astype(str)
            .str.strip()
            .str.replace('.', ':', regex=False)
            .str.replace('：', ':', regex=False)
        )
        time_parsed = pd.to_datetime(time_str, format='%H:%M:%S', errors='coerce')
        mask = time_parsed.isna()
        for fmt in ('%H:%M', '%H%M%S', '%H%M'):
            if not mask.any():
                break
            time_parsed.loc[mask] = pd.to_datetime(time_str[mask], format=fmt, errors='coerce')
            mask = time_parsed.isna()

        seconds = (
            time_parsed.dt.hour.fillna(0) * 3600
            + time_parsed.dt.minute.fillna(0) * 60
            + time_parsed.dt.second.fillna(0)
        )
        df['时间对象'] = df['日期对象'] + pd.to_timedelta(seconds, unit='s')

        # 处理 Excel 分数时间（数值型且未被字符串解析覆盖）
        numeric_mask = time_parsed.isna() & pd.api.types.is_numeric_dtype(df[time_col])
        if numeric_mask.any():
            frac = pd.to_numeric(df.loc[numeric_mask, time_col], errors='coerce')
            valid = frac.between(0, 1)
            if valid.any():
                idx = frac[valid].index
                df.loc[idx, '时间对象'] = df.loc[idx, '日期对象'] + pd.to_timedelta(
                    (frac[valid] * 86400).round(), unit='s'
                )
    else:
        df['时间对象'] = df['日期对象']

    # 提取数据来源标识
    file_label = os.path.splitext(os.path.basename(file_path))[0]
    account_num = ''
    for c in ('账号', '帐号'):
        if c in df.columns and len(df) > 0:
            v = df[c].iloc[0]
            if pd.notna(v):
                if isinstance(v, (int, float)):
                    account_num = f'{int(float(v))}'
                else:
                    account_num = str(v).strip()
                break

    # ================================================================
    # 2. 划转账户归一化：支持单个字符串或列表
    # ================================================================
    if isinstance(designated_account, str):
        designated_accounts = {designated_account.strip()}
    else:
        designated_accounts = {a.strip() for a in designated_account}

    # ================================================================
    # 3. 匹配逻辑：
    #    - 待查资金 = 非指定账户汇入
    #    - 匹配资金 = 汇往指定账户的出账
    #    - 在匹配资金发生前，先检查是否有待查资金已超截止日（deadline < 当前日期）
    #      超期记录只标记不移出，仍可被后续匹配资金匹配
    #    - 结合日期+时间戳筛选：仅匹配资金发生时间点之前收到的待查资金
    #    - 根据附言中的退款笔数 N：
    #      N >= 剩余待查笔数 → 全部退回
    #      N < 剩余待查笔数 → 金额匹配确定具体 N 笔
    #    - 无附言笔数时回退到自由 FIFO 金额匹配
    # ================================================================
    df = df.sort_values('时间对象').reset_index(drop=True)

    # 提取列为 Python 列表遍历，避免 to_dict/iterrows 的开销
    amounts = df['交易金额'].tolist()
    dates = df['日期对象'].tolist()
    dts = df['时间对象'].tolist()
    counterparties = (
        df['对方户名'].fillna('').astype(str).str.strip().values
        if '对方户名' in df.columns else [''] * len(df)
    )
    summaries = (
        df['摘要'].fillna('').astype(str).str.strip().values
        if '摘要' in df.columns else [''] * len(df)
    )
    remarks_col = (
        df['附言'].fillna('').astype(str).str.strip().values
        if '附言' in df.columns else [''] * len(df)
    )
    acct_col = None
    for c in ('对方账号', '对方帐号'):
        if c in df.columns:
            acct_col = c
            break
    accts = df[acct_col].values if acct_col else [None] * len(df)

    credits = []           # 所有待查资金记录
    pending = []           # 待查资金索引（按时间顺序）
    pending_times = []     # 对应时间戳
    pending_deadlines = [] # 对应截止日期
    active = []            # 是否未匹配
    matched = set()        # 已退回的待查资金索引
    expired = set()        # 超期未退回的待查资金索引

    def _expire_overdue(current_date):
        """将截止日早于 current_date 且未退回的待查资金标记为违规。

        只把超期记录放入 expired 集合，不移出待查池。
        利用 pending_deadlines 有序性二分查找定位。
        """
        idx = bisect.bisect_left(pending_deadlines, current_date)
        for i in range(idx):
            if active[i]:
                expired.add(pending[i])

    def _find_refund_subset(eligible_items, count, target_amount):
        """
        从 eligible_items 中找出 count 笔，使其金额之和等于 target_amount。
        返回在 eligible_items 中的索引列表，找不到返回 None。

        为避免组合爆炸，规模较大时直接按时间顺序取前 count 笔。
        """
        n = len(eligible_items)
        if count > n:
            return None
        if count == 0:
            return [] if abs(target_amount) < 0.005 else None

        # 全部匹配
        if count == n:
            total = sum(item['amount'] for item in eligible_items)
            if abs(total - target_amount) < 0.005:
                return list(range(n))
            return None

        # 优先尝试按时间顺序的前 count 笔
        first_total = sum(eligible_items[i]['amount'] for i in range(count))
        if abs(first_total - target_amount) < 0.005:
            return list(range(count))

        # 规模太大时避免回溯爆炸，直接按时间顺序取前 count 笔
        if n > 50 or count > 10:
            return list(range(count))

        # 小规模时用回溯精确匹配
        amounts_fen = [round(item['amount'] * 100) for item in eligible_items]
        target_fen = round(target_amount * 100)

        suffix_sum = [0] * (n + 1)
        for i in range(n - 1, -1, -1):
            suffix_sum[i] = suffix_sum[i + 1] + amounts_fen[i]

        def backtrack(start, remaining_count, current_sum, path):
            if remaining_count == 0:
                return path[:] if current_sum == target_fen else None
            if start >= n or n - start < remaining_count:
                return None
            if current_sum > target_fen:
                return None
            if current_sum + suffix_sum[start] < target_fen:
                return None
            for i in range(start, n):
                r = backtrack(
                    i + 1, remaining_count - 1,
                    current_sum + amounts_fen[i],
                    path + [i],
                )
                if r is not None:
                    return r
            return None

        return backtrack(0, count, 0, [])

    for i in range(len(df)):
        amount = amounts[i]
        if pd.isna(amount) or amount == 0:
            continue
        date = dates[i]
        dt = dts[i]
        counterparty = counterparties[i]

        # ── 日期切换时输出进度 ──
        if i == 0 or date != dates[i - 1] if i > 0 else True:
            date_str = date.strftime('%Y-%m-%d')
            print(f"[清算核查] 处理 {date_str} 待查池:{sum(active)} 已退回:{len(matched)} 超期:{len(expired)}")

        # 先处理已超期的待查资金
        _expire_overdue(date)

        if amount > 0:
            # ── 待查资金：非指定账户汇入 → 入队 ──
            if counterparty in designated_accounts:
                continue
            ci = len(credits)
            deadline = add_working_days(date, days_threshold)
            credits.append({
                'date': date, 'datetime': dt, 'amount': round(amount, 2),
                'counterparty': counterparty,
                'summary': summaries[i],
                'deadline': deadline, 'row_idx': i,
                'pending_idx': len(pending),
            })
            pending.append(ci)
            pending_times.append(dt)
            pending_deadlines.append(deadline)
            active.append(True)

        elif amount < 0:
            # ── 匹配资金：汇往指定账户 → 尝试匹配待查资金 ──
            if counterparty not in designated_accounts:
                continue
            debit_amt = round(abs(amount), 2)

            remarks = remarks_col[i] or summaries[i]
            # 优先匹配"授权支付退款共xx笔"，再回退到通用模式
            m = re.search(r'共\s*(\d+)\s*笔', remarks)
            if not m:
                m = re.search(r'(?:退|共)?\s*(\d+)\s*笔', remarks)
            batch_n = int(m.group(1)) if m else None

            # 二分查找定位匹配资金时间点之前的待查资金
            idx_end = bisect.bisect_left(pending_times, dt)
            eligible = [pending[j] for j in range(idx_end) if active[j]]

            if batch_n is not None and eligible:
                if batch_n >= len(eligible):
                    # 附言笔数 >= 剩余笔数，认为全部退回
                    for ci in eligible:
                        matched.add(ci)
                        active[credits[ci]['pending_idx']] = False
                else:
                    # 附言笔数 < 剩余笔数，按金额匹配找出具体 N 笔
                    print(
                        f"[清算核查] 条数({len(eligible)}) > 笔数({batch_n})，需金额匹配: "
                        f"{date.strftime('%Y-%m-%d')} {dt.strftime('%H:%M:%S')} "
                        f"金额:{debit_amt}"
                    )
                    eligible_items = [credits[ci] for ci in eligible]
                    subset = _find_refund_subset(eligible_items, batch_n, debit_amt)
                    if subset is not None:
                        for idx in subset:
                            ci = eligible[idx]
                            matched.add(ci)
                            active[credits[ci]['pending_idx']] = False
                    else:
                        # 金额匹配失败时，按时间顺序默认退前 N 笔
                        for k in range(batch_n):
                            ci = eligible[k]
                            matched.add(ci)
                            active[credits[ci]['pending_idx']] = False
            elif eligible:
                # 无附言笔数，回退到自由 FIFO 金额匹配
                remaining_debit = debit_amt
                for ci in eligible:
                    if remaining_debit < 0.005:
                        break
                    avail = credits[ci]['amount']
                    if avail <= remaining_debit + 0.005:
                        matched.add(ci)
                        active[credits[ci]['pending_idx']] = False
                        remaining_debit -= avail

    # 数据遍历结束后，以最后日期为界再检查一次超期待查资金
    if len(df) > 0:
        last_date = df['日期对象'].max()
        _expire_overdue(last_date)

    # ── 按日期输出汇总 ──
    daily = {}
    for ci, credit in enumerate(credits):
        dk = credit['date'].strftime('%Y-%m-%d')
        if dk not in daily:
            daily[dk] = {'+待查': 0, '✓退回': 0, '✗超期': 0}
        daily[dk]['+待查'] += 1
        if ci in matched:
            daily[dk]['✓退回'] += 1
        if ci in expired:
            daily[dk]['✗超期'] += 1
    if daily:
        print(f"\n{'日期':<12} {'新增待查':>6} {'已退回':>6} {'已超期':>6}")
        print("-" * 34)
        for dk in sorted(daily):
            d = daily[dk]
            print(f"{dk:<12} {d['+待查']:>6} {d['✓退回']:>6} {d['✗超期']:>6}")
        print()

    # ================================================================
    # 4. 生成结果：超期未退回的待查资金 → 违规
    # ================================================================
    results = []

    for ci in expired:
        credit = credits[ci]
        row_idx = credit['row_idx']
        deadline_str = credit['deadline'].strftime('%Y-%m-%d')

        cpty_acct = ''
        if acct_col:
            v = accts[row_idx]
            if pd.notna(v):
                cpty_acct = f'{int(float(v))}' if isinstance(v, (int, float)) else str(v).strip()

        results.append({
            '文件来源': file_label,
            '账号': account_num,
            '来源日期': credit['date'].strftime('%Y-%m-%d'),
            '来源金额': credit['amount'],
            '摘要': credit['summary'],
            '对方账号': cpty_acct,
            '来账对方户名': credit['counterparty'],
            '窗口截止': deadline_str,
            '窗口工作日': days_threshold,
            '状态': '违规',
        })

    results_df = pd.DataFrame(results)
    if len(results_df) > 0:
        results_df = results_df.sort_values('来源日期', ascending=False).reset_index(drop=True)

    print(f"[清算核查] 待查:{len(credits)} 已退回:{len(matched)} 违规:{len(results_df)}")

    if confirm_file_path and len(results_df) > 0:
        results_df = confirm_settlement_suspicious(results_df, confirm_file_path)

    return results_df


# =========================
# 独立运行入口
# =========================

if __name__ == "__main__":
    print("=" * 50)
    print("资金匹配检测")
    print("=" * 50)
    fund_df = detect_fund_matching()
    if len(fund_df) > 0:
        print(f"匹配记录数: {len(fund_df)}")
        print(fund_df.head(10).to_string(index=False))

    print("\n" + "=" * 50)
    print("延迟清算积数检测")
    print("=" * 50)
    delay_df = detect_delayed_settlement()
    if len(delay_df) > 0:
        print(f"延迟清算记录数: {len(delay_df)}")
        print(delay_df.head(10).to_string(index=False))
    else:
        print("未找到延迟清算记录")
