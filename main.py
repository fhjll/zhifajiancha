# -*- coding: utf-8 -*-
"""
执法检测系统 — 主入口

输入流水账户报表 -> 发现违规记录 -> 生成执法文书

用法:
  python main.py
  python main.py --zero-balance 零余额账户.xlsx
  python main.py --skip-report
  python main.py --report-mode llm
  python main.py --output-dir ./结果输出
"""

import argparse
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
from process_zero_balance import detect_fund_matching, detect_non_tax_verification, detect_settlement_verification
from generate_report import batch_generate
from logger import setup_logging, TeeWriter, get_log_path


def _convert_matching_to_report_rows(matching_df):
    """
    将资金匹配检测结果转换为报告生成器所需的格式。

    列名映射:
      检测输出 -> 报告输入
      退款日期  -> 退款日期
      清算日期  -> 清算日期
      退至垫款户日期 -> 退至过渡户日期
      退至金库日期 -> 退回日期
      账号 -> 账户名称
      交易金额 -> 金额

    只保留退款日期和退至金库日期均不为空的完整违规记录。
    """
    rows = []
    for _, r in matching_df.iterrows():
        if pd.isna(r.get("退款日期")) or pd.isna(r.get("退至金库日期")):
            continue
        rows.append(
            {
                "清算日期": r.get("清算日期", ""),
                "退款日期": r.get("退款日期", ""),
                "退至过渡户日期": r.get("退至垫款户日期", ""),
                "退回日期": r.get("退至金库日期", ""),
                "账户名称": r.get("sheet名称", ""),
                "金额": r.get("交易金额", 0),
            }
        )
    return rows


def _summary_stats(df):
    """输出统计摘要"""
    stats = {}
    for col in ["清算日期", "退款日期", "退至垫款户日期", "退至金库日期"]:
        if col in df.columns:
            stats[col] = int(df[col].notna().sum())
    return stats


def run_pipeline(
    zero_balance_path="零余额账户.xlsx",
    qing_suan_path="集中支付零余额清算待转.xlsx",
    advance_acct_name="集中支付零余额清算待转",
    output_dir="output",
    skip_report=False,
    report_mode="template",
    template_path=None,
    non_tax=False,
    non_tax_file=None,
    non_tax_account=None,
    non_tax_days=2,
    settlement_check=None,
    settlement_account="待报解预算收入",
    settlement_days=2,
    settlement_confirm=None,
    max_workers=1,
):
    """
    运行检测-报告全流程。

    参数
    ----------
    max_workers : int
        并行处理进程数（默认 1=顺序处理）。
        - 第1步（资金匹配）中多个 sheet 并行
        - 文书生成（模板填充）使用多线程并行
    """
    os.makedirs(output_dir, exist_ok=True)

    # ========== 步骤1: 资金匹配检测（核心检测逻辑） ==========
    # detect_fund_matching 实现三阶段匹配：
    #   阶段1 — 预建索引（O(N log N)）：提取清算日期候选、负向交易金额索引
    #   阶段2 — 提取待查记录（正向非垫款户交易）
    #   阶段3 — 逐条匹配（O(P log N)）：用二分查找找清算日期 + 金额索引找退至待转户日期
    # 对每个 Sheet（每个账号）独立处理，支持 ProcessPoolExecutor 并行
    print("=" * 50)
    print("步骤1/2: 资金匹配检测（违规记录）")
    print("=" * 50)
    print(f"  零余额账户: {zero_balance_path}")
    print(f"  清算待转文件: {qing_suan_path}")
    print(f"  垫款户名称: {advance_acct_name}")

    matching_results = detect_fund_matching(
        zero_balance_path=zero_balance_path,
        qing_suan_path=qing_suan_path,
        advance_acct_name=advance_acct_name,
        max_workers=max_workers,
    )

    matching_output = os.path.join(output_dir, "违规记录.xlsx")
    matching_results.to_excel(matching_output, index=False)

    stats = _summary_stats(matching_results)
    print(f"  总记录数: {len(matching_results)}")
    print(f"  其中找到清算日期: {stats.get('清算日期', 0)}")
    print(f"  其中找到退款日期: {stats.get('退款日期', 0)}")
    print(f"  其中找到退至垫款户日期: {stats.get('退至垫款户日期', 0)}")
    print(f"  其中找到退至金库日期: {stats.get('退至金库日期', 0)}")
    print(f"  结果已保存: {matching_output}")

    # ========== 步骤2: 生成执法文书 ==========
    if not skip_report:
        print()
        print("=" * 50)
        print("步骤2/2: 生成执法文书")
        print("=" * 50)

        report_rows = _convert_matching_to_report_rows(matching_results)
        if report_rows:
            filepaths = batch_generate(
                rows=report_rows,
                mode=report_mode,
                output_dir=output_dir,
                template_path=template_path,
                max_workers=max(1, max_workers * 2),  # 文书生成用更多线程（I/O 密集型）
            )
            print()
            print(f"  共生成 {len(filepaths)} 份文书:")
            for fp in filepaths:
                print(f"    {fp}")
        else:
            print("  未找到完整的违规记录（需要同时有退款日期和退至金库日期），跳过文书生成")
    else:
        print()
        print("=" * 50)
        print("步骤2/2: 已跳过文书生成")
        print("=" * 50)

    # ========== 步骤4（可选）: 非税核查 ==========
    if non_tax:
        print()
        print("=" * 50)
        print("步骤3/3: 非税专户核查（1:1 精确金额匹配 + 工作日计算）")
        print("=" * 50)

        nt_file = non_tax_file or zero_balance_path
        nt_account = non_tax_account or advance_acct_name

        nt_results = detect_non_tax_verification(
            file_path=nt_file,
            designated_account=nt_account,
            days_threshold=non_tax_days,
        )

        nt_output = os.path.join(output_dir, "非税核查结果.xlsx")
        # 仅保存有问题的记录（延迟划转/未划转/可疑）+ 末尾汇总行
        # 仅保存不正常的记录（排除已划转和汇总行），去掉备注列
        nt_save = nt_results[
            (nt_results['状态'] != '已划转') &
            (nt_results['来源日期'] != '--- 合计 ---')
        ].reset_index(drop=True)
        if '备注' in nt_save.columns:
            nt_save = nt_save.drop(columns=['备注'])
        nt_save.to_excel(nt_output, index=False)

        suspicious = nt_results[nt_results['状态'] == '可疑']
        pending = nt_results[nt_results['状态'].isin(['未划转'])]
        delayed = nt_results[nt_results['状态'] == '延迟划转']
        zb_cleared = nt_results[nt_results['备注'].str.contains('余额归零', na=False)] if '备注' in nt_results.columns else pd.DataFrame()
        split_transfers = nt_results[nt_results['备注'].str.contains('分批划转', na=False)] if '备注' in nt_results.columns else pd.DataFrame()
        print(f"  来账总数: {len(nt_results)}")
        print(f"  可疑（超{non_tax_days}工作日未划转）: {len(suspicious)}")
        print(f"  未划转（在途）: {len(pending)}")
        print(f"  延迟划转: {len(delayed)}")
        if len(zb_cleared) > 0:
            print(f"  余额归零: {len(zb_cleared)}")
        if len(split_transfers) > 0:
            print(f"  分批划转: {len(split_transfers)}")
        if len(nt_results) > 0:
            summary = nt_results.iloc[-1]
            print(f"  流入合计: {summary['来源金额']}")
            print(f"  已匹配合计: {summary['划转金额']}")
            if '来源金额' in summary and '划转金额' in summary:
                try:
                    unmatched = float(summary['来源金额']) - float(summary['划转金额'])
                    print(f"  未匹配合计: {unmatched:.2f}")
                except (ValueError, TypeError):
                    pass
        print(f"  结果已保存: {nt_output}")

    # ========== 步骤5（可选）: 清算退款确认（逐笔匹配确认CSV） ==========
    if settlement_check:
        if not settlement_confirm:
            print("错误: 启用清算退款确认时必须提供确认CSV文件路径", file=sys.stderr)
            return
        print()
        print("=" * 50)
        print("步骤: 清算退款确认（逐笔匹配确认CSV）")
        print("=" * 50)

        sc_results = detect_settlement_verification(
            file_path=settlement_check,
            designated_account=settlement_account,
            days_threshold=settlement_days,
            confirm_file_path=settlement_confirm,
        )

        sc_output = os.path.join(output_dir, "清算核查结果.xlsx")
        print(f"  确认CSV: {settlement_confirm}")
        if len(sc_results) == 0:
            print("  未发现可疑记录（所有退款均在窗口内确认）")
            pd.DataFrame(columns=[
                '文件来源', '账号', '来源日期', '来源金额', '摘要',
                '对方账号', '来账对方户名', '窗口截止', '窗口工作日', '状态'
            ]).to_excel(sc_output, index=False)
        else:
            sc_results.to_excel(sc_output, index=False)
            print(f"  可疑记录数: {len(sc_results)}")
            for _, row in sc_results.iterrows():
                print(f"    {row['来源日期']} | 金额:{row['来源金额']} | "
                      f"来账方:{row['来账对方户名']} | "
                      f"窗口截止:{row['窗口截止']}")
        print(f"  结果已保存: {sc_output}")

    print()
    print("=" * 50)
    print("所有步骤处理完成！")
    print("=" * 50)


def main():
    # ── 初始化日志系统（所有 print() 自动写入日志文件） ──
    setup_logging()
    sys.stdout = TeeWriter(sys.stdout)
    sys.stderr = TeeWriter(sys.stderr)
    print(f"[日志文件] {get_log_path()}")
    print()

    parser = argparse.ArgumentParser(
        description="执法检测系统 - 输入零余额账户+清算待转 -> 发现违规记录 -> 生成执法文书",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  python main.py                                               # 使用默认文件路径\n"

            "  python main.py --zero-balance ./data/零余额.xlsx              # 指定零余额账户\n"
            "  python main.py --template-path ./模板/事实认定书模板.docx       # 自定义模板\n"
            "  python main.py --output-dir ./结果输出 --report-mode llm       # LLM模式生成文书\n"
            "  python main.py --skip-report                                  # 只检测不生成文书\n"
        ),
    )
    parser.add_argument(
        "--zero-balance",
        default="零余额账户.xlsx",
        help="零余额账户文件路径 (默认: 零余额账户.xlsx)",
    )
    parser.add_argument(
        "--qing-suan",
        default="集中支付零余额清算待转.xlsx",
        help="集中支付零余额清算待转文件路径 (默认: 集中支付零余额清算待转.xlsx)",
    )
    parser.add_argument(
        "--advance-name",
        default="集中支付零余额清算待转",
        help="垫款户名称（对方户名），用于识别清算交易 (默认: 集中支付零余额清算待转)",
    )
    parser.add_argument(
        "--non-tax",
        action="store_true",
        help="启用非税核查（滚动匹配划转资金，标记超2日未划转记录）",
    )
    parser.add_argument(
        "--non-tax-file",
        default="",
        help="非税核查的流水文件路径（默认同 --zero-balance）",
    )
    parser.add_argument(
        "--non-tax-account",
        default="",
        help="非税核查的指定划转账户名称（默认同 --advance-name）",
    )
    parser.add_argument(
        "--non-tax-days",
        type=int,
        default=2,
        help="非税核查未划转阈值天数 (默认: 2)",
    )
    parser.add_argument(
        "--settlement-check",
        default="",
        help="启用清算退款确认，指定流水文件路径",
    )
    parser.add_argument(
        "--settlement-account",
        default="待报解预算收入",
        help="兼容参数，新清算退款确认逻辑不再使用",
    )
    parser.add_argument(
        "--settlement-days",
        type=int,
        default=2,
        help="清算退款确认匹配窗口工作日数 (默认: 2)",
    )
    parser.add_argument(
        "--settlement-confirm",
        default="",
        help="清算退款确认 CSV 文件路径（必填，凭证类型编号 2302）",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="输出目录 (默认: output)",
    )
    parser.add_argument(
        "--template-path",
        default="",
        help="事实认定书模板文件路径 (留空则使用默认模板)",
    )
    parser.add_argument(
        "--report-mode",
        choices=["template", "llm"],
        default="template",
        help="文书生成模式: template=模板填充, llm=大模型生成 (默认: template)",
    )
    parser.add_argument(
        "--skip-report",
        action="store_true",
        help="跳过文书生成步骤",
    )
    parser.add_argument(
        "--from-results",
        default="",
        help="从已有的违规记录文件生成文书（跳过检测）",
    )
    parser.add_argument(
        "--workers", "-w",
        type=int,
        default=1,
        help="并行处理进程数（默认 1）。建议设为 CPU 核心数，如 --workers 4",
    )

    args = parser.parse_args()

    # 检查输入文件是否存在
    for name, path in [
        ("零余额账户文件", args.zero_balance),
        ("清算待转文件", args.qing_suan),
    ]:
        if not os.path.exists(path):
            print(f"错误: {name} 不存在: {path}", file=sys.stderr)
            sys.exit(1)

    if args.settlement_check and not args.settlement_confirm:
        print("错误: 启用清算退款确认时必须提供 --settlement-confirm", file=sys.stderr)
        sys.exit(1)

    if args.from_results:
        # 直接由结果文件生成文书
        from generate_report import batch_generate_from_file
        print("=" * 50)
        print("从结果文件生成文书")
        print("=" * 50)
        if not os.path.exists(args.from_results):
            print(f"错误: 结果文件不存在: {args.from_results}", file=sys.stderr)
            sys.exit(1)
        filepaths = batch_generate_from_file(
            results_path=args.from_results,
            mode=args.report_mode,
            output_dir=args.output_dir,
            template_path=args.template_path or None,
        )
        print(f"\n共生成 {len(filepaths)} 份文书:")
        for fp in filepaths:
            print(f"  {fp}")
        return

    # 解析逗号分隔的多账户
    settlement_account_list = None
    if args.settlement_account:
        parts = [a.strip() for a in args.settlement_account.replace('，', ',').split(',') if a.strip()]
        settlement_account_list = parts if len(parts) > 1 else args.settlement_account

    run_pipeline(
        zero_balance_path=args.zero_balance,
        qing_suan_path=args.qing_suan,
        advance_acct_name=args.advance_name,
        output_dir=args.output_dir,
        skip_report=args.skip_report,
        report_mode=args.report_mode,
        template_path=args.template_path or None,
        non_tax=args.non_tax,
        non_tax_file=args.non_tax_file or None,
        non_tax_account=args.non_tax_account or None,
        non_tax_days=args.non_tax_days,
        settlement_check=args.settlement_check or None,
        settlement_account=settlement_account_list or args.settlement_account,
        settlement_days=args.settlement_days,
        settlement_confirm=args.settlement_confirm or None,
        max_workers=args.workers,
    )


if __name__ == "__main__":
    main()
