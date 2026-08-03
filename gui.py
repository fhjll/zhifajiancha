# -*- coding: utf-8 -*-
"""
执法检测系统 — 图形界面（三标签页版）

标签页：
  1. 清算退款 — 资金匹配 + 违规记录生成
  2. 非税收入 — 滚动匹配划转资金核查
  3. 文书生成 — 从结果文件生成执法文书
"""

import os
import sys
import threading
import io
from datetime import datetime
from tkinter import filedialog
import pandas as pd

import customtkinter as ctk

from process_zero_balance import detect_fund_matching, detect_non_tax_verification, detect_settlement_verification
from generate_report import batch_generate
from preprocess import preprocess_transactions, detect_end_of_day_nonzero
from logger import setup_logging, log_message, get_log_path

# ── 外观设置 ──
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class TextRedirector(io.StringIO):
    """将 stdout 重定向到 GUI 日志区域，同时写入文件日志"""

    def __init__(self, text_widget, tag="info"):
        super().__init__()
        self.text_widget = text_widget
        self.tag = tag
        self._lock = threading.Lock()

    def write(self, s):
        if s.strip():
            log_message(s.rstrip('\n'))  # 写入文件日志
        with self._lock:
            self.text_widget.after(0, self._append, s)

    def _append(self, s):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", s, self.tag)
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")

    def flush(self):
        pass


# ── 工具函数 ──

def _safe_float(v: str, default: float) -> float:
    try:
        return float(v.strip())
    except (ValueError, AttributeError):
        return default


def _parse_multi_account(account_str: str):
    """解析逗号分隔的多账户字符串，返回列表（多个）或字符串（单个）"""
    parts = [a.strip() for a in account_str.replace('，', ',').split(',') if a.strip()]
    return parts if len(parts) > 1 else (parts[0] if parts else account_str)


def _safe_int(v: str, default: int) -> int:
    try:
        return int(v.strip())
    except (ValueError, AttributeError):
        return default


class App(ctk.CTk):
    TITLE = "执法检测系统"
    WIDTH = 900
    HEIGHT = 820

    LABEL_W = 80

    def __init__(self):
        super().__init__()

        # ── 初始化日志系统 ──
        setup_logging()

        self.title(self.TITLE)
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.minsize(700, 700)

        # ── 状态变量 ──
        # 清算退款
        self.zero_balance_path = ctk.StringVar(value="零余额账户.xlsx")
        self.qing_suan_path = ctk.StringVar(value="集中支付零余额清算待转.xlsx")
        self.advance_acct_name = ctk.StringVar(value="集中执法垫款户")
        self.output_dir = ctk.StringVar(value="output")
        self.skip_report = ctk.BooleanVar(value=False)

        # 非税收入
        self.non_tax_file_path = ctk.StringVar(value="零余额账户.xlsx")
        self.non_tax_account_name = ctk.StringVar(value="待报解预算收入")
        self.non_tax_days = ctk.StringVar(value="2")

        # 清算核查
        self.settlement_check_file = ctk.StringVar(value="")
        self.settlement_check_account = ctk.StringVar(value="待报解预算收入")
        self.settlement_check_days = ctk.StringVar(value="2")
        self.settlement_confirm_file = ctk.StringVar(value="")

        # ── 流水预处理 ──
        self.preprocess_input_path = ctk.StringVar(value="")
        self.preprocess_output_path = ctk.StringVar(value="")

        # ── 延迟清算 ──
        self.delayed_settlement_input = ctk.StringVar(value="")
        self.delayed_settlement_output = ctk.StringVar(value="")

        # 文书生成
        self.results_path = ctk.StringVar(value="")
        self.template_path = ctk.StringVar(value="")
        self.report_mode = ctk.StringVar(value="template")

        # ── LLM 配置 ──
        self.llm_api_key = ctk.StringVar(value=os.environ.get("DEEPSEEK_API_KEY", ""))
        self.llm_api_base = ctk.StringVar(value="https://api.deepseek.com")
        self.llm_model = ctk.StringVar(value="deepseek-chat")
        self.llm_temperature = ctk.StringVar(value="0.3")
        self.llm_max_tokens = ctk.StringVar(value="2048")
        self.llm_top_p = ctk.StringVar(value="1.0")

        # 运行状态
        self.running = False
        self.running_nt = False
        self.running_sc = False

        # 并行控制
        self.max_workers = ctk.StringVar(value="1")

        self._build_ui()

        # 居中
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    # ──────────── UI 构建 ────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)  # 日志区域可伸缩

        # ── 标题栏 ──
        title_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="ew", padx=24, pady=(16, 6))
        title_frame.grid_columnconfigure(1, weight=1)

        # LLM 配置按钮 + 版本号（左上角）
        self.llm_config_btn = ctk.CTkButton(
            title_frame, text="⚙  大模型配置",
            font=ctk.CTkFont(size=12),
            width=110, height=28,
            corner_radius=6,
            command=self._open_llm_config,
        )
        self.llm_config_btn.grid(row=0, column=0, sticky="w", padx=(0, 8))

        ctk.CTkLabel(
            title_frame, text="v1.0",
            font=ctk.CTkFont(size=11), text_color="gray",
        ).grid(row=0, column=1, sticky="w", padx=(0, 12))

        ctk.CTkLabel(
            title_frame,
            text=self.TITLE,
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=2, sticky="w", padx=(0, 12))

        ctk.CTkLabel(
            title_frame,
            text="清算退款 | 非税收入 | 文书生成",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).grid(row=0, column=3, sticky="w")

        # ── 三标签页 ──
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 8))
        self.grid_rowconfigure(1, weight=0)  # 不伸缩

        self.tab_clear = self.tabview.add("清算退款")
        self.tab_non_tax = self.tabview.add("非税收入")
        self.tab_report = self.tabview.add("文书生成")
        self.tab_preprocess = self.tabview.add("流水预处理")

        # 为每个标签页构建内容
        self._build_tab_clear()
        self._build_tab_non_tax()
        self._build_tab_report()
        self._build_tab_preprocess()

        # ── 日志标签 ──
        log_label = ctk.CTkLabel(
            self, text="📋 运行日志",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        )
        log_label.grid(row=2, column=0, sticky="ew", padx=24, pady=(4, 2))

        # ── 日志区域（可伸缩） ──
        log_frame = ctk.CTkFrame(self, corner_radius=8)
        log_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 16))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
            state="disabled",
        )
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        # 日志颜色标签
        self.log_text.tag_config("info", foreground="#888888")
        self.log_text.tag_config("step", foreground="#0066CC")
        self.log_text.tag_config("success", foreground="#339933")
        self.log_text.tag_config("error", foreground="#CC3333")
        self.log_text.tag_config("result", foreground="#CC6600")

        # 重定向 stdout（同时写入文件日志）
        sys.stdout = TextRedirector(self.log_text, "info")
        print(f"📁 日志文件: {get_log_path()}")
        print("执法检测系统 GUI 已启动")
        print("请选择对应标签页中的文件后点击运行")
        self._original_stdout = sys.stdout

    # ── 标签页1: 清算退款 ──

    def _build_tab_clear(self):
        tab = self.tab_clear
        LW = self.LABEL_W
        tab.grid_columnconfigure(1, weight=1)

        r = 0

        # Section: 输入文件
        ctk.CTkLabel(tab, text="📂 输入文件",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(12, 4), padx=(12, 0))
        r += 1

        for label, var in [
            ("零余额账户", self.zero_balance_path),
            ("清算待转", self.qing_suan_path),
        ]:
            row_f = ctk.CTkFrame(tab, fg_color="transparent")
            row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
            row_f.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(row_f, text=label, width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
            ctk.CTkEntry(row_f, textvariable=var).grid(row=0, column=1, sticky="ew", padx=(0, 8))
            ctk.CTkButton(row_f, text="浏览", width=64,
                           command=lambda v=var: self._browse_file(v)
                           ).grid(row=0, column=2, padx=(0, 6))
            r += 1

        # 垫款户名称
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="垫款户名称", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.advance_acct_name).grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=(0, 6))
        r += 1

        # Section: 输出设置
        ctk.CTkLabel(tab, text="⚙️ 输出设置",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(14, 4), padx=(12, 0))
        r += 1

        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="输出目录", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.output_dir).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=64,
                       command=self._browse_output_dir).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # 跳过文书
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(4, 4), padx=(6, 6))
        ctk.CTkLabel(row_f, text="", width=LW).grid(row=0, column=0, padx=(6, 8))
        ctk.CTkCheckBox(row_f, text="跳过文书生成（只进行检测）",
                         variable=self.skip_report).grid(row=0, column=1, sticky="w")
        r += 1

        # 并行进程数
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(0, 8), padx=(6, 6))
        ctk.CTkLabel(row_f, text="并行进程数", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.max_workers, width=60).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(row_f, text="设为1则顺序处理，多sheet/文件时可提高",
                     text_color="gray", font=ctk.CTkFont(size=11)).grid(row=0, column=2, sticky="w")
        r += 1

        # 运行按钮
        self.run_btn = ctk.CTkButton(
            tab, text="▶  开始检测",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=44, corner_radius=8,
            command=self._run_pipeline_clear,
        )
        self.run_btn.grid(row=r, column=0, columnspan=3, sticky="ew", padx=12, pady=(8, 12))
        r += 1

        # ── 清算核查（仅输出可疑记录） ──
        ctk.CTkLabel(tab, text="",
                     font=ctk.CTkFont(size=1)).grid(row=r, column=0, columnspan=3)
        r += 1

        ctk.CTkLabel(tab, text="🔍 清算核查（仅输出可疑记录）",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(12, 4), padx=(12, 0))
        r += 1

        # 流水文件
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="流水文件", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.settlement_check_file,
                     placeholder_text="选择需要核查的流水文件...").grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=64,
                       command=lambda: self._browse_file(self.settlement_check_file)
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # 二次确认 CSV
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="确认CSV", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.settlement_confirm_file,
                     placeholder_text="可选：2302凭证二次确认文件...").grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=64,
                       command=lambda: self._browse_file(self.settlement_confirm_file)
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # 指定账户名称
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="划转账户", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.settlement_check_account,
                     placeholder_text="多账户用逗号分隔，如: 待报解预算收入,国库经收处").grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=(0, 6))
        r += 1

        # 阈值天数
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        ctk.CTkLabel(row_f, text="阈值天数", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.settlement_check_days, width=60).grid(
            row=0, column=1, sticky="w")
        ctk.CTkLabel(row_f, text="超过此工作日数未划转即标记为可疑（跳过周末）",
                     text_color="gray", font=ctk.CTkFont(size=11)).grid(
            row=0, column=2, sticky="w")
        r += 1

        # 运行按钮
        self.run_sc_btn = ctk.CTkButton(
            tab, text="▶  清算核查",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=38, corner_radius=8,
            command=self._run_settlement_check,
        )
        self.run_sc_btn.grid(row=r, column=0, columnspan=3, sticky="ew", padx=12, pady=(12, 12))
        r += 1

    # ── 标签页2: 非税收入 ──

    def _build_tab_non_tax(self):
        tab = self.tab_non_tax
        LW = self.LABEL_W
        tab.grid_columnconfigure(1, weight=1)

        r = 0

        ctk.CTkLabel(tab, text="📂 非税账户流水文件",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(12, 4), padx=(12, 0))
        r += 1

        # 非税账户流水文件选择
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="流水文件", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.non_tax_file_path).grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=64,
                       command=lambda: self._browse_file(self.non_tax_file_path)
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # Section: 核查参数
        ctk.CTkLabel(tab, text="⚙️ 核查参数",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(14, 4), padx=(12, 0))
        r += 1

        # 划转账户名称
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="划转账户", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.non_tax_account_name).grid(
            row=0, column=1, columnspan=2, sticky="ew", padx=(0, 6))
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(0, 4), padx=(6, 6))
        ctk.CTkLabel(row_f, text="", width=LW).grid(row=0, column=0, padx=(6, 8))
        ctk.CTkLabel(row_f, text="对方户名为此名称的支出将被视为划转（1:1精确金额匹配，不可拆分）",
                     text_color="gray", font=ctk.CTkFont(size=11)).grid(row=0, column=1, sticky="w")
        r += 1

        # 阈值天数
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        ctk.CTkLabel(row_f, text="阈值天数", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.non_tax_days, width=60).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(row_f, text="超过此工作日数未划转即标记为可疑（跳过周末）",
                     text_color="gray", font=ctk.CTkFont(size=11)).grid(row=0, column=2, sticky="w")
        r += 1

        # 运行按钮
        self.run_nt_btn = ctk.CTkButton(
            tab, text="▶  开始核查",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=44, corner_radius=8,
            command=self._run_non_tax,
        )
        self.run_nt_btn.grid(row=r, column=0, columnspan=3, sticky="ew", padx=12, pady=(20, 12))
        r += 1

    # ── 标签页3: 文书生成 ──

    def _build_tab_report(self):
        tab = self.tab_report
        LW = self.LABEL_W
        tab.grid_columnconfigure(1, weight=1)

        r = 0

        ctk.CTkLabel(tab, text="📄 从已有结果文件生成文书",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(12, 4), padx=(12, 0))
        r += 1

        # 结果文件
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="结果文件", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.results_path).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=64,
                       command=lambda: self._browse_file(self.results_path)
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # Section: 输出设置
        ctk.CTkLabel(tab, text="⚙️ 输出设置",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(14, 4), padx=(12, 0))
        r += 1

        for label, var, browse_fn, ph in [
            ("模板文件", self.template_path,
             lambda: self._browse_docx(self.template_path),
             "留空则使用默认模板 (事实认定书模板.docx)"),
            ("输出目录", self.output_dir,
             self._browse_output_dir,
             "文书将保存在此目录"),
        ]:
            row_f = ctk.CTkFrame(tab, fg_color="transparent")
            row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
            row_f.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(row_f, text=label, width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
            ctk.CTkEntry(row_f, textvariable=var, placeholder_text=ph).grid(
                row=0, column=1, sticky="ew", padx=(0, 8))
            ctk.CTkButton(row_f, text="浏览", width=64,
                           command=browse_fn).grid(row=0, column=2, padx=(0, 6))
            r += 1

        # 文书模式
        ctk.CTkLabel(tab, text="📝 文书模式",
                     font=ctk.CTkFont(size=14, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(14, 4), padx=(12, 0))
        r += 1

        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        ctk.CTkLabel(row_f, text="", width=LW).grid(row=0, column=0, padx=(6, 8))
        opt_f = ctk.CTkFrame(row_f, fg_color="transparent")
        opt_f.grid(row=0, column=1, sticky="w")
        ctk.CTkRadioButton(opt_f, text="模板填充 (template)",
                           variable=self.report_mode, value="template",
                           ).pack(side="left", padx=(0, 16))
        ctk.CTkRadioButton(opt_f, text="LLM 生成 (需设置 API Key)",
                           variable=self.report_mode, value="llm",
                           ).pack(side="left")
        r += 1

        # 生成按钮
        ctk.CTkButton(
            tab, text="📄  生成文书",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=44, corner_radius=8,
            command=self._run_report_from_file,
        ).grid(row=r, column=0, columnspan=3, sticky="ew", padx=12, pady=(14, 12))

    # ── 标签页4: 流水预处理 + 延迟清算 ──

    def _build_tab_preprocess(self):
        tab = self.tab_preprocess
        LW = self.LABEL_W
        tab.grid_columnconfigure(1, weight=1)

        r = 0

        # ── Section 1: 流水文件预处理 ──
        ctk.CTkLabel(tab, text="📂 流水文件预处理",
                     font=ctk.CTkFont(size=15, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(14, 4), padx=(12, 0))
        r += 1

        ctk.CTkLabel(tab,
                     text="将多个账户混合在同一个Sheet中的流水，按账户拆分到独立Sheet，并按交易日期+时间排序保存为副本",
                     font=ctk.CTkFont(size=11), text_color="gray",
                     wraplength=600, justify="left"
                     ).grid(row=r, column=0, columnspan=3, sticky="w", padx=(12, 0), pady=(0, 8))
        r += 1

        # 输入文件
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="流水文件", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.preprocess_input_path).grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=64,
                       command=lambda: self._browse_file(self.preprocess_input_path)
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # 输出路径（可选）
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="输出文件", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.preprocess_output_path,
                     placeholder_text="留空则自动生成（原文件名_预处理.xlsx）").grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="另存为", width=64,
                       command=lambda: self._browse_save_as(self.preprocess_output_path,
                                                            "Excel 文件", ".xlsx")
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # 预处理运行按钮
        self.run_preprocess_btn = ctk.CTkButton(
            tab, text="▶  开始预处理",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=42, corner_radius=8,
            command=self._run_preprocess,
        )
        self.run_preprocess_btn.grid(row=r, column=0, columnspan=3, sticky="ew",
                                     padx=12, pady=(12, 6))
        r += 1

        # ── 分隔 ──
        ctk.CTkLabel(tab, text="",
                     font=ctk.CTkFont(size=1)).grid(row=r, column=0, columnspan=3)
        r += 1

        sep = ctk.CTkFrame(tab, height=2, fg_color="#CCCCCC")
        sep.grid(row=r, column=0, columnspan=3, sticky="ew", padx=12, pady=4)
        r += 1

        # ── Section 2: 延迟清算检测 ──
        ctk.CTkLabel(tab, text="🔍 延迟清算检测",
                     font=ctk.CTkFont(size=15, weight="bold")
                     ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(14, 4), padx=(12, 0))
        r += 1

        ctk.CTkLabel(tab,
                     text="对所有账户、所有日期，筛选出当日最后一笔交易余额不为零的记录（即存在延迟清算嫌疑）",
                     font=ctk.CTkFont(size=11), text_color="gray",
                     wraplength=600, justify="left"
                     ).grid(row=r, column=0, columnspan=3, sticky="w", padx=(12, 0), pady=(0, 8))
        r += 1

        # 输入文件（预处理后的Excel）
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="输入文件", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.delayed_settlement_input,
                     placeholder_text="选择预处理后的Excel（或任意流水Excel）").grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="浏览", width=64,
                       command=lambda: self._browse_file(self.delayed_settlement_input)
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # 输出路径（可选）
        row_f = ctk.CTkFrame(tab, fg_color="transparent")
        row_f.grid(row=r, column=0, columnspan=3, sticky="ew", pady=2, padx=(6, 6))
        row_f.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(row_f, text="输出文件", width=LW, anchor="w").grid(row=0, column=0, padx=(6, 8))
        ctk.CTkEntry(row_f, textvariable=self.delayed_settlement_output,
                     placeholder_text="留空则自动生成（延迟清算结果.xlsx）").grid(
            row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row_f, text="另存为", width=64,
                       command=lambda: self._browse_save_as(self.delayed_settlement_output,
                                                            "Excel 文件", ".xlsx")
                       ).grid(row=0, column=2, padx=(0, 6))
        r += 1

        # 延迟清算运行按钮
        self.run_delayed_btn = ctk.CTkButton(
            tab, text="▶  延迟清算检测",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=42, corner_radius=8,
            command=self._run_delayed_settlement,
        )
        self.run_delayed_btn.grid(row=r, column=0, columnspan=3, sticky="ew",
                                  padx=12, pady=(14, 12))
        r += 1

    # ── 文件选择（另存为） ──

    def _browse_save_as(self, var, description="Excel 文件", ext=".xlsx"):
        path = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=ext,
            filetypes=[(description, f"*{ext}"), ("所有文件", "*.*")],
        )
        if path:
            var.set(path)

    # ──────────── 流水预处理运行 ────────────

    def _run_preprocess(self):
        if getattr(self, '_preprocess_running', False):
            return

        input_path = self.preprocess_input_path.get().strip()
        if not input_path or not os.path.exists(input_path):
            self._log_error(f"输入文件不存在: {input_path}")
            return

        self._preprocess_running = True
        self.run_preprocess_btn.configure(text="⏳  处理中...", state="disabled")

        thread = threading.Thread(target=self._preprocess_worker, daemon=True)
        thread.start()

    def _preprocess_worker(self):
        try:
            self._run_preprocess_internal()
        except Exception as e:
            self._log_error(f"预处理出错: {e}")
            import traceback
            self._log_error(traceback.format_exc())
        finally:
            self._preprocess_running = False
            self.after(0, lambda: self.run_preprocess_btn.configure(
                text="▶  开始预处理", state="normal"))

    def _run_preprocess_internal(self):
        input_path = self.preprocess_input_path.get().strip()
        output_path = self.preprocess_output_path.get().strip() or None

        self._log_step("═══ 流水文件预处理 ═══")
        self._log_result(f"  输入文件: {input_path}")

        result_path = preprocess_transactions(input_path, output_path)

        if result_path:
            self._log_success(f"预处理完成！已保存: {result_path}")
            # 自动填入延迟清算的输入
            self.delayed_settlement_input.set(result_path)
            self._log_result('  已自动将结果填入下方"延迟清算检测"的输入文件')
        else:
            self._log_error("预处理失败")

        self._log_step("═══════════════════════════════")

    # ──────────── 延迟清算检测运行 ────────────

    def _run_delayed_settlement(self):
        if getattr(self, '_delayed_running', False):
            return

        input_path = self.delayed_settlement_input.get().strip()
        if not input_path or not os.path.exists(input_path):
            self._log_error(f"输入文件不存在: {input_path}")
            return

        self._delayed_running = True
        self.run_delayed_btn.configure(text="⏳  检测中...", state="disabled")

        thread = threading.Thread(target=self._delayed_settlement_worker, daemon=True)
        thread.start()

    def _delayed_settlement_worker(self):
        try:
            self._run_delayed_settlement_internal()
        except Exception as e:
            self._log_error(f"延迟清算检测出错: {e}")
            import traceback
            self._log_error(traceback.format_exc())
        finally:
            self._delayed_running = False
            self.after(0, lambda: self.run_delayed_btn.configure(
                text="▶  延迟清算检测", state="normal"))

    def _run_delayed_settlement_internal(self):
        input_path = self.delayed_settlement_input.get().strip()
        output_path = self.delayed_settlement_output.get().strip() or None

        self._log_step("═══ 延迟清算检测 ═══")
        self._log_result(f"  输入文件: {input_path}")

        results = detect_end_of_day_nonzero(input_path, output_path)

        if results is not None and len(results) > 0:
            self._log_result(f"  发现 {len(results)} 条余额未归零的记录:")
            for _, row in results.head(20).iterrows():
                acct = row.get('账户', '')
                d = row.get('日期', '')
                bal = row.get('余额', '')
                self._log_result(f"    {acct} | {d} | 余额: {bal}")
            if len(results) > 20:
                self._log_result(f"    ... 共 {len(results)} 条，详细结果请查看保存的文件")
            out_path = self.delayed_settlement_output.get().strip()
            if not out_path:
                base_dir = os.path.dirname(input_path) or '.'
                out_path = os.path.join(base_dir, "延迟清算结果.xlsx")
            self._log_success(f"延迟清算检测完成！已保存: {out_path}")
        elif results is not None:
            self._log_success("未发现余额未归零的记录（所有账户每日末余额均为零）")
        else:
            self._log_error("检测失败")

        self._log_step("═══════════════════════════════")

    def _browse_file(self, var):
        path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[("Excel/CSV 文件", "*.xlsx *.xls *.csv *.txt"), ("所有文件", "*.*")],
        )
        if path:
            var.set(path)

    def _browse_docx(self, var):
        path = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            var.set(path)

    def _browse_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    # ──────────── LLM 配置菜单 ────────────

    @property
    def llm_config(self) -> dict:
        """将 GUI 中 LLM 配置字段汇总为一个 dict，供 generate_report 使用。"""
        return {
            "api_key": self.llm_api_key.get(),
            "api_base": self.llm_api_base.get() or "https://api.deepseek.com",
            "model": self.llm_model.get() or "deepseek-chat",
            "temperature": _safe_float(self.llm_temperature.get(), 0.3),
            "max_tokens": _safe_int(self.llm_max_tokens.get(), 2048),
            "top_p": _safe_float(self.llm_top_p.get(), 1.0),
        }

    def _open_llm_config(self):
        """弹出大模型配置对话框"""
        dlg = ctk.CTkToplevel(self)
        dlg.title("大模型配置")
        dlg.geometry("480x420")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()  # 模态

        # 居中
        dlg.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - dlg.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{x}+{y}")

        main = ctk.CTkFrame(dlg, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=(16, 12))
        main.grid_columnconfigure(1, weight=1)

        # ── 帮助提示 ──
        ctk.CTkLabel(
            main, text="配置用于生成执法文书的大模型参数",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))

        ctk.CTkLabel(
            main, text="💡 Key 留空则自动传占位符：Ollama / 内网自部署可直接留空；远程 API 会报认证错误",
            font=ctk.CTkFont(size=11),
            text_color="#888888",
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12), padx=(0, 0))

        # ── 字段 ──
        entries = []
        r = 1
        for label, var, kwargs in [
            ("API Key", self.llm_api_key, {"show": "*"}),
            ("API Base URL", self.llm_api_base, {}),
            ("Model", self.llm_model, {}),
            ("Temperature", self.llm_temperature, {}),
            ("Max Tokens", self.llm_max_tokens, {}),
            ("Top P", self.llm_top_p, {}),
        ]:
            ctk.CTkLabel(main, text=label, width=90, anchor="w").grid(
                row=r, column=0, padx=(0, 8), pady=4, sticky="w")
            entry = ctk.CTkEntry(main, textvariable=var, **kwargs)
            entry.grid(row=r, column=1, columnspan=2, sticky="ew", pady=4)
            entries.append(entry)
            r += 1

        # ── API Key 显示/隐藏切换 ──
        def toggle_key_visibility():
            api_entry = entries[0]
            current = api_entry.cget("show")
            api_entry.configure(show="" if current == "*" else "*")

        ctk.CTkButton(
            main, text="👁", width=32,
            command=toggle_key_visibility,
        ).grid(row=1, column=3, padx=(4, 0), pady=4)

        # ── 底部按钮 ──
        btn_frame = ctk.CTkFrame(main, fg_color="transparent")
        btn_frame.grid(row=r, column=0, columnspan=4, pady=(16, 0))
        ctk.CTkButton(btn_frame, text="关闭", width=100,
                       command=dlg.destroy).pack(side="right", padx=4)

    # ──────────── 日志方法 ────────────

    def _log_step(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"\n{msg}\n", "step")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log_result(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{msg}\n", "result")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log_success(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"  {msg}\n", "success")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log_error(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"  {msg}\n", "error")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ──────────── 清算退款：运行管线 ────────────

    def _run_pipeline_clear(self):
        if self.running:
            return

        # 文件检查
        files = [
            ("零余额账户文件", self.zero_balance_path.get()),
            ("清算待转文件", self.qing_suan_path.get()),
        ]
        missing = []
        for name, path in files:
            if not os.path.exists(path):
                missing.append(f"{name}: {path}")

        if missing:
            msg = "输入文件不存在:\n" + "\n".join(f"  • {m}" for m in missing)
            self._log_error(msg)
            return

        self.running = True
        self.run_btn.configure(text="⏳  运行中...", state="disabled")

        thread = threading.Thread(target=self._pipeline_clear_worker, daemon=True)
        thread.start()

    def _pipeline_clear_worker(self):
        try:
            self._run_clear_internal()
        except Exception as e:
            self._log_error(f"运行出错: {e}")
            import traceback
            self._log_error(traceback.format_exc())
        finally:
            self.running = False
            self.after(0, lambda: self.run_btn.configure(
                text="▶  开始检测", state="normal"))

    def _run_clear_internal(self):
        zero_balance = self.zero_balance_path.get()
        qing_suan = self.qing_suan_path.get()
        advance_name = self.advance_acct_name.get().strip()
        out_dir = self.output_dir.get()
        skip = self.skip_report.get()
        try:
            workers = int(self.max_workers.get())
        except ValueError:
            workers = 1

        # 自动匹配垫款户文件
        if qing_suan == "集中支付零余额清算待转.xlsx" and advance_name:
            for ext in ['.xlsx', '.csv']:
                candidate = advance_name + ext
                if os.path.exists(candidate):
                    qing_suan = candidate
                    self._log_result(f"自动匹配垫款户文件: {candidate}")
                    break

        os.makedirs(out_dir, exist_ok=True)

        # 步骤1: 资金匹配
        self._log_step("═══ 步骤 1/2: 资金匹配检测 ═══")
        matching_results = detect_fund_matching(
            zero_balance_path=zero_balance,
            qing_suan_path=qing_suan,
            advance_acct_name=advance_name,
            max_workers=workers,
        )

        matching_output = os.path.join(out_dir, "违规记录.xlsx")
        matching_results.to_excel(matching_output, index=False)

        total = len(matching_results)
        matched = int(matching_results["退款日期"].notna().sum())
        cleared = int(matching_results["清算日期"].notna().sum())
        tui_zhuan = int(matching_results["退至垫款户日期"].notna().sum())
        tui_guo = int(matching_results["退至金库日期"].notna().sum())

        print(f"  总记录数: {total}")
        print(f"  找到清算日期: {cleared}")
        print(f"  找到退款日期: {matched}")
        print(f"  找到退至垫款户日期: {tui_zhuan}")
        print(f"  找到退至金库日期: {tui_guo}")
        self._log_success(f"结果已保存: {matching_output}")

        # 步骤2（可选）：文书生成
        if not skip:
            self._log_step("═══ 步骤 2/2: 生成执法文书 ═══")
            report_rows = []
            for _, r in matching_results.iterrows():
                if pd.isna(r.get("退款日期")) or pd.isna(r.get("退至金库日期")):
                    continue
                report_rows.append({
                    "清算日期": r.get("清算日期", ""),
                    "退款日期": r.get("退款日期", ""),
                    "退至过渡户日期": r.get("退至垫款户日期", ""),
                    "退回日期": r.get("退至金库日期", ""),
                    "账户名称": r.get("sheet名称", ""),
                    "金额": r.get("交易金额", 0),
                })

            if report_rows:
                filepaths = batch_generate(
                    rows=report_rows,
                    mode=self.report_mode.get(),
                    output_dir=out_dir,
                    template_path=self.template_path.get() or None,
                    max_workers=max(1, workers * 2),
                    llm_config=self.llm_config,
                )
                print(f"\n  共生成 {len(filepaths)} 份文书:")
                for fp in filepaths:
                    print(f"    {fp}")
                self._log_success("文书生成完成")
            else:
                self._log_result("  未找到完整违规记录，跳过文书生成")
        else:
            print("  已跳过文书生成")

        self._log_step("═══════════════════════════════")
        self._log_success("清算退款检测完成！")

    # ──────────── 非税收入：运行核查 ────────────

    def _run_non_tax(self):
        if self.running_nt:
            return

        file_path = self.non_tax_file_path.get()
        if not os.path.exists(file_path):
            self._log_error(f"非税账户流水文件不存在: {file_path}")
            return

        self.running_nt = True
        self.run_nt_btn.configure(text="⏳  核查中...", state="disabled")

        thread = threading.Thread(target=self._non_tax_worker, daemon=True)
        thread.start()

    def _non_tax_worker(self):
        try:
            self._run_non_tax_internal()
        except Exception as e:
            self._log_error(f"非税核查出错: {e}")
            import traceback
            self._log_error(traceback.format_exc())
        finally:
            self.running_nt = False
            self.after(0, lambda: self.run_nt_btn.configure(
                text="▶  开始核查", state="normal"))

    def _run_non_tax_internal(self):
        file_path = self.non_tax_file_path.get()
        account_name = self.non_tax_account_name.get().strip() or "集中支付零余额清算待转"
        try:
            threshold = int(self.non_tax_days.get())
        except ValueError:
            threshold = 2

        out_dir = self.output_dir.get()
        os.makedirs(out_dir, exist_ok=True)

        self._log_step("═══ 非税收入：滚动匹配划转核查 ═══")
        self._log_result(f"  流水文件: {file_path}")
        self._log_result(f"  指定账户: {account_name}")
        self._log_result(f"  阈值天数: {threshold}")

        results = detect_non_tax_verification(
            file_path=file_path,
            designated_account=account_name,
            days_threshold=threshold,
        )

        output_path = os.path.join(out_dir, "非税核查结果.xlsx")
        # 仅保存有问题的记录（延迟划转/未划转/可疑）+ 末尾汇总行
        # 仅保存不正常的记录（排除已划转和汇总行），去掉备注列
        save_df = results[
            (results['状态'] != '已划转') &
            (results['来源日期'] != '--- 合计 ---')
        ].reset_index(drop=True)
        if '备注' in save_df.columns:
            save_df = save_df.drop(columns=['备注'])
        save_df.to_excel(output_path, index=False)

        suspicious = results[results['状态'] == '可疑']
        pending = results[results['状态'].isin(['未划转'])]
        delayed = results[results['状态'] == '延迟划转']
        zb_cleared = results[results['备注'].str.contains('余额归零', na=False)] if '备注' in results.columns else pd.DataFrame()
        split_transfers = results[results['备注'].str.contains('分批划转', na=False)] if '备注' in results.columns else pd.DataFrame()

        print(f"  来账总数: {len(results)}")
        print(f"  可疑（超{threshold}工作日未划转）: {len(suspicious)}")
        print(f"  未划转（在途）: {len(pending)}")
        print(f"  延迟划转: {len(delayed)}")
        if len(zb_cleared) > 0:
            print(f"  余额归零: {len(zb_cleared)}")
        if len(split_transfers) > 0:
            print(f"  分批划转: {len(split_transfers)}")
        if len(results) > 0:
            s = results.iloc[-1]
            print(f"  流入合计: {s['来源金额']}")
            print(f"  已匹配合计: {s['划转金额']}")
            if '来源金额' in s and '划转金额' in s:
                try:
                    unmatched = float(s['来源金额']) - float(s['划转金额'])
                    print(f"  未匹配合计: {unmatched:.2f}")
                except (ValueError, TypeError):
                    pass
        self._log_success(f"非税核查结果已保存: {output_path}")

        self._log_step("═══════════════════════════════")
        self._log_success("非税核查完成！")

    # ──────────── 清算核查（仅输出可疑记录） ────────────

    def _run_settlement_check(self):
        if self.running_sc:
            return

        file_path = self.settlement_check_file.get().strip()
        if not file_path or not os.path.exists(file_path):
            self._log_error(f"流水文件不存在: {file_path}")
            return

        self.running_sc = True
        self.run_sc_btn.configure(text="⏳  核查中...", state="disabled")

        thread = threading.Thread(target=self._settlement_check_worker, daemon=True)
        thread.start()

    def _settlement_check_worker(self):
        try:
            self._run_settlement_check_internal()
        except Exception as e:
            self._log_error(f"清算核查出错: {e}")
            import traceback
            self._log_error(traceback.format_exc())
        finally:
            self.running_sc = False
            self.after(0, lambda: self.run_sc_btn.configure(
                text="▶  清算核查", state="normal"))

    def _run_settlement_check_internal(self):
        file_path = self.settlement_check_file.get().strip()
        confirm_path = self.settlement_confirm_file.get().strip()
        if confirm_path and not os.path.exists(confirm_path):
            self._log_error(f"二次确认CSV不存在: {confirm_path}")
            return

        account_name = self.settlement_check_account.get().strip() or "待报解预算收入"
        try:
            threshold = int(self.settlement_check_days.get())
        except ValueError:
            threshold = 2

        out_dir = self.output_dir.get()
        os.makedirs(out_dir, exist_ok=True)

        self._log_step("═══ 清算核查：仅输出可疑记录 ═══")
        self._log_result(f"  流水文件: {file_path}")
        self._log_result(f"  确认CSV: {confirm_path or '未启用'}")
        self._log_result(f"  划转账户: {account_name}")
        self._log_result(f"  阈值天数: {threshold}")

        results = detect_settlement_verification(
            file_path=file_path,
            designated_account=_parse_multi_account(account_name),
            days_threshold=threshold,
            confirm_file_path=confirm_path or None,
        )

        output_path = os.path.join(out_dir, "清算核查结果.xlsx")

        if len(results) == 0:
            self._log_success("未发现可疑记录（所有来账后均有划转）")
            # 仍然保存空结果文件
            pd.DataFrame(columns=['文件来源', '账号', '来源日期', '来源金额', '摘要',
                                  '对方账号', '来账对方户名', '窗口截止', '窗口工作日', '状态']
                        ).to_excel(output_path, index=False)
        else:
            results.to_excel(output_path, index=False)
            self._log_result(f"  可疑记录数: {len(results)}")
            for _, row in results.iterrows():
                self._log_result(
                    f"    {row['来源日期']} | 金额:{row['来源金额']} | "
                    f"来账方:{row['来账对方户名']} | 窗口截止:{row['窗口截止']}"
                )

        self._log_success(f"清算核查结果已保存: {output_path}")
        self._log_step("═══════════════════════════════")
        self._log_success("清算核查完成！")

    # ──────────── 文书生成 ────────────

    def _run_report_from_file(self):
        path = self.results_path.get()
        if not path or not os.path.exists(path):
            self._log_error("结果文件不存在: " + str(path))
            return

        from generate_report import batch_generate_from_file

        self._log_step("═══ 从结果文件生成文书 ═══")
        out_dir = self.output_dir.get()
        mode = self.report_mode.get()
        template = self.template_path.get() or None

        try:
            filepaths = batch_generate_from_file(
                results_path=path,
                mode=mode,
                output_dir=out_dir,
                template_path=template,
                llm_config=self.llm_config,
            )
            print(f"\n共生成 {len(filepaths)} 份文书:")
            for fp in filepaths:
                print(f"  {fp}")
            self._log_success("文书生成完成")
        except Exception as e:
            self._log_error("生成失败: " + str(e))
            import traceback
            self._log_error(traceback.format_exc())


if __name__ == "__main__":
    app = App()
    app.mainloop()
