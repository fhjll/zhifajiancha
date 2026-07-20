# -*- coding: utf-8 -*-
"""
运行日志模块 — 将所有 print() 输出同时写入文件，并捕获未处理异常。

用法（CLI 模式）：
    from logger import setup_logging, TeeWriter
    setup_logging()
    sys.stdout = TeeWriter(sys.stdout)   # print() 自动双向输出
    sys.stderr = TeeWriter(sys.stderr)   # 错误信息同样落盘

用法（GUI 模式）：
    from logger import setup_logging, log_message
    setup_logging()
    # 在 TextRedirector.write() 中额外调用 log_message(text)

日志文件位置：logs/执法检测系统_YYYY-MM-DD_HHMMSS.log
"""

import os
import sys
import threading
from datetime import datetime

# ── 全局状态 ──

_LOG_FILE = None          # 当前日志文件路径
_LOG_LOCK = threading.Lock()
_ORIGINAL_EXCEPTHOOK = sys.excepthook  # 保存原始异常钩子


# ── 公开函数 ──

def setup_logging(log_dir="logs"):
    """
    初始化日志系统。

    1. 创建 logs/ 目录（如不存在）
    2. 按启动时间生成日志文件名
    3. 写入文件头（启动时间、分隔线）
    4. 设置 sys.excepthook 捕获未处理异常

    参数
    ----------
    log_dir : str
        日志文件存放目录（默认 logs/）
    """
    global _LOG_FILE
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    _LOG_FILE = os.path.join(log_dir, f"执法检测系统_{timestamp}.log")

    # 写入文件头
    with open(_LOG_FILE, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write(f"  执法检测系统 — 运行日志\n")
        f.write(f"  启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 60 + "\n")

    # 设置未捕获异常钩子
    def _exception_hook(exc_type, exc_value, exc_traceback):
        import traceback
        log_message(f"[严重错误] 未捕获的异常: {exc_type.__name__}: {exc_value}")
        tb_text = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        log_message(tb_text)
        # 调用原始钩子（确保终端也显示）
        _ORIGINAL_EXCEPTHOOK(exc_type, exc_value, exc_traceback)

    sys.excepthook = _exception_hook


def get_log_path():
    """返回当前日志文件路径，未初始化时返回 None。"""
    return _LOG_FILE


def log_message(message):
    """
    向日志文件追加一行（带时间戳）。线程安全。

    参数
    ----------
    message : str
        要写入的消息（不含换行符，函数自动追加 \\n）
    """
    if not _LOG_FILE:
        return
    timestamp = datetime.now().strftime("%H:%M:%S")
    with _LOG_LOCK:
        try:
            with open(_LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] {message}\n")
        except Exception:
            pass  # 写入失败不阻塞程序


# ── TeeWriter（CLI 模式用） ──

class TeeWriter:
    """
    同时向原始 stdout/stderr 和日志文件写入的 writer。

    用法：
        sys.stdout = TeeWriter(sys.stdout)
        print("hello")  # 同时显示在终端和写入日志文件
    """

    def __init__(self, original_stream):
        """
        参数
        ----------
        original_stream : IO
            原始输出流（如 sys.stdout 或 sys.stderr）
        """
        self._original = original_stream

    def write(self, text):
        # 写入原始流（控制台 / GUI）
        self._original.write(text)
        # 同时写入日志文件
        if text.strip():
            log_message(text.rstrip('\n'))

    def flush(self):
        self._original.flush()

    def isatty(self):
        return self._original.isatty()

    @property
    def encoding(self):
        return self._original.encoding
