"""DataPolish · 数据清洗工坊

一个轻量级的 Streamlit 应用：
1. 上传 CSV / Excel 并快速了解数据质量；
2. 按列交互式处理缺失值；
3. 自动生成单变量分布，并提供散点、KDE、折线、堆叠柱、分组箱线、
   小提琴与相关性热力图等多维分析。

运行方式：
    streamlit run app.py
"""

from __future__ import annotations

import csv
import hashlib
import html
import io
import json
import math
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots


# ---------------------------------------------------------------------------
# 常量与样式
# ---------------------------------------------------------------------------

APP_TITLE = "DataPolish · 数据清洗工坊"
SUPPORTED_TYPES = ["csv", "xlsx", "xls"]
MAX_PLOT_POINTS = 20_000
MAX_SCATTER_POINTS = 5_000
MAX_KDE_POINTS = 5_000
MAX_LINE_POINTS = 20_000
MISSING_LABEL = "（缺失）"

STRATEGY_LABELS: Dict[str, str] = {
    "keep": "保持不变",
    "drop": "删除缺失行",
    "mean": "均值填充",
    "median": "中位数填充",
    "mode": "众数填充",
    "constant": "自定义值填充",
}
NUMERIC_STRATEGIES = ["keep", "drop", "mean", "median", "mode", "constant"]
OTHER_STRATEGIES = ["keep", "drop", "mode", "constant"]

COLORS = [
    "#4F6BED",
    "#22A699",
    "#F2A93B",
    "#E5678B",
    "#7C6BE8",
    "#3BA7D8",
    "#88B04B",
    "#E07A5F",
    "#6C7A89",
    "#B58B5A",
    "#8FB8E8",
    "#D7A1C4",
    "#6BBF9A",
    "#E5B36A",
    "#9A8FE0",
    "#4CA3A6",
    "#C97C8D",
    "#7B9E5A",
    "#B28C6C",
    "#8393A6",
]
PLOT_CONFIG = {
    "displayModeBar": False,
    "responsive": True,
}
PLOTLY_FONT = "Microsoft YaHei, PingFang SC, Noto Sans CJK SC, Arial, sans-serif"

CUSTOM_CSS = """
<style>
    :root {
        --ink: #0F172A;
        --muted: #64748B;
        --line: #E6ECF5;
        --primary: #4F6BED;
        --surface: #FFFFFF;
        --canvas: #F4F7FC;
    }

    html, body, [class*="css"] {
        font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC",
                     "Segoe UI", Arial, sans-serif;
    }

    .stApp {
        background:
            radial-gradient(circle at 7% 0%, rgba(79, 107, 237, 0.08), transparent 30rem),
            radial-gradient(circle at 95% 12%, rgba(34, 166, 153, 0.07), transparent 28rem),
            var(--canvas);
    }

    .block-container {
        max-width: 1500px;
        padding-top: 1.7rem;
        padding-bottom: 3rem;
    }

    /* 顶部品牌区 */
    .app-hero {
        position: relative;
        overflow: hidden;
        border-radius: 22px;
        padding: 2.1rem 2.4rem;
        margin-bottom: 1.5rem;
        color: #FFFFFF;
        background:
            radial-gradient(circle at 86% 16%, rgba(88, 212, 196, 0.26), transparent 18rem),
            linear-gradient(120deg, #111C39 0%, #1D2C5C 52%, #3455B8 100%);
        box-shadow: 0 18px 45px rgba(26, 43, 91, 0.18);
    }

    .app-hero::after {
        content: "";
        position: absolute;
        width: 18rem;
        height: 18rem;
        border: 1px solid rgba(255, 255, 255, 0.13);
        border-radius: 50%;
        right: -4rem;
        bottom: -9rem;
    }

    .hero-kicker {
        display: inline-flex;
        align-items: center;
        gap: 0.45rem;
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.16em;
        color: #9FE3DB;
        text-transform: uppercase;
        margin-bottom: 0.7rem;
    }

    .hero-kicker::before {
        content: "";
        width: 1.6rem;
        height: 2px;
        background: #5ED6C5;
        border-radius: 2px;
    }

    .app-hero h1 {
        margin: 0;
        color: #FFFFFF;
        font-size: clamp(1.9rem, 3.1vw, 2.8rem);
        line-height: 1.2;
        letter-spacing: -0.035em;
    }

    .app-hero h1 span {
        display: inline-block;
        margin-left: 0.18em;
        color: #B8C7FF;
        font-weight: 500;
    }

    .app-hero p {
        margin: 0.8rem 0 0;
        max-width: 52rem;
        color: #D6E1FF;
        font-size: 0.98rem;
        line-height: 1.75;
    }

    .hero-tags {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
        margin-top: 1.2rem;
    }

    .hero-tags span {
        padding: 0.25rem 0.6rem;
        border: 1px solid rgba(255, 255, 255, 0.19);
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.08);
        color: #EEF3FF;
        font-size: 0.7rem;
        letter-spacing: 0.04em;
    }

    /* 侧边栏 */
    [data-testid="stSidebar"] {
        background: #FFFFFF;
        border-right: 1px solid var(--line);
    }

    [data-testid="stSidebar"] .block-container {
        padding-top: 1.4rem;
    }

    .side-brand {
        display: flex;
        align-items: center;
        gap: 0.72rem;
        padding: 0.2rem 0 0.9rem;
    }

    .side-brand-mark {
        display: grid;
        place-items: center;
        width: 2.5rem;
        height: 2.5rem;
        border-radius: 13px;
        background: linear-gradient(145deg, #4F6BED, #2E49B4);
        color: #FFFFFF;
        font-size: 1.2rem;
        box-shadow: 0 8px 18px rgba(79, 107, 237, 0.28);
    }

    .side-brand strong {
        display: block;
        color: var(--ink);
        font-size: 1rem;
        letter-spacing: -0.01em;
    }

    .side-brand small {
        color: var(--muted);
        font-size: 0.7rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    .file-card {
        margin: 0.55rem 0 0.8rem;
        padding: 0.9rem 1rem;
        border: 1px solid #E4EAF4;
        border-radius: 14px;
        background: linear-gradient(160deg, #F8FAFF, #FFFFFF);
    }

    .file-card-label {
        color: var(--muted);
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    .file-card-name {
        margin: 0.28rem 0 0.32rem;
        color: var(--ink);
        font-size: 0.88rem;
        font-weight: 700;
        word-break: break-all;
        line-height: 1.45;
    }

    .file-card-meta {
        color: var(--muted);
        font-size: 0.75rem;
    }

    .status-pill {
        display: inline-block;
        margin-top: 0.6rem;
        padding: 0.18rem 0.55rem;
        border-radius: 999px;
        font-size: 0.68rem;
        font-weight: 700;
    }

    .status-raw {
        color: #8A5A12;
        background: #FFF5DD;
        border: 1px solid #F6E1B3;
    }

    .status-clean {
        color: #0B756D;
        background: #E7FAF6;
        border: 1px solid #BEEFE7;
    }

    /* 章节标题 */
    .section-head {
        display: flex;
        align-items: flex-start;
        gap: 0.85rem;
        margin: 0.5rem 0 1rem;
    }

    .section-num {
        display: grid;
        place-items: center;
        flex: 0 0 2.15rem;
        width: 2.15rem;
        height: 2.15rem;
        border-radius: 11px;
        color: #FFFFFF;
        background: linear-gradient(145deg, #4F6BED, #334CB4);
        font-size: 0.75rem;
        font-weight: 800;
        letter-spacing: 0.04em;
        box-shadow: 0 8px 18px rgba(79, 107, 237, 0.22);
    }

    .section-head h2 {
        margin: 0;
        color: var(--ink);
        font-size: 1.28rem;
        line-height: 1.35;
        letter-spacing: -0.02em;
    }

    .section-head p {
        margin: 0.22rem 0 0;
        color: var(--muted);
        font-size: 0.84rem;
        line-height: 1.55;
    }

    /* 指标与容器 */
    [data-testid="stMetric"] {
        min-height: 7.2rem;
        padding: 1rem 1.15rem;
        border: 1px solid #E6ECF5;
        border-radius: 16px;
        background: #FFFFFF;
        box-shadow: 0 8px 22px rgba(24, 43, 89, 0.045);
    }

    [data-testid="stMetricLabel"] {
        color: #64748B;
    }

    [data-testid="stMetricValue"] {
        color: #12203E;
        letter-spacing: -0.025em;
    }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-color: #E6ECF5 !important;
    }

    /* 输入控件与按钮 */
    .stButton > button, .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 600;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #4F6BED, #3B55CF);
        border: none;
        box-shadow: 0 7px 16px rgba(79, 107, 237, 0.22);
    }

    .stButton > button[kind="primary"]:hover {
        transform: translateY(-1px);
        box-shadow: 0 9px 20px rgba(79, 107, 237, 0.3);
    }

    [data-testid="stFileUploaderDropzone"] {
        border-radius: 14px;
        border: 1px dashed #B9C7E4;
        background: #F8FAFF;
    }

    /* 空状态与功能卡 */
    .empty-panel {
        padding: 2.3rem 1.8rem;
        border: 1px solid #E5EBF5;
        border-radius: 20px;
        background: #FFFFFF;
        text-align: center;
        box-shadow: 0 12px 30px rgba(24, 43, 89, 0.05);
    }

    .empty-icon {
        display: grid;
        place-items: center;
        width: 3.5rem;
        height: 3.5rem;
        margin: 0 auto 0.8rem;
        border-radius: 18px;
        background: #EEF2FF;
        font-size: 1.55rem;
    }

    .empty-panel h2 {
        margin: 0;
        color: var(--ink);
        font-size: 1.35rem;
        letter-spacing: -0.02em;
    }

    .empty-panel p {
        max-width: 34rem;
        margin: 0.55rem auto 0;
        color: var(--muted);
        font-size: 0.9rem;
        line-height: 1.7;
    }

    .feature-card {
        height: 100%;
        padding: 1.2rem 1.25rem;
        border: 1px solid #E6ECF5;
        border-radius: 16px;
        background: #FFFFFF;
        box-shadow: 0 7px 20px rgba(24, 43, 89, 0.04);
    }

    .feature-card .feature-index {
        color: #4F6BED;
        font-size: 0.7rem;
        font-weight: 800;
        letter-spacing: 0.12em;
    }

    .feature-card h3 {
        margin: 0.4rem 0 0.35rem;
        color: var(--ink);
        font-size: 1rem;
    }

    .feature-card p {
        margin: 0;
        color: var(--muted);
        font-size: 0.8rem;
        line-height: 1.65;
    }

    /* 可视化分析标签页 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.35rem;
        padding: 0.35rem;
        border: 1px solid #E5EBF5;
        border-radius: 13px;
        background: #FFFFFF;
        box-shadow: 0 5px 16px rgba(24, 43, 89, 0.035);
    }

    .stTabs [data-baseweb="tab"] {
        height: 2.4rem;
        padding: 0 1rem;
        border-radius: 9px;
        color: #61708A;
        font-weight: 650;
    }

    .stTabs [aria-selected="true"] {
        color: #3048B5 !important;
        background: #EEF2FF;
    }

    .stTabs [data-baseweb="tab-highlight"],
    .stTabs [data-baseweb="tab-border"] {
        display: none;
    }

    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 1rem;
    }

    /* 图表卡片 */
    .chart-title {
        margin: 0;
        color: #14213D;
        font-size: 0.96rem;
        font-weight: 750;
        line-height: 1.4;
        word-break: break-word;
    }

    .chart-kicker {
        display: inline-block;
        margin-bottom: 0.2rem;
        color: #7C8CA8;
        font-size: 0.66rem;
        font-weight: 800;
        letter-spacing: 0.1em;
    }

    .app-footer {
        margin-top: 2.2rem;
        padding-top: 1.1rem;
        border-top: 1px solid #E7ECF4;
        color: #8A97AC;
        font-size: 0.73rem;
        text-align: center;
        line-height: 1.7;
    }

    @media (max-width: 760px) {
        .app-hero { padding: 1.5rem 1.25rem; border-radius: 18px; }
        .hero-tags { display: none; }
        .block-container { padding-top: 1rem; }
    }
</style>
"""


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------


class DataLoadError(Exception):
    """文件无法被解析时抛出的可读异常。"""


@dataclass
class CleaningResult:
    """一次清洗操作的完整结果，便于在界面中展示清洗报告。"""

    data: pd.DataFrame
    rows_before: int
    rows_after: int
    missing_before: int
    missing_after: int
    actions: List[Dict[str, str]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------


def _file_size_label(size: int) -> str:
    """把字节数转换为适合展示的大小。"""
    if size >= 1024 * 1024:
        return f"{size / 1024 / 1024:.1f} MB"
    if size >= 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size} B"


def _format_int(value: Any) -> str:
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return "—"


def _format_stat(value: Any) -> str:
    """数值统计量的简短格式化，兼顾大数、小数和空值。"""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "—"
    if not math.isfinite(number):
        return "—"
    if number == 0:
        return "0"
    if abs(number) >= 1_000_000:
        return f"{number:,.3g}"
    if abs(number) < 0.01:
        return f"{number:.3g}"
    return f"{number:,.2f}"


def _widget_key(token: str, *parts: Any) -> str:
    """生成稳定且不会过长的组件 key。"""
    raw = "|".join(str(part) for part in parts)
    digest = hashlib.md5(raw.encode("utf-8")).hexdigest()[:10]
    return f"w_{token}_{digest}"


def _safe_download_name(file_name: str) -> str:
    """清理下载文件名中的非法字符，同时保留中文。"""
    stem = Path(file_name or "dataset").stem or "dataset"
    safe = re.sub(r'[\\/:*?"<>|]+', "_", stem).strip(" .")
    return f"{safe or 'dataset'}_cleaned.csv"


def _sample_path() -> Path:
    return Path(__file__).resolve().parent / "sample_data" / "dirty_customer_data.csv"


def _friendly_dtype(series: pd.Series) -> str:
    """把 Pandas dtype 转成用户能理解的中文类型。"""
    dtype = series.dtype
    if pd.api.types.is_bool_dtype(dtype):
        return "布尔型"
    if pd.api.types.is_numeric_dtype(dtype):
        return "数值型"
    if pd.api.types.is_datetime64_any_dtype(dtype):
        return "日期时间"
    if isinstance(dtype, pd.CategoricalDtype):
        return "分类型"
    if pd.api.types.is_string_dtype(dtype) or dtype == object:
        return "文本/分类型"
    return "其他"


def _is_numeric_column(series: pd.Series) -> bool:
    """布尔值不参与均值/中位数填充，也不绘制数值图。"""
    return bool(
        pd.api.types.is_numeric_dtype(series)
        and not pd.api.types.is_bool_dtype(series)
    )


def _missing_counts(df: pd.DataFrame) -> pd.Series:
    return df.isna().sum()


def _preview_styler(df: pd.DataFrame, rows: int = 5):
    """仅用于展示的表格样式：统一数字千分位和小数位，不改动原始数据。"""
    return df.head(rows).style.format(
        na_rep="—",
        precision=2,
        thousands=",",
        escape="html",
    )


# ---------------------------------------------------------------------------
# 文件读取
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner=False, max_entries=8)
def load_data(file_name: str, payload: bytes) -> pd.DataFrame:
    """读取 CSV / XLSX / XLS 的第一个工作表。

    CSV 会依次尝试常见中文编码，并通过 Python 引擎自动识别分隔符；
    错误统一转换为 DataLoadError，方便界面友好展示。
    """
    suffix = Path(file_name).suffix.lower()
    if suffix not in {".csv", ".xlsx", ".xls"}:
        raise DataLoadError("仅支持 .csv、.xlsx 和 .xls 文件。")

    if suffix == ".csv":
        encodings = ("utf-8-sig", "utf-8", "gb18030", "big5")
        decode_errors: List[str] = []
        parse_errors: List[str] = []

        def read_csv_frame(encoding: str) -> pd.DataFrame:
            """优先自动嗅探分隔符；单列 CSV 则回退到逗号解析。"""
            try:
                return pd.read_csv(
                    io.BytesIO(payload),
                    encoding=encoding,
                    sep=None,
                    engine="python",
                )
            except (pd.errors.ParserError, TypeError, csv.Error) as auto_error:
                try:
                    return pd.read_csv(
                        io.BytesIO(payload),
                        encoding=encoding,
                        sep=",",
                        engine="python",
                    )
                except Exception:
                    raise auto_error

        for encoding in encodings:
            try:
                frame = read_csv_frame(encoding)
                if frame.shape[1] == 0:
                    raise DataLoadError("CSV 中没有可用的数据列。")
                return frame
            except UnicodeDecodeError:
                decode_errors.append(f"{encoding}: 编码不匹配")
            except pd.errors.EmptyDataError as exc:
                raise DataLoadError("CSV 文件为空，未找到任何数据。") from exc
            except pd.errors.ParserError as exc:
                parse_errors.append(f"{encoding}: {exc}")
            except DataLoadError:
                raise
            except Exception as exc:  # 保留底层信息但不在界面暴露 traceback
                parse_errors.append(f"{encoding}: {exc}")

        detail = (parse_errors or decode_errors or ["未知解析错误"])[-1]
        raise DataLoadError(
            "无法解析这个 CSV 文件。请检查分隔符、引号和文件是否损坏。"
            f"（{detail}）"
        )

    try:
        if suffix == ".xlsx":
            return pd.read_excel(
                io.BytesIO(payload),
                sheet_name=0,
                engine="openpyxl",
            )
        return pd.read_excel(
            io.BytesIO(payload),
            sheet_name=0,
            engine="xlrd",
        )
    except ImportError as exc:
        package = "openpyxl" if suffix == ".xlsx" else "xlrd"
        raise DataLoadError(
            f"读取 {suffix} 文件需要安装 {package}。"
            f"请运行：pip install {package}"
        ) from exc
    except Exception as exc:
        raise DataLoadError(
            "Excel 文件读取失败。请确认文件未损坏，且第一个工作表可正常打开。"
            f"（{exc}）"
        ) from exc


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """统一列名为字符串，并处理重复列名，避免组件 key 冲突。"""
    frame = df.copy()
    seen: Dict[str, int] = {}
    columns: List[str] = []
    for index, column in enumerate(frame.columns):
        base = str(column).strip() or f"未命名列_{index + 1}"
        count = seen.get(base, 0)
        name = base if count == 0 else f"{base}_{count + 1}"
        while name in seen:
            count += 1
            name = f"{base}_{count + 1}"
        seen[name] = 1
        columns.append(name)
    frame.columns = columns
    return frame


def build_column_profile(df: pd.DataFrame) -> pd.DataFrame:
    """各字段的基础体检信息。"""
    total_rows = max(len(df), 1)
    records: List[Dict[str, Any]] = []
    for column in df.columns:
        series = df[column]
        missing = int(series.isna().sum())
        try:
            unique = int(series.nunique(dropna=True))
        except TypeError:
            unique = int(series.astype(str).nunique(dropna=True))
        records.append(
            {
                "字段": column,
                "数据类型": _friendly_dtype(series),
                "Pandas dtype": str(series.dtype),
                "非空值": int(series.notna().sum()),
                "缺失数": missing,
                "缺失率": missing / total_rows,
                "唯一值": unique,
            }
        )
    return pd.DataFrame(records)


# ---------------------------------------------------------------------------
# 缺失值清洗
# ---------------------------------------------------------------------------


def _parse_constant(series: pd.Series, raw_value: str) -> Any:
    """按目标列类型解析自定义填充值，类型不匹配时给出可读错误。"""
    text = str(raw_value).strip()
    if not text:
        raise ValueError("自定义填充值不能为空")

    if pd.api.types.is_bool_dtype(series.dtype):
        normalized = text.lower()
        truthy = {"true", "1", "yes", "y", "是", "有", "开启"}
        falsy = {"false", "0", "no", "n", "否", "无", "关闭"}
        if normalized in truthy:
            return True
        if normalized in falsy:
            return False
        raise ValueError("布尔列请填写 是/否 或 true/false")

    if pd.api.types.is_integer_dtype(series.dtype):
        try:
            numeric = float(text.replace(",", ""))
        except ValueError as exc:
            raise ValueError("该列是整数型，请填写整数") from exc
        if not math.isfinite(numeric) or not numeric.is_integer():
            raise ValueError("该列是整数型，请填写整数")
        return int(numeric)

    if pd.api.types.is_numeric_dtype(series.dtype):
        try:
            numeric = float(text.replace(",", ""))
        except ValueError as exc:
            raise ValueError("该列是数值型，请填写有效数字") from exc
        if not math.isfinite(numeric):
            raise ValueError("请填写有限数值，不能使用 NaN 或无穷大")
        return numeric

    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        try:
            timestamp = pd.Timestamp(text)
        except Exception as exc:
            raise ValueError("日期格式无法识别，例如：2025-01-01") from exc
        timezone = getattr(series.dt, "tz", None)
        if timezone is not None:
            if timestamp.tzinfo is None:
                timestamp = timestamp.tz_localize(timezone)
            else:
                timestamp = timestamp.tz_convert(timezone)
        return timestamp

    return text


def _constant_placeholder(series: pd.Series) -> str:
    if pd.api.types.is_bool_dtype(series.dtype):
        return "例如：是 / 否"
    if pd.api.types.is_integer_dtype(series.dtype):
        return "例如：0"
    if pd.api.types.is_numeric_dtype(series.dtype):
        return "例如：0 或 3.5"
    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        return "例如：2025-01-01"
    return "例如：未知"


def clean_data(
    df: pd.DataFrame,
    configs: Dict[str, Dict[str, Optional[str]]],
) -> CleaningResult:
    """根据按列策略清洗 DataFrame。

    删除行会先统一执行：任意一个被设置为 ``drop`` 的字段缺失时，该行即删除。
    其余字段随后按列填充，因此填充不会影响已经被删除行的判定。
    """
    working = df.copy(deep=True)
    rows_before = int(len(working))
    missing_before = int(working.isna().sum().sum())
    actions: List[Dict[str, str]] = []
    warnings: List[str] = []

    # 1) 统一处理“删除缺失行”，保证多列选择时语义一致。
    drop_columns = [
        column
        for column, config in configs.items()
        if config.get("strategy") == "drop" and column in working.columns
    ]
    if drop_columns:
        before_drop = len(working)
        try:
            working = working.dropna(subset=drop_columns, how="any").copy()
            removed = before_drop - len(working)
            actions.append(
                {
                    "字段": "、".join(str(column) for column in drop_columns),
                    "处理方式": "删除缺失行",
                    "说明": f"联合删除 {removed:,} 行（任一选中列缺失即删除）",
                }
            )
        except Exception as exc:
            warnings.append(f"删除缺失行失败：{exc}")

    # 2) 逐列执行填充。
    for column, config in configs.items():
        if column not in working.columns:
            warnings.append(f"字段「{column}」已不存在，跳过。")
            continue

        strategy = str(config.get("strategy") or "keep")
        if strategy in {"keep", "drop"}:
            continue

        series = working[column]
        missing_now = int(series.isna().sum())
        if missing_now == 0:
            continue

        if strategy in {"mean", "median"}:
            if not _is_numeric_column(series):
                warnings.append(
                    f"字段「{column}」不是数值型，{STRATEGY_LABELS[strategy]}已跳过。"
                )
                continue
            try:
                fill_value = series.mean() if strategy == "mean" else series.median()
                if pd.isna(fill_value) or not math.isfinite(float(fill_value)):
                    warnings.append(
                        f"字段「{column}」没有可用于计算的有效数值，"
                        f"{STRATEGY_LABELS[strategy]}已跳过。"
                    )
                    continue
                working[column] = series.fillna(fill_value)
                actions.append(
                    {
                        "字段": str(column),
                        "处理方式": STRATEGY_LABELS[strategy],
                        "说明": f"用 {_format_stat(fill_value)} 填充 {missing_now:,} 个缺失值",
                    }
                )
            except Exception as exc:
                warnings.append(
                    f"字段「{column}」{STRATEGY_LABELS[strategy]}失败：{exc}"
                )
            continue

        if strategy == "mode":
            try:
                modes = series.mode(dropna=True)
                if modes.empty:
                    warnings.append(
                        f"字段「{column}」没有可计算的众数，众数填充已跳过。"
                    )
                    continue
                fill_value = modes.iloc[0]
                working[column] = series.fillna(fill_value)
                actions.append(
                    {
                        "字段": str(column),
                        "处理方式": "众数填充",
                        "说明": (
                            f"用「{fill_value}」填充 {missing_now:,} 个缺失值"
                            "（并列众数取首个）"
                        ),
                    }
                )
            except Exception as exc:
                warnings.append(f"字段「{column}」众数填充失败：{exc}")
            continue

        if strategy == "constant":
            try:
                fill_value = _parse_constant(series, str(config.get("constant") or ""))
                if isinstance(series.dtype, pd.CategoricalDtype):
                    if fill_value not in series.cat.categories:
                        working[column] = series.cat.add_categories([fill_value])
                        series = working[column]
                working[column] = series.fillna(fill_value)
                actions.append(
                    {
                        "字段": str(column),
                        "处理方式": "自定义值填充",
                        "说明": f"用「{fill_value}」填充 {missing_now:,} 个缺失值",
                    }
                )
            except Exception as exc:
                warnings.append(f"字段「{column}」自定义填充失败：{exc}")
            continue

        warnings.append(f"字段「{column}」使用了未知策略「{strategy}」，已跳过。")

    missing_after = int(working.isna().sum().sum())
    return CleaningResult(
        data=working,
        rows_before=rows_before,
        rows_after=int(len(working)),
        missing_before=missing_before,
        missing_after=missing_after,
        actions=actions,
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# 可视化数据准备
# ---------------------------------------------------------------------------


def _numeric_values(series: pd.Series) -> Tuple[pd.Series, int, int]:
    """返回有限数值、原始缺失数、无穷值数。"""
    missing = int(series.isna().sum())
    numeric = pd.to_numeric(series, errors="coerce").astype("float64")
    infinity = int(np.isinf(numeric).sum())
    finite = numeric[np.isfinite(numeric)]
    return finite, missing, infinity


def _category_counts(
    series: pd.Series, top_n: int
) -> Optional[Tuple[pd.DataFrame, int, int, bool, int]]:
    """计算分类频数，并合并长尾类别、补充缺失类别。

    返回 (绘图数据, 总记录数, 缺失数, 是否发生合并, 非缺失类别数)。
    """
    missing = int(series.isna().sum())
    raw = series.dropna()
    if raw.empty and missing == 0:
        return None

    try:
        text = raw.astype(str).str.strip()
    except Exception:
        text = raw.astype(str)

    text = text.replace("", "（空字符串）")
    counts = text.value_counts(dropna=False)
    category_total = int(len(counts))
    total = int(counts.sum()) + missing
    merged = False

    if len(counts) > top_n:
        top = counts.head(top_n).copy()
        rest_count = int(counts.iloc[top_n:].sum())
        rest_label = f"其他（{len(counts) - top_n} 类）"
        while rest_label in top.index:
            rest_label += "·"
        top.loc[rest_label] = rest_count
        counts = top
        merged = True

    if missing > 0:
        label = MISSING_LABEL
        while label in counts.index:
            label += "·"
        counts.loc[label] = missing

    if counts.empty:
        return None

    plot_df = counts.rename_axis("类别").reset_index(name="数量")
    plot_df["占比"] = plot_df["数量"] / max(total, 1)
    plot_df["标签"] = plot_df.apply(
        lambda row: f"{int(row['数量']):,}  ({row['占比']:.1%})", axis=1
    )
    plot_df = plot_df.sort_values("数量", ascending=False).reset_index(drop=True)
    return plot_df, total, missing, merged, category_total


def _base_layout(height: int) -> Dict[str, Any]:
    return {
        "height": height,
        "template": "plotly_white",
        "font": {"family": PLOTLY_FONT, "size": 12, "color": "#3B4A63"},
        "margin": {"l": 10, "r": 12, "t": 10, "b": 10},
        "paper_bgcolor": "rgba(0,0,0,0)",
        "plot_bgcolor": "rgba(0,0,0,0)",
        "hoverlabel": {
            "bgcolor": "#101A33",
            "font": {"family": PLOTLY_FONT, "color": "#FFFFFF", "size": 12},
            "bordercolor": "#101A33",
        },
        "showlegend": False,
    }


def build_numeric_figure(
    values: pd.Series,
    column_name: str,
    chart_mode: str,
    mean_value: float,
    median_value: float,
) -> go.Figure:
    """生成直方图、箱线图，或两者组合的交互式图表。"""
    column_label = str(column_name)
    if len(column_label) > 46:
        column_label = column_label[:43] + "..."

    if chart_mode == "both":
        fig = make_subplots(
            rows=2,
            cols=1,
            shared_xaxes=True,
            row_heights=[0.72, 0.28],
            vertical_spacing=0.07,
        )
        fig.add_trace(
            go.Histogram(
                x=values,
                nbinsx=min(44, max(12, int(math.sqrt(len(values)) * 1.4))),
                marker={
                    "color": COLORS[0],
                    "line": {"color": "rgba(255,255,255,0.55)", "width": 1},
                },
                opacity=0.88,
                name="频数",
                hovertemplate="区间：%{x}<br>频数：%{y:,}<extra></extra>",
            ),
            row=1,
            col=1,
        )
        fig.add_trace(
            go.Box(
                x=values,
                name="分布",
                orientation="h",
                boxpoints="outliers",
                jitter=0.35,
                marker={"color": COLORS[1], "size": 5, "opacity": 0.85},
                line={"color": COLORS[1], "width": 1.6},
                fillcolor="rgba(34,166,153,0.16)",
                hovertemplate="值：%{x}<extra></extra>",
            ),
            row=2,
            col=1,
        )
        fig.update_yaxes(
            title_text="频数",
            gridcolor="#EDF1F8",
            zeroline=False,
            row=1,
            col=1,
        )
        fig.update_yaxes(showticklabels=False, row=2, col=1)
        fig.update_xaxes(
            title_text=column_label,
            gridcolor="#EDF1F8",
            zeroline=False,
            row=2,
            col=1,
        )
        height = 410
    elif chart_mode == "hist":
        fig = go.Figure(
            go.Histogram(
                x=values,
                nbinsx=min(50, max(12, int(math.sqrt(len(values)) * 1.6))),
                marker={
                    "color": COLORS[0],
                    "line": {"color": "rgba(255,255,255,0.55)", "width": 1},
                },
                opacity=0.88,
                hovertemplate="区间：%{x}<br>频数：%{y:,}<extra></extra>",
            )
        )
        fig.update_xaxes(title_text=column_label)
        fig.update_yaxes(
            title_text="频数", gridcolor="#EDF1F8", zeroline=False
        )
        height = 350
    else:
        fig = go.Figure(
            go.Box(
                x=values,
                name="分布",
                orientation="h",
                boxpoints="outliers",
                jitter=0.35,
                marker={"color": COLORS[1], "size": 5, "opacity": 0.85},
                line={"color": COLORS[1], "width": 1.6},
                fillcolor="rgba(34,166,153,0.16)",
                hovertemplate="值：%{x}<extra></extra>",
            )
        )
        fig.update_xaxes(title_text=column_label)
        fig.update_yaxes(showticklabels=False)
        height = 300

    layout = _base_layout(height)
    layout["bargap"] = 0.04
    fig.update_layout(**layout)
    fig.update_xaxes(showgrid=False, zeroline=False)
    fig.update_yaxes(gridcolor="#EDF1F8", zeroline=False)

    # 参考线不放进图例，避免在窄卡片中显得拥挤。
    if chart_mode in {"both", "hist"}:
        for value, color, dash in (
            (mean_value, "#F2A93B", "dash"),
            (median_value, "#E5678B", "dot"),
        ):
            if pd.notna(value) and math.isfinite(float(value)):
                fig.add_vline(
                    x=float(value),
                    line={"color": color, "width": 1.3, "dash": dash},
                    row=1 if chart_mode == "both" else None,
                    col=1 if chart_mode == "both" else None,
                )

    return fig


def build_category_bar(plot_df: pd.DataFrame, total: int) -> go.Figure:
    """水平柱状图：更适合展示较长的中文类别名。"""
    ordered = plot_df.sort_values("数量", ascending=False)
    max_count = max(int(ordered["数量"].max()), 1)
    height = max(320, min(600, 95 + 35 * len(ordered)))
    fig = px.bar(
        ordered,
        x="数量",
        y="类别",
        orientation="h",
        color="类别",
        color_discrete_sequence=COLORS,
        text="标签",
        custom_data=["数量", "占比"],
    )
    fig.update_traces(
        textposition="outside",
        cliponaxis=False,
        textfont={"size": 10, "color": "#53627A"},
        hovertemplate=(
            "<b>%{y}</b><br>数量：%{customdata[0]:,}<br>"
            "占比：%{customdata[1]:.1%}<extra></extra>"
        ),
    )
    layout = _base_layout(height)
    layout.update(
        {
            "xaxis_title": "数量",
            "yaxis_title": None,
            "xaxis": {
                "range": [0, max_count * 1.22],
                "gridcolor": "#EDF1F8",
                "zeroline": False,
            },
            "yaxis": {"showgrid": False},
            "margin": {"l": 8, "r": 22, "t": 8, "b": 8},
        }
    )
    fig.update_layout(**layout)
    fig.update_xaxes(title_font={"size": 11})
    return fig


def build_category_pie(plot_df: pd.DataFrame, total: int) -> go.Figure:
    """环形饼图：突出各类别占比。"""
    height = 410
    fig = px.pie(
        plot_df,
        values="数量",
        names="类别",
        hole=0.46,
        color_discrete_sequence=COLORS,
        custom_data=["数量", "占比"],
    )
    fig.update_traces(
        textposition="inside",
        textinfo="percent",
        insidetextorientation="radial",
        hovertemplate=(
            "<b>%{label}</b><br>数量：%{customdata[0]:,}<br>"
            "占比：%{customdata[1]:.1%}<extra></extra>"
        ),
        marker={"line": {"color": "#FFFFFF", "width": 2}},
    )
    layout = _base_layout(height)
    layout.update(
        {
            "showlegend": True,
            "legend": {
                "orientation": "h",
                "yanchor": "top",
                "y": -0.02,
                "xanchor": "center",
                "x": 0.5,
                "font": {"size": 10},
            },
            "margin": {"l": 12, "r": 12, "t": 12, "b": 62},
        }
    )
    fig.update_layout(**layout)
    fig.add_annotation(
        text=f"<b>{total:,}</b><br><span style='font-size:10px;color:#7C8CA8'>总记录</span>",
        showarrow=False,
        font={"size": 16, "color": "#1C2A46"},
    )
    return fig


# ---------------------------------------------------------------------------
# 高级可视化数据准备与图表构建
# ---------------------------------------------------------------------------


def _rgba(hex_color: str, alpha: float) -> str:
    """把十六进制颜色转为带透明度的 rgba 字符串。"""
    value = str(hex_color).lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        return f"rgba(79, 107, 237, {alpha})"
    try:
        red, green, blue = (int(value[index : index + 2], 16) for index in (0, 2, 4))
    except ValueError:
        return f"rgba(79, 107, 237, {alpha})"
    return f"rgba({red}, {green}, {blue}, {alpha})"


def _kde_curve(
    values: pd.Series,
    grid_size: int = 320,
) -> Optional[Tuple[np.ndarray, np.ndarray, float, int]]:
    """使用高斯核与 Silverman 经验带宽计算一维 KDE。

    返回 (网格点, 密度值, 带宽, 实际参与计算的有效样本数)。
    为控制浏览器渲染成本，最多使用 5,000 个有效值。
    """
    numeric = pd.to_numeric(values, errors="coerce").astype("float64")
    array = numeric[np.isfinite(numeric)].to_numpy()
    if len(array) < 3:
        return None

    rng = np.random.default_rng(42)
    if len(array) > MAX_KDE_POINTS:
        array = rng.choice(array, size=MAX_KDE_POINTS, replace=False)

    sample_count = int(len(array))
    std = float(np.std(array, ddof=1))
    if not math.isfinite(std) or std <= 0:
        return None

    q1, q3 = np.percentile(array, [25, 75])
    iqr = float(q3 - q1)
    spread = min(std, iqr / 1.349) if iqr > 0 else std
    if not math.isfinite(spread) or spread <= 0:
        spread = std

    bandwidth = 0.9 * spread * (sample_count ** (-0.2))
    if not math.isfinite(bandwidth) or bandwidth <= 0:
        return None

    grid = np.linspace(
        float(array.min() - 3 * bandwidth),
        float(array.max() + 3 * bandwidth),
        grid_size,
    )
    density = np.zeros_like(grid)
    for start in range(0, sample_count, 1_000):
        chunk = array[start : start + 1_000]
        standardized = (grid[:, None] - chunk[None, :]) / bandwidth
        density += np.exp(-0.5 * standardized**2).sum(axis=1)
    density /= sample_count * bandwidth * math.sqrt(2 * math.pi)
    return grid, density, bandwidth, sample_count


def build_scatter_figure(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    color_col: Optional[str] = None,
    show_trend: bool = True,
) -> Tuple[go.Figure, Dict[str, Any]]:
    """构建双数值散点图，可附带分类着色和线性趋势线。"""
    for column in (x_col, y_col):
        if column not in df.columns or not _is_numeric_column(df[column]):
            raise ValueError(f"字段「{column}」不是数值型字段，无法绘制散点图。")
    if x_col == y_col:
        raise ValueError("X 轴与 Y 轴不能是同一个字段。")
    if color_col and color_col not in df.columns:
        raise ValueError(f"着色字段「{color_col}」不存在。")

    numeric = df[[x_col, y_col]].copy()
    for column in (x_col, y_col):
        numeric[column] = pd.to_numeric(numeric[column], errors="coerce")
    numeric = numeric.replace([np.inf, -np.inf], np.nan).dropna()
    if numeric.empty:
        raise ValueError("所选字段没有同时有效的数值记录。")

    x_values = numeric[x_col].to_numpy(dtype="float64")
    y_values = numeric[y_col].to_numpy(dtype="float64")
    correlation = float("nan")
    if (
        len(numeric) >= 2
        and float(np.std(x_values)) > 0
        and float(np.std(y_values)) > 0
    ):
        correlation = float(np.corrcoef(x_values, y_values)[0, 1])

    plot_frame = numeric.copy()
    if color_col:
        color_values = df.loc[numeric.index, color_col].astype("string").str.strip()
        color_values = color_values.fillna(MISSING_LABEL).replace("", "（空字符串）")
        plot_frame[color_col] = color_values.astype(str)

    total_count = int(len(plot_frame))
    if total_count > MAX_SCATTER_POINTS:
        plot_frame = plot_frame.sample(n=MAX_SCATTER_POINTS, random_state=42)

    kwargs: Dict[str, Any] = {
        "data_frame": plot_frame,
        "x": x_col,
        "y": y_col,
        "opacity": 0.72,
        "color_discrete_sequence": COLORS,
        "render_mode": "svg",
    }
    if color_col:
        kwargs["color"] = color_col
    figure = px.scatter(**kwargs)
    figure.update_traces(
        marker={
            "size": 7,
            "line": {"width": 0.45, "color": "rgba(255,255,255,0.55)"},
        },
        selector={"mode": "markers"},
    )

    trend_slope: Optional[float] = None
    if (
        show_trend
        and len(numeric) >= 3
        and float(np.std(x_values)) > 0
    ):
        coefficients = np.polyfit(x_values, y_values, 1)
        trend_slope = float(coefficients[0])
        line_x = np.linspace(float(x_values.min()), float(x_values.max()), 220)
        line_y = np.polyval(coefficients, line_x)
        figure.add_trace(
            go.Scatter(
                x=line_x,
                y=line_y,
                mode="lines",
                name=f"线性趋势（斜率 {trend_slope:.3g}）",
                line={"color": "#E5678B", "width": 2.2, "dash": "dash"},
                hovertemplate="趋势线<br>x：%{x:.4g}<br>y：%{y:.4g}<extra></extra>",
            )
        )

    layout = _base_layout(515)
    layout.update(
        {
            "showlegend": bool(color_col or show_trend),
            "legend": {
                "orientation": "h",
                "yanchor": "bottom",
                "y": 1.02,
                "xanchor": "left",
                "x": 0,
                "font": {"size": 10},
            },
            "margin": {
                "l": 10,
                "r": 10,
                "t": 52 if (color_col or show_trend) else 10,
                "b": 10,
            },
            "hovermode": "closest",
        }
    )
    figure.update_layout(**layout)
    figure.update_xaxes(
        title_text=x_col,
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    figure.update_yaxes(
        title_text=y_col,
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    return figure, {
        "total": total_count,
        "shown": int(len(plot_frame)),
        "correlation": correlation,
        "slope": trend_slope,
    }


def build_density_figure(
    df: pd.DataFrame,
    columns: Sequence[str],
) -> Tuple[go.Figure, List[str], List[Dict[str, Any]]]:
    """为一组数值字段绘制叠加的核密度估计曲线。"""
    figure = go.Figure()
    skipped: List[str] = []
    profiles: List[Dict[str, Any]] = []

    for index, column in enumerate(columns):
        if column not in df.columns or not _is_numeric_column(df[column]):
            skipped.append(str(column))
            continue
        values, missing, infinity = _numeric_values(df[column])
        kde = _kde_curve(values)
        if kde is None:
            skipped.append(str(column))
            continue
        grid, density, bandwidth, sample_count = kde
        color = COLORS[index % len(COLORS)]
        figure.add_trace(
            go.Scatter(
                x=grid,
                y=density,
                mode="lines",
                name=str(column),
                line={"color": color, "width": 2.5, "shape": "spline"},
                fill="tozeroy",
                fillcolor=_rgba(color, 0.09),
                hovertemplate=(
                    "<b>%{fullData.name}</b><br>"
                    "值：%{x:.4g}<br>密度：%{y:.5f}<extra></extra>"
                ),
            )
        )
        profiles.append(
            {
                "column": str(column),
                "n": sample_count,
                "missing": missing,
                "infinity": infinity,
                "bandwidth": bandwidth,
            }
        )

    if not figure.data:
        raise ValueError("所选字段都没有足够的离散数值用于 KDE 计算。")

    layout = _base_layout(470)
    layout.update(
        {
            "showlegend": True,
            "legend": {
                "orientation": "h",
                "yanchor": "bottom",
                "y": 1.02,
                "xanchor": "left",
                "x": 0,
                "font": {"size": 10},
            },
            "margin": {"l": 10, "r": 10, "t": 58, "b": 12},
            "hovermode": "x unified",
        }
    )
    figure.update_layout(**layout)
    figure.update_xaxes(title_text="数值", gridcolor="#EDF1F8", zeroline=False)
    figure.update_yaxes(
        title_text="概率密度",
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    return figure, skipped, profiles


def build_line_figure(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    use_index: bool,
    sort_by_x: bool,
    show_markers: bool,
) -> Tuple[go.Figure, Dict[str, Any]]:
    """绘制按索引或数值 X 排序的折线趋势图。"""
    if y_col not in df.columns or not _is_numeric_column(df[y_col]):
        raise ValueError(f"Y 轴字段「{y_col}」必须是数值型。")
    if not use_index and (
        x_col not in df.columns or not _is_numeric_column(df[x_col])
    ):
        raise ValueError(f"X 轴字段「{x_col}」必须是数值型，或改用数据索引。")

    if use_index:
        x_series = pd.Series(
            np.arange(1, len(df) + 1),
            index=df.index,
            dtype="float64",
        )
        x_label = "数据索引（行号）"
    else:
        x_series = pd.to_numeric(df[x_col], errors="coerce")
        x_label = str(x_col)
    y_series = pd.to_numeric(df[y_col], errors="coerce")

    working = pd.DataFrame({"__x__": x_series, "__y__": y_series})
    working = working.replace([np.inf, -np.inf], np.nan).dropna()
    if working.empty:
        raise ValueError("所选字段没有可绘制的有效记录。")
    if sort_by_x and not use_index:
        working = working.sort_values("__x__", kind="mergesort")

    total_count = int(len(working))
    sampled = False
    if total_count > MAX_LINE_POINTS:
        positions = np.linspace(0, total_count - 1, MAX_LINE_POINTS).round().astype(int)
        working = working.iloc[positions]
        sampled = True

    figure = go.Figure(
        go.Scatter(
            x=working["__x__"],
            y=working["__y__"],
            mode="lines+markers" if show_markers else "lines",
            line={"color": COLORS[0], "width": 2.3},
            marker={"size": 4.5, "color": COLORS[0], "opacity": 0.72},
            name=str(y_col),
            hovertemplate=(
                f"<b>{html.escape(x_label)}</b>：%{{x}}<br>"
                f"{html.escape(str(y_col))}：%{{y:,.5g}}<extra></extra>"
            ),
        )
    )
    layout = _base_layout(430)
    layout.update(
        {
            "margin": {"l": 10, "r": 12, "t": 18, "b": 12},
            "hovermode": "x unified",
        }
    )
    figure.update_layout(**layout)
    figure.update_xaxes(
        title_text=x_label,
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    figure.update_yaxes(
        title_text=str(y_col),
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    return figure, {
        "total": total_count,
        "shown": int(len(working)),
        "sampled": sampled,
        "x_label": x_label,
    }


def _cap_categories(
    series: pd.Series,
    limit: int,
) -> Tuple[pd.Series, bool, int]:
    """保留最高频的若干类别，把长尾类别合并为“其他”。"""
    text = series.astype("string").str.strip()
    text = text.fillna(MISSING_LABEL).replace("", "（空字符串）")
    counts = text.value_counts(dropna=False)
    if len(counts) <= limit:
        return text, False, 0

    keep = set(counts.head(limit).index)
    other_label = f"其他（{len(counts) - limit} 类）"
    existing = set(counts.index)
    while other_label in existing:
        other_label += "·"
    capped = text.where(text.isin(keep), other_label)
    return capped, True, len(counts) - limit


def _other_label_last(order: Sequence[str]) -> List[str]:
    """把“其他”类别移动到分类顺序的末尾。"""
    items = [str(item) for item in order]
    others = [item for item in items if item.startswith("其他（")]
    return [item for item in items if item not in others] + others


def build_stacked_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    color_col: str,
    normalize: bool,
    top_n: int,
) -> Tuple[go.Figure, Dict[str, Any]]:
    """构建两个分类字段的堆叠柱状图，可切换计数或行百分比。"""
    if x_col == color_col:
        raise ValueError("堆叠柱状图需要两个不同的分类型字段。")
    for column in (x_col, color_col):
        if column not in df.columns or _is_numeric_column(df[column]):
            raise ValueError(f"字段「{column}」不是分类型字段。")

    missing = int(df[[x_col, color_col]].isna().any(axis=1).sum())
    working = df[[x_col, color_col]].dropna()
    if working.empty:
        raise ValueError("所选字段没有可用的组合记录。")

    x_text = working[x_col]
    color_text = working[color_col]
    x_text, x_merged, x_extra = _cap_categories(x_text, top_n)
    color_text, color_merged, color_extra = _cap_categories(color_text, top_n)

    x_counts = x_text.value_counts()
    color_counts = color_text.value_counts()
    x_order = _other_label_last(list(x_counts.index))
    color_order = _other_label_last(list(color_counts.index))

    cross = pd.crosstab(x_text, color_text)
    cross = cross.reindex(index=x_order, columns=color_order, fill_value=0)
    long_frame = pd.DataFrame(
        [
            {
                "__x__": str(x_value),
                "__color__": str(color_value),
                "__value__": int(cross.at[x_value, color_value]),
            }
            for x_value in x_order
            for color_value in color_order
        ]
    )

    if normalize:
        totals = long_frame.groupby("__x__")["__value__"].transform("sum")
        long_frame["__value__"] = (
            long_frame["__value__"] / totals.replace(0, np.nan) * 100
        ).fillna(0)
        value_title = "占比 (%)"
    else:
        value_title = "记录数"

    figure = px.bar(
        long_frame,
        x="__x__",
        y="__value__",
        color="__color__",
        barmode="stack",
        category_orders={"__x__": x_order, "__color__": color_order},
        color_discrete_sequence=COLORS,
        labels={
            "__x__": str(x_col),
            "__color__": str(color_col),
            "__value__": value_title,
        },
    )
    value_format = "%{y:.1f}%" if normalize else "%{y:,.0f}"
    figure.update_traces(
        marker={"line": {"width": 0.6, "color": "rgba(255,255,255,0.8)"}},
        hovertemplate=(
            "<b>%{x}</b><br>"
            f"{html.escape(str(color_col))}：%{{fullData.name}}<br>"
            f"{value_title}：{value_format}<extra></extra>"
        ),
    )
    layout = _base_layout(470)
    layout.update(
        {
            "showlegend": True,
            "legend": {
                "orientation": "h",
                "yanchor": "top",
                "y": -0.2,
                "xanchor": "center",
                "x": 0.5,
                "font": {"size": 10},
            },
            "margin": {"l": 10, "r": 10, "t": 18, "b": 92},
            "hovermode": "closest",
            "bargap": 0.18,
        }
    )
    figure.update_layout(**layout)
    figure.update_xaxes(
        title_text=str(x_col),
        tickangle=-25,
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    figure.update_yaxes(
        title_text=value_title,
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    return figure, {
        "missing": missing,
        "x_merged": x_merged,
        "y_merged": color_merged,
        "x_extra": x_extra,
        "y_extra": color_extra,
    }


def _prepare_group_data(
    df: pd.DataFrame,
    category_col: str,
    numeric_col: str,
    color_col: Optional[str],
    top_n: int,
) -> Tuple[pd.DataFrame, List[str], List[str], Dict[str, Any]]:
    """准备分组箱线图与小提琴图共用的数据。"""
    if category_col not in df.columns or _is_numeric_column(df[category_col]):
        raise ValueError(f"字段「{category_col}」必须是分类型字段。")
    if numeric_col not in df.columns or not _is_numeric_column(df[numeric_col]):
        raise ValueError(f"字段「{numeric_col}」必须是数值型字段。")
    if color_col and color_col == category_col:
        raise ValueError("分组着色字段不能与 X 轴类别字段相同。")
    if color_col and color_col not in df.columns:
        raise ValueError(f"着色字段「{color_col}」不存在。")

    needed = [category_col, numeric_col]
    if color_col:
        needed.append(color_col)
    raw = df[needed].copy()
    raw[numeric_col] = pd.to_numeric(raw[numeric_col], errors="coerce")
    raw = raw.replace([np.inf, -np.inf], np.nan)
    missing_dropped = int(len(df) - len(raw.dropna(subset=needed)))
    working = raw.dropna(subset=needed).copy()
    if working.empty:
        raise ValueError("所选字段没有可用的分组记录。")

    category_text, category_merged, category_extra = _cap_categories(
        working[category_col], top_n
    )
    working[category_col] = category_text.astype(str)
    category_order = _other_label_last(
        list(working[category_col].value_counts().index)
    )

    color_order: List[str] = []
    color_merged = False
    color_extra = 0
    if color_col:
        color_text, color_merged, color_extra = _cap_categories(
            working[color_col], top_n
        )
        working[color_col] = color_text.astype(str)
        color_order = _other_label_last(
            list(working[color_col].value_counts().index)
        )

    return working, category_order, color_order, {
        "missing_dropped": missing_dropped,
        "category_merged": category_merged,
        "category_extra": category_extra,
        "color_merged": color_merged,
        "color_extra": color_extra,
    }


def build_group_box_figure(
    working: pd.DataFrame,
    category_col: str,
    numeric_col: str,
    color_col: Optional[str],
    category_order: Sequence[str],
) -> go.Figure:
    """构建按类别分组的箱线图。"""
    color_argument = color_col if color_col else category_col
    figure = px.box(
        working,
        x=category_col,
        y=numeric_col,
        color=color_argument,
        points="outliers",
        color_discrete_sequence=COLORS,
        category_orders={category_col: list(category_order)},
    )
    figure.update_traces(
        boxmean=True,
        marker={"size": 5, "opacity": 0.8},
        selector={"type": "box"},
    )
    layout = _base_layout(485)
    layout.update(
        {
            "showlegend": bool(color_col),
            "legend": {
                "orientation": "h",
                "yanchor": "top",
                "y": -0.2,
                "xanchor": "center",
                "x": 0.5,
                "font": {"size": 10},
            },
            "margin": {"l": 10, "r": 12, "t": 24, "b": 92 if color_col else 70},
        }
    )
    figure.update_layout(**layout)
    figure.update_xaxes(
        title_text=str(category_col),
        tickangle=-25,
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    figure.update_yaxes(
        title_text=str(numeric_col),
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    return figure


def build_group_violin_figure(
    working: pd.DataFrame,
    category_col: str,
    numeric_col: str,
    color_col: Optional[str],
    category_order: Sequence[str],
) -> go.Figure:
    """构建结合箱线信息与核密度轮廓的小提琴图。"""
    color_argument = color_col if color_col else category_col
    figure = px.violin(
        working,
        x=category_col,
        y=numeric_col,
        color=color_argument,
        box=True,
        points="outliers",
        violinmode="group" if color_col else "overlay",
        color_discrete_sequence=COLORS,
        category_orders={category_col: list(category_order)},
    )
    figure.update_traces(
        meanline_visible=True,
        spanmode="soft",
        selector={"type": "violin"},
    )
    layout = _base_layout(485)
    layout.update(
        {
            "showlegend": bool(color_col),
            "legend": {
                "orientation": "h",
                "yanchor": "top",
                "y": -0.2,
                "xanchor": "center",
                "x": 0.5,
                "font": {"size": 10},
            },
            "margin": {"l": 10, "r": 12, "t": 24, "b": 92 if color_col else 70},
            "violinmode": "group" if color_col else "overlay",
        }
    )
    figure.update_layout(**layout)
    figure.update_xaxes(
        title_text=str(category_col),
        tickangle=-25,
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    figure.update_yaxes(
        title_text=str(numeric_col),
        gridcolor="#EDF1F8",
        zeroline=False,
    )
    return figure


def build_correlation_figure(
    df: pd.DataFrame,
    columns: Sequence[str],
    method: str,
    annotate: bool,
) -> Tuple[go.Figure, pd.DataFrame, List[str]]:
    """构建 Pearson / Spearman 相关系数热力图。"""
    if method not in {"pearson", "spearman"}:
        raise ValueError("相关系数方法只支持 Pearson 或 Spearman。")
    usable: List[str] = []
    dropped: List[str] = []
    numeric_frame = df[list(columns)].apply(pd.to_numeric, errors="coerce")
    numeric_frame = numeric_frame.replace([np.inf, -np.inf], np.nan)
    for column in columns:
        series = numeric_frame[column]
        if series.notna().sum() < 2 or series.nunique(dropna=True) < 2:
            dropped.append(str(column))
        else:
            usable.append(str(column))
    if len(usable) < 2:
        raise ValueError("至少需要两个有效数值字段才能计算相关系数矩阵。")

    correlation = numeric_frame[usable].corr(method=method)
    if correlation.empty:
        raise ValueError("无法计算相关系数矩阵。")
    labels = [str(column) for column in correlation.columns]
    matrix = correlation.to_numpy(dtype="float64")
    text_matrix = [
        ["" if pd.isna(value) else f"{float(value):.2f}" for value in row]
        for row in matrix
    ]
    height = max(440, min(880, 190 + 36 * len(labels)))
    heatmap_kwargs: Dict[str, Any] = {
        "z": matrix,
        "x": labels,
        "y": labels,
        "zmin": -1,
        "zmax": 1,
        "colorscale": [
            [0.0, "#3767A8"],
            [0.5, "#F7F9FC"],
            [1.0, "#D94F64"],
        ],
        "hovertemplate": (
            "<b>%{y}</b> × <b>%{x}</b><br>"
            "相关系数：%{z:.3f}<extra></extra>"
        ),
        "colorbar": {
            "title": "r",
            "thickness": 13,
            "len": 0.72,
            "tickvals": [-1, -0.5, 0, 0.5, 1],
        },
    }
    if annotate:
        heatmap_kwargs["text"] = text_matrix
        heatmap_kwargs["texttemplate"] = "%{text}"
        heatmap_kwargs["textfont"] = {
            "size": 10 if len(labels) <= 14 else 8,
            "color": "#26334A",
        }
    figure = go.Figure(go.Heatmap(**heatmap_kwargs))
    layout = _base_layout(height)
    layout.update(
        {
            "margin": {"l": 12, "r": 18, "t": 95, "b": 18},
            "hovermode": "closest",
        }
    )
    figure.update_layout(**layout)
    figure.update_xaxes(
        side="top",
        tickangle=-35,
        showgrid=False,
        zeroline=False,
    )
    figure.update_yaxes(
        autorange="reversed",
        showgrid=False,
        zeroline=False,
    )
    return figure, correlation, dropped


def _correlation_pairs(correlation: pd.DataFrame, limit: int = 8) -> pd.DataFrame:
    """提取绝对值最大的若干相关字段对。"""
    columns = list(correlation.columns)
    records: List[Dict[str, Any]] = []
    for left_index, left in enumerate(columns):
        for right in columns[left_index + 1 :]:
            value = correlation.loc[left, right]
            if pd.isna(value):
                continue
            absolute = abs(float(value))
            if absolute >= 0.8:
                strength = "强"
            elif absolute >= 0.5:
                strength = "中等"
            elif absolute >= 0.3:
                strength = "弱"
            else:
                strength = "很弱"
            records.append(
                {
                    "字段 A": str(left),
                    "字段 B": str(right),
                    "相关系数": float(value),
                    "方向": "正相关" if float(value) >= 0 else "负相关",
                    "强度": strength,
                }
            )
    if not records:
        return pd.DataFrame()
    frame = pd.DataFrame(records)
    frame["_abs"] = frame["相关系数"].abs()
    frame = frame.sort_values("_abs", ascending=False).drop(columns=["_abs"])
    return frame.head(limit).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Session State
# ---------------------------------------------------------------------------


def init_state() -> None:
    defaults: Dict[str, Any] = {
        "raw_df": None,
        "clean_df": None,
        "clean_meta": None,
        "clean_version": 0,
        "applied_config_fingerprint": None,
        "file_name": "",
        "file_signature": "",
        "file_size": 0,
        "loaded_at": "",
        "nav_page": "overview",
        "uploader_version": 0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if st.session_state.nav_page not in {"overview", "clean", "visuals"}:
        st.session_state.nav_page = "overview"


def _clear_derived_widget_state() -> None:
    """删除属于当前数据集的清洗/可视化组件状态。"""
    for key in list(st.session_state.keys()):
        if str(key).startswith("w_"):
            del st.session_state[key]


def activate_data(
    df: pd.DataFrame,
    file_name: str,
    signature: str,
    file_size: int,
) -> None:
    """把新读入的数据放入 Session State，并重置依赖数据的界面状态。"""
    _clear_derived_widget_state()
    st.session_state.raw_df = df
    st.session_state.clean_df = None
    st.session_state.clean_meta = None
    st.session_state.applied_config_fingerprint = None
    st.session_state.file_name = file_name
    st.session_state.file_signature = signature
    st.session_state.file_size = file_size
    st.session_state.loaded_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    st.session_state.clean_version = int(st.session_state.clean_version) + 1
    st.session_state.nav_page = "overview"


def clear_loaded_data() -> None:
    """侧边栏“清空数据”回调：在下一轮运行前释放当前数据。"""
    _clear_derived_widget_state()
    st.session_state.raw_df = None
    st.session_state.clean_df = None
    st.session_state.clean_meta = None
    st.session_state.applied_config_fingerprint = None
    st.session_state.file_name = ""
    st.session_state.file_signature = ""
    st.session_state.file_size = 0
    st.session_state.loaded_at = ""
    st.session_state.clean_version = int(st.session_state.clean_version) + 1
    st.session_state.nav_page = "overview"
    st.session_state.uploader_version = int(st.session_state.uploader_version) + 1


def reset_cleaning_state() -> None:
    """清洗页“重置”按钮回调：在组件重新创建前清理策略与结果。"""
    _clear_derived_widget_state()
    st.session_state.clean_df = None
    st.session_state.clean_meta = None
    st.session_state.applied_config_fingerprint = None
    st.session_state.clean_version = int(st.session_state.clean_version) + 1


@st.cache_data(show_spinner=False, max_entries=6)
def dataframe_to_csv(df: pd.DataFrame) -> bytes:
    """使用 UTF-8 BOM 导出，保证 Excel 打开中文不乱码。"""
    return df.to_csv(index=False).encode("utf-8-sig")


# ---------------------------------------------------------------------------
# 侧边栏与全局布局
# ---------------------------------------------------------------------------


def inject_css() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def render_hero() -> None:
    st.markdown(
        """
        <section class="app-hero">
            <div class="hero-kicker">Data Quality Studio</div>
            <h1>DataPolish <span>数据清洗工坊</span></h1>
            <p>
                从一份“脏数据”出发，自动体检缺失情况，按字段精细处理，
                再用多维交互图表看清数据分布、分组差异与变量关联。
            </p>
            <div class="hero-tags">
                <span>CSV · XLSX · XLS</span>
                <span>缺失值策略</span>
                <span>Plotly 交互图表</span>
                <span>多维分析</span>
                <span>一键导出</span>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> None:
    with st.sidebar:
        st.markdown(
            """
            <div class="side-brand">
                <div class="side-brand-mark">✦</div>
                <div>
                    <strong>DataPolish</strong>
                    <small>Clean · Explore</small>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        uploaded = st.file_uploader(
            "上传 CSV / Excel",
            type=SUPPORTED_TYPES,
            key=f"uploader_{st.session_state.uploader_version}",
            help="支持 .csv、.xlsx、.xls；Excel 默认读取第一个工作表。",
        )
        if uploaded is not None:
            _handle_uploaded_file(uploaded)

        raw_df = st.session_state.raw_df
        if raw_df is not None:
            clean_meta = st.session_state.clean_meta
            clean_df = st.session_state.clean_df
            status_class = "status-clean" if clean_df is not None else "status-raw"
            status_text = "已应用清洗" if clean_df is not None else "原始数据"
            st.markdown(
                f"""
                <div class="file-card">
                    <div class="file-card-label">当前数据</div>
                    <div class="file-card-name">{html.escape(str(st.session_state.file_name))}</div>
                    <div class="file-card-meta">
                        {len(raw_df):,} 行 × {len(raw_df.columns):,} 列 ·
                        {_file_size_label(int(st.session_state.file_size or 0))}
                    </div>
                    <span class="status-pill {status_class}">{status_text}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if clean_df is not None:
                st.download_button(
                    "下载清洗后的 CSV",
                    data=dataframe_to_csv(clean_df),
                    file_name=_safe_download_name(st.session_state.file_name),
                    mime="text/csv; charset=utf-8",
                    icon="📥",
                    width="stretch",
                    key=f"sidebar_download_{st.session_state.clean_version}",
                )

            st.divider()
            st.radio(
                "工作区导航",
                options=["overview", "clean", "visuals"],
                format_func=lambda value: {
                    "overview": "◈  数据概览",
                    "clean": "✦  缺失值清洗",
                    "visuals": "▦  分析与可视化",
                }[value],
                key="nav_page",
            )

            st.divider()
            st.button(
                "清空当前数据",
                on_click=clear_loaded_data,
                icon="🔄",
                width="stretch",
                key="clear_data_button",
            )
            st.caption("上传新文件会自动替换当前数据。")
        else:
            st.caption("上传文件后解锁清洗与可视化工作区。")

        st.markdown("<div style='height: 1.1rem;'></div>", unsafe_allow_html=True)
        st.caption("本地 Streamlit 进程处理 · 不写入第三方服务")
        st.caption("Pandas + Plotly · 轻量版 v1.1")


def _handle_uploaded_file(uploaded: Any) -> None:
    """解析上传文件；若读取失败则保留原数据并展示错误。"""
    payload = uploaded.getvalue()
    signature = hashlib.sha256(
        uploaded.name.encode("utf-8") + payload
    ).hexdigest()
    if signature == st.session_state.file_signature:
        return

    try:
        with st.spinner("正在解析文件..."):
            frame = load_data(uploaded.name, payload)
            frame = normalize_columns(frame)
            if len(frame.columns) == 0:
                raise DataLoadError("文件中没有可用的数据列。")
        activate_data(frame, uploaded.name, signature, len(payload))
        st.success("文件读取成功，已切换到数据概览。")
    except DataLoadError as exc:
        st.error(f"文件读取失败：{exc}")
    except Exception as exc:
        st.error(f"文件读取失败：{exc}")


def load_sample_data() -> bool:
    """加载内置演示数据，返回是否成功。"""
    path = _sample_path()
    if not path.exists():
        st.error("示例数据缺失，请直接上传自己的 CSV / Excel 文件。")
        return False
    try:
        payload = path.read_bytes()
        frame = load_data(path.name, payload)
        frame = normalize_columns(frame)
        signature = hashlib.sha256(
            path.name.encode("utf-8") + payload
        ).hexdigest()
        activate_data(frame, path.name, signature, len(payload))
        return True
    except Exception as exc:
        st.error(f"示例数据加载失败：{exc}")
        return False


# ---------------------------------------------------------------------------
# 页面：欢迎页
# ---------------------------------------------------------------------------


def render_welcome() -> None:
    st.markdown(
        """
        <div class="empty-panel">
            <div class="empty-icon">📂</div>
            <h2>把数据拖进左侧，开始一次干净的分析</h2>
            <p>
                上传 CSV 或 Excel 文件。应用会自动完成字段体检、缺失值统计，
                并针对每一个缺失字段提供可定制的清洗策略。
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")
    columns = st.columns(3, gap="large")
    cards = [
        (
            "01",
            "自动体检",
            "一次看清行列规模、缺失比例、字段类型与重复行，避免在脏数据上直接下结论。",
        ),
        (
            "02",
            "精细清洗",
            "按列选择删除、均值、中位数、众数或自定义填充；每次应用都从原始数据重新计算。",
        ),
        (
            "03",
            "即时洞察",
            "数值字段自动出直方图与箱线图，分类字段生成频数柱状图或环形饼图。",
        ),
    ]
    for column, (index, title, description) in zip(columns, cards):
        with column:
            st.markdown(
                f"""
                <div class="feature-card">
                    <div class="feature-index">{index}</div>
                    <h3>{title}</h3>
                    <p>{description}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.write("")
    button_column, note_column = st.columns([1, 2])
    with button_column:
        if st.button(
            "加载示例数据",
            type="primary",
            icon="🚀",
            width="stretch",
            key="load_sample_button",
        ):
            if load_sample_data():
                st.rerun()
    with note_column:
        st.caption("想先看看效果？内置 30 行客户数据，包含数值、分类和日期字段的缺失值。")


# ---------------------------------------------------------------------------
# 页面：数据概览
# ---------------------------------------------------------------------------


def section_header(number: str, title: str, subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="section-head">
            <div class="section-num">{number}</div>
            <div>
                <h2>{title}</h2>
                <p>{subtitle}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_overview(df: pd.DataFrame, clean_df: Optional[pd.DataFrame]) -> None:
    section_header(
        "01",
        "数据概览",
        "先确认数据规模、类型和缺失情况，再决定是否清洗。",
    )

    total_rows, total_columns = df.shape
    total_cells = total_rows * total_columns
    missing_total = int(df.isna().sum().sum())
    missing_ratio = missing_total / total_cells if total_cells else 0.0
    complete_rows = int((~df.isna().any(axis=1)).sum())
    duplicate_rows = int(df.duplicated().sum())

    metric_columns = st.columns(4)
    metric_columns[0].metric("总行数", _format_int(total_rows), help="上传文件中的记录数")
    metric_columns[1].metric("总列数", _format_int(total_columns), help="字段数量")
    metric_columns[2].metric(
        "缺失值总数",
        _format_int(missing_total),
        help=f"占全部数据单元格的 {missing_ratio:.1%}",
    )
    metric_columns[3].metric(
        "完整行数",
        _format_int(complete_rows),
        help=f"无任何缺失值的行；重复行 {duplicate_rows:,} 行",
    )
    st.caption(
        f"缺失率 {missing_ratio:.1%} · 重复行 {duplicate_rows:,} 行 · "
        f"数据于 {st.session_state.loaded_at} 载入"
    )

    profile = build_column_profile(df)
    type_summary = (
        profile.groupby("数据类型", as_index=False)
        .agg(
            字段数=("字段", "size"),
            字段列表=("字段", lambda values: "、".join(str(v) for v in values)),
        )
        .sort_values("字段数", ascending=False)
    )
    type_summary["占比"] = type_summary["字段数"] / max(total_columns, 1)

    with st.container(border=True):
        st.markdown("#### 字段类型统计")
        st.dataframe(
            type_summary,
            hide_index=True,
            width="stretch",
            column_config={
                "字段数": st.column_config.NumberColumn(format="%d"),
                "占比": st.column_config.ProgressColumn(
                    "占比",
                    format="percent",
                    min_value=0.0,
                    max_value=1.0,
                ),
                "字段列表": st.column_config.TextColumn("包含字段", width="large"),
            },
        )

    with st.container(border=True):
        st.markdown("#### 原始数据预览")
        st.caption(
            "按要求展示前 5 行；表中 None 表示缺失值。"
            "完整数据可通过清洗结果或字段明细查看。"
        )
        if clean_df is not None:
            raw_tab, clean_tab = st.tabs(
                ["原始数据（前 5 行）", "清洗后数据（前 5 行）"]
            )
            with raw_tab:
                st.dataframe(_preview_styler(df), hide_index=True, width="stretch")
            with clean_tab:
                st.dataframe(_preview_styler(clean_df), hide_index=True, width="stretch")
        else:
            st.dataframe(_preview_styler(df), hide_index=True, width="stretch")

    with st.expander("查看全部字段明细", expanded=False):
        st.dataframe(
            profile,
            hide_index=True,
            width="stretch",
            column_config={
                "缺失数": st.column_config.NumberColumn(format="%d"),
                "缺失率": st.column_config.ProgressColumn(
                    "缺失率",
                    format="percent",
                    min_value=0.0,
                    max_value=1.0,
                ),
                "非空值": st.column_config.NumberColumn(format="%d"),
                "唯一值": st.column_config.NumberColumn(format="%d"),
            },
            height=min(520, 70 + 35 * len(profile)),
        )


# ---------------------------------------------------------------------------
# 页面：缺失值清洗
# ---------------------------------------------------------------------------


def _strategy_options(series: pd.Series) -> List[str]:
    return NUMERIC_STRATEGIES if _is_numeric_column(series) else OTHER_STRATEGIES


def _strategy_fingerprint(configs: Dict[str, Dict[str, Optional[str]]]) -> str:
    serializable = {
        str(column): {
            "strategy": config.get("strategy", "keep"),
            "constant": str(config.get("constant") or ""),
        }
        for column, config in sorted(configs.items(), key=lambda item: str(item[0]))
    }
    return hashlib.md5(
        json.dumps(serializable, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()


def _render_missing_settings(
    df: pd.DataFrame, token: str
) -> Tuple[Dict[str, Dict[str, Optional[str]]], pd.DataFrame]:
    """渲染缺失值表格和按列策略卡片，返回策略配置与缺失摘要。"""
    missing_counts = _missing_counts(df)
    missing_counts = missing_counts[missing_counts > 0].sort_values(ascending=False)
    total_rows = max(len(df), 1)
    summary = pd.DataFrame(
        {
            "字段": [str(column) for column in missing_counts.index],
            "数据类型": [
                _friendly_dtype(df[column]) for column in missing_counts.index
            ],
            "缺失数": [int(value) for value in missing_counts.values],
            "缺失率": [int(value) / total_rows for value in missing_counts.values],
        }
    )

    st.dataframe(
        summary,
        hide_index=True,
        width="stretch",
        height=min(520, 72 + 35 * len(summary)),
        column_config={
            "缺失数": st.column_config.NumberColumn(format="%d"),
            "缺失率": st.column_config.ProgressColumn(
                "缺失率",
                format="percent",
                min_value=0.0,
                max_value=1.0,
            ),
        },
    )

    st.markdown("#### 按字段设置清洗策略")
    st.caption(
        "每个字段都可以使用不同策略。点击下方按钮后才会真正执行；"
        "再次应用会从原始数据重新计算。"
    )
    st.info(
        "“删除缺失行”会在所有字段上联合执行：任意一个设置为该策略的字段缺失，"
        "对应行都会被删除。",
        icon="💡",
    )

    configs: Dict[str, Dict[str, Optional[str]]] = {}
    missing_columns = [str(column) for column in missing_counts.index]

    for start in range(0, len(missing_columns), 2):
        row_columns = st.columns(2, gap="large")
        for offset, column in enumerate(missing_columns[start : start + 2]):
            series = df[column]
            missing = int(series.isna().sum())
            ratio = missing / total_rows
            options = _strategy_options(series)

            with row_columns[offset]:
                with st.container(border=True):
                    st.markdown(f"**{column}**")
                    st.caption(
                        f"{_friendly_dtype(series)} · 缺失 {missing:,} / {len(df):,} "
                        f"({ratio:.1%})"
                    )
                    strategy_key = _widget_key(
                        token, "clean", "strategy", column
                    )
                    if (
                        strategy_key in st.session_state
                        and st.session_state[strategy_key] not in options
                    ):
                        del st.session_state[strategy_key]
                    strategy = st.selectbox(
                        "清洗策略",
                        options=options,
                        format_func=lambda value: STRATEGY_LABELS[value],
                        key=strategy_key,
                    )

                    constant_value: Optional[str] = None
                    if strategy == "constant":
                        constant = st.text_input(
                            "自定义填充值",
                            placeholder=_constant_placeholder(series),
                            key=_widget_key(token, "clean", "constant", column),
                        )
                        constant_value = constant
                    configs[column] = {
                        "strategy": strategy,
                        "constant": constant_value,
                    }

    return configs, summary


def _render_cleaning_result(
    df: pd.DataFrame,
    result: CleaningResult,
    token: str,
) -> None:
    missing_ratio_before = (
        result.missing_before / (result.rows_before * len(df.columns))
        if result.rows_before and len(df.columns)
        else 0.0
    )
    missing_ratio_after = (
        result.missing_after / (result.rows_after * len(df.columns))
        if result.rows_after and len(df.columns)
        else 0.0
    )

    st.markdown("#### 清洗结果")
    metric_columns = st.columns(4)
    metric_columns[0].metric(
        "数据行数",
        _format_int(result.rows_after),
        delta=f"{result.rows_after - result.rows_before:+,}",
        delta_color="off",
        help=f"清洗前 {result.rows_before:,} 行",
    )
    metric_columns[1].metric(
        "缺失值",
        _format_int(result.missing_after),
        delta=f"{result.missing_after - result.missing_before:+,}",
        delta_color="inverse",
        help=f"清洗前 {result.missing_before:,} 个",
    )
    metric_columns[2].metric(
        "剩余缺失率",
        f"{missing_ratio_after:.1%}",
        help=f"清洗前 {missing_ratio_before:.1%}",
    )
    metric_columns[3].metric(
        "已执行操作",
        _format_int(len(result.actions)),
        help=f"包含 {len(result.warnings)} 条提示",
    )

    if result.rows_after == 0:
        st.warning("清洗后没有剩余数据行，请调整“删除缺失行”策略后重试。")

    st.caption("清洗后数据预览（前 5 行）；None 表示仍未处理的缺失值。")
    st.dataframe(
        _preview_styler(result.data),
        hide_index=True,
        width="stretch",
    )

    download_column, info_column = st.columns([1, 2])
    with download_column:
        st.download_button(
            "下载清洗后的 CSV",
            data=dataframe_to_csv(result.data),
            file_name=_safe_download_name(st.session_state.file_name),
            mime="text/csv; charset=utf-8",
            type="primary",
            icon="📥",
            width="stretch",
            key=f"main_download_{token}_{st.session_state.clean_version}",
        )
    with info_column:
        st.caption(
            f"文件名：{_safe_download_name(st.session_state.file_name)} · "
            "UTF-8 BOM 编码，Excel 打开中文不乱码"
        )

    with st.expander("查看清洗日志与提示", expanded=not result.actions and bool(result.warnings)):
        if result.actions:
            action_df = pd.DataFrame(result.actions)
            st.dataframe(
                action_df,
                hide_index=True,
                width="stretch",
                column_config={
                    "字段": st.column_config.TextColumn("字段", width="medium"),
                    "处理方式": st.column_config.TextColumn("处理方式", width="small"),
                    "说明": st.column_config.TextColumn("说明", width="large"),
                },
            )
        else:
            st.caption("本次没有字段被修改。")

        for warning in result.warnings:
            st.warning(warning, icon="⚠️")
        if not result.warnings:
            st.success("全部策略均已顺利执行。", icon="✅")


def render_cleaning(df: pd.DataFrame, token: str) -> None:
    section_header(
        "02",
        "缺失值清洗",
        "先看缺失分布，再为每个字段指定策略；清洗结果会保存在会话中。",
    )

    missing_counts = _missing_counts(df)
    missing_counts = missing_counts[missing_counts > 0]
    if missing_counts.empty:
        st.success("当前数据没有缺失值，可以直接进入可视化分析。", icon="✅")
        st.caption("如后续换用含缺失值的数据，此页面会自动生成清洗设置。")
        st.download_button(
            "下载当前数据 CSV",
            data=dataframe_to_csv(df),
            file_name=_safe_download_name(st.session_state.file_name),
            mime="text/csv; charset=utf-8",
            icon="📥",
            width="stretch",
            key=f"no_missing_download_{token}",
        )
        return

    st.caption(
        f"共 {len(missing_counts):,} 个字段存在缺失，"
        f"涉及 {int(missing_counts.sum()):,} 个缺失单元格。"
    )

    configs, _ = _render_missing_settings(df, token)
    current_fingerprint = _strategy_fingerprint(configs)
    applied_fingerprint = st.session_state.applied_config_fingerprint
    if (
        st.session_state.clean_meta is not None
        and applied_fingerprint != current_fingerprint
    ):
        st.warning(
            "清洗策略已修改。当前预览和下载仍为上次应用的结果，请重新点击"
            "“应用清洗”以更新。",
            icon="⚠️",
        )

    st.divider()
    apply_column, reset_column, note_column = st.columns([1.15, 1.15, 2.7])
    with apply_column:
        apply_clicked = st.button(
            "应用清洗",
            type="primary",
            icon="✨",
            width="stretch",
            key=f"apply_cleaning_{token}",
        )
    with reset_column:
        st.button(
            "重置策略与结果",
            on_click=reset_cleaning_state,
            icon="🔄",
            width="stretch",
            key=f"reset_cleaning_{token}",
        )
    with note_column:
        st.caption("清洗始终基于原始数据执行，反复调整策略不会产生叠加误差。")

    if apply_clicked:
        try:
            with st.spinner("正在应用清洗策略..."):
                result = clean_data(df, configs)
            st.session_state.clean_df = result.data
            st.session_state.clean_meta = result
            st.session_state.applied_config_fingerprint = current_fingerprint
            st.session_state.clean_version = int(st.session_state.clean_version) + 1
            st.success(
                f"清洗完成：{result.rows_before:,} → {result.rows_after:,} 行，"
                f"缺失值 {result.missing_before:,} → {result.missing_after:,} 个。",
                icon="✅",
            )
            st.toast("清洗结果已保存到当前会话", icon="✅")
        except Exception as exc:
            st.error(f"清洗失败：{exc}")

    result = st.session_state.clean_meta
    if result is not None:
        _render_cleaning_result(df, result, token)


# ---------------------------------------------------------------------------
# 页面：多维可视化
# ---------------------------------------------------------------------------


def _ensure_multiselect(
    key: str,
    options: Sequence[str],
    default: Optional[Sequence[str]] = None,
) -> None:
    """在组件创建前同步选项，避免切换数据后保留失效字段。"""
    option_list = list(options)
    if key not in st.session_state:
        st.session_state[key] = list(default if default is not None else option_list)
        return
    selected = st.session_state.get(key)
    if not isinstance(selected, list):
        st.session_state[key] = option_list
        return
    valid = [value for value in selected if value in option_list]
    if valid != selected:
        st.session_state[key] = valid


def _render_numeric_grid(
    df: pd.DataFrame,
    columns: Sequence[str],
    token: str,
    chart_mode: str,
) -> None:
    if not columns:
        st.info("当前数据中没有数值型字段，因此不会生成数值分布图。")
        return

    for start in range(0, len(columns), 2):
        grid = st.columns(2, gap="large")
        for offset, column in enumerate(columns[start : start + 2]):
            series = df[column]
            with grid[offset]:
                with st.container(border=True):
                    st.markdown(
                        '<div class="chart-kicker">NUMERICAL</div>'
                        f'<div class="chart-title">{html.escape(str(column))}</div>',
                        unsafe_allow_html=True,
                    )
                    values, missing, infinity = _numeric_values(series)
                    if values.empty:
                        st.info("该字段没有可绘制的有效数值。")
                        continue

                    mean_value = float(values.mean())
                    median_value = float(values.median())
                    std_value = float(values.std(ddof=1)) if len(values) > 1 else 0.0
                    st.caption(
                        f"均值 {_format_stat(mean_value)} · "
                        f"中位数 {_format_stat(median_value)} · "
                        f"标准差 {_format_stat(std_value)} · "
                        f"缺失 {missing:,} · 无穷值 {infinity:,}"
                    )

                    plot_values = values
                    if len(values) > MAX_PLOT_POINTS:
                        plot_values = values.sample(
                            n=MAX_PLOT_POINTS, random_state=42
                        )
                        st.caption(
                            f"为保持交互流畅，图表对 {len(values):,} 个有效值"
                            f"随机抽样 {MAX_PLOT_POINTS:,} 个；统计量仍基于全量数据。"
                        )

                    try:
                        figure = build_numeric_figure(
                            plot_values,
                            str(column),
                            chart_mode,
                            mean_value,
                            median_value,
                        )
                        st.plotly_chart(
                            figure,
                            use_container_width=True,
                            theme=None,
                            config=PLOT_CONFIG,
                            key=_widget_key(
                                token,
                                "viz",
                                "numeric",
                                column,
                                chart_mode,
                                st.session_state.clean_version,
                            ),
                        )
                    except Exception as exc:
                        st.warning(f"图表生成失败：{exc}")


def _render_category_grid(
    df: pd.DataFrame,
    columns: Sequence[str],
    token: str,
    chart_mode: str,
    top_n: int,
) -> None:
    if not columns:
        st.info("当前数据中没有分类型/文本型字段，因此不会生成频数图。")
        return

    for start in range(0, len(columns), 2):
        grid = st.columns(2, gap="large")
        for offset, column in enumerate(columns[start : start + 2]):
            series = df[column]
            with grid[offset]:
                with st.container(border=True):
                    st.markdown(
                        '<div class="chart-kicker">CATEGORICAL</div>'
                        f'<div class="chart-title">{html.escape(str(column))}</div>',
                        unsafe_allow_html=True,
                    )
                    prepared = _category_counts(series, top_n)
                    if prepared is None:
                        st.info("该字段没有可绘制的有效类别。")
                        continue
                    plot_df, total, missing, merged, category_total = prepared
                    detail = (
                        f"{category_total:,} 个类别 · 总计 {total:,} 条 · "
                        f"缺失 {missing:,}"
                    )
                    if merged:
                        detail += f" · 仅显示 Top {top_n}，其余合并为“其他”"
                    if missing:
                        detail += " · 缺失值作为独立类别展示"
                    st.caption(detail)

                    try:
                        if chart_mode == "bar":
                            figure = build_category_bar(plot_df, total)
                        else:
                            figure = build_category_pie(plot_df, total)
                        st.plotly_chart(
                            figure,
                            use_container_width=True,
                            theme=None,
                            config=PLOT_CONFIG,
                            key=_widget_key(
                                token,
                                "viz",
                                "category",
                                column,
                                chart_mode,
                                top_n,
                                st.session_state.clean_version,
                            ),
                        )
                    except Exception as exc:
                        st.warning(f"图表生成失败：{exc}")


def _ensure_option_state(key: str, options: Sequence[str]) -> None:
    """在创建组件前移除已失效的选项状态，避免切换字段后报错。"""
    option_list = [str(option) for option in options]
    if key in st.session_state and st.session_state[key] not in option_list:
        del st.session_state[key]


def _chart_card_header(
    kicker: str,
    title: str,
    description: Optional[str] = None,
) -> None:
    """为高级分析图表绘制统一的卡片标题。"""
    st.markdown(
        f'<div class="chart-kicker">{html.escape(kicker)}</div>'
        f'<div class="chart-title">{html.escape(title)}</div>',
        unsafe_allow_html=True,
    )
    if description:
        st.caption(description)


def _render_scatter_section(
    df: pd.DataFrame,
    numeric_columns: Sequence[str],
    category_columns: Sequence[str],
    token: str,
) -> None:
    """散点图：支持数值 X/Y、分类着色与线性趋势。"""
    with st.container(border=True):
        _chart_card_header(
            "BIVARIATE · NUMERICAL",
            "散点图",
            "选择两个数值字段查看相关性、点云分布和线性趋势；可叠加分类着色。",
        )
        if len(numeric_columns) < 2:
            st.warning("散点图至少需要 2 个数值型字段。当前数据不满足条件。")
            return

        control_columns = st.columns([1.1, 1.1, 1.1])
        x_key = _widget_key(
            token,
            "viz",
            "scatter_x",
            st.session_state.clean_version,
        )
        _ensure_option_state(x_key, numeric_columns)
        x_col = control_columns[0].selectbox(
            "X 轴（数值型）",
            options=list(numeric_columns),
            key=x_key,
        )

        y_key = _widget_key(
            token,
            "viz",
            "scatter_y",
            st.session_state.clean_version,
        )
        if y_key not in st.session_state:
            st.session_state[y_key] = numeric_columns[1]
        _ensure_option_state(y_key, numeric_columns)
        y_col = control_columns[1].selectbox(
            "Y 轴（数值型）",
            options=list(numeric_columns),
            key=y_key,
        )

        color_key = _widget_key(
            token,
            "viz",
            "scatter_color",
            st.session_state.clean_version,
        )
        color_options = ["__none__"] + list(category_columns)
        _ensure_option_state(color_key, color_options)
        color_choice = control_columns[2].selectbox(
            "颜色分组（可选）",
            options=color_options,
            format_func=lambda value: "不分组" if value == "__none__" else value,
            key=color_key,
        )
        color_col = None if color_choice == "__none__" else color_choice
        show_trend = st.checkbox(
            "显示线性趋势线",
            value=True,
            help="使用最小二乘线性拟合；趋势线仅表示总体方向，不代表因果关系。",
            key=_widget_key(
                token,
                "viz",
                "scatter_trend",
                st.session_state.clean_version,
            ),
        )

        if x_col == y_col:
            st.warning("X 轴和 Y 轴选择了同一个字段，请选择两个不同的数值字段。")
            return

        try:
            figure, meta = build_scatter_figure(
                df,
                str(x_col),
                str(y_col),
                color_col=str(color_col) if color_col else None,
                show_trend=show_trend,
            )
        except ValueError as exc:
            st.warning(str(exc))
            return
        except Exception as exc:
            st.warning(f"散点图生成失败：{exc}")
            return

        correlation = meta["correlation"]
        correlation_text = (
            f"Pearson r = {correlation:+.3f}"
            if pd.notna(correlation)
            else "Pearson r 不可计算（存在常量字段）"
        )
        sample_text = (
            f" · 图中抽样 {meta['shown']:,} / {meta['total']:,} 条"
            if meta["shown"] < meta["total"]
            else f" · 有效记录 {meta['total']:,} 条"
        )
        st.caption(correlation_text + sample_text)
        st.plotly_chart(
            figure,
            use_container_width=True,
            theme=None,
            config=PLOT_CONFIG,
            key=_widget_key(
                token,
                "viz",
                "scatter_chart",
                st.session_state.clean_version,
            ),
        )


def _render_density_section(
    df: pd.DataFrame,
    numeric_columns: Sequence[str],
    token: str,
) -> None:
    """密度图：对多个数值字段叠加 KDE 曲线。"""
    with st.container(border=True):
        _chart_card_header(
            "DISTRIBUTION",
            "核密度图（KDE）",
            "用平滑曲线呈现单一数值字段的概率密度，适合观察分布形态、峰值与偏态。",
        )
        if not numeric_columns:
            st.warning("密度图至少需要 1 个数值型字段。当前数据不满足条件。")
            return

        density_key = _widget_key(
            token,
            "viz",
            "density_column",
            st.session_state.clean_version,
        )
        _ensure_option_state(density_key, numeric_columns)
        selected = st.selectbox(
            "密度字段（数值型）",
            options=list(numeric_columns),
            key=density_key,
        )

        try:
            figure, skipped, profiles = build_density_figure(df, [str(selected)])
        except ValueError as exc:
            st.warning(str(exc))
            return
        except Exception as exc:
            st.warning(f"密度图生成失败：{exc}")
            return

        if skipped:
            st.warning(
                "该字段有效值不足或为常量，无法计算核密度估计："
                + "、".join(skipped)
            )
            return
        profile = profiles[0]
        st.caption(
            "Silverman 经验带宽 · "
            f"有效样本 {profile['n']:,} · "
            f"缺失 {profile['missing']:,} · "
            f"带宽 {profile['bandwidth']:,.4g}"
        )
        st.plotly_chart(
            figure,
            use_container_width=True,
            theme=None,
            config=PLOT_CONFIG,
            key=_widget_key(
                token,
                "viz",
                "density_chart",
                st.session_state.clean_version,
            ),
        )


def _render_line_section(
    df: pd.DataFrame,
    numeric_columns: Sequence[str],
    token: str,
) -> None:
    """折线图：按数据索引或数值 X 轴查看趋势。"""
    with st.container(border=True):
        _chart_card_header(
            "TREND",
            "折线图",
            "按行顺序或数值字段排序，观察单一数值指标的变化趋势。",
        )
        if not numeric_columns:
            st.warning("折线图至少需要 1 个数值型字段。当前数据不满足条件。")
            return

        x_key = _widget_key(
            token,
            "viz",
            "line_x",
            st.session_state.clean_version,
        )
        x_options = ["__index__"] + list(numeric_columns)
        _ensure_option_state(x_key, x_options)
        field_columns = st.columns([1.2, 1.2, 0.9, 0.9])
        x_choice = field_columns[0].selectbox(
            "X 轴",
            options=x_options,
            format_func=lambda value: (
                "数据索引（行顺序）" if value == "__index__" else value
            ),
            key=x_key,
        )

        y_key = _widget_key(
            token,
            "viz",
            "line_y",
            st.session_state.clean_version,
        )
        if y_key not in st.session_state:
            st.session_state[y_key] = numeric_columns[0]
        _ensure_option_state(y_key, numeric_columns)
        y_col = field_columns[1].selectbox(
            "Y 轴（数值型）",
            options=list(numeric_columns),
            key=y_key,
        )

        use_index = x_choice == "__index__"
        sort_by_x = field_columns[2].checkbox(
            "按 X 排序",
            value=True,
            disabled=use_index,
            key=_widget_key(
                token,
                "viz",
                "line_sort",
                st.session_state.clean_version,
            ),
        )
        show_markers = field_columns[3].checkbox(
            "显示数据点",
            value=len(df) <= 200,
            key=_widget_key(
                token,
                "viz",
                "line_markers",
                st.session_state.clean_version,
            ),
        )

        if not use_index and str(y_col) == str(x_choice):
            st.warning("X 轴和 Y 轴不能是同一个字段，请选择不同的数值字段。")
            return

        try:
            figure, meta = build_line_figure(
                df,
                str(x_choice),
                str(y_col),
                use_index=use_index,
                sort_by_x=sort_by_x,
                show_markers=show_markers,
            )
        except ValueError as exc:
            st.warning(str(exc))
            return
        except Exception as exc:
            st.warning(f"折线图生成失败：{exc}")
            return

        if meta["sampled"]:
            st.caption(
                f"共 {meta['total']:,} 条，图中等距抽样 {meta['shown']:,} 条展示；"
                "完整数据仍保留在会话中。"
            )
        else:
            st.caption(f"X 轴：{meta['x_label']} · 有效记录 {meta['total']:,} 条")
        st.plotly_chart(
            figure,
            use_container_width=True,
            theme=None,
            config=PLOT_CONFIG,
            key=_widget_key(
                token,
                "viz",
                "line_chart",
                st.session_state.clean_version,
            ),
        )


def _render_stacked_section(
    df: pd.DataFrame,
    category_columns: Sequence[str],
    token: str,
) -> None:
    """堆叠柱状图：展示两个分类字段的组合结构与占比。"""
    with st.container(border=True):
        _chart_card_header(
            "COMPOSITION",
            "堆叠柱状图",
            "选择两个分类型字段，查看组合构成；可切换计数或每一列归一化后的百分比。",
        )
        if len(category_columns) < 2:
            st.warning("堆叠柱状图至少需要 2 个分类型字段。当前数据不满足条件。")
            return

        x_key = _widget_key(
            token,
            "viz",
            "stack_x",
            st.session_state.clean_version,
        )
        _ensure_option_state(x_key, category_columns)
        control_columns = st.columns([1.2, 1.2, 1.0, 1.0])
        x_col = control_columns[0].selectbox(
            "X 轴（分类型）",
            options=list(category_columns),
            key=x_key,
        )

        color_options = [column for column in category_columns if column != x_col]
        if not color_options:
            st.warning("需要另一个分类型字段作为堆叠颜色。")
            return
        color_key = _widget_key(
            token,
            "viz",
            "stack_color",
            st.session_state.clean_version,
        )
        _ensure_option_state(color_key, color_options)
        if color_key not in st.session_state:
            st.session_state[color_key] = color_options[0]
        color_col = control_columns[1].selectbox(
            "堆叠颜色（分类型）",
            options=color_options,
            key=color_key,
        )

        display_key = _widget_key(
            token,
            "viz",
            "stack_display",
            st.session_state.clean_version,
        )
        display = control_columns[2].selectbox(
            "数值呈现",
            options=["count", "percent"],
            format_func=lambda value: "记录数" if value == "count" else "行内百分比",
            key=display_key,
        )
        top_n = int(
            control_columns[3].slider(
                "类别上限",
                min_value=5,
                max_value=20,
                value=10,
                step=1,
                help="X 轴和颜色字段中超出上限的长尾类别会合并为“其他”。",
                key=_widget_key(
                    token,
                    "viz",
                    "stack_top_n",
                    st.session_state.clean_version,
                ),
            )
        )

        try:
            figure, meta = build_stacked_bar_figure(
                df,
                str(x_col),
                str(color_col),
                normalize=display == "percent",
                top_n=top_n,
            )
        except ValueError as exc:
            st.warning(str(exc))
            return
        except Exception as exc:
            st.warning(f"堆叠柱状图生成失败：{exc}")
            return

        detail_parts = []
        if meta["missing"]:
            detail_parts.append(f"忽略 {meta['missing']:,} 条字段缺失记录")
        if meta["x_merged"]:
            detail_parts.append(f"X 轴 {meta['x_extra']} 个长尾类别已合并")
        if meta["y_merged"]:
            detail_parts.append(f"颜色字段 {meta['y_extra']} 个长尾类别已合并")
        if display == "percent":
            detail_parts.append("每个 X 类别的堆叠高度为 100%")
        if detail_parts:
            st.caption(" · ".join(detail_parts))
        st.plotly_chart(
            figure,
            use_container_width=True,
            theme=None,
            config=PLOT_CONFIG,
            key=_widget_key(
                token,
                "viz",
                "stack_chart",
                st.session_state.clean_version,
            ),
        )


def _render_grouped_section(
    df: pd.DataFrame,
    numeric_columns: Sequence[str],
    category_columns: Sequence[str],
    token: str,
) -> None:
    """分组箱线图与小提琴图，共享字段选择和类别范围。"""
    if not numeric_columns or not category_columns:
        st.warning("分组对比图需要至少 1 个数值型字段和 1 个分类型字段。")
        return

    with st.container(border=True):
        _chart_card_header(
            "GROUP COMPARISON",
            "分组箱线图 & 小提琴图",
            "以分类型字段为 X 轴、数值字段为 Y 轴，对比不同类别的分布差异；"
            "可再选择一个分类字段进行二重分组。",
        )
        control_columns = st.columns([1.2, 1.2, 1.2, 0.9])
        category_key = _widget_key(
            token,
            "viz",
            "group_category",
            st.session_state.clean_version,
        )
        _ensure_option_state(category_key, category_columns)
        category_col = control_columns[0].selectbox(
            "X 轴（分类型）",
            options=list(category_columns),
            key=category_key,
        )

        numeric_key = _widget_key(
            token,
            "viz",
            "group_numeric",
            st.session_state.clean_version,
        )
        _ensure_option_state(numeric_key, numeric_columns)
        numeric_col = control_columns[1].selectbox(
            "Y 轴（数值型）",
            options=list(numeric_columns),
            key=numeric_key,
        )

        color_options = ["__none__"] + [
            column for column in category_columns if column != category_col
        ]
        color_key = _widget_key(
            token,
            "viz",
            "group_color",
            st.session_state.clean_version,
        )
        _ensure_option_state(color_key, color_options)
        color_choice = control_columns[2].selectbox(
            "二次分组（可选）",
            options=color_options,
            format_func=lambda value: "不分组" if value == "__none__" else value,
            key=color_key,
        )
        color_col = None if color_choice == "__none__" else color_choice
        top_n = int(
            control_columns[3].slider(
                "类别上限",
                min_value=5,
                max_value=20,
                value=10,
                step=1,
                help="X 轴与二次分组字段中超出上限的长尾类别会合并为“其他”。",
                key=_widget_key(
                    token,
                    "viz",
                    "group_top_n",
                    st.session_state.clean_version,
                ),
            )
        )

        try:
            working, category_order, _, meta = _prepare_group_data(
                df,
                str(category_col),
                str(numeric_col),
                str(color_col) if color_col else None,
                top_n,
            )
        except ValueError as exc:
            st.warning(str(exc))
            return
        except Exception as exc:
            st.warning(f"分组数据准备失败：{exc}")
            return

        detail_parts = [f"有效记录 {len(working):,} 条"]
        if meta["missing_dropped"]:
            detail_parts.append(f"忽略 {meta['missing_dropped']:,} 条缺失记录")
        if meta["category_merged"]:
            detail_parts.append(f"X 轴 {meta['category_extra']} 个长尾类别已合并")
        if meta["color_merged"]:
            detail_parts.append(f"二次分组 {meta['color_extra']} 个长尾类别已合并")
        st.caption(" · ".join(detail_parts))

    chart_working = working
    sampled = len(chart_working) > MAX_PLOT_POINTS
    if sampled:
        chart_working = chart_working.sample(n=MAX_PLOT_POINTS, random_state=42)

    chart_columns = st.columns(2, gap="large")
    chart_specs = [
        ("BOX PLOT", "分组箱线图", build_group_box_figure),
        ("VIOLIN PLOT", "小提琴图", build_group_violin_figure),
    ]
    for chart_column, (kicker, title, builder) in zip(chart_columns, chart_specs):
        with chart_column:
            with st.container(border=True):
                _chart_card_header(kicker, title)
                try:
                    figure = builder(
                        chart_working,
                        str(category_col),
                        str(numeric_col),
                        str(color_col) if color_col else None,
                        category_order,
                    )
                    if sampled:
                        st.caption(
                            f"为提升交互性能，图中抽样 {len(chart_working):,} / "
                            f"{len(working):,} 条记录。"
                        )
                    st.plotly_chart(
                        figure,
                        use_container_width=True,
                        theme=None,
                        config=PLOT_CONFIG,
                        key=_widget_key(
                            token,
                            "viz",
                            f"chart_{kicker}",
                            category_col,
                            numeric_col,
                            color_col or "",
                            top_n,
                            st.session_state.clean_version,
                        ),
                    )
                except Exception as exc:
                    st.warning(f"{title}生成失败：{exc}")


def _render_correlation_section(
    df: pd.DataFrame,
    numeric_columns: Sequence[str],
    token: str,
) -> None:
    """相关性热力图：支持 Pearson / Spearman 与数值标注。"""
    with st.container(border=True):
        _chart_card_header(
            "MULTIVARIATE",
            "相关性热力图",
            "自动使用全部有效数值字段计算相关系数矩阵；红色为正相关，蓝色为负相关。",
        )
        if len(numeric_columns) < 2:
            st.warning("相关性分析至少需要 2 个数值型字段。当前数据不满足条件。")
            return

        method_column, annotate_column = st.columns([1, 1.8])
        method = method_column.selectbox(
            "相关系数方法",
            options=["pearson", "spearman"],
            format_func=lambda value: (
                "Pearson · 线性相关" if value == "pearson" else "Spearman · 秩相关"
            ),
            key=_widget_key(
                token,
                "viz",
                "correlation_method",
                st.session_state.clean_version,
            ),
        )
        annotate = annotate_column.checkbox(
            "在格子里标注相关系数",
            value=len(numeric_columns) <= 14,
            help="字段很多时可以关闭标注，悬停依然可以查看精确数值。",
            key=_widget_key(
                token,
                "viz",
                "correlation_annotate",
                st.session_state.clean_version,
            ),
        )

        try:
            figure, correlation, dropped = build_correlation_figure(
                df,
                numeric_columns,
                method,
                annotate,
            )
        except ValueError as exc:
            st.warning(str(exc))
            return
        except Exception as exc:
            st.warning(f"相关性热力图生成失败：{exc}")
            return

        if dropped:
            st.warning(
                "以下字段只有 1 个唯一值或有效观测不足，已从矩阵中排除："
                + "、".join(dropped)
            )
        st.caption(
            f"矩阵包含 {len(correlation.columns)} 个字段 · "
            "按成对完整观测计算 · 相关系数不代表因果关系"
        )
        st.plotly_chart(
            figure,
            use_container_width=True,
            theme=None,
            config=PLOT_CONFIG,
            key=_widget_key(
                token,
                "viz",
                "correlation_chart",
                st.session_state.clean_version,
            ),
        )

    pairs = _correlation_pairs(correlation, limit=8)
    if not pairs.empty:
        with st.expander("查看绝对值最大的相关字段对", expanded=False):
            st.dataframe(
                pairs,
                hide_index=True,
                width="stretch",
                column_config={
                    "相关系数": st.column_config.NumberColumn(
                        "相关系数",
                        format="%.3f",
                    )
                },
            )


def _render_single_variable_tab(
    df: pd.DataFrame,
    numeric_columns: Sequence[str],
    category_columns: Sequence[str],
    token: str,
) -> None:
    """保留原有单变量分布图，并放入独立标签页。"""
    st.caption(
        "自动为全部数值型和分类型字段出图。数值字段展示分布与离群点，"
        "分类字段展示频数与占比。"
    )
    with st.container(border=True):
        control_a, control_b, control_c = st.columns([1.1, 1.1, 1.1])
        with control_a:
            chart_mode = st.selectbox(
                "数值型图表",
                options=["both", "hist", "box"],
                format_func=lambda value: {
                    "both": "直方图 + 箱线图",
                    "hist": "仅直方图",
                    "box": "仅箱线图",
                }[value],
                key=_widget_key(
                    token,
                    "viz",
                    "numeric_mode",
                    st.session_state.clean_version,
                ),
            )
        with control_b:
            category_mode = st.selectbox(
                "分类型图表",
                options=["bar", "pie"],
                format_func=lambda value: {
                    "bar": "频数柱状图",
                    "pie": "环形饼图",
                }[value],
                key=_widget_key(
                    token,
                    "viz",
                    "category_mode",
                    st.session_state.clean_version,
                ),
            )
        with control_c:
            top_n = int(
                st.slider(
                    "分类图最多类别数",
                    min_value=5,
                    max_value=20,
                    value=10,
                    step=1,
                    help="超出的长尾类别会合并为“其他”，缺失值单独展示。",
                    key=_widget_key(
                        token,
                        "viz",
                        "top_n",
                        st.session_state.clean_version,
                    ),
                )
            )

        with st.expander("筛选要展示的字段", expanded=False):
            numeric_key = _widget_key(
                token,
                "viz",
                "numeric_columns",
                st.session_state.clean_version,
            )
            category_key = _widget_key(
                token,
                "viz",
                "category_columns",
                st.session_state.clean_version,
            )
            _ensure_multiselect(numeric_key, numeric_columns)
            _ensure_multiselect(category_key, category_columns)
            selected_numeric = st.multiselect(
                "数值型字段",
                options=list(numeric_columns),
                key=numeric_key,
                placeholder="选择数值字段",
            )
            selected_category = st.multiselect(
                "分类型 / 文本型字段",
                options=list(category_columns),
                key=category_key,
                placeholder="选择分类字段",
            )
        st.caption(
            "数值图默认使用 2 万条以内样本保证交互流畅；分类图会聚合长尾类别，"
            "缺失值作为单独类别标注。"
        )

    st.markdown("### 数值型字段")
    st.caption(f"{len(selected_numeric)} / {len(numeric_columns)} 个字段已选择")
    if numeric_columns and not selected_numeric:
        st.info("已隐藏全部数值型字段。可在上方“筛选要展示的字段”中重新选择。")
    else:
        _render_numeric_grid(df, selected_numeric, token, chart_mode)

    st.markdown("### 分类型 / 文本型字段")
    st.caption(f"{len(selected_category)} / {len(category_columns)} 个字段已选择")
    if category_columns and not selected_category:
        st.info("已隐藏全部分类型字段。可在上方“筛选要展示的字段”中重新选择。")
    else:
        _render_category_grid(
            df,
            selected_category,
            token,
            category_mode,
            top_n,
        )


def render_visualizations(
    raw_df: pd.DataFrame,
    clean_df: Optional[pd.DataFrame],
    token: str,
) -> None:
    section_header(
        "03",
        "多维可视化",
        "按分析层次组织图表：单变量看分布，双变量看关系，分组看差异，热力图看全局相关。",
    )

    if clean_df is not None:
        source_key = _widget_key(
            token, "viz", "source", st.session_state.clean_version
        )
        source = st.radio(
            "可视化数据源",
            options=["清洗后数据", "原始数据"],
            horizontal=True,
            key=source_key,
            help="清洗结果会保存在会话中，可随时切换查看清洗前后差异。",
        )
        viz_df = clean_df if source == "清洗后数据" else raw_df
    else:
        source = "原始数据"
        viz_df = raw_df
        st.caption("尚未应用清洗，当前直接使用原始数据进行可视化。")

    if viz_df.empty:
        st.warning("当前数据源没有数据行，无法生成图表。")
        return

    numeric_columns = [
        str(column)
        for column in viz_df.columns
        if _is_numeric_column(viz_df[column])
    ]
    category_columns = [
        str(column)
        for column in viz_df.columns
        if not _is_numeric_column(viz_df[column])
    ]

    missing_total = int(viz_df.isna().sum().sum())
    summary_columns = st.columns(4)
    summary_columns[0].metric("当前数据源", source)
    summary_columns[1].metric("数值型字段", _format_int(len(numeric_columns)))
    summary_columns[2].metric("分类型字段", _format_int(len(category_columns)))
    summary_columns[3].metric("剩余缺失值", _format_int(missing_total))

    st.write("")
    tabs = st.tabs(
        [
            "📊  单变量分布",
            "🔗  双变量关系",
            "🎻  分组对比",
            "🧮  相关性矩阵",
        ]
    )

    with tabs[0]:
        _render_single_variable_tab(
            viz_df,
            numeric_columns,
            category_columns,
            token,
        )

    with tabs[1]:
        st.caption(
            "三个分析视角共享当前数据源：散点图关注数值关系，KDE 关注分布形状，"
            "折线图关注序列趋势。"
        )
        _render_scatter_section(
            viz_df,
            numeric_columns,
            category_columns,
            token,
        )
        _render_density_section(viz_df, numeric_columns, token)
        _render_line_section(viz_df, numeric_columns, token)

    with tabs[2]:
        st.caption(
            "堆叠柱状图关注分类组合结构；箱线图和小提琴图从不同角度对比类别间的"
            "数值分布、中位数与离群点。"
        )
        _render_stacked_section(viz_df, category_columns, token)
        _render_grouped_section(
            viz_df,
            numeric_columns,
            category_columns,
            token,
        )

    with tabs[3]:
        _render_correlation_section(viz_df, numeric_columns, token)


def render_footer() -> None:
    st.markdown(
        """
        <div class="app-footer">
            DataPolish · Streamlit + Pandas + Plotly<br>
            文件在本地 Streamlit 进程中解析；清洗结果保存在当前浏览器会话中。
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 应用入口
# ---------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="🧹",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_state()
    inject_css()
    render_hero()
    render_sidebar()

    raw_df = st.session_state.raw_df
    if raw_df is None:
        render_welcome()
        render_footer()
        return

    # 新数据使用签名生成短 token，隔离不同文件的组件状态。
    token = hashlib.sha1(
        str(st.session_state.file_signature).encode("utf-8")
    ).hexdigest()[:10]
    page = st.session_state.nav_page

    try:
        if page == "clean":
            render_cleaning(raw_df, token)
        elif page == "visuals":
            render_visualizations(
                raw_df,
                st.session_state.clean_df,
                token,
            )
        else:
            render_overview(raw_df, st.session_state.clean_df)
    except Exception as exc:
        st.error(f"页面渲染出现异常：{exc}")
        with st.expander("查看错误详情", expanded=False):
            st.exception(exc)

    render_footer()


if __name__ == "__main__":
    main()
