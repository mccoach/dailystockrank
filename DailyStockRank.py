import os
import re
import sys
import json
import glob
import struct
import sqlite3
import threading
import queue
import calendar
import ntpath
from copy import deepcopy
from datetime import datetime, time, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ============================================================
# 基础常量
# ============================================================

CONFIG_FILE_NAME = "DailyStockRank_config.json"
QR_CODE_FILE_NAME = "wechat_qr.png"
ICON_FILE_NAME = "icon.ico"
DEFAULT_TIMEZONE_NAME = "Asia/Shanghai"
TDX_DAY_RECORD_SIZE = 32
TDX_TNF_HEADER_SIZE = 50
TDX_TNF_RECORD_SIZE = 360

# ============================================================
# 字段定义
# ============================================================

FIELD_DEFS = [
    {
        "key": "MARKET",
        "title": "市场",
        "db_expr": "c.market"
    },
    {
        "key": "SYMBOL",
        "title": "标的代码",
        "db_expr": "c.symbol"
    },
    {
        "key": "NAME",
        "title": "标的名称",
        "db_expr": "s.name"
    },
    {
        "key": "DATE",
        "title": "交易日期",
        "db_expr": None
    },
    {
        "key": "OPEN",
        "title": "开盘价",
        "db_expr": "c.open"
    },
    {
        "key": "HIGH",
        "title": "最高价",
        "db_expr": "c.high"
    },
    {
        "key": "LOW",
        "title": "最低价",
        "db_expr": "c.low"
    },
    {
        "key": "CLOSE",
        "title": "收盘价",
        "db_expr": "c.close"
    },
    {
        "key": "VOLUME",
        "title": "成交量",
        "db_expr": "c.volume"
    },
    {
        "key": "AMOUNT",
        "title": "成交额",
        "db_expr": "c.amount"
    },
]

DATA_SOURCE_OPTIONS = [
    {
        "key": "local_db",
        "title": "本地DB库",
        "path_kind": "file"
    },
    {
        "key": "tdx_day",
        "title": "通达信盘后数据",
        "path_kind": "directory"
    },
]

SORT_ORDER_OPTIONS = [
    {
        "key": "ASC",
        "title": "升序"
    },
    {
        "key": "DESC",
        "title": "降序"
    },
]

# ============================================================
# 默认配置
# ============================================================


def one_month_before(d):
    year = d.year
    month = d.month - 1

    if month == 0:
        year -= 1
        month = 12

    day = min(d.day, calendar.monthrange(year, month)[1])
    return d.replace(year=year, month=month, day=day)


def default_start_end_dates():
    today = datetime.now(ZoneInfo(DEFAULT_TIMEZONE_NAME)).date()
    start = one_month_before(today)
    return start.strftime("%Y%m%d"), today.strftime("%Y%m%d")


def make_default_config():
    start_date, end_date = default_start_end_dates()

    return {
        "config_version": 1,
        "ui": {
            "window_title": "个股扫描排序助手 v1.0 20260524",
            "window_geometry": "1360x760",
        },
        "data_source": "local_db",
        "paths": {
            "local_db": "",
            "tdx_day": "",
        },
        "date_range": {
            "start_date": start_date,
            "end_date": end_date,
        },
        "output": {
            "output_dir": "",
            "keep_all_rows": False,
            "max_rows_per_sheet": 20,
            "filename_custom_part": "个股排序",
        },
        "query": {
            "stock_class_value": "stock",
            "timezone": DEFAULT_TIMEZONE_NAME,
        },
        "tdx": {
            "vipdoc_subdir": "vipdoc",
            "hq_cache_subdir": win_join("T0002", "hq_cache"),
            "market_folders": {
                "sh": "SH",
                "sz": "SZ",
                "bj": "BJ",
            },
            "tnf_files": {
                "SH": "shs.tnf",
                "SZ": "szs.tnf",
                "BJ": "bjs.tnf",
            },
            "lday_subdir": "lday",
            "day_file_ext": ".day",
            "record_size": TDX_DAY_RECORD_SIZE,
            "price_divisor": 100.0,
            "stock_prefixes": {
                "SH": [
                    "689",
                    "688",
                    "605",
                    "603",
                    "601",
                    "600",
                    "900",
                    "360",
                    "330",
                ],
                "SZ": [
                    "30",
                    "004",
                    "003",
                    "002",
                    "001",
                    "000",
                    "20",
                    "140",
                ],
                "BJ": [
                    "92",
                    "88",
                    "87",
                    "83",
                    "820",
                    "420",
                    "400",
                ],
            },
        },
        "fields": {
            "order": [
                "MARKET",
                "SYMBOL",
                "NAME",
                "DATE",
                "OPEN",
                "HIGH",
                "LOW",
                "CLOSE",
                "VOLUME",
                "AMOUNT",
            ],
            "selected": [
                "MARKET",
                "SYMBOL",
                "NAME",
                "DATE",
                "OPEN",
                "HIGH",
                "LOW",
                "CLOSE",
                "VOLUME",
                "AMOUNT",
            ],
        },
        "sort": {
            "rules": [
                {
                    "field": "AMOUNT",
                    "order": "DESC"
                },
            ],
        },
    }


# ============================================================
# 路径工具
# ============================================================


def strip_outer_path_quotes(path):
    """
    去掉用户复制路径时可能带上的外层引号。

    支持：
    - "D:\\TDX_new\\vipdoc"
    - 'D:\\TDX_new\\vipdoc'
    - “D:\\TDX_new\\vipdoc”
    - ‘D:\\TDX_new\\vipdoc’
    """
    text = str(path or "").strip()

    quote_pairs = [
        ('"', '"'),
        ("'", "'"),
        ("“", "”"),
        ("‘", "’"),
    ]

    changed = True

    while changed and len(text) >= 2:
        changed = False

        for left, right in quote_pairs:
            if text.startswith(left) and text.endswith(right):
                text = text[1:-1].strip()
                changed = True
                break

    return text


def normalize_path_slashes(path):
    if path is None:
        return ""

    return str(path).replace("/", "\\")


def is_windows_root_path(path):
    """
    判断是否 Windows 根路径。

    需要保护这些路径，不能把末尾 \\ 去掉：
    - D:\\
    - C:\\
    - \\\\server\\share\\
    """
    path = normalize_path_slashes(path).strip()

    if not path:
        return False

    drive, tail = ntpath.splitdrive(path)

    if drive and tail == "\\":
        return True

    if path.startswith("\\\\"):
        parts = [part for part in path.split("\\") if part]
        return len(parts) == 2

    return False


def strip_redundant_trailing_slashes(path):
    """
    去掉多余的末尾反斜杠。

    示例：
    - D:\\TDX_new\\      -> D:\\TDX_new
    - D:\\TDX_new\\\\    -> D:\\TDX_new
    - D:\\               -> D:\\
    - \\\\server\\share\\ -> \\\\server\\share\\
    """
    text = normalize_path_slashes(path).strip()

    while len(text) > 1 and text.endswith(
            "\\") and not is_windows_root_path(text):
        text = text[:-1]

    return text


def normalize_user_path(path, trim_trailing_slash=True):
    """
    标准化用户输入路径。

    统一处理：
    - 去外层引号
    - 去首尾空白
    - / 转为 \\
    - 可选去掉目录末尾多余反斜杠

    注意：
    - 不会破坏 D:\\ 这种磁盘根目录
    - 不会破坏 \\\\server\\share\\ 这种 UNC 根目录
    """
    text = strip_outer_path_quotes(path)
    text = normalize_path_slashes(text).strip()

    if trim_trailing_slash:
        text = strip_redundant_trailing_slashes(text)

    return text


def win_join(*parts):
    return normalize_path_slashes(ntpath.join(*parts))


def win_dirname(path):
    return normalize_path_slashes(ntpath.dirname(path))


def get_program_dir():
    """
    返回程序所在目录。

    普通 Python 运行：
    - 返回 .py 文件所在目录

    PyInstaller onefile 运行：
    - 返回 exe 文件所在目录
    """
    if getattr(sys, "frozen", False):
        return normalize_path_slashes(os.path.dirname(sys.executable))

    return normalize_path_slashes(os.path.dirname(os.path.abspath(__file__)))


def get_resource_path(filename):
    """
    获取资源文件路径。

    普通运行：
    - 从 .py 所在目录读取

    PyInstaller onefile：
    - 从临时解包目录读取
    """
    if getattr(sys, "frozen", False):
        base_dir = getattr(sys, "_MEIPASS", get_program_dir())
    else:
        base_dir = get_program_dir()

    return normalize_path_slashes(os.path.join(base_dir, filename))


def get_config_path():
    return win_join(get_program_dir(), CONFIG_FILE_NAME)


# ============================================================
# 配置读写
# ============================================================


def deep_merge_known(default_obj, user_obj):
    """
    只合并默认配置中存在的键，忽略旧配置文件里的未知废弃字段。
    """
    result = deepcopy(default_obj)

    if not isinstance(user_obj, dict):
        return result

    for key in result.keys():
        if key not in user_obj:
            continue

        default_value = result[key]
        user_value = user_obj[key]

        if user_value is None:
            continue

        if isinstance(default_value, dict) and isinstance(user_value, dict):
            result[key] = deep_merge_known(default_value, user_value)
        else:
            result[key] = user_value

    return result


def normalize_config_paths(config):
    config["paths"]["local_db"] = normalize_user_path(
        config["paths"].get("local_db", ""),
        trim_trailing_slash=True,
    )

    config["paths"]["tdx_day"] = normalize_user_path(
        config["paths"].get("tdx_day", ""),
        trim_trailing_slash=True,
    )

    config["output"]["output_dir"] = normalize_user_path(
        config["output"].get("output_dir", ""),
        trim_trailing_slash=True,
    )

    return config


def load_app_config():
    default_config = make_default_config()
    config_path = get_config_path()

    if not os.path.exists(config_path):
        return normalize_config_paths(default_config)

    with open(config_path, "r", encoding="utf-8") as f:
        saved_config = json.load(f)

    allowed_saved_config = {
        "data_source": saved_config.get("data_source"),
        "paths": saved_config.get("paths"),
        "output": saved_config.get("output"),
        "fields": saved_config.get("fields"),
        "sort": saved_config.get("sort"),
    }

    config = deep_merge_known(default_config, allowed_saved_config)

    # 日期每次启动动态计算，绝不从配置文件恢复。
    start_date, end_date = default_start_end_dates()
    config["date_range"]["start_date"] = start_date
    config["date_range"]["end_date"] = end_date

    return normalize_config_paths(config)


def is_file_open(filepath):
    """检查文件是否正在被打开"""
    try:
        # 尝试以写入模式打开文件
        with open(filepath, 'a'):
            return False
    except IOError:
        return True


def save_app_config(config):
    """
    只保存用户设置。
    不保存日期、TDX 内部规则、DB 查询规则。
    """
    saved_config = {
        "config_version": 1,
        "data_source": config["data_source"],
        "paths": {
            "local_db": normalize_path_slashes(config["paths"]["local_db"]),
            "tdx_day": normalize_path_slashes(config["paths"]["tdx_day"]),
        },
        "output": {
            "output_dir":
            normalize_path_slashes(config["output"]["output_dir"]),
            "keep_all_rows":
            bool(config["output"]["keep_all_rows"]),
            "max_rows_per_sheet":
            int(config["output"]["max_rows_per_sheet"]),
            "filename_custom_part":
            str(config["output"].get("filename_custom_part", "个股排序")),
        },
        "fields": {
            "order": list(config["fields"]["order"]),
            "selected": list(config["fields"]["selected"]),
        },
        "sort": {
            "rules": deepcopy(config["sort"]["rules"]),
        },
    }

    config_path = get_config_path()

    try:
        # 尝试打开文件进行写入
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(saved_config, f, ensure_ascii=False, indent=2)
    except PermissionError:
        # 文件正在被打开，弹出提示框
        choice = messagebox.askyesnocancel(
            "文件已打开",
            f"配置文件已被打开：\n{config_path}\n请关闭文件后继续保存，或选择另存为。",
        )
        if choice is True:
            # 用户选择确认，尝试再次保存
            save_app_config(config)  # 递归调用
        elif choice is False:
            # 用户选择另存为
            new_path = filedialog.asksaveasfilename(
                title="另存为配置文件",
                defaultextension=".json",
                filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
            )
            if new_path:
                with open(new_path, "w", encoding="utf-8") as f:
                    json.dump(saved_config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        # 捕获其他异常
        print(f"保存配置失败：{e}")  # 错误处理


# ============================================================
# 日期工具
# ============================================================


def normalize_date_text(value):
    """
    将常见日期格式修正为 yyyymmdd。
    """
    raw = str(value).strip()

    if not raw:
        raise ValueError("日期不能为空。")

    raw = (raw.replace("，", ",").replace("。", ".").replace("／", "/").replace(
        "－", "-").replace("—", "-").replace("年",
                                            "-").replace("月",
                                                         "-").replace("日", ""))

    if re.fullmatch(r"\d{8}", raw):
        d = datetime.strptime(raw, "%Y%m%d").date()
        return d.strftime("%Y%m%d")

    if re.fullmatch(r"\d{6}", raw):
        year = 2000 + int(raw[0:2])
        month = int(raw[2:4])
        day = int(raw[4:6])
        d = datetime(year, month, day).date()
        return d.strftime("%Y%m%d")

    parts = re.findall(r"\d+", raw)
    current_year = datetime.now(ZoneInfo(DEFAULT_TIMEZONE_NAME)).year

    if len(parts) == 3:
        year = int(parts[0])
        month = int(parts[1])
        day = int(parts[2])

        if year < 100:
            year = 2000 + year

        d = datetime(year, month, day).date()
        return d.strftime("%Y%m%d")

    if len(parts) == 2:
        month = int(parts[0])
        day = int(parts[1])
        d = datetime(current_year, month, day).date()
        return d.strftime("%Y%m%d")

    raise ValueError(f"无法识别日期格式：{value}")


def parse_yyyymmdd(value):
    normalized = normalize_date_text(value)
    return datetime.strptime(normalized, "%Y%m%d").date()


def format_yyyymmdd(d):
    return d.strftime("%Y%m%d")


def date_range_desc(start_date, end_date):
    current = end_date

    while current >= start_date:
        yield current
        current -= timedelta(days=1)


def day_range_ms(d, timezone_name):
    tz = ZoneInfo(timezone_name)
    start_dt = datetime.combine(d, time.min).replace(tzinfo=tz)
    end_dt = start_dt + timedelta(days=1)

    return int(start_dt.timestamp() * 1000), int(end_dt.timestamp() * 1000)


# ============================================================
# 字段与排序工具
# ============================================================


def field_keys():
    return [item["key"] for item in FIELD_DEFS]


def field_title_map():
    return {item["key"]: item["title"] for item in FIELD_DEFS}


def field_key_by_title(title):
    for item in FIELD_DEFS:
        if item["title"] == title:
            return item["key"]

    raise ValueError(f"未知字段标题：{title}")


def sort_order_title_map():
    return {item["key"]: item["title"] for item in SORT_ORDER_OPTIONS}


def sort_order_key_by_title(title):
    for item in SORT_ORDER_OPTIONS:
        if item["title"] == title:
            return item["key"]

    raise ValueError(f"未知排序方向：{title}")


def sort_dataframe(df, sort_rules):
    if df.empty or not sort_rules:
        return df

    by = [rule["field"] for rule in sort_rules]
    ascending = [rule["order"] == "ASC" for rule in sort_rules]

    return df.sort_values(
        by=by,
        ascending=ascending,
        na_position="last",
        kind="mergesort",
    )


def apply_row_limit(df, config):
    if df.empty:
        return df

    if config["output"]["keep_all_rows"]:
        return df

    return df.head(config["output"]["max_rows_per_sheet"])


def output_dataframe(df, selected_fields):
    out_df = df[selected_fields].copy()
    out_df.rename(columns=field_title_map(), inplace=True)
    return out_df


# ============================================================
# 输出文件工具
# ============================================================


def filename_date_part(start_str, end_str):
    return f"{start_str}-{end_str}"


def split_output_filename(filename):
    """
    拆分输出文件名。

    返回：
    - 日期范围部分：20250101-20250131
    - 自定义部分：日期后面、扩展名前面的内容
    - 扩展名：.xlsx

    如果文件名开头不是标准日期范围，则：
    - 日期范围部分返回空
    - 整个主文件名视为自定义部分
    """
    filename = str(filename or "").strip()
    stem, ext = ntpath.splitext(filename)

    if not ext:
        ext = ".xlsx"

    match = re.match(r"^(\d{8}-\d{8})(.*)$", stem)

    if match:
        return match.group(1), match.group(2), ext

    return "", stem, ext


def build_output_filename(start_str, end_str, custom_part, ext=".xlsx"):
    custom_part = str(custom_part or "").strip()

    if not custom_part:
        custom_part = "个股排序"

    ext = str(ext or ".xlsx").strip()

    if not ext.startswith("."):
        ext = "." + ext

    return f"{filename_date_part(start_str, end_str)}{custom_part}{ext}"


def default_output_path(config):
    start_str = config["date_range"]["start_date"]
    end_str = config["date_range"]["end_date"]

    custom_part = config["output"].get(
        "filename_custom_part",
        "个股排序",
    )

    filename = build_output_filename(
        start_str,
        end_str,
        custom_part,
        ".xlsx",
    )

    output_dir = normalize_user_path(
        config["output"]["output_dir"],
        trim_trailing_slash=True,
    )

    if output_dir:
        return normalize_user_path(
            win_join(output_dir, filename),
            trim_trailing_slash=True,
        )

    return normalize_user_path(
        win_join(get_program_dir(), filename),
        trim_trailing_slash=True,
    )


def make_unique_output_path(path):
    path = normalize_user_path(path, trim_trailing_slash=True)
    folder = win_dirname(path)
    filename = ntpath.basename(path)
    stem, ext = ntpath.splitext(filename)

    if not ext:
        ext = ".xlsx"

    index = 1

    while True:
        candidate = win_join(folder, f"{stem}_{index}{ext}")

        if not os.path.exists(candidate):
            return candidate

        index += 1


def set_sheet_style(writer, sheet_name, df):
    ws = writer.sheets[sheet_name]

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col in enumerate(df.columns, start=1):
        col_letter = ws.cell(row=1, column=idx).column_letter

        if df.empty:
            max_len = len(str(col))
        else:
            sample_values = df[col].astype(str).head(500)
            max_len = max(len(str(col)), sample_values.map(len).max())

        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 28)


# ============================================================
# 配置校验
# ============================================================


def validate_runtime_config(config, output_path):
    output_path = normalize_user_path(output_path, trim_trailing_slash=True)

    source = config["data_source"]
    valid_sources = [item["key"] for item in DATA_SOURCE_OPTIONS]

    if source not in valid_sources:
        raise ValueError(f"未知数据源：{source}")

    source_path = normalize_user_path(
        config["paths"][source],
        trim_trailing_slash=True,
    )

    if not source_path:
        raise ValueError(f"当前数据源路径缺失：{source}")

    if source == "local_db" and not os.path.isfile(source_path):
        raise FileNotFoundError(f"本地 DB 文件不存在：{source_path}")

    if source == "tdx_day":
        if not os.path.isdir(source_path):
            raise FileNotFoundError(f"通达信数据目录不存在：{source_path}")

        vipdoc_dir = resolve_tdx_vipdoc_dir(config)
        hq_cache_dir = resolve_tdx_hq_cache_dir(config)

        if not os.path.isdir(vipdoc_dir):
            raise FileNotFoundError("通达信 vipdoc 目录不存在：\n"
                                    f"{vipdoc_dir}\n\n"
                                    "请填写通达信根目录，例如 D:\\TDX_new，"
                                    "或直接填写 vipdoc 目录，例如 D:\\TDX_new\\vipdoc。")

        if not os.path.isdir(hq_cache_dir):
            raise FileNotFoundError("通达信名称目录不存在：\n"
                                    f"{hq_cache_dir}\n\n"
                                    "请确认存在 T0002\\hq_cache 目录。")

        for market, filename in config["tdx"]["tnf_files"].items():
            tnf_path = tdx_tnf_file_path(config, market)

            if not os.path.isfile(tnf_path):
                raise FileNotFoundError(f"通达信 {market} 名称文件不存在：\n{tnf_path}")

    start_date = parse_yyyymmdd(config["date_range"]["start_date"])
    end_date = parse_yyyymmdd(config["date_range"]["end_date"])

    if start_date > end_date:
        raise ValueError("开始日期不能晚于结束日期。")

    if not output_path:
        raise ValueError("输出文件路径缺失。")

    output_dir = win_dirname(output_path)

    if output_dir and not os.path.isdir(output_dir):
        raise FileNotFoundError(f"输出目录不存在：{output_dir}")

    if not output_path.lower().endswith(".xlsx"):
        raise ValueError("输出文件必须是 .xlsx 格式。")

    if not isinstance(config["output"]["keep_all_rows"], bool):
        raise ValueError("配置错误：output.keep_all_rows 必须是布尔值。")

    max_rows = config["output"]["max_rows_per_sheet"]

    if not isinstance(max_rows, int) or max_rows <= 0:
        raise ValueError("配置错误：output.max_rows_per_sheet 必须是正整数。")

    valid_fields = field_keys()
    valid_field_set = set(valid_fields)

    field_order = config["fields"]["order"]
    selected_fields = config["fields"]["selected"]

    if not field_order:
        raise ValueError("字段顺序配置缺失。")

    if not selected_fields:
        raise ValueError("至少需要选择一个输出字段。")

    if len(field_order) != len(valid_fields):
        raise ValueError("字段顺序配置存在缺失或重复。")

    unknown_order_fields = set(field_order) - valid_field_set
    if unknown_order_fields:
        raise ValueError(f"字段顺序中包含未知字段：{sorted(unknown_order_fields)}")

    missing_order_fields = valid_field_set - set(field_order)
    if missing_order_fields:
        raise ValueError(f"字段顺序中缺少字段：{sorted(missing_order_fields)}")

    unknown_selected_fields = set(selected_fields) - valid_field_set
    if unknown_selected_fields:
        raise ValueError(f"已选字段中包含未知字段：{sorted(unknown_selected_fields)}")

    for rule in config["sort"]["rules"]:
        if rule["field"] not in valid_fields:
            raise ValueError(f"未知排序字段：{rule['field']}")

        if rule["order"] not in ["ASC", "DESC"]:
            raise ValueError(f"未知排序方向：{rule['order']}")


# ============================================================
# 本地 DB 数据源
# ============================================================


def build_local_db_sql():
    select_parts = []

    for item in FIELD_DEFS:
        key = item["key"]

        if key == "DATE":
            select_parts.append('? AS "DATE"')
        else:
            select_parts.append(f'{item["db_expr"]} AS "{key}"')

    select_sql = ",\n        ".join(select_parts)

    return f"""
    SELECT
        {select_sql}
    FROM symbol_index AS s
    JOIN candles_day_raw AS c
      ON c.market = s.market
     AND c.symbol = s.symbol
    WHERE s."class" = ?
      AND c.ts >= ?
      AND c.ts < ?
    """


def read_local_db_one_day(conn, config, current_date):
    yyyymmdd = format_yyyymmdd(current_date)
    start_ms, end_ms = day_range_ms(current_date, config["query"]["timezone"])

    return pd.read_sql_query(
        build_local_db_sql(),
        conn,
        params=(
            yyyymmdd,
            config["query"]["stock_class_value"],
            start_ms,
            end_ms,
        ),
    )


# ============================================================
# 通达信 TDX 数据源
# ============================================================


def resolve_tdx_root_dir(config):
    """
    解析通达信根目录。

    支持用户输入：
    - D:\\TDX_new
    - D:\\TDX_new\\
    - "D:\\TDX_new"
    - "D:\\TDX_new\\"
    - D:\\TDX_new\\vipdoc
    - D:\\TDX_new\\vipdoc\\
    - "D:\\TDX_new\\vipdoc"
    """
    input_path = normalize_user_path(
        config["paths"]["tdx_day"],
        trim_trailing_slash=True,
    )

    vipdoc_name = config["tdx"]["vipdoc_subdir"]

    if ntpath.basename(input_path).lower() == vipdoc_name.lower():
        return win_dirname(input_path)

    return input_path


def resolve_tdx_vipdoc_dir(config):
    """
    返回 vipdoc 目录。
    """
    return win_join(
        resolve_tdx_root_dir(config),
        config["tdx"]["vipdoc_subdir"],
    )


def resolve_tdx_hq_cache_dir(config):
    """
    返回 T0002\\hq_cache 目录。
    """
    return win_join(
        resolve_tdx_root_dir(config),
        config["tdx"]["hq_cache_subdir"],
    )


def tdx_tnf_file_path(config, market):
    """
    返回指定市场的 .tnf 文件路径。
    """
    market = str(market or "").strip().upper()
    filename = config["tdx"]["tnf_files"].get(market)

    if not filename:
        raise ValueError(f"不支持的 TDX TNF 市场：{market}")

    return win_join(resolve_tdx_hq_cache_dir(config), filename)


def decode_tdx_ascii(raw):
    try:
        return raw.decode("ascii", errors="ignore").strip("\x00 ").strip()
    except Exception:
        return ""


def decode_tdx_gbk(raw):
    try:
        return raw.decode("gbk", errors="ignore").strip("\x00 ").strip()
    except Exception:
        return ""


def load_tdx_tnf_name_map_for_market(config, market):
    """
    解析单个市场的 .tnf 文件，只提取当前工具需要的名称映射。

    返回：
    {
        "600000": "浦发银行",
        ...
    }

    只保留必要字段：
    - symbol
    - name
    """
    market = str(market or "").strip().upper()
    path = tdx_tnf_file_path(config, market)

    if not os.path.exists(path):
        raise FileNotFoundError(f"TDX 名称文件不存在：{path}")

    with open(path, "rb") as f:
        raw = f.read()

    if len(raw) <= TDX_TNF_HEADER_SIZE:
        return {}

    payload = raw[TDX_TNF_HEADER_SIZE:]
    total = len(payload) // TDX_TNF_RECORD_SIZE

    name_map = {}

    for idx in range(total):
        start = idx * TDX_TNF_RECORD_SIZE
        rec = payload[start:start + TDX_TNF_RECORD_SIZE]

        if len(rec) < TDX_TNF_RECORD_SIZE:
            continue

        symbol = decode_tdx_ascii(rec[0:20])
        name = decode_tdx_gbk(rec[31:63])

        if symbol:
            name_map[symbol] = name

    return name_map


def load_tdx_name_map(config, log_func):
    """
    加载 TDX 三个市场的标的名称映射。

    返回：
    {
        ("SH", "600000"): "浦发银行",
        ("SZ", "000001"): "平安银行",
        ("BJ", "830799"): "...",
    }
    """
    result = {}

    for market in config["tdx"]["tnf_files"].keys():
        market_name_map = load_tdx_tnf_name_map_for_market(config, market)

        for symbol, name in market_name_map.items():
            result[(market, symbol)] = name

        log_func(f"TDX 名称文件解析完成：{market}，名称数量：{len(market_name_map)}")

    return result


def tdx_day_files(config):
    vipdoc_dir = resolve_tdx_vipdoc_dir(config)
    files = []

    for market_folder in config["tdx"]["market_folders"].keys():
        folder = win_join(
            vipdoc_dir,
            market_folder,
            config["tdx"]["lday_subdir"],
        )

        pattern = win_join(folder, f"*{config['tdx']['day_file_ext']}")
        files.extend(glob.glob(pattern))

    return files


def parse_tdx_market_and_symbol(file_path, config):
    filename = ntpath.basename(file_path)
    stem = ntpath.splitext(filename)[0].lower()

    for market_folder, market in config["tdx"]["market_folders"].items():
        if stem.startswith(market_folder):
            symbol = stem[len(market_folder):]
            return market, symbol

    raise ValueError(f"无法识别通达信文件市场和代码：{file_path}")


def is_tdx_stock_symbol(market, symbol, config):
    market = str(market or "").strip().upper()
    symbol = str(symbol or "").strip()

    prefixes = config["tdx"]["stock_prefixes"].get(market, [])
    prefixes = sorted(prefixes, key=len, reverse=True)

    return any(symbol.startswith(prefix) for prefix in prefixes)


def parse_tdx_day_records_in_range(file_path, config, start_int, end_int):
    """
    倒序解析单个通达信 .day 文件，只读取用户日期范围内的数据。

    依赖 TDX .day 常规结构：
    - 记录按交易日期升序存储
    - 越靠近文件尾部日期越新
    """
    path = normalize_user_path(file_path, trim_trailing_slash=True)

    if not os.path.exists(path):
        raise FileNotFoundError(f".day 文件不存在：{path}")

    record_size = int(config["tdx"]["record_size"])
    price_divisor = float(config["tdx"]["price_divisor"])

    if record_size != TDX_DAY_RECORD_SIZE:
        raise ValueError(f"TDX .day 记录长度必须是 {TDX_DAY_RECORD_SIZE} 字节。")

    file_size = os.path.getsize(path)

    if file_size == 0:
        return []

    if file_size % record_size != 0:
        raise ValueError(
            f".day 文件大小非法：{path}，bytes={file_size}，不能被 {record_size} 整除")

    total_records = file_size // record_size
    rows = []
    seen_dates = set()

    with open(path, "rb") as f:
        for idx in range(total_records - 1, -1, -1):
            f.seek(idx * record_size)
            rec = f.read(record_size)

            if len(rec) != record_size:
                raise ValueError(f".day 记录读取失败：file={path}，idx={idx}，"
                                 f"期望字节={record_size}，实际字节={len(rec)}")

            try:
                trade_date, open_i, high_i, low_i, close_i, amount_f, volume_i, _ = struct.unpack(
                    "<IIIIIfII",
                    rec,
                )
            except Exception as e:
                raise ValueError(
                    f".day 记录解析失败：file={path}，idx={idx}，原因={e}") from e

            trade_date = int(trade_date)

            if not (19000101 <= trade_date <= 21001231):
                raise ValueError(
                    f".day 文件存在非法交易日期：file={path}，idx={idx}，date={trade_date}")

            if trade_date < start_int:
                break

            if trade_date > end_int:
                continue

            yyyymmdd = str(trade_date)

            if yyyymmdd in seen_dates:
                continue

            seen_dates.add(yyyymmdd)

            rows.append({
                "DATE": yyyymmdd,
                "OPEN": float(open_i) / price_divisor,
                "HIGH": float(high_i) / price_divisor,
                "LOW": float(low_i) / price_divisor,
                "CLOSE": float(close_i) / price_divisor,
                "VOLUME": int(volume_i),
                "AMOUNT": float(amount_f),
            })

    return rows


def read_tdx_range(config, start_date, end_date, log_func):
    """
    读取 TDX 指定日期范围内的个股日线数据。

    数据来源：
    - .day：行情数据
    - .tnf：标的名称
    """
    start_int = int(format_yyyymmdd(start_date))
    end_int = int(format_yyyymmdd(end_date))

    vipdoc_dir = resolve_tdx_vipdoc_dir(config)
    hq_cache_dir = resolve_tdx_hq_cache_dir(config)

    log_func(f"通达信 vipdoc 目录：{vipdoc_dir}")
    log_func(f"通达信名称目录：{hq_cache_dir}")

    name_map = load_tdx_name_map(config, log_func)

    files = tdx_day_files(config)
    result = {}

    scanned_count = 0
    stock_file_count = 0
    effective_file_count = 0
    matched_row_count = 0
    skipped_file_count = 0

    log_func(f"扫描通达信 .day 文件数量：{len(files)}")

    for file_path in files:
        scanned_count += 1

        try:
            market, symbol = parse_tdx_market_and_symbol(file_path, config)
        except Exception as e:
            skipped_file_count += 1
            log_func(f"识别文件失败，跳过：{file_path}，原因：{e}")
            continue

        if not is_tdx_stock_symbol(market, symbol, config):
            continue

        stock_file_count += 1

        try:
            rows = parse_tdx_day_records_in_range(
                file_path,
                config,
                start_int,
                end_int,
            )
        except Exception as e:
            skipped_file_count += 1
            log_func(f"解析 .day 失败，跳过：{file_path}，原因：{e}")
            continue

        file_has_rows = False
        symbol_name = name_map.get((market, symbol), "")

        for row in rows:
            yyyymmdd = row["DATE"]

            result.setdefault(yyyymmdd, []).append({
                "MARKET": market,
                "SYMBOL": symbol,
                "NAME": symbol_name,
                "DATE": yyyymmdd,
                "OPEN": row["OPEN"],
                "HIGH": row["HIGH"],
                "LOW": row["LOW"],
                "CLOSE": row["CLOSE"],
                "VOLUME": row["VOLUME"],
                "AMOUNT": row["AMOUNT"],
            })

            matched_row_count += 1
            file_has_rows = True

        if file_has_rows:
            effective_file_count += 1

        if scanned_count % 500 == 0:
            log_func(f"已扫描文件：{scanned_count}/{len(files)}，"
                     f"股票文件：{stock_file_count}，"
                     f"有效文件：{effective_file_count}，"
                     f"匹配记录：{matched_row_count}")

    log_func(f"通达信扫描完成：总文件 {scanned_count}，"
             f"股票文件 {stock_file_count}，"
             f"有效文件 {effective_file_count}，"
             f"匹配记录 {matched_row_count}，"
             f"异常跳过 {skipped_file_count}")

    return result


# ============================================================
# 导出服务
# ============================================================


def export_to_excel(config, output_path, log_func=print):
    validate_runtime_config(config, output_path)

    start_date = parse_yyyymmdd(config["date_range"]["start_date"])
    end_date = parse_yyyymmdd(config["date_range"]["end_date"])

    selected_fields = config["fields"]["selected"]
    sort_rules = config["sort"]["rules"]
    source = config["data_source"]

    log_func(f"数据源：{source}")
    log_func(f"开始日期：{format_yyyymmdd(start_date)}")
    log_func(f"结束日期：{format_yyyymmdd(end_date)}")
    log_func(f"输出文件：{output_path}")

    wrote_sheet_count = 0
    total_rows = 0
    conn = None
    tdx_cache = None

    # 用于统计名次出现次数的字典
    rank_count = {}
    max_ranks = min(config["output"]["max_rows_per_sheet"], 30)

    try:
        if source == "local_db":
            conn = sqlite3.connect(config["paths"]["local_db"])

        if source == "tdx_day":
            tdx_cache = read_tdx_range(config, start_date, end_date, log_func)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # 输出汇总表
            for current_date in date_range_desc(start_date, end_date):
                yyyymmdd = format_yyyymmdd(current_date)

                log_func(f"处理日期：{yyyymmdd}")

                if source == "local_db":
                    df = read_local_db_one_day(conn, config, current_date)
                elif source == "tdx_day":
                    df = pd.DataFrame(tdx_cache.get(yyyymmdd, []),
                                      columns=field_keys())
                else:
                    raise ValueError(f"未知数据源：{source}")

                if df.empty:
                    log_func(f"  无数据，跳过 Sheet：{yyyymmdd}")
                    continue

                # 统计名次
                df = sort_dataframe(df, sort_rules)
                for rank in range(1, max_ranks + 1):
                    ranked_df = df.head(rank)  # 取前rank名次的个股
                    for index, row in ranked_df.iterrows():
                        symbol = row['SYMBOL']
                        if symbol not in rank_count:
                            rank_count[symbol] = {
                                "NAME": row['NAME'],
                                "RANKS": [0] * max_ranks  # 初始化名次计数
                            }
                        rank_count[symbol]["RANKS"][rank - 1] += 1  # 增加该名次的计数

            # 汇总表数据准备
            summary_data = []
            for symbol, data in rank_count.items():
                summary_data.append([symbol, data["NAME"]] + data["RANKS"])

            # 创建汇总 DataFrame
            summary_df = pd.DataFrame(
                summary_data,
                columns=["标的代码", "标的名称"] +
                [str(i) for i in range(1, max_ranks + 1)])
            summary_df.sort_values(
                by=[f"{i}" for i in range(1, max_ranks + 1)],
                ascending=False,
                inplace=True)

            # 输出汇总表
            summary_df.to_excel(writer, sheet_name="汇总", index=False)
            set_sheet_style(writer, "汇总", summary_df)

            # 设置列宽
            worksheet = writer.sheets["汇总"]
            worksheet.column_dimensions['A'].width = 10  # 标的代码列宽
            worksheet.column_dimensions['B'].width = 10  # 标的名称列宽
            for i in range(2, max_ranks + 2):  # 从第3列开始设置宽度
                worksheet.column_dimensions[chr(65 + i)].width = 4  # 排名列宽

            # 继续写入每日数据
            for current_date in date_range_desc(start_date, end_date):
                yyyymmdd = format_yyyymmdd(current_date)

                log_func(f"处理日期：{yyyymmdd}")

                if source == "local_db":
                    df = read_local_db_one_day(conn, config, current_date)
                elif source == "tdx_day":
                    df = pd.DataFrame(tdx_cache.get(yyyymmdd, []),
                                      columns=field_keys())
                else:
                    raise ValueError(f"未知数据源：{source}")

                if df.empty:
                    log_func(f"  无数据，跳过 Sheet：{yyyymmdd}")
                    continue

                df = sort_dataframe(df, sort_rules)
                df = apply_row_limit(df, config)
                out_df = output_dataframe(df, selected_fields)

                out_df.to_excel(writer, sheet_name=yyyymmdd, index=False)
                set_sheet_style(writer, yyyymmdd, out_df)

                wrote_sheet_count += 1
                total_rows += len(out_df)

        log_func("-" * 50)
        log_func("导出完成。")
        log_func(f"Sheet 数量：{wrote_sheet_count}")
        log_func(f"总行数：{total_rows}")
        log_func(f"文件位置：{output_path}")

    finally:
        if conn is not None:
            conn.close()


# ============================================================
# 横向字段选择组件
# ============================================================


class HorizontalDraggableFieldCheckList(tk.Frame):

    def __init__(self, parent, field_defs, order_keys, selected_keys):
        super().__init__(parent)

        self.field_defs = field_defs
        self.field_map = {item["key"]: item for item in field_defs}
        self.valid_keys = [item["key"] for item in field_defs]

        self.order_keys = list(order_keys)
        self.selected_keys = set(selected_keys)

        self.items = []
        self.drag_key = None

        self.normal_bg = "#ffffff"
        self.selected_bg = "#dcfce7"
        self.drag_bg = "#dbeafe"
        self.drop_bg = "#fff7cc"
        self.handle_normal_bg = "#eeeeee"
        self.handle_selected_bg = "#bbf7d0"

        self.validate_initial_state()

        self.vars = {
            key: tk.BooleanVar(value=key in self.selected_keys)
            for key in self.valid_keys
        }

        self.build_toolbar()
        self.build_scroll_area()

    def validate_initial_state(self):
        valid_set = set(self.valid_keys)
        order_set = set(self.order_keys)

        if len(self.order_keys) != len(self.valid_keys):
            raise ValueError("字段顺序配置存在缺失或重复。")

        unknown_order = order_set - valid_set
        if unknown_order:
            raise ValueError(f"字段顺序中包含未知字段：{sorted(unknown_order)}")

        missing_order = valid_set - order_set
        if missing_order:
            raise ValueError(f"字段顺序中缺少字段：{sorted(missing_order)}")

        unknown_selected = set(self.selected_keys) - valid_set
        if unknown_selected:
            raise ValueError(f"已选字段中包含未知字段：{sorted(unknown_selected)}")

    def build_toolbar(self):
        button_width = 12

        toolbar = tk.Frame(self)
        toolbar.pack(fill="x", padx=4, pady=(2, 0))

        tk.Label(
            toolbar,
            text="字段从左到右即为 Excel 输出列顺序；勾选输出，拖动右侧 ☰ 调整顺序",
            fg="#555555",
        ).pack(side="left", anchor="center")

        button_frame = tk.Frame(toolbar)
        button_frame.pack(side="right", anchor="e")

        tk.Button(
            button_frame,
            text="全选",
            width=button_width,
            command=self.select_all,
        ).pack(side="left", padx=(0, 4))

        tk.Button(
            button_frame,
            text="全不选",
            width=button_width,
            command=self.unselect_all,
        ).pack(side="left", padx=(4, 0))

    def build_scroll_area(self):
        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(container, height=40, highlightthickness=0)

        self.h_scrollbar = tk.Scrollbar(
            container,
            orient="horizontal",
            command=self.canvas.xview,
        )

        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        self.canvas.pack(side="top", fill="x", expand=False)
        self.h_scrollbar.pack(side="bottom", fill="x")

        self.list_frame = tk.Frame(self.canvas)

        self.window_id = self.canvas.create_window(
            (0, 0),
            window=self.list_frame,
            anchor="nw",
        )

        self.list_frame.bind("<Configure>", self.on_list_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        self.canvas.bind("<Enter>", self.bind_mousewheel)
        self.canvas.bind("<Leave>", self.unbind_mousewheel)

        self.render_items()

    def bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

    def unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def on_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_list_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window_id, height=event.height)

    def render_items(self):
        for child in self.list_frame.winfo_children():
            child.destroy()

        self.items = []

        for key in self.order_keys:
            title = self.field_map[key]["title"]

            item_frame = tk.Frame(
                self.list_frame,
                bd=1,
                relief="solid",
                bg=self.normal_bg,
            )
            item_frame.pack(side="left", fill="y", padx=2, pady=3)

            handle = tk.Label(
                item_frame,
                text="☰",
                width=2,
                cursor="fleur",
                bg=self.handle_normal_bg,
                fg="#333333",
            )
            handle.pack(side="right", fill="y", padx=(1, 2), pady=1)

            cb = tk.Checkbutton(
                item_frame,
                text=title,
                variable=self.vars[key],
                anchor="w",
                width=10,
                bg=self.normal_bg,
                activebackground=self.selected_bg,
                selectcolor=self.selected_bg,
                command=self.refresh_item_colors,
            )
            cb.pack(side="left", fill="both", expand=True, padx=(3, 1), pady=1)

            handle.bind("<ButtonPress-1>",
                        lambda event, k=key: self.on_drag_start(event, k))
            handle.bind("<B1-Motion>",
                        lambda event, k=key: self.on_drag_motion(event, k))
            handle.bind("<ButtonRelease-1>",
                        lambda event, k=key: self.on_drag_release(event, k))

            self.items.append({
                "key": key,
                "widget": item_frame,
                "checkbox": cb,
                "handle": handle,
            })

        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.refresh_item_colors()

    def item_base_color(self, key):
        if self.vars[key].get():
            return self.selected_bg

        return self.normal_bg

    def item_handle_color(self, key):
        if self.vars[key].get():
            return self.handle_selected_bg

        return self.handle_normal_bg

    def refresh_item_colors(self):
        for item in self.items:
            key = item["key"]
            base_color = self.item_base_color(key)

            item["widget"].configure(bg=base_color)
            item["checkbox"].configure(
                bg=base_color,
                activebackground=base_color,
                selectcolor=base_color,
            )
            item["handle"].configure(bg=self.item_handle_color(key))

    def on_drag_start(self, event, key):
        self.drag_key = key
        self.highlight_drag_item(key)

    def on_drag_motion(self, event, key):
        self.highlight_drop_position()

    def on_drag_release(self, event, key):
        if self.drag_key is None:
            return

        drop_index = self.calculate_drop_index()
        old_index = self.order_keys.index(self.drag_key)

        if drop_index > old_index:
            drop_index -= 1

        moved_key = self.order_keys.pop(old_index)
        self.order_keys.insert(drop_index, moved_key)

        self.drag_key = None
        self.render_items()

    def calculate_drop_index(self):
        pointer_x = self.winfo_pointerx()

        for idx, item in enumerate(self.items):
            widget = item["widget"]
            middle = widget.winfo_rootx() + widget.winfo_width() / 2

            if pointer_x < middle:
                return idx

        return len(self.items)

    def set_item_color(self, item, color, force_handle=False):
        item["widget"].configure(bg=color)
        item["checkbox"].configure(
            bg=color,
            activebackground=color,
            selectcolor=color,
        )

        if force_handle:
            item["handle"].configure(bg=color)
        else:
            item["handle"].configure(bg=self.item_handle_color(item["key"]))

    def highlight_drag_item(self, key):
        for item in self.items:
            if item["key"] == key:
                self.set_item_color(item, self.drag_bg, force_handle=True)
            else:
                self.set_item_color(item, self.item_base_color(item["key"]))

    def highlight_drop_position(self):
        drop_index = self.calculate_drop_index()

        for idx, item in enumerate(self.items):
            if idx == drop_index:
                self.set_item_color(item, self.drop_bg, force_handle=True)
            elif item["key"] == self.drag_key:
                self.set_item_color(item, self.drag_bg, force_handle=True)
            else:
                self.set_item_color(item, self.item_base_color(item["key"]))

    def select_all(self):
        for key in self.valid_keys:
            self.vars[key].set(True)

        self.refresh_item_colors()

    def unselect_all(self):
        for key in self.valid_keys:
            self.vars[key].set(False)

        self.refresh_item_colors()

    def select_keys(self, keys):
        for key in keys:
            if key in self.vars:
                self.vars[key].set(True)

        self.refresh_item_colors()

    def get_ordered_keys(self):
        return list(self.order_keys)

    def get_selected_keys_in_order(self):
        return [key for key in self.order_keys if self.vars[key].get()]


# ============================================================
# GUI 应用
# ============================================================


class ExportApp:

    def __init__(self, root):
        self.root = root
        self.config = load_app_config()

        self.root.title(self.config["ui"]["window_title"])
        self.set_window_icon()
        self.root.geometry(self.config["ui"]["window_geometry"])
        self.root.state("zoomed")

        self.log_queue = queue.Queue()
        self.sort_rules = deepcopy(self.config["sort"]["rules"])
        self.output_path_auto_editing = False
        self.allow_next_date_auto_update = True

        self.build_variables()
        self.build_ui()
        self.refresh_sort_listbox()
        self.update_output_path_by_dates()

        self.root.after(100, self.poll_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_variables(self):
        self.data_source_var = tk.StringVar(value="tdx_day")  # 设置默认为通达信盘后数据

        self.path_vars = {}

        for item in DATA_SOURCE_OPTIONS:
            key = item["key"]
            self.path_vars[key] = tk.StringVar(value=normalize_user_path(
                self.config["paths"][key],
                trim_trailing_slash=True,
            ))
            self.path_vars[key].trace_add(
                "write", lambda *_: self.save_current_config())  # 添加保存配置的监听

        self.start_date_var = tk.StringVar(
            value=self.config["date_range"]["start_date"])
        self.end_date_var = tk.StringVar(
            value=self.config["date_range"]["end_date"])
        self.start_date_var.trace_add(
            "write", lambda *_: self.save_current_config())  # 添加保存配置的监听
        self.end_date_var.trace_add(
            "write", lambda *_: self.save_current_config())  # 添加保存配置的监听

        self.output_path_var = tk.StringVar(value=normalize_user_path(
            default_output_path(self.config),
            trim_trailing_slash=True,
        ))

        self.output_path_var.trace_add(
            "write", lambda *_: self.on_output_path_changed())

        self.keep_all_rows_var = tk.BooleanVar(
            value=self.config["output"]["keep_all_rows"])
        self.keep_all_rows_var.trace_add(
            "write", lambda *_: self.save_current_config())  # 添加保存配置的监听

        self.max_rows_per_sheet_var = tk.StringVar(
            value=str(self.config["output"]["max_rows_per_sheet"]))
        self.max_rows_per_sheet_var.trace_add(
            "write", lambda *_: self.save_current_config())  # 添加保存配置的监听

        self.sort_field_title_var = tk.StringVar(value="成交额")
        self.sort_order_title_var = tk.StringVar(value="降序")

        # 新增：自动打开输出目录的复选框
        self.auto_open_var = tk.BooleanVar(value=True)

    def build_ui(self):
        main = tk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=14, pady=10)

        title_frame = tk.Frame(main)
        title_frame.pack(fill="x", pady=(0, 8))

        tk.Label(
            title_frame,
            text="个股排序导出为多页Excel表——汇总表及每日明细表",
            font=("Microsoft YaHei", 15, "bold"),
        ).pack(side="left", anchor="w")

        tk.Button(
            title_frame,
            text="反馈交流",
            width=10,
            command=self.show_feedback_qr,
        ).pack(side="right")

        self.build_source_frame(main)
        self.build_date_output_frame(main)
        self.build_field_frame(main)
        self.build_sort_log_frame(main)

    def build_source_frame(self, parent):
        frame = tk.LabelFrame(parent, text="数据源")
        frame.pack(fill="x", pady=5)

        button_width = 12  # 统一按钮宽度

        for row, item in enumerate(DATA_SOURCE_OPTIONS):
            key = item["key"]

            tk.Radiobutton(
                frame,
                text=item["title"],
                variable=self.data_source_var,
                value=key,
            ).grid(row=row, column=0, sticky="w", padx=10, pady=5)

            tk.Entry(
                frame,
                textvariable=self.path_vars[key],
                width=90,  # 调整宽度
            ).grid(row=row, column=1, sticky="we", padx=8, pady=5)

            tk.Button(
                frame,
                text="选择目录",
                command=lambda k=key: self.choose_source_path(k),
                width=button_width,
            ).grid(row=row, column=2, sticky="w", padx=(8, 4), pady=5)

            tk.Button(
                frame,
                text="打开目录",
                command=lambda k=key: self.open_source_folder(k),
                width=button_width,
            ).grid(row=row, column=3, sticky="w", padx=(4, 8), pady=5)

        frame.columnconfigure(1, weight=1)

    def build_date_output_frame(self, parent):
        frame = tk.LabelFrame(parent, text="日期与输出")
        frame.pack(fill="x", pady=5)

        button_width = 12  # 统一按钮宽度

        tk.Label(frame, text="开始日期：").grid(row=0,
                                           column=0,
                                           sticky="e",
                                           padx=8,
                                           pady=5)

        self.start_date_entry = tk.Entry(frame,
                                         textvariable=self.start_date_var,
                                         width=14)
        self.start_date_entry.grid(row=0, column=1, sticky="w", padx=6, pady=5)
        self.start_date_entry.bind(
            "<FocusOut>",
            lambda event: self.normalize_date_entry(self.start_date_var))

        tk.Label(frame, text="结束日期：").grid(row=0,
                                           column=2,
                                           sticky="e",
                                           padx=8,
                                           pady=5)

        self.end_date_entry = tk.Entry(frame,
                                       textvariable=self.end_date_var,
                                       width=14)
        self.end_date_entry.grid(row=0, column=3, sticky="w", padx=6, pady=5)
        self.end_date_entry.bind(
            "<FocusOut>",
            lambda event: self.normalize_date_entry(self.end_date_var))

        frame.columnconfigure(4, weight=1)

        tk.Label(frame, text="每页保留标的数：").grid(row=0,
                                              column=5,
                                              sticky="e",
                                              padx=(8, 4),
                                              pady=5)

        vcmd = (self.root.register(self.validate_positive_integer_input), "%P")

        self.max_rows_entry = tk.Entry(
            frame,
            textvariable=self.max_rows_per_sheet_var,
            width=10,
            validate="key",
            validatecommand=vcmd)
        self.max_rows_entry.grid(row=0,
                                 column=6,
                                 sticky="e",
                                 padx=(4, 6),
                                 pady=5)

        tk.Checkbutton(frame,
                       text="保留全部标的",
                       variable=self.keep_all_rows_var,
                       command=self.toggle_max_rows_entry).grid(row=0,
                                                                column=7,
                                                                sticky="w",
                                                                padx=(0, 4),
                                                                pady=5)

        tk.Checkbutton(frame, text="导出完成后自动打开",
                       variable=self.auto_open_var).grid(row=0,
                                                         column=8,
                                                         columnspan=3,
                                                         sticky="e",
                                                         padx=(8, 8),
                                                         pady=5)

        tk.Label(frame, text="输出文件：").grid(row=1,
                                           column=0,
                                           sticky="e",
                                           padx=8,
                                           pady=5)

        self.output_path_entry = tk.Entry(frame,
                                          textvariable=self.output_path_var,
                                          width=90)  # 调整宽度
        self.output_path_entry.grid(row=1,
                                    column=1,
                                    columnspan=6,
                                    sticky="we",
                                    padx=6,
                                    pady=5)

        tk.Button(frame,
                  text="另存为",
                  command=self.choose_output_file,
                  width=button_width).grid(row=1,
                                           column=7,
                                           sticky="w",
                                           padx=(4, 4),
                                           pady=5)
        tk.Button(frame,
                  text="选择目录",
                  command=self.choose_output_dir,
                  width=button_width).grid(row=1,
                                           column=8,
                                           sticky="w",
                                           padx=(6, 4),
                                           pady=5)
        tk.Button(frame,
                  text="打开目录",
                  command=self.open_output_folder,
                  width=button_width).grid(row=1,
                                           column=9,
                                           sticky="w",
                                           padx=(4, 8),
                                           pady=5)

        self.start_date_var.trace_add("write",
                                      lambda *_: self.on_date_range_changed())
        self.end_date_var.trace_add("write",
                                    lambda *_: self.on_date_range_changed())

        self.toggle_max_rows_entry()

    def build_field_frame(self, parent):
        field_frame = tk.LabelFrame(parent, text="输出字段，横向顺序即为 Excel 列顺序")
        field_frame.pack(fill="x", pady=5)

        self.field_checklist = HorizontalDraggableFieldCheckList(
            field_frame,
            field_defs=FIELD_DEFS,
            order_keys=self.config["fields"]["order"],
            selected_keys=self.config["fields"]["selected"],
        )
        self.field_checklist.pack(fill="x", expand=False, padx=6, pady=2)

    def build_sort_log_frame(self, parent):
        container = tk.Frame(parent)
        container.pack(fill="both", expand=True, pady=5)

        container.columnconfigure(0, weight=4, uniform="sortlog")
        container.columnconfigure(1, weight=20, uniform="sortlog")
        container.rowconfigure(0, weight=1)

        self.build_sort_frame(container)
        self.build_log_frame(container)

    def build_sort_frame(self, parent):
        sort_frame = tk.LabelFrame(parent, text="排序条件")
        sort_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        sort_frame.columnconfigure(0, weight=1)
        sort_frame.rowconfigure(3, weight=1)

        control_frame = tk.Frame(sort_frame)
        control_frame.grid(row=0, column=0, sticky="we", padx=6, pady=(6, 3))
        control_frame.columnconfigure(1, weight=1)

        tk.Label(control_frame, text="字段：").grid(row=0,
                                                 column=0,
                                                 sticky="e",
                                                 padx=(0, 4),
                                                 pady=3)

        self.sort_field_combo = ttk.Combobox(
            control_frame,
            textvariable=self.sort_field_title_var,
            values=[item["title"] for item in FIELD_DEFS],
            width=12,
            state="readonly",
        )
        self.sort_field_combo.grid(row=0, column=1, sticky="we", pady=3)

        tk.Label(control_frame, text="方向：").grid(row=1,
                                                 column=0,
                                                 sticky="e",
                                                 padx=(0, 4),
                                                 pady=3)

        self.sort_order_combo = ttk.Combobox(
            control_frame,
            textvariable=self.sort_order_title_var,
            values=[item["title"] for item in SORT_ORDER_OPTIONS],
            width=8,
            state="readonly",
        )
        self.sort_order_combo.grid(row=1, column=1, sticky="we", pady=3)

        tk.Button(
            sort_frame,
            text="添加排序条件",
            command=self.add_sort_rule,
        ).grid(row=1, column=0, sticky="we", padx=6, pady=4)

        list_container = tk.Frame(sort_frame)
        list_container.grid(row=3, column=0, sticky="nsew", padx=6, pady=4)
        list_container.rowconfigure(0, weight=1)
        list_container.columnconfigure(0, weight=1)

        self.sort_listbox = tk.Listbox(list_container,
                                       height=8,
                                       exportselection=False)
        self.sort_listbox.grid(row=0, column=0, sticky="nsew")

        sort_v_scrollbar = tk.Scrollbar(
            list_container,
            orient="vertical",
            command=self.sort_listbox.yview,
        )
        sort_v_scrollbar.grid(row=0, column=1, sticky="ns")

        sort_h_scrollbar = tk.Scrollbar(
            list_container,
            orient="horizontal",
            command=self.sort_listbox.xview,
        )
        sort_h_scrollbar.grid(row=1, column=0, sticky="we")

        self.sort_listbox.configure(
            yscrollcommand=sort_v_scrollbar.set,
            xscrollcommand=sort_h_scrollbar.set,
        )

        btn_frame = tk.Frame(sort_frame)
        btn_frame.grid(row=4, column=0, sticky="we", padx=6, pady=(4, 6))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        tk.Button(btn_frame, text="上移",
                  command=self.move_sort_rule_up).grid(row=0,
                                                       column=0,
                                                       sticky="we",
                                                       padx=2,
                                                       pady=2)
        tk.Button(btn_frame, text="下移",
                  command=self.move_sort_rule_down).grid(row=0,
                                                         column=1,
                                                         sticky="we",
                                                         padx=2,
                                                         pady=2)
        tk.Button(btn_frame, text="删除",
                  command=self.delete_sort_rule).grid(row=1,
                                                      column=0,
                                                      sticky="we",
                                                      padx=2,
                                                      pady=2)
        tk.Button(btn_frame, text="清空",
                  command=self.clear_sort_rules).grid(row=1,
                                                      column=1,
                                                      sticky="we",
                                                      padx=2,
                                                      pady=2)

    def build_log_frame(self, parent):
        log_frame = tk.LabelFrame(parent, text="日志")
        log_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        text_container = tk.Frame(log_frame)
        text_container.grid(row=0,
                            column=0,
                            sticky="nsew",
                            padx=6,
                            pady=(6, 4))
        text_container.rowconfigure(0, weight=1)
        text_container.columnconfigure(0, weight=1)

        self.log_text = tk.Text(text_container, height=18, wrap="none")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        log_v_scrollbar = tk.Scrollbar(
            text_container,
            orient="vertical",
            command=self.log_text.yview,
        )
        log_v_scrollbar.grid(row=0, column=1, sticky="ns")

        log_h_scrollbar = tk.Scrollbar(
            text_container,
            orient="horizontal",
            command=self.log_text.xview,
        )
        log_h_scrollbar.grid(row=1, column=0, sticky="we")

        self.log_text.configure(
            yscrollcommand=log_v_scrollbar.set,
            xscrollcommand=log_h_scrollbar.set,
        )

        action_frame = tk.Frame(log_frame)
        action_frame.grid(row=1, column=0, sticky="e", padx=6, pady=(2, 6))

        self.export_btn = tk.Button(
            action_frame,
            text="开始导出",
            width=14,
            command=self.start_export,
        )
        self.export_btn.pack(side="left", padx=4)

    def set_window_icon(self):
        icon_path = get_resource_path(ICON_FILE_NAME)

        if not os.path.exists(icon_path):
            return

        try:
            self.root.iconbitmap(icon_path)
        except Exception:
            pass

    def bind_popup_escape_close(self, popup):
        popup.bind("<Escape>", lambda event: popup.destroy())
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def show_feedback_qr(self):
        qr_path = get_resource_path(QR_CODE_FILE_NAME)

        if not os.path.exists(qr_path):
            messagebox.showwarning(
                "二维码不存在", "未找到微信二维码图片：\n\n"
                f"{qr_path}\n\n"
                "请将图片命名为 wechat_qr.png，并放在程序目录下。")
            return

        popup = tk.Toplevel(self.root)
        popup.title("反馈交流")
        popup.resizable(False, False)
        self.bind_popup_escape_close(popup)

        tk.Label(
            popup,
            text="扫码添加作者微信，反馈交流",
            font=("Microsoft YaHei", 12, "bold"),
        ).pack(padx=18, pady=(16, 8))

        try:
            original_image = tk.PhotoImage(file=qr_path)
        except Exception as e:
            messagebox.showerror(
                "图片加载失败", f"二维码图片加载失败：\n\n{qr_path}\n\n原因：{e}\n\n"
                "建议使用 PNG 格式图片。")
            popup.destroy()
            return

        # 控制二维码最大显示尺寸
        max_qr_size = 360

        image_width = original_image.width()
        image_height = original_image.height()
        max_side = max(image_width, image_height)

        scale = max(1, (max_side + max_qr_size - 1) // max_qr_size)

        qr_image = original_image.subsample(scale, scale)

        image_label = tk.Label(popup, image=qr_image)
        image_label.image = qr_image
        image_label.pack(padx=18, pady=8)

        tk.Button(
            popup,
            text="关闭",
            width=10,
            command=popup.destroy,
        ).pack(pady=(4, 16))

        popup.transient(self.root)
        popup.grab_set()
        popup.update_idletasks()

        x = self.root.winfo_rootx() + (self.root.winfo_width() -
                                       popup.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() -
                                       popup.winfo_height()) // 2
        popup.geometry(f"+{x}+{y}")
        popup.focus_set()

    # ------------------------------------------------------------
    # 路径打开复用逻辑
    # ------------------------------------------------------------

    def get_data_source_option(self, source_key):
        for item in DATA_SOURCE_OPTIONS:
            if item["key"] == source_key:
                return item

        raise ValueError(f"未知数据源：{source_key}")

    def open_existing_folder(self, folder_path):
        folder_path = normalize_user_path(folder_path,
                                          trim_trailing_slash=True)

        if not folder_path:
            messagebox.showwarning("无法打开", "路径为空。")
            return

        if not os.path.isdir(folder_path):
            messagebox.showwarning("无法打开", f"文件夹不存在：\n{folder_path}")
            return

        try:
            os.startfile(folder_path)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开文件夹：\n{folder_path}\n\n原因：{e}")

    def open_folder_for_path(self, path, path_kind):
        path = normalize_user_path(path, trim_trailing_slash=True)

        if not path:
            messagebox.showwarning("无法打开", "路径为空。")
            return

        if path_kind == "file":
            folder_path = win_dirname(path)
        elif path_kind == "directory":
            folder_path = path
        else:
            messagebox.showerror("错误", f"未知路径类型：{path_kind}")
            return

        self.open_existing_folder(folder_path)

    def open_source_folder(self, source_key):
        option = self.get_data_source_option(source_key)
        source_path = self.path_vars[source_key].get().strip()

        self.open_folder_for_path(
            source_path,
            option["path_kind"],
        )

    def open_output_folder(self):
        output_path = normalize_user_path(
            self.output_path_var.get(),
            trim_trailing_slash=True,
        )

        if output_path:
            self.open_folder_for_path(output_path, "file")
            return

        output_dir = self.config["output"]["output_dir"] or get_program_dir()
        self.open_existing_folder(output_dir)

    # ------------------------------------------------------------
    # 日期、路径选择、配置收集
    # ------------------------------------------------------------

    def set_output_path(self, path):
        self.output_path_auto_editing = True

        try:
            self.output_path_var.set(
                normalize_user_path(path, trim_trailing_slash=True))
        finally:
            self.output_path_auto_editing = False

    def on_output_path_changed(self):
        if self.output_path_auto_editing:
            return

        path = normalize_user_path(
            self.output_path_var.get(),
            trim_trailing_slash=True,
        )

        filename = ntpath.basename(path)
        _, custom_part, _ = split_output_filename(filename)

        custom_part = custom_part.strip()

        if custom_part:
            self.config["output"]["filename_custom_part"] = custom_part

        self.allow_next_date_auto_update = False
        self.save_current_config()

    def on_date_range_changed(self):
        self.allow_next_date_auto_update = True
        self.update_output_path_by_dates()
        self.allow_next_date_auto_update = False
        self.save_current_config()

    def normalize_date_entry(self, date_var):
        raw = date_var.get().strip()

        if not raw:
            return

        try:
            normalized = normalize_date_text(raw)
            date_var.set(normalized)
        except Exception:
            pass

    def normalize_date_inputs_or_raise(self):
        try:
            start_normalized = normalize_date_text(self.start_date_var.get())
        except Exception as e:
            raise ValueError(f"开始日期格式错误：{e}")

        try:
            end_normalized = normalize_date_text(self.end_date_var.get())
        except Exception as e:
            raise ValueError(f"结束日期格式错误：{e}")

        self.start_date_var.set(start_normalized)
        self.end_date_var.set(end_normalized)

        start_date = parse_yyyymmdd(start_normalized)
        end_date = parse_yyyymmdd(end_normalized)

        if start_date > end_date:
            raise ValueError("开始日期不能晚于结束日期。")

        self.update_output_path_by_dates()

    def choose_source_path(self, source_key):
        option = self.get_data_source_option(source_key)

        if option["path_kind"] == "file":
            path = filedialog.askopenfilename(
                title=f"选择{option['title']}文件",
                filetypes=[
                    ("SQLite 数据库", "*.sqlite *.db"),
                    ("所有文件", "*.*"),
                ],
                initialdir=self.path_vars[source_key].get().strip()
                or None  # 优先使用设置路径
            )
        else:
            path = filedialog.askdirectory(
                title=f"选择{option['title']}目录",
                initialdir=self.path_vars[source_key].get().strip()
                or None  # 优先使用设置路径
            )
        if path:
            path = normalize_user_path(path, trim_trailing_slash=True)
            self.path_vars[source_key].set(path)
            self.data_source_var.set(source_key)

    def choose_output_dir(self):
        path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.config["output"]["output_dir"] or None)

        if path:
            path = normalize_user_path(path, trim_trailing_slash=True)
            self.config["output"]["output_dir"] = path

            self.allow_next_date_auto_update = True
            self.update_output_path_by_dates(force_dir=path)
            self.allow_next_date_auto_update = False

            self.save_current_config()

    def choose_output_file(self):
        current_path = normalize_user_path(
            self.output_path_var.get(),
            trim_trailing_slash=True,
        )

        current_filename = ntpath.basename(current_path)

        if not current_filename:
            try:
                start_str = normalize_date_text(self.start_date_var.get())
                end_str = normalize_date_text(self.end_date_var.get())
            except Exception:
                start_str = self.start_date_var.get().strip()
                end_str = self.end_date_var.get().strip()

            custom_part = self.config["output"].get(
                "filename_custom_part",
                "个股排序",
            )
            current_filename = build_output_filename(
                start_str,
                end_str,
                custom_part,
                ".xlsx",
            )

        path = filedialog.asksaveasfilename(
            title="选择输出 Excel 文件",
            defaultextension=".xlsx",
            initialfile=current_filename,
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("所有文件", "*.*"),
            ],
        )

        if path:
            path = normalize_user_path(path, trim_trailing_slash=True)
            self.set_output_path(path)
            self.on_output_path_changed()

            output_dir = normalize_user_path(
                win_dirname(path),
                trim_trailing_slash=True,
            )

            if output_dir:
                self.config["output"]["output_dir"] = output_dir

            self.save_current_config()

    def update_output_path_by_dates(self, force_dir=None):
        if not self.allow_next_date_auto_update and force_dir is None:
            return

        try:
            start_str = normalize_date_text(self.start_date_var.get())
            end_str = normalize_date_text(self.end_date_var.get())
        except Exception:
            return

        current_path = normalize_user_path(
            self.output_path_var.get(),
            trim_trailing_slash=True,
        )

        current_dir = win_dirname(current_path)
        current_filename = ntpath.basename(current_path)

        _, custom_part, ext = split_output_filename(current_filename)

        if force_dir is not None:
            output_dir = force_dir
        elif current_dir:
            output_dir = current_dir
        else:
            output_dir = self.config["output"]["output_dir"]

        output_dir = normalize_user_path(output_dir, trim_trailing_slash=True)

        if not output_dir:
            output_dir = get_program_dir()

        custom_part = custom_part.strip()

        if not custom_part:
            custom_part = self.config["output"].get(
                "filename_custom_part",
                "个股排序",
            )

        self.config["output"]["filename_custom_part"] = custom_part

        filename = build_output_filename(
            start_str,
            end_str,
            custom_part,
            ext or ".xlsx",
        )

        self.set_output_path(win_join(output_dir, filename))

    def validate_positive_integer_input(self, value):
        if value == "":
            return True

        return value.isdigit() and int(value) > 0

    def toggle_max_rows_entry(self):
        if self.keep_all_rows_var.get():
            self.max_rows_entry.config(state="disabled")
        else:
            self.max_rows_entry.config(state="normal")

    # ------------------------------------------------------------
    # 排序条件
    # ------------------------------------------------------------

    def add_sort_rule(self):
        field = field_key_by_title(self.sort_field_title_var.get())
        order = sort_order_key_by_title(self.sort_order_title_var.get())

        new_rule = {
            "field": field,
            "order": order,
        }

        if new_rule in self.sort_rules:
            messagebox.showwarning("排序条件已存在", "该排序条件已存在，无需重复添加。")
            return

        self.sort_rules.append(new_rule)

        self.ensure_sort_fields_selected()
        self.refresh_sort_listbox()

    def selected_sort_index(self):
        selection = self.sort_listbox.curselection()

        if not selection:
            return None

        return selection[0]

    def move_sort_rule_up(self):
        idx = self.selected_sort_index()

        if idx is None or idx <= 0:
            return

        self.sort_rules[idx - 1], self.sort_rules[idx] = self.sort_rules[
            idx], self.sort_rules[idx - 1]

        self.ensure_sort_fields_selected()
        self.refresh_sort_listbox()
        self.sort_listbox.selection_set(idx - 1)

    def move_sort_rule_down(self):
        idx = self.selected_sort_index()

        if idx is None or idx >= len(self.sort_rules) - 1:
            return

        self.sort_rules[idx + 1], self.sort_rules[idx] = self.sort_rules[
            idx], self.sort_rules[idx + 1]

        self.ensure_sort_fields_selected()
        self.refresh_sort_listbox()
        self.sort_listbox.selection_set(idx + 1)

    def delete_sort_rule(self):
        idx = self.selected_sort_index()

        if idx is None:
            return

        del self.sort_rules[idx]
        self.ensure_sort_fields_selected()
        self.refresh_sort_listbox()

    def clear_sort_rules(self):
        self.sort_rules = []
        self.refresh_sort_listbox()
        self.save_current_config()

    def ensure_sort_fields_selected(self):
        sort_fields = [rule["field"] for rule in self.sort_rules]

        if sort_fields:
            self.field_checklist.select_keys(sort_fields)
            self.save_current_config()

    def refresh_sort_listbox(self):
        self.sort_listbox.delete(0, "end")

        title_map = field_title_map()
        order_map = sort_order_title_map()

        for idx, rule in enumerate(self.sort_rules, start=1):
            self.sort_listbox.insert(
                "end",
                f"{idx}. {title_map[rule['field']]} - {order_map[rule['order']]}",
            )

    # ------------------------------------------------------------
    # 配置与导出
    # ------------------------------------------------------------

    def collect_config_from_ui(self):
        self.normalize_date_inputs_or_raise()

        selected_fields = self.field_checklist.get_selected_keys_in_order()

        if not selected_fields:
            raise ValueError("至少需要选择一个输出字段。")

        max_rows_raw = self.max_rows_per_sheet_var.get().strip()

        if not max_rows_raw or not max_rows_raw.isdigit() or int(
                max_rows_raw) <= 0:
            raise ValueError("每页保留标的数量必须是正整数。")

        config = deepcopy(self.config)

        config["data_source"] = self.data_source_var.get()

        for item in DATA_SOURCE_OPTIONS:
            key = item["key"]
            config["paths"][key] = normalize_user_path(
                self.path_vars[key].get(),
                trim_trailing_slash=True,
            )

        config["date_range"]["start_date"] = self.start_date_var.get().strip()
        config["date_range"]["end_date"] = self.end_date_var.get().strip()

        output_path = normalize_user_path(
            self.output_path_var.get(),
            trim_trailing_slash=True,
        )

        config["output"]["output_dir"] = normalize_user_path(
            win_dirname(output_path),
            trim_trailing_slash=True,
        )

        filename = ntpath.basename(output_path)
        _, custom_part, _ = split_output_filename(filename)

        custom_part = custom_part.strip()

        if custom_part:
            config["output"]["filename_custom_part"] = custom_part
        else:
            config["output"]["filename_custom_part"] = self.config[
                "output"].get(
                    "filename_custom_part",
                    "个股排序",
                )

        config["output"]["keep_all_rows"] = bool(self.keep_all_rows_var.get())
        config["output"]["max_rows_per_sheet"] = int(max_rows_raw)

        config["fields"]["order"] = self.field_checklist.get_ordered_keys()
        config["fields"]["selected"] = selected_fields

        config["sort"]["rules"] = deepcopy(self.sort_rules)

        return normalize_config_paths(config)

    def resolve_existing_output_file(self, output_path):
        output_path = normalize_user_path(output_path,
                                          trim_trailing_slash=True)

        if not os.path.exists(output_path):
            return output_path

        choice = messagebox.askyesnocancel(
            "文件已存在",
            "输出文件已存在：\n\n"
            f"{output_path}\n\n"
            "请选择处理方式：\n\n"
            "是：覆盖原文件\n"
            "否：自动改名保存\n"
            "取消：取消本次导出",
        )

        if choice is True:
            return output_path

        if choice is False:
            new_path = make_unique_output_path(output_path)
            self.set_output_path(new_path)
            self.on_output_path_changed()
            return new_path

        return None

    def save_current_config(self):
        try:
            self.config = self.collect_config_from_ui()
            save_app_config(self.config)
            # print("配置已自动保存。")  # 调试信息

        except Exception as e:
            print(f"保存配置失败：{e}")  # 错误处理

    def log(self, message):
        self.log_queue.put(str(message))

    def poll_log_queue(self):
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break

            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")

        self.root.after(100, self.poll_log_queue)

    def start_export(self):
        try:
            config = self.collect_config_from_ui()
            output_path = normalize_user_path(
                self.output_path_var.get(),
                trim_trailing_slash=True,
            )

            if output_path and not output_path.lower().endswith(".xlsx"):
                output_path += ".xlsx"
                output_path = normalize_user_path(output_path,
                                                  trim_trailing_slash=True)
                self.set_output_path(output_path)
                self.on_output_path_changed()

            validate_runtime_config(config, output_path)

            # 检查输出文件是否正在被打开
            while is_file_open(output_path):
                choice = messagebox.askyesnocancel(
                    "文件已打开", f"输出文件已被打开：\n{output_path}\n"
                    "请关闭文件后选择【是】继续导出，或选择【否】更改文件名，"
                    "选择【取消】以终止导出。",
                    icon='warning')
                if choice is True:
                    # 用户选择已关闭，继续检查
                    continue
                elif choice is False:
                    # 用户选择另存为
                    new_path = filedialog.asksaveasfilename(
                        title="另存为输出文件",
                        defaultextension=".xlsx",
                        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
                        initialfile=os.path.basename(output_path),
                    )
                    if new_path:
                        output_path = normalize_user_path(
                            new_path, trim_trailing_slash=True)
                        self.set_output_path(output_path)
                        self.on_output_path_changed()
                    else:
                        return

                else:
                    return  # 用户取消

            resolved_output_path = self.resolve_existing_output_file(
                output_path)

            if resolved_output_path is None:
                return

            output_path = normalize_user_path(resolved_output_path,
                                              trim_trailing_slash=True)
            self.set_output_path(output_path)
            self.on_output_path_changed()

            self.config = config
            save_app_config(self.config)

        except Exception as e:
            messagebox.showerror("输入错误", str(e))
            return

        self.export_btn.config(state="disabled")
        self.log_text.delete("1.0", "end")

        thread = threading.Thread(
            target=self.run_export_thread,
            args=(self.config, output_path),
            daemon=True,
        )
        thread.start()

    def run_export_thread(self, config, output_path):
        try:
            export_to_excel(config, output_path, log_func=self.log)

            self.log("任务成功完成。")

            # 新增：检查复选框状态
            if self.auto_open_var.get():
                # 直接打开生成的 Excel 文件
                if os.path.exists(output_path):  # 确保文件存在
                    self.log(f"准备打开输出文件: {output_path}")  # 打印输出文件路径以进行调试
                    os.startfile(output_path)  # 打开文件
                else:
                    self.log(f"输出文件不存在: {output_path}")

            self.root.after(
                0,
                lambda: messagebox.showinfo("完成", "导出完成。"),
            )

        except Exception as e:
            msg = str(e)
            self.log(f"导出失败：{msg}")

            self.root.after(
                0,
                lambda m=msg: messagebox.showerror("导出失败", m),
            )

        finally:
            self.root.after(
                0,
                lambda: self.export_btn.config(state="normal"),
            )

    def on_close(self):
        try:
            self.config = self.collect_config_from_ui()
            save_app_config(self.config)
        except Exception:
            pass

        self.root.destroy()


def main():
    root = tk.Tk()
    ExportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
    

# 【运行虚拟环境】
# PS E:\AppProject\DailyStockRank> .venv\\Scripts\activate

# 【封装exe文件】
# Set-Location "E:\AppProject\DailyStockRank"
# pyinstaller --noconfirm --onefile --windowed --name DailyStockRank --collect-data tzdata --hidden-import openpyxl --hidden-import openpyxl.styles --hidden-import openpyxl.utils --hidden-import openpyxl.writer.excel --hidden-import openpyxl.reader.excel --hidden-import pandas._libs.tslibs.np_datetime --hidden-import pandas._libs.tslibs.nattype --hidden-import pandas._libs.tslibs.timedeltas --hidden-import pandas._libs.missing --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module notebook --exclude-module jupyter --exclude-module pytest --exclude-module unittest --exclude-module requests --exclude-module httpx --exclude-module mootdx --exclude-module tdxpy --exclude-module eltdx --exclude-module OpenCC "E:\AppProject\DailyStockRank\DailyStockRank.py" --add-data "E:\AppProject\DailyStockRank\wechat_qr.png;." --icon "E:\AppProject\DailyStockRank\icon.ico" --add-data "E:\AppProject\DailyStockRank\icon.ico;." --upx-dir "D:\upx-5.1.1-win64"

# 文件会生成在"E:\AppProject\DailyStockRank\dist\DailyStockRank.exe"

# 【github版本管理】

# Set-Location "E:\AppProject\DailyStockRank"
# git init #初始化，仅最初运行一次，后续不再运行

# git add . #把所有代码加入暂存区，需要每次运行
# git commit -m "提交说明" #提交代码到本地仓库

# git remote add origin https://github.com/mccoach/dailystockrank.git #关联远程仓库（最关键一步），仅第一次推送前运行

# git push origin main #推送代码到远程仓库，只推代码不推版本号

# git tag 版本号（不能含空格） #提交版本号，不能包含空格，如v1.0.0.20260524

# git push origin --tags #推送所有版本号标签到远程仓库，只推版本号不推代码

# #最常用的 5 条命令（每次推送都用）
# git add .            # 保存改动
# git commit -m "说明"  # 提交到本地
# git push origin main # 推送到远程
# git tag 版本号      # 打版本号标签
# git push origin --tags  # 推送版本号
