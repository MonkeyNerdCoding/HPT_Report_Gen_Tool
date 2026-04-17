from __future__ import annotations

import glob
import os
import sys
from collections.abc import Callable
from pathlib import Path

import pandas as pd

from .name_detect import extract_sheet_name


LogCallback = Callable[[str], None]


def merge_sql_csv(
    input_folder: str | Path,
    output_file: str | Path,
    log_callback: LogCallback | None = None,
) -> str | None:
    """Merge all CSV files in one DB folder into a multi-sheet Excel workbook."""
    input_path = Path(input_folder)
    output_path = Path(output_file)
    log = _make_logger(log_callback)

    csv_files = sorted(glob.glob(os.path.join(str(input_path), "*.csv")))
    if not csv_files:
        log(f"⚠️ Không có CSV trong {input_path}, bỏ qua.\n")
        return None

    all_data = {}

    for file in csv_files:
        filename = os.path.basename(file)
        sheet_name = extract_sheet_name(filename)

        if not sheet_name:
            continue

        try:
            dataframe = pd.read_csv(file)
            log(f"   ✅ {filename} ({len(dataframe)} dòng)")
        except Exception as exc:
            log(f"   ❌ Lỗi đọc {filename}: {exc}")
            continue

        if sheet_name not in all_data:
            all_data[sheet_name] = []
        all_data[sheet_name].append(dataframe)

    if not all_data:
        return None

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, dataframes in all_data.items():
            merged_df = pd.concat(dataframes, ignore_index=True)
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            log(f"📝 Đã ghi sheet: {sheet_name} ({len(merged_df)} dòng)")

    log(f"✅ Done! File Excel sinh ra: {output_path}\n")
    return str(output_path)


def merge_sql_root_csv(
    input_root: str | Path,
    output_file: str | Path,
    log_callback: LogCallback | None = None,
) -> str | None:
    """Merge CSV files from every direct DB subfolder into one Excel workbook."""
    input_path = Path(input_root)
    output_path = Path(output_file)
    log = _make_logger(log_callback)

    db_folders = sorted(child for child in input_path.iterdir() if child.is_dir())
    if not db_folders:
        log(f"⚠️ Không có DB folder trong {input_path}, bỏ qua.\n")
        return None

    all_data: dict[str, list[pd.DataFrame]] = {}
    for db_folder in db_folders:
        csv_files = sorted(glob.glob(os.path.join(str(db_folder), "*.csv")))
        if not csv_files:
            log(f"⚠️ Không có CSV trong {db_folder}, bỏ qua.\n")
            continue

        log("")
        log(f"🚀 Đang xử lý DB folder: {db_folder.name}")
        folder_data = _read_csv_files(csv_files, log)
        for sheet_name, dataframes in folder_data.items():
            all_data.setdefault(sheet_name, []).extend(dataframes)

    if not all_data:
        log(f"⚠️ Không có CSV hợp lệ trong SQL root folder: {input_path}")
        return None

    return _write_excel(all_data, output_path, log)


def merge_sql_root_healthcheck(
    input_root: str | Path,
    output_file: str | Path,
    log_callback: LogCallback | None = None,
) -> str | None:
    """Merge one DB folder or a root of DB subfolders with CSV files into one workbook."""
    input_path = Path(input_root)
    log = _make_logger(log_callback)

    direct_csv_files = sorted(glob.glob(os.path.join(str(input_path), "*.csv")))
    if direct_csv_files:
        return merge_sql_csv(input_path, output_file, log_callback=log_callback)

    db_folders_with_csv = [
        child for child in input_path.iterdir()
        if child.is_dir() and any(csv_file.is_file() for csv_file in child.glob("*.csv"))
    ]
    if db_folders_with_csv:
        return merge_sql_root_csv(input_path, output_file, log_callback=log_callback)

    log(f"⚠️ Không tìm thấy CSV hoặc DB folder chứa CSV trong {input_path}, bỏ qua.\n")
    return None


def _read_csv_files(
    csv_files: list[str],
    log: LogCallback,
) -> dict[str, list[pd.DataFrame]]:
    all_data: dict[str, list[pd.DataFrame]] = {}
    for csv_file in csv_files:
        filename = os.path.basename(csv_file)
        sheet_name = extract_sheet_name(filename)
        if not sheet_name:
            continue

        try:
            dataframe = pd.read_csv(csv_file)
        except Exception as exc:
            log(f"   ❌ Lỗi đọc {filename}: {exc}")
            continue

        if sheet_name not in all_data:
            all_data[sheet_name] = []
        all_data[sheet_name].append(dataframe)
        log(f"   ✅ {filename} ({len(dataframe)} dòng)")

    return all_data


def _write_excel(
    all_data: dict[str, list[pd.DataFrame]],
    output_path: Path,
    log: LogCallback,
) -> str:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, dataframes in all_data.items():
            merged = pd.concat(dataframes, ignore_index=True)
            merged.to_excel(writer, sheet_name=sheet_name, index=False)
            log(f"📝 Đã ghi sheet: {sheet_name} ({len(merged)} dòng)")

    log(f"✅ Done! File Excel sinh ra: {output_path}\n")
    return str(output_path)


def _make_logger(log_callback: LogCallback | None) -> LogCallback:
    def log(message: str) -> None:
        _safe_print(message)
        if log_callback:
            log_callback(message)

    return log


def _safe_print(message: str) -> None:
    try:
        print(message)
    except UnicodeEncodeError:
        encoding = sys.stdout.encoding or "utf-8"
        safe_message = message.encode(encoding, errors="backslashreplace").decode(encoding, errors="replace")
        print(safe_message)
