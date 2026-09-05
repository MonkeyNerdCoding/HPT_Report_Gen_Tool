from __future__ import annotations

import glob
import os
import re
import sys
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .name_detect import extract_sheet_name


LogCallback = Callable[[str], None]


_VOLUME_INFO_FILENAME = re.compile(
    r"^(?P<source>.+?)-DQ-0*26-Volume\s+Info(?:-(?P<suffix>.*))?$",
    re.IGNORECASE,
)
_FILENAME_TIMESTAMP = re.compile(r"(?<!\d)(\d{14,})(?!\d)")
_COLLECTION_TIMESTAMP_COLUMNS = {
    "collectiontimestamp",
    "collectiontime",
    "collectedat",
    "collectiondate",
    "timestamp",
}


def merge_sql_csv(
    input_folder: str | Path,
    output_file: str | Path,
    log_callback: LogCallback | None = None,
    cancel_check: Callable[[], bool] | None = None,
) -> str | None:
    """Merge all CSV files in one DB folder into a multi-sheet Excel workbook."""
    input_path = Path(input_folder)
    output_path = Path(output_file)
    log = _make_logger(log_callback)

    csv_files = sorted(glob.glob(os.path.join(str(input_path), "*.csv")))
    if not csv_files:
        log(f"⚠️ Không có CSV trong {input_path}, bỏ qua.\n")
        return None

    csv_files = _select_latest_volume_info_files(csv_files, log)
    all_data = {}

    for file in csv_files:
        if cancel_check and cancel_check():
            return None
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
            merged_df = _merge_dataframes(sheet_name, dataframes)
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_excel_sheet(writer, sheet_name, merged_df)
            log(f"📝 Đã ghi sheet: {sheet_name} ({len(merged_df)} dòng)")

    log(f"✅ Done! File Excel sinh ra: {output_path}\n")
    return str(output_path)


def merge_sql_root_csv(
    input_root: str | Path,
    output_file: str | Path,
    log_callback: LogCallback | None = None,
    cancel_check: Callable[[], bool] | None = None,
) -> str | None:
    """Merge CSV files from every direct DB subfolder into one Excel workbook."""
    input_path = Path(input_root)
    output_path = Path(output_file)
    log = _make_logger(log_callback)

    db_folders = sorted(child for child in input_path.iterdir() if child.is_dir())
    if not db_folders:
        log(f"⚠️ Không có DB folder trong {input_path}, bỏ qua.\n")
        return None

    all_csv_files = [
        csv_file
        for db_folder in db_folders
        for csv_file in glob.glob(os.path.join(str(db_folder), "*.csv"))
    ]
    selected_csv_files = set(_select_latest_volume_info_files(all_csv_files, log))

    all_data: dict[str, list[pd.DataFrame]] = {}
    for db_folder in db_folders:
        if cancel_check and cancel_check():
            return None
        csv_files = sorted(
            csv_file
            for csv_file in glob.glob(os.path.join(str(db_folder), "*.csv"))
            if csv_file in selected_csv_files
        )
        if not csv_files:
            log(f"⚠️ Không có CSV trong {db_folder}, bỏ qua.\n")
            continue

        log("")
        log(f"🚀 Đang xử lý DB folder: {db_folder.name}")
        folder_data = _read_csv_files(csv_files, log, cancel_check=cancel_check)
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
    cancel_check: Callable[[], bool] | None = None,
) -> str | None:
    """Merge one DB folder or a root of DB subfolders with CSV files into one workbook."""
    input_path = Path(input_root)
    log = _make_logger(log_callback)

    direct_csv_files = sorted(glob.glob(os.path.join(str(input_path), "*.csv")))
    if direct_csv_files:
        return merge_sql_csv(input_path, output_file, log_callback=log_callback, cancel_check=cancel_check)

    db_folders_with_csv = [
        child for child in input_path.iterdir()
        if child.is_dir() and any(csv_file.is_file() for csv_file in child.glob("*.csv"))
    ]
    if db_folders_with_csv:
        return merge_sql_root_csv(input_path, output_file, log_callback=log_callback, cancel_check=cancel_check)

    log(f"⚠️ Không tìm thấy CSV hoặc DB folder chứa CSV trong {input_path}, bỏ qua.\n")
    return None


def _read_csv_files(
    csv_files: list[str],
    log: LogCallback,
    cancel_check: Callable[[], bool] | None = None,
) -> dict[str, list[pd.DataFrame]]:
    all_data: dict[str, list[pd.DataFrame]] = {}
    for csv_file in _select_latest_volume_info_files(csv_files, log):
        if cancel_check and cancel_check():
            return all_data
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
            merged = _merge_dataframes(sheet_name, dataframes)
            merged.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_excel_sheet(writer, sheet_name, merged)
            log(f"📝 Đã ghi sheet: {sheet_name} ({len(merged)} dòng)")

    log(f"✅ Done! File Excel sinh ra: {output_path}\n")
    return str(output_path)


def _select_latest_volume_info_files(csv_files: list[str], log: LogCallback) -> list[str]:
    """Keep only the newest DQ-26 snapshot for each server/instance source."""
    regular_files: list[str] = []
    volume_groups: dict[str, list[tuple[int, str]]] = {}

    for csv_file in csv_files:
        path = Path(csv_file)
        match = _VOLUME_INFO_FILENAME.match(path.stem)
        if not match:
            regular_files.append(csv_file)
            continue

        source_key = match.group("source").strip().casefold()
        timestamp = _collection_timestamp(path, match.group("suffix") or "")
        volume_groups.setdefault(source_key, []).append((timestamp, csv_file))

    selected_volume_files: list[str] = []
    for snapshots in volume_groups.values():
        newest_timestamp = max(timestamp for timestamp, _ in snapshots)
        newest_files = [file for timestamp, file in snapshots if timestamp == newest_timestamp]
        selected_volume_files.extend(newest_files)

        skipped_count = len(snapshots) - len(newest_files)
        if skipped_count:
            log(
                "   DQ-26 Volume Info: "
                f"dung snapshot moi nhat {Path(newest_files[0]).name}, "
                f"bo qua {skipped_count} snapshot cu"
            )

    return sorted(regular_files + selected_volume_files)


def _collection_timestamp(path: Path, filename_suffix: str) -> int:
    timestamp_matches = _FILENAME_TIMESTAMP.findall(filename_suffix)
    if timestamp_matches:
        timestamp = timestamp_matches[-1]
        try:
            collected_at = datetime.strptime(timestamp[:14], "%Y%m%d%H%M%S")
            base_nanoseconds = int(pd.Timestamp(collected_at).value)
            fractional = timestamp[14:]
            fractional_nanoseconds = int(fractional[:9].ljust(9, "0")) if fractional else 0
            return base_nanoseconds + fractional_nanoseconds
        except ValueError:
            pass

    metadata_timestamp = _collection_timestamp_from_csv(path)
    if metadata_timestamp is not None:
        return metadata_timestamp

    return path.stat().st_mtime_ns


def _collection_timestamp_from_csv(path: Path) -> int | None:
    try:
        sample = pd.read_csv(path, nrows=20)
    except Exception:
        return None

    for column in sample.columns:
        normalized = re.sub(r"[^a-z0-9]", "", str(column).casefold())
        if normalized not in _COLLECTION_TIMESTAMP_COLUMNS:
            continue
        timestamps = pd.to_datetime(sample[column], errors="coerce").dropna()
        if not timestamps.empty:
            return int(timestamps.max().value)
    return None


def _merge_dataframes(sheet_name: str, dataframes: list[pd.DataFrame]) -> pd.DataFrame:
    merged = pd.concat(dataframes, ignore_index=True)
    if sheet_name.strip().casefold() == "volume info":
        merged = merged.drop_duplicates(ignore_index=True)
    return merged


def _format_excel_sheet(writer: pd.ExcelWriter, sheet_name: str, dataframe: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for cell in worksheet[1]:
        cell.font = Font(name="Cambria", size=12, bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for column_index, column_name in enumerate(dataframe.columns, start=1):
        column_letter = get_column_letter(column_index)
        values = dataframe[column_name].head(300).tolist()
        lengths = [len(str(column_name))]
        lengths.extend(len(str(value)) for value in values if pd.notna(value))
        width = max(10, min(65, max(lengths, default=10) + 2))
        worksheet.column_dimensions[column_letter].width = width

        wrap = width >= 35
        for cell in worksheet[column_letter]:
            if cell.row != 1:
                cell.font = Font(name="Cambria", size=12)
            cell.alignment = Alignment(wrap_text=wrap, vertical="top")


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
