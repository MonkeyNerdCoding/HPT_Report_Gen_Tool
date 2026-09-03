from __future__ import annotations

from pathlib import Path
from collections import defaultdict
import re

from .html_parser import parse_html_file
from .table_extractor import extract_tables


CACHE_HIT_ASSESSMENT = "Tỉ lệ buffer cache hit và library cache hit đang ở ngưỡng tối ưu 99 – 100%."
BUFFER_CACHE_HIT_ASSESSMENT = "Tỉ lệ buffer cache hit đang ở ngưỡng tối ưu 99 – 100%."
LIBRARY_CACHE_HIT_ASSESSMENT = "Tỉ lệ library cache hit đang ở ngưỡng tối ưu 99 – 100%."
CACHE_HIT_RECOMMENDATION = "N/A"
LOG_SWITCH_RECOMMENDATION = "Tăng thêm dung lượng cho redo log file để giảm tần suất switch."
PGA_ASSESSMENT = "Nhìn chung, vùng nhớ PGA được CSDL sử dụng vẫn nằm trong mức an toàn."
PGA_RECOMMENDATION = "N/A"
ASM_FREE_WARNING_GB = 500
ASM_FREE_WARNING_PERCENT = 10
MULTIPLEXED_REDO_ASSESSMENT = (
    "Theo như cấu hình hiện tại, các redo log group đang được multiplexing, tức là mỗi redo log group có 2 members, "
    "mỗi member được đặt ở một disk controller khác nhau. Điều này tăng tính sẵn sàng của redo log group, đảm bảo database "
    "luôn được vận hành khi có sự cố ảnh hưởng đến một trong những member trong redo log group."
)


def build_edb360_assessment_mapping(input_root: str | Path) -> dict[str, str]:
    root = Path(input_root)
    mapping: dict[str, str] = {}

    all_parameters = _first_table(root, "*all_parameters.html")
    memory_configuration = _first_table(root, "*memory_configuration.html")
    redo_log = _first_table(root, "*redo_log.html")
    redo_log_files = _first_table(root, "*redo_log_files.html")
    registry_sql_patch = _first_table(root, "*registry_sql_patch.html")
    rman_backup = _first_table_any(root, ["*rman_backup_job_details.html", "*_rman_backup.html", "*rman_backup.html"])
    tablespace_usage = _first_table(root, "*tablespace_usage.html")
    log_switch_tables = _tables_for(root, "*log_switch_frequency_for_instance_*.html", exclude=("_line_chart",))
    cpu_busy_tables = _tables_for(root, "*cpu_busy_and_idle_times_percent_for_instance_*.html", exclude=("_line_chart",))
    asm_disk_group = _first_table(root, "*asm_disk_group.html")
    scheduler_jobs = _first_table(root, "*scheduler_jobs.html")
    no_index = _first_table(root, "*tables_without_indexes.html")
    no_pk = _first_table(root, "*tables_without_primary_key_constraints.html")
    invalid_objects = _first_table(root, "*invalid_objects.html")
    table_stats = _first_table(root, "*tables_with_stale_stats.html")
    index_stats = _first_table(root, "*indexes_with_stale_stats.html")

    mapping.update(_control_file_assessment(all_parameters))
    mapping.update(_redo_assessment(redo_log, redo_log_files))
    mapping.update(_memory_assessment(memory_configuration))
    mapping.update(_patching_backup_assessment(registry_sql_patch, rman_backup))
    mapping.update(_tablespace_assessment(tablespace_usage))
    mapping.update(_cache_hit_assessment(log_switch_tables, cpu_busy_tables))
    mapping.update(_asm_assessment(asm_disk_group))
    mapping.update(_scheduler_jobs_assessment(scheduler_jobs))
    mapping.update(_count_assessments(no_index, no_pk, invalid_objects, table_stats, index_stats))
    return {key: value for key, value in mapping.items() if value is not None}


def _control_file_assessment(rows: list[list[str]]) -> dict[str, str]:
    raw_values = [row.get("VALUE", "") for row in _dict_rows(rows) if row.get("NAME", "").lower() == "control_files"]
    values: list[str] = []
    for raw_value in raw_values:
        values.extend([value.strip() for value in str(raw_value).split(",") if value.strip()])
    if not values:
        return {
            "{{assessment_control_file}}": "",
            "{{recommendation_control_file}}": "",
        }

    locations = {_storage_root(value) for value in values}
    count = len(values)
    if len(locations) >= 2:
        assessment = (
            f"Hiện tại, database đang có {count} control files và các control files này được đặt trên "
            f"{len(locations)} vị trí lưu trữ khác nhau ({', '.join(sorted(locations))}), đảm bảo tính an toàn cho control files."
        )
        recommendation = "N/A"
    else:
        location = next(iter(locations), "")
        assessment = (
            f"Hiện tại, database đang có {count} control files. Tuy nhiên, các control files này đang nằm cùng một vị trí "
            f"{location}, không đảm bảo tính an toàn nếu vị trí lưu trữ này gặp sự cố."
        )
        recommendation = "Đưa các control files ra nhiều phân vùng/disk group khác nhau hoặc bổ sung control file ở vị trí lưu trữ độc lập."
    return {
        "{{assessment_control_file}}": assessment,
        "{{recommendation_control_file}}": recommendation,
    }


def _redo_assessment(redo_rows: list[list[str]], redo_file_rows: list[list[str]]) -> dict[str, str]:
    redo = _dict_rows(redo_rows)
    redo_files = _dict_rows(redo_file_rows)
    if not redo and not redo_files:
        return {"{{assessment_redo_log}}": "", "{{recommendation_redo_log}}": ""}

    member_counts = []
    for row in redo:
        member = _to_float(row.get("MEMBERS", ""))
        if member is not None:
            member_counts.append(int(member))
    if not member_counts and redo_files:
        groups: dict[str, set[str]] = defaultdict(set)
        for row in redo_files:
            group = row.get("GROUP#", "")
            member = row.get("MEMBER", "")
            if group and member:
                groups[group].add(member)
        member_counts = [len(items) for items in groups.values()]

    group_count = len(member_counts)
    min_members = min(member_counts) if member_counts else 0
    if min_members >= 2:
        assessment = MULTIPLEXED_REDO_ASSESSMENT
        recommendation = "N/A"
    else:
        assessment = (
            f"Theo cấu hình hiện tại, database có {group_count} redo log groups nhưng có redo log group chưa được multiplexing đầy đủ."
        )
        recommendation = "Bổ sung redo log member trên disk group/phân vùng khác để tăng tính sẵn sàng."
    return {
        "{{assessment_redo_log}}": assessment,
        "{{recommendation_redo_log}}": recommendation,
    }


def _memory_assessment(rows: list[list[str]]) -> dict[str, str]:
    data = _dict_rows(rows)
    if not data:
        return {"{{assessment_memory_configuration}}": "", "{{recommendation_memory_configuration}}": ""}

    values: dict[str, list[str]] = defaultdict(list)
    for row in data:
        name = row.get("NAME", "").lower()
        value = _preferred_memory_value(row)
        if name and value:
            values[name].append(value)

    memory_target = _max_numeric(values.get("memory_target", []))
    sga = _positive_values(values.get("sga_target", [])) or _positive_values(values.get("sga_max_size", []))
    pga = values.get("pga_aggregate_target") or []
    mode = "Automatic Memory Management (AMM)" if memory_target and memory_target > 0 else "Automatic Shared Memory Management (ASMM)"

    parts = []
    if sga:
        parts.append(f"SGA {', '.join(dict.fromkeys(sga))}/instance")
    if pga:
        parts.append(f"PGA {', '.join(dict.fromkeys(pga))}/instance")
    detail = "; ".join(parts) if parts else "chưa xác định được SGA/PGA từ EDB360"
    return {
        "{{assessment_memory_configuration}}": f"Cơ sở dữ liệu hiện tại đang được cấu hình vùng nhớ ở chế độ {mode}. Trong đó: {detail}.",
        "{{recommendation_memory_configuration}}": "N/A",
    }


def _patching_backup_assessment(registry_rows: list[list[str]], backup_rows: list[list[str]]) -> dict[str, str]:
    patches = _dict_rows(registry_rows)
    backups = _dict_rows(backup_rows)
    patch_desc = ""
    if patches:
        latest = patches[-1]
        patch_desc = latest.get("DESCRIPTION") or latest.get("VERSION") or latest.get("PATCH_ID", "")

    completed = [row for row in backups if "COMPLETED" in row.get("STATUS", "").upper()]
    if backups:
        completed_text = f" Trong đó có {len(completed)} backup job hoàn thành." if completed else ""
        backup_assessment = f"Đã có RMAN backup. EDB360 ghi nhận {len(backups)} backup job trong dữ liệu thu thập.{completed_text}"
        backup_recommendation = (
            "Khuyến nghị chuẩn bị môi trường thực hiện kiểm thử restore các bản backup. "
            "Việc không có môi trường khôi phục kiểm thử bản backup sẽ không đảm bảo bản backup có thể khôi phục thành công khi cần thiết."
        )
    else:
        backup_assessment = "Hiện tại chưa có RMAN backup"
        backup_recommendation = "Khuyến nghị tạo thêm RMAN backup để thực hiện khôi phục hoàn toàn khi hệ thống\ngặp sự cố."

    return {
        "{{assessment_patching}}": f"Phiên bản/patch hiện tại: {patch_desc}" if patch_desc else "",
        "{{recommendation_patching}}": "Đánh giá kế hoạch nâng cấp patch theo chính sách vận hành và khuyến nghị bảo mật của Oracle.",
        "{{assessment_backup}}": backup_assessment,
        "{{recommendation_backup}}": backup_recommendation,
    }


def _tablespace_assessment(rows: list[list[str]]) -> dict[str, str]:
    data = _dict_rows(rows)
    over_threshold = []
    for row in data:
        pct = _to_float(row.get("PCT_USED", "") or row.get("USED_%", "") or row.get("USED", ""))
        size_gb = _to_float(row.get("SIZE_GB", ""))
        max_size_gb = _to_float(row.get("MAX_SIZE_GB", ""))
        name = row.get("TABLESPACE_NAME", "") or row.get("NAME", "")
        is_maxed = size_gb is not None and max_size_gb is not None and abs(size_gb - max_size_gb) < 0.01
        if pct is not None and pct > 85 and is_maxed and name.lower() != "total":
            over_threshold.append(name)
    if over_threshold:
        return {
            "{{assessment_tablespace_usage}}": f"Một số tablespace có dung lượng sử dụng ở mức nguy hiểm (>=85%): {', '.join(over_threshold[:10])}.",
            "{{recommendation_tablespace_usage}}": "Cung cấp thêm datafile hoặc extend datafile có sẵn cho các tablespace trên.",
        }
    return {
        "{{assessment_tablespace_usage}}": "Dung lượng của các tablespace đang ở ngưỡng an toàn.",
        "{{recommendation_tablespace_usage}}": "N/A",
    }


def _cache_hit_assessment(log_switch_tables: list[list[list[str]]], cpu_busy_tables: list[list[list[str]]]) -> dict[str, str]:
    log_switch_assessment = _log_switch_assessment(log_switch_tables)
    foreground_cpu_assessment = _foreground_cpu_assessment(cpu_busy_tables)
    return {
        "{{assessment_cache_hit}}": CACHE_HIT_ASSESSMENT,
        "{{assessment_buffer_cache_hit}}": BUFFER_CACHE_HIT_ASSESSMENT,
        "{{assessment_library_cache_hit}}": LIBRARY_CACHE_HIT_ASSESSMENT,
        "{{recommendation_cache_hit}}": CACHE_HIT_RECOMMENDATION,
        "{{recommendation_buffer_cache_hit}}": CACHE_HIT_RECOMMENDATION,
        "{{recommendation_library_cache_hit}}": CACHE_HIT_RECOMMENDATION,
        "{{assessment_log_switch}}": log_switch_assessment,
        "{{recommendation_log_switch}}": LOG_SWITCH_RECOMMENDATION,
        "{{assessment_oracle_foreground_process}}": foreground_cpu_assessment,
        "{{recommendation_oracle_foreground_process}}": CACHE_HIT_RECOMMENDATION,
        "{{assessment_pga}}": PGA_ASSESSMENT,
        "{{recommendation_pga}}": PGA_RECOMMENDATION,
    }


def _log_switch_assessment(tables: list[list[list[str]]]) -> str:
    rows = _numeric_column_rows(tables, "LOG_SWITCHES")
    values = [value for value, _row in rows]
    if not values:
        return ""
    average = sum(values) / len(values)
    minimum = min(values)
    maximum = max(values)
    display_minimum = 1 if minimum == 0 and maximum > 0 else minimum
    assessment = (
        f"Nhìn chung, tần suất log switch của các instance dao động khoảng {_format_number(display_minimum)} - {_format_number(maximum)} lần/giờ, "
        f"trung bình {_format_integer(average)} lần/giờ."
    )
    if maximum >= 30:
        peak_windows = _log_switch_peak_windows(rows)
        if peak_windows:
            assessment += (
                f" Một số thời điểm xuất hiện đột biến cao tại {', '.join(peak_windows)}, "
                "cho thấy hệ thống có hiện tượng phát sinh redo lớn trong các khung giờ cao điểm."
            )
        else:
            assessment += " Một số thời điểm xuất hiện đột biến cao, cho thấy hệ thống có hiện tượng phát sinh redo lớn trong các khung giờ cao điểm."
    return assessment


def _foreground_cpu_assessment(tables: list[list[list[str]]]) -> str:
    values = _numeric_column_values(tables, "BUSY_TIME_PERC")
    if not values:
        return ""
    average = sum(values) / len(values)
    minimum = min(values)
    maximum = max(values)
    return (
        f"Nhìn chung, các instance cơ sở dữ liệu sử dụng CPU server trung bình {_format_number(average)}%, "
        f"dao động khoảng {_format_number(minimum)}% - {_format_number(maximum)}%."
    )


def _asm_assessment(rows: list[list[str]]) -> dict[str, str]:
    warnings = []
    for row in _dict_rows(rows):
        name = row.get("NAME", "").strip()
        total_mb = _to_float(row.get("TOTAL_MB", ""))
        free_mb = _to_float(row.get("FREE_MB", "") or row.get("USABLE_FILE_MB", ""))
        if not name or free_mb is None:
            continue
        free_gb = free_mb / 1024
        free_percent = (free_mb / total_mb * 100) if total_mb and total_mb > 0 else None
        if free_gb <= ASM_FREE_WARNING_GB or (free_percent is not None and free_percent <= ASM_FREE_WARNING_PERCENT):
            warnings.append((name, free_gb))
    if not warnings:
        return {"{{assessment_asm_disk_group}}": "", "{{recommendation_asm_disk_group}}": ""}
    group, free_gb = next((item for item in warnings if item[0].upper() == "DATA"), sorted(warnings, key=lambda item: item[1])[0])
    return {
        "{{assessment_asm_disk_group}}": (
            f"Disk group {group} chỉ còn trống {_format_number(free_gb)}GB. "
            "Nếu không đủ dung lượng cung cấp cho hệ thống sẽ gây gián đoạn và ảnh hưởng đến hoạt động hệ thống."
        ),
        "{{recommendation_asm_disk_group}}": (
            f"Cấp thêm đĩa 500GB cho disk group {group}. "
            f"Sau đó HPT sẽ tiến hành thêm đĩa mới vào disk group {group}."
        ),
    }


def _scheduler_jobs_assessment(rows: list[list[str]]) -> dict[str, str]:
    failed_enabled_jobs = []
    for row in _dict_rows(rows):
        enabled = _first_present(row, ("ENABLED", "ENABLE", "ENABL"))
        failure_count = _to_float(row.get("FAILURE_COUNT", ""))
        if enabled.strip().upper() == "TRUE" and failure_count is not None and failure_count > 0:
            failed_enabled_jobs.append(row)
    if not failed_enabled_jobs:
        return {"{{assessment_sche_job}}": "", "{{recommendation_sche_job}}": ""}
    return {
        "{{assessment_sche_job}}": "Một số job đang enable nhưng có ghi nhận lỗi trong quá trình chạy.",
        "{{recommendation_sche_job}}": (
            "Kiểm tra lại các job đang enable và có FAILURE_COUNT > 0 để tránh ảnh hưởng đến hoạt động của hệ thống/ ứng dụng."
        ),
    }


def _count_assessments(
    no_index: list[list[str]],
    no_pk: list[list[str]],
    invalid_objects: list[list[str]],
    table_stats: list[list[str]],
    index_stats: list[list[str]],
) -> dict[str, str]:
    table_count = max(0, len(table_stats) - 1)
    index_count = max(0, len(index_stats) - 1)
    return {
        "{{assessment_no_index}}": f"Hệ thống đang có {max(0, len(no_index) - 1)} bảng không có index.",
        "{{recommendation_no_index}}": "Xem xét tạo index cho các bảng cần thiết để tăng tốc độ truy vấn.",
        "{{assessment_no_pk}}": f"Hệ thống đang có {max(0, len(no_pk) - 1)} bảng không có khóa chính.",
        "{{recommendation_no_pk}}": "Xem xét khởi tạo khoá chính hoặc unique index cho các bảng cần thiết để tăng tốc truy xuất dữ liệu.",
        "{{assessment_invalid_objects}}": f"Hệ thống có {max(0, len(invalid_objects) - 1)} invalid objects.",
        "{{recommendation_invalid_objects}}": "Thực hiện recompile lại các đối tượng để hợp lệ hoá các đối tượng, tránh ảnh hưởng đến ứng dụng hoặc hệ thống.",
        "{{assessment_stale_stats}}": f"Hệ thống có tổng cộng {table_count} tables with stale stats và {index_count} indexes with stale stats.",
        "{{recommendation_stale_stats}}": "Thu thập lại (re-gather) statistic của các đối tượng có stale statistics.",
    }


def _first_table(root: Path, pattern: str) -> list[list[str]]:
    path = next(iter(sorted(root.rglob(pattern))), None)
    if not path:
        return []
    page, soup, _html = parse_html_file(path)
    tables = extract_tables(page, soup)
    return tables[0].rows if tables else []


def _first_table_any(root: Path, patterns: list[str]) -> list[list[str]]:
    for pattern in patterns:
        rows = _first_table(root, pattern)
        if rows:
            return rows
    return []


def _tables_for(root: Path, pattern: str, exclude: tuple[str, ...] = ()) -> list[list[list[str]]]:
    result = []
    for path in sorted(root.rglob(pattern)):
        name = path.name.lower()
        if any(token in name for token in exclude):
            continue
        page, soup, _html = parse_html_file(path)
        tables = extract_tables(page, soup)
        if tables:
            result.append(tables[0].rows)
    return result


def _dict_rows(rows: list[list[str]]) -> list[dict[str, str]]:
    if len(rows) < 2:
        return []
    headers = [_normalize_header(item) for item in rows[0]]
    result = []
    for row in rows[1:]:
        result.append({headers[index]: row[index] for index in range(min(len(headers), len(row)))})
    return result


def _normalize_header(value: str) -> str:
    return re.sub(r"\s+", "_", value.strip().upper())


def _storage_root(value: str) -> str:
    if value.startswith("+"):
        return value.split("/", 1)[0]
    normalized = value.replace("\\", "/")
    parts = [part for part in normalized.split("/") if part]
    if len(parts) >= 1:
        return f"/{parts[0]}"
    return value


def _preferred_memory_value(row: dict[str, str]) -> str:
    current = str(row.get("CURRENT_GB", "") or "").strip()
    spfile = str(row.get("SPFILE_VALUE", "") or "").strip()
    current_number = _to_float(current)
    if current and current_number is not None and current_number > 0:
        return current
    if spfile:
        return spfile
    return current


def _to_float(value: str) -> float | None:
    text = str(value or "").strip().replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


def _max_numeric(values: list[str]) -> float | None:
    numbers = [number for value in values if (number := _to_float(value)) is not None]
    return max(numbers) if numbers else None


def _positive_values(values: list[str]) -> list[str]:
    return [value for value in values if (number := _to_float(value)) is not None and number > 0]


def _first_present(row: dict[str, str], keys: tuple[str, ...]) -> str:
    for key in keys:
        value = row.get(key, "")
        if value:
            return str(value)
    return ""


def _numeric_column_values(tables: list[list[list[str]]], column_name: str) -> list[float]:
    return [value for value, _row in _numeric_column_rows(tables, column_name)]


def _numeric_column_rows(tables: list[list[list[str]]], column_name: str) -> list[tuple[float, dict[str, str]]]:
    result = []
    for rows in tables:
        for row in _dict_rows(rows):
            number = _to_float(row.get(column_name, ""))
            if number is not None:
                result.append((number, row))
    return result


def _format_number(value: float) -> str:
    if abs(value - round(value)) < 0.05:
        return str(int(round(value)))
    return f"{value:.1f}"


def _format_integer(value: float) -> str:
    return str(int(round(value)))


def _log_switch_peak_windows(rows: list[tuple[float, dict[str, str]]]) -> list[str]:
    peak_rows = sorted((item for item in rows if item[0] >= 30), key=lambda item: item[0], reverse=True)
    windows = []
    for value, row in peak_rows[:3]:
        begin_time = row.get("BEGIN_TIME", "").strip()
        end_time = row.get("END_TIME", "").strip()
        if begin_time and end_time:
            windows.append(f"{begin_time} - {end_time} ({_format_integer(value)} lần/giờ)")
        elif begin_time:
            windows.append(f"{begin_time} ({_format_integer(value)} lần/giờ)")
    return windows
