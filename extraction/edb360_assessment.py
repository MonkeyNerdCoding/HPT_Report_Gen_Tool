from __future__ import annotations

from pathlib import Path
from collections import defaultdict
from datetime import datetime
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
EN_CACHE_HIT_ASSESSMENT = "The buffer cache hit and library cache hit ratios are currently at an optimal level of 99 - 100%."
EN_BUFFER_CACHE_HIT_ASSESSMENT = "The buffer cache hit ratio is currently at an optimal level of 99 - 100%."
EN_LIBRARY_CACHE_HIT_ASSESSMENT = "The library cache hit ratio is currently at an optimal level of 99 - 100%."
EN_LOG_SWITCH_RECOMMENDATION = "Increase the redo log file size to reduce log switch frequency."
EN_PGA_ASSESSMENT = "Overall, PGA memory usage remains within a safe range."
ASM_FREE_WARNING_PERCENT = 10
MULTIPLEXED_REDO_ASSESSMENT = (
    "Theo như cấu hình hiện tại, các redo log group đang được multiplexing, tức là mỗi redo log group có 2 members, "
    "mỗi member được đặt ở một disk controller khác nhau. Điều này tăng tính sẵn sàng của redo log group, đảm bảo database "
    "luôn được vận hành khi có sự cố ảnh hưởng đến một trong những member trong redo log group."
)
EN_MULTIPLEXED_REDO_ASSESSMENT = (
    "Based on the current configuration, the redo log groups are multiplexed. Each redo log group has 2 members, "
    "and each member is placed on a different disk controller. This improves redo log availability and helps ensure "
    "the database can continue operating if one redo log member is affected."
)


def build_edb360_assessment_mapping(input_root: str | Path, language: str = "vi") -> dict[str, str]:
    root = Path(input_root)
    english = _is_english(language)
    mapping: dict[str, str] = {}

    all_parameters = _first_table(root, "*all_parameters.html")
    memory_configuration = _first_table(root, "*memory_configuration.html")
    redo_log = _first_table(root, "*redo_log.html")
    redo_log_files = _first_table(root, "*redo_log_files.html")
    registry_sql_patch = _first_table(root, "*registry_sql_patch.html")
    rman_backup = _rman_backup_table(root)
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

    mapping.update(_control_file_assessment(all_parameters, english))
    mapping.update(_redo_assessment(redo_log, redo_log_files, english))
    mapping.update(_memory_assessment(memory_configuration, english))
    mapping.update(_patching_backup_assessment(registry_sql_patch, rman_backup, english))
    mapping.update(_tablespace_assessment(tablespace_usage, english))
    mapping.update(_cache_hit_assessment(log_switch_tables, cpu_busy_tables, english))
    mapping.update(_asm_assessment(asm_disk_group, english))
    mapping.update(_scheduler_jobs_assessment(scheduler_jobs, english))
    mapping.update(_count_assessments(no_index, no_pk, invalid_objects, table_stats, index_stats, english))
    return {key: value for key, value in mapping.items() if value is not None}


def _control_file_assessment(rows: list[list[str]], english: bool = False) -> dict[str, str]:
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
        if english:
            assessment = (
                f"The database currently has {count} control files, and these control files are located on "
                f"{len(locations)} separate storage locations ({', '.join(sorted(locations))}), ensuring control file protection."
            )
        else:
            assessment = (
                f"Hiện tại, database đang có {count} control files và các control files này được đặt trên "
                f"{len(locations)} vị trí lưu trữ khác nhau ({', '.join(sorted(locations))}), đảm bảo tính an toàn cho control files."
            )
        recommendation = "N/A"
    else:
        location = next(iter(locations), "")
        if english:
            assessment = (
                f"The database currently has {count} control files. However, these control files are located in the same storage location "
                f"{location}, which does not provide sufficient protection if that storage location has an incident."
            )
            recommendation = "Place the control files on different partitions/disk groups or add a control file on independent storage."
        else:
            assessment = (
                f"Hiện tại, database đang có {count} control files. Tuy nhiên, các control files này đang nằm cùng một vị trí "
                f"{location}, không đảm bảo tính an toàn nếu vị trí lưu trữ này gặp sự cố."
            )
            recommendation = "Đưa các control files ra nhiều phân vùng/disk group khác nhau hoặc bổ sung control file ở vị trí lưu trữ độc lập."
    return {
        "{{assessment_control_file}}": assessment,
        "{{recommendation_control_file}}": recommendation,
    }


def _redo_assessment(redo_rows: list[list[str]], redo_file_rows: list[list[str]], english: bool = False) -> dict[str, str]:
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
        assessment = EN_MULTIPLEXED_REDO_ASSESSMENT if english else MULTIPLEXED_REDO_ASSESSMENT
        recommendation = "N/A"
    else:
        if english:
            assessment = f"Based on the current configuration, the database has {group_count} redo log groups, but at least one redo log group is not fully multiplexed."
            recommendation = "Add redo log members on another disk group/partition to improve availability."
        else:
            assessment = (
                f"Theo cấu hình hiện tại, database có {group_count} redo log groups nhưng có redo log group chưa được multiplexing đầy đủ."
            )
            recommendation = "Bổ sung redo log member trên disk group/phân vùng khác để tăng tính sẵn sàng."
    return {
        "{{assessment_redo_log}}": assessment,
        "{{recommendation_redo_log}}": recommendation,
    }


def _memory_assessment(rows: list[list[str]], english: bool = False) -> dict[str, str]:
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
    detail = "; ".join(parts) if parts else ("SGA/PGA could not be identified from the system data" if english else "chưa xác định được SGA/PGA từ hệ thống")
    assessment = (
        f"The database memory is currently configured in {mode} mode. Details: {detail}."
        if english
        else f"Cơ sở dữ liệu hiện tại đang được cấu hình vùng nhớ ở chế độ {mode}. Trong đó: {detail}."
    )
    return {
        "{{assessment_memory_configuration}}": assessment,
        "{{recommendation_memory_configuration}}": "N/A",
    }


def _patching_backup_assessment(registry_rows: list[list[str]], backup_rows: list[list[str]], english: bool = False) -> dict[str, str]:
    patches = _dict_rows(registry_rows)
    backups = _dict_rows(backup_rows)
    patch_desc = ""
    if patches:
        latest = patches[-1]
        patch_desc = latest.get("DESCRIPTION") or latest.get("VERSION") or latest.get("PATCH_ID", "")
    patch_major_version = _oracle_major_version(latest) if patches else None

    completed = [row for row in backups if "COMPLETED" in row.get("STATUS", "").upper()]
    if backups:
        if english:
            completed_text = f" {len(completed)} backup job(s) completed." if completed else ""
            backup_assessment = f"RMAN backup data is available. The system recorded {len(backups)} backup job(s) in the collected data.{completed_text}"
            backup_recommendation = (
                "Prepare an environment to perform backup restore testing. "
                "Without a restore test environment, backup recoverability cannot be confirmed when needed."
            )
        else:
            completed_text = f" Trong đó có {len(completed)} backup job hoàn thành." if completed else ""
            backup_assessment = f"Đã có RMAN backup. Hệ thống ghi nhận {len(backups)} backup job trong dữ liệu thu thập.{completed_text}"
            backup_recommendation = (
                "Khuyến nghị chuẩn bị môi trường thực hiện kiểm thử restore các bản backup. "
                "Việc không có môi trường khôi phục kiểm thử bản backup sẽ không đảm bảo bản backup có thể khôi phục thành công khi cần thiết."
            )
    else:
        backup_assessment = "No RMAN backup is currently available" if english else "Hiện tại chưa có RMAN backup"
        backup_recommendation = (
            "Create RMAN backups to enable full recovery when the system encounters an incident."
            if english
            else "Khuyến nghị tạo thêm RMAN backup để thực hiện khôi phục hoàn toàn khi hệ thống\ngặp sự cố."
        )

    patching_19c_recommendation = ""
    if patch_major_version is not None and patch_major_version < 19:
        patching_19c_recommendation = (
            "Recommendation: upgrade to Oracle Database 19C to take advantage of performance, security, and the best vendor support."
            if english
            else "Khuyến nghị: nâng cấp lên phiên bản 19C để tận dụng các tính năng về performance, bảo mật và sự hỗ trợ tốt nhất từ hãng"
        )

    return {
        "{{assessment_patching}}": (f"Current version/patch: {patch_desc}" if english else f"Phiên bản/patch hiện tại: {patch_desc}") if patch_desc else "",
        "{{recommendation_patching_19c}}": patching_19c_recommendation,
        "{{recommendation_patching}}": (
            "Review the patch upgrade plan according to operational policy and Oracle security recommendations."
            if english
            else "Đánh giá kế hoạch nâng cấp patch theo chính sách vận hành và khuyến nghị bảo mật của Oracle."
        ),
        "{{assessment_backup}}": backup_assessment,
        "{{recommendation_backup}}": backup_recommendation,
    }


def _tablespace_assessment(rows: list[list[str]], english: bool = False) -> dict[str, str]:
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
            "{{assessment_tablespace_usage}}": (
                f"Some tablespaces are at a critical usage level (>=85%): {', '.join(over_threshold[:10])}."
                if english
                else f"Một số tablespace có dung lượng sử dụng ở mức nguy hiểm (>=85%): {', '.join(over_threshold[:10])}."
            ),
            "{{recommendation_tablespace_usage}}": (
                "Add datafiles or extend existing datafiles for the tablespaces above."
                if english
                else "Cung cấp thêm datafile hoặc extend datafile có sẵn cho các tablespace trên."
            ),
        }
    return {
        "{{assessment_tablespace_usage}}": "Tablespace usage is within a safe range." if english else "Dung lượng của các tablespace đang ở ngưỡng an toàn.",
        "{{recommendation_tablespace_usage}}": "N/A",
    }


def _cache_hit_assessment(log_switch_tables: list[list[list[str]]], cpu_busy_tables: list[list[list[str]]], english: bool = False) -> dict[str, str]:
    log_switch_assessment = _log_switch_assessment(log_switch_tables, english)
    foreground_cpu_assessment = _foreground_cpu_assessment(cpu_busy_tables, english)
    return {
        "{{assessment_cache_hit}}": EN_CACHE_HIT_ASSESSMENT if english else CACHE_HIT_ASSESSMENT,
        "{{assessment_buffer_cache_hit}}": EN_BUFFER_CACHE_HIT_ASSESSMENT if english else BUFFER_CACHE_HIT_ASSESSMENT,
        "{{assessment_library_cache_hit}}": EN_LIBRARY_CACHE_HIT_ASSESSMENT if english else LIBRARY_CACHE_HIT_ASSESSMENT,
        "{{recommendation_cache_hit}}": CACHE_HIT_RECOMMENDATION,
        "{{recommendation_buffer_cache_hit}}": CACHE_HIT_RECOMMENDATION,
        "{{recommendation_library_cache_hit}}": CACHE_HIT_RECOMMENDATION,
        "{{assessment_log_switch}}": log_switch_assessment,
        "{{recommendation_log_switch}}": EN_LOG_SWITCH_RECOMMENDATION if english else LOG_SWITCH_RECOMMENDATION,
        "{{assessment_oracle_foreground_process}}": foreground_cpu_assessment,
        "{{recommendation_oracle_foreground_process}}": CACHE_HIT_RECOMMENDATION,
        "{{assessment_pga}}": EN_PGA_ASSESSMENT if english else PGA_ASSESSMENT,
        "{{recommendation_pga}}": PGA_RECOMMENDATION,
    }


def _log_switch_assessment(tables: list[list[list[str]]], english: bool = False) -> str:
    rows = _numeric_column_rows(tables, "LOG_SWITCHES")
    if not rows:
        return ""

    summaries = _log_switch_instance_summaries(rows)
    if not summaries:
        return ""
    assessment = _log_switch_combined_summary_text(summaries, english)
    notable_windows = _log_switch_peak_windows(rows)
    if notable_windows:
        assessment += (
            f" Notable peak windows: {', '.join(notable_windows)}."
            if english
            else f" Các khung giờ nổi bật: {', '.join(notable_windows)}."
        )
    return assessment


def _foreground_cpu_assessment(tables: list[list[list[str]]], english: bool = False) -> str:
    rows = _foreground_cpu_rows(tables)
    if not rows:
        return ""
    values_by_instance: dict[str, list[float]] = defaultdict(list)
    for value, row in rows:
        instance = _log_switch_instance_label(row) or "instance không xác định"
        values_by_instance[instance].append(value)

    parts = []
    for instance, values in sorted(values_by_instance.items()):
        average = sum(values) / len(values)
        minimum = min(values)
        maximum = max(values)
        if english:
            parts.append(
                f"{instance}: average {_format_number(average)}%, ranging from {_format_number(minimum)}% to {_format_number(maximum)}%"
            )
        else:
            parts.append(
                f"{instance}: trung bình {_format_number(average)}%, dao động khoảng {_format_number(minimum)}% - {_format_number(maximum)}%"
            )

    if english:
        return f"Overall, server CPU used by database foreground processes by instance is {', '.join(parts)}."
    return f"Nhìn chung, các instance cơ sở dữ liệu sử dụng CPU server theo từng instance như sau: {', '.join(parts)}."


def _asm_assessment(rows: list[list[str]], english: bool = False) -> dict[str, str]:
    warnings = []
    for row in _dict_rows(rows):
        name = row.get("NAME", "").strip()
        total_mb = _to_float(row.get("TOTAL_MB", ""))
        free_mb = _to_float(row.get("FREE_MB", "") or row.get("USABLE_FILE_MB", ""))
        if not name or free_mb is None:
            continue
        free_gb = free_mb / 1024
        free_percent = (free_mb / total_mb * 100) if total_mb and total_mb > 0 else None
        if free_percent is not None and free_percent <= ASM_FREE_WARNING_PERCENT:
            warnings.append((name, free_gb))
    if not warnings:
        return {"{{assessment_asm_disk_group}}": "", "{{recommendation_asm_disk_group}}": ""}
    warnings = sorted(warnings, key=lambda item: item[0].upper())
    group_names = ", ".join(group for group, _free_gb in warnings)
    english_details = ", ".join(f"disk group {group} has only {_format_number(free_gb)}GB free" for group, free_gb in warnings)
    vietnamese_details = ", ".join(f"Disk group {group} chỉ còn trống {_format_number(free_gb)}GB" for group, free_gb in warnings)
    return {
        "{{assessment_asm_disk_group}}": (
            f"{english_details}. "
            "Insufficient capacity may interrupt and affect system operations."
            if english
            else (
                f"{vietnamese_details}. "
                "Nếu không đủ dung lượng cung cấp cho hệ thống sẽ gây gián đoạn và ảnh hưởng đến hoạt động hệ thống."
            )
        ),
        "{{recommendation_asm_disk_group}}": (
            f"Add disk capacity to disk group {group_names}. HPT will then add the new disk to disk group {group_names}."
            if english
            else (
                f"Cấp thêm đĩa cho disk group {group_names}. "
                f"Sau đó HPT sẽ tiến hành thêm đĩa mới vào disk group {group_names}."
            )
        ),
    }


def _scheduler_jobs_assessment(rows: list[list[str]], english: bool = False) -> dict[str, str]:
    failed_enabled_jobs = []
    for row in _dict_rows(rows):
        enabled = _first_present(row, ("ENABLED", "ENABLE", "ENABL"))
        failure_count = _to_float(row.get("FAILURE_COUNT", ""))
        if enabled.strip().upper() == "TRUE" and failure_count is not None and failure_count > 0:
            failed_enabled_jobs.append(row)
    if not failed_enabled_jobs:
        return {"{{assessment_sche_job}}": "", "{{recommendation_sche_job}}": ""}
    return {
        "{{assessment_sche_job}}": "Some enabled jobs have recorded failures during execution." if english else "Một số job đang enable nhưng có ghi nhận lỗi trong quá trình chạy.",
        "{{recommendation_sche_job}}": (
            "Review enabled jobs with FAILURE_COUNT > 0 to avoid impact on system/application operations."
            if english
            else "Kiểm tra lại các job đang enable và có FAILURE_COUNT > 0 để tránh ảnh hưởng đến hoạt động của hệ thống/ ứng dụng."
        ),
    }


def _count_assessments(
    no_index: list[list[str]],
    no_pk: list[list[str]],
    invalid_objects: list[list[str]],
    table_stats: list[list[str]],
    index_stats: list[list[str]],
    english: bool = False,
) -> dict[str, str]:
    table_count = max(0, len(table_stats) - 1)
    index_count = max(0, len(index_stats) - 1)
    if english:
        return {
            "{{assessment_no_index}}": f"The system currently has {max(0, len(no_index) - 1)} table(s) without indexes.",
            "{{recommendation_no_index}}": "Consider creating indexes for necessary tables to improve query performance.",
            "{{assessment_no_pk}}": f"The system currently has {max(0, len(no_pk) - 1)} table(s) without primary keys.",
            "{{recommendation_no_pk}}": "Consider creating primary keys or unique indexes for necessary tables to improve data access performance.",
            "{{assessment_invalid_objects}}": f"The system has {max(0, len(invalid_objects) - 1)} invalid object(s).",
            "{{recommendation_invalid_objects}}": "Recompile invalid objects to validate them and avoid impact on applications or the system.",
            "{{assessment_stale_stats}}": f"The system has a total of {table_count} tables with stale stats and {index_count} indexes with stale stats.",
            "{{recommendation_stale_stats}}": "Re-gather statistics for objects with stale statistics.",
        }
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


def _is_english(language: str | None) -> bool:
    return str(language or "").strip().lower() in {"en", "eng", "english"}


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


def _rman_backup_table(root: Path) -> list[list[str]]:
    candidates = []
    for path in sorted(root.rglob("*rman_backup*.html")):
        name = path.name.lower()
        if name.endswith("_line_chart.html"):
            continue
        page, soup, _html = parse_html_file(path)
        title = page.title.strip().lower()
        stem = path.stem.lower()
        is_backup_page = stem.endswith("_rman_backup") or stem.endswith("_rman_backup_job_details")
        is_backup_title = title in {"rman backup", "rman backup job details"}
        if not is_backup_page and not is_backup_title:
            continue
        tables = extract_tables(page, soup)
        if tables:
            priority = 0 if "job_details" in stem or title == "rman backup job details" else 1
            candidates.append((priority, path.name, tables[0].rows))
    if not candidates:
        return []
    return sorted(candidates, key=lambda item: (item[0], item[1]))[0][2]


def _tables_for(root: Path, pattern: str, exclude: tuple[str, ...] = ()) -> list[list[list[str]]]:
    result = []
    for path in sorted(root.rglob(pattern)):
        name = path.name.lower()
        if any(token in name for token in exclude):
            continue
        page, soup, _html = parse_html_file(path)
        tables = extract_tables(page, soup)
        if tables:
            result.append(_rows_with_source_file(tables[0].rows, path.name))
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
    match = re.search(r"-?(?:\d+(?:\.\d+)?|\.\d+)", text)
    if not match:
        return None
    return float(match.group(0))


def _oracle_major_version(row: dict[str, str]) -> int | None:
    candidates = [
        row.get("VERSION", ""),
        row.get("DESCRIPTION", ""),
        row.get("ACTION_TIME", ""),
        row.get("BUNDLE_SERIES", ""),
    ]
    for value in candidates:
        text = str(value or "")
        match = re.search(r"\b(1[0-9]|2[0-9])(?:c|\.\d)", text, flags=re.IGNORECASE)
        if match:
            return int(match.group(1))
    return None


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


def _value_at(row: list[str], index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    return row[index]


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


def _foreground_cpu_rows(tables: list[list[list[str]]]) -> list[tuple[float, dict[str, str]]]:
    result = []
    for rows in tables:
        if len(rows) < 2:
            continue
        headers = [_normalize_header(item) for item in rows[0]]
        if "BUSY_TIME_PERC" not in headers:
            continue
        busy_index = headers.index("BUSY_TIME_PERC")
        idle_index = headers.index("IDLE_TIME_PERC") if "IDLE_TIME_PERC" in headers else None
        source_index = headers.index("SOURCE_FILE") if "SOURCE_FILE" in headers else None
        instance_indexes = {
            header: headers.index(header)
            for header in ("INSTANCE_NUMBER", "INSTANCE_NAME", "INSTANCE", "INST_ID", "INSTANCE_ID")
            if header in headers
        }
        for row_values in rows[1:]:
            busy = _to_float(_value_at(row_values, busy_index))
            idle = _to_float(_value_at(row_values, idle_index))
            if busy is None:
                continue
            if idle is not None and busy >= 90 and idle <= 10 and abs((busy + idle) - 100) <= 1:
                busy = idle
            row = {key: _value_at(row_values, index) for key, index in instance_indexes.items()}
            if source_index is not None:
                row["SOURCE_FILE"] = _value_at(row_values, source_index)
            result.append((busy, row))
    return result


def _rows_with_source_file(rows: list[list[str]], source_file: str) -> list[list[str]]:
    if not rows:
        return rows
    headers = [_normalize_header(item) for item in rows[0]]
    if "SOURCE_FILE" in headers:
        return rows
    return [rows[0] + ["SOURCE_FILE"]] + [row + [source_file] for row in rows[1:]]


def _format_number(value: float) -> str:
    if abs(value - round(value)) < 0.05:
        return str(int(round(value)))
    return f"{value:.1f}"


def _format_integer(value: float) -> str:
    return str(int(round(value)))


def _format_switch_range(minimum: float, maximum: float, english: bool = False) -> str:
    unit = "times/hour" if english else "lần/giờ"
    minimum = _display_log_switch_value(minimum, maximum)
    maximum = _display_log_switch_value(maximum, maximum)
    if abs(minimum - maximum) < 0.05:
        return f"{_format_number(maximum)} {unit}"
    return f"{_format_number(minimum)} - {_format_number(maximum)} {unit}"


def _display_log_switch_value(value: float, maximum: float) -> float:
    if value == 0 and maximum >= 0:
        return 1
    return value


def _log_switch_level(value: float) -> str:
    if value <= 6:
        return "optimal"
    if value <= 10:
        return "stable"
    if value <= 20:
        return "elevated"
    return "high"


def _log_switch_level_label(level: str, english: bool = False) -> str:
    labels = {
        "optimal": ("optimal", "tối ưu"),
        "stable": ("stable", "ổn định"),
        "elevated": ("quite high, should be monitored", "khá cao, nên theo dõi"),
        "high": ("high, should be checked", "cao, nên kiểm tra"),
    }
    english_label, vietnamese_label = labels[level]
    return english_label if english else vietnamese_label


def _log_switch_instance_summaries(rows: list[tuple[float, dict[str, str]]]) -> list[dict[str, object]]:
    values_by_instance: dict[str, list[float]] = defaultdict(list)
    for value, row in rows:
        instance = _log_switch_instance_label(row) or "instance không xác định"
        values_by_instance[instance].append(value)

    summaries = []
    for instance, values in sorted(values_by_instance.items()):
        grouped_values: dict[str, list[float]] = defaultdict(list)
        for value in values:
            grouped_values[_log_switch_level(value)].append(value)
        total_hours = len(values)
        ratios = {
            level: len(grouped_values[level]) / total_hours
            for level in ("optimal", "stable", "elevated", "high")
        }
        healthy_ratio = ratios["optimal"] + ratios["stable"]
        risk_ratio = ratios["elevated"] + ratios["high"]
        if ratios["optimal"] >= 0.70:
            overall = "optimal"
        elif ratios["stable"] >= 0.70:
            overall = "stable"
        elif ratios["elevated"] >= 0.70:
            overall = "elevated"
        elif ratios["high"] >= 0.70:
            overall = "high"
        elif healthy_ratio >= 0.70:
            overall = "optimal_to_stable"
        elif risk_ratio >= 0.50:
            overall = "frequent_elevated"
        else:
            overall = "mixed"

        summaries.append(
            {
                "instance": instance,
                "overall": overall,
                "max": max(values),
                "ratios": ratios,
                "ranges": {
                    level: _range_for_values(grouped_values[level])
                    for level in ("optimal", "stable", "elevated", "high")
                },
                "healthy_range": _range_for_values(grouped_values["optimal"] + grouped_values["stable"]),
                "risk_range": _range_for_values(grouped_values["elevated"] + grouped_values["high"]),
            }
        )
    return summaries


def _log_switch_peak_windows(rows: list[tuple[float, dict[str, str]]]) -> list[str]:
    peak_rows = sorted(rows, key=lambda item: item[0], reverse=True)
    windows = []
    for value, row in peak_rows[:3]:
        begin_time = row.get("BEGIN_TIME", "").strip()
        end_time = row.get("END_TIME", "").strip()
        instance = _log_switch_instance_label(row)
        prefix = f"{instance} " if instance else ""
        if begin_time and end_time:
            windows.append(f"{prefix}{_format_log_switch_window(begin_time, end_time)} ({_format_integer(value)} lần/giờ)")
        elif begin_time:
            windows.append(f"{prefix}{_format_log_switch_time(begin_time)} ({_format_integer(value)} lần/giờ)")
    return windows


def _range_for_values(values: list[float]) -> tuple[float, float] | None:
    if not values:
        return None
    return min(values), max(values)


def _format_log_switch_window(begin_time: str, end_time: str) -> str:
    begin = _parse_log_switch_datetime(begin_time)
    end = _parse_log_switch_datetime(end_time)
    if not begin:
        return f"{begin_time} - {end_time}"
    date_text = f"{begin.day}/{begin.month}/{begin:%y}"
    begin_hour = _format_hour_minute(begin)
    if not end:
        return f"{date_text} ({begin_hour})"
    end_hour = _format_hour_minute(end)
    return f"{date_text} ({begin_hour}-{end_hour})"


def _format_log_switch_time(value: str) -> str:
    parsed = _parse_log_switch_datetime(value)
    if not parsed:
        return value
    return f"{parsed.day}/{parsed.month}/{parsed:%y} ({_format_hour_minute(parsed)})"


def _parse_log_switch_datetime(value: str) -> datetime | None:
    text = value.strip()
    for pattern in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%m/%d/%y %H:%M", "%m/%d/%Y %H:%M"):
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            continue
    return None


def _format_hour_minute(value: datetime) -> str:
    return f"{value.hour}h{value.minute:02d}"


def _format_switch_range_value(value_range: tuple[float, float] | None, english: bool = False) -> str:
    if value_range is None:
        return ""
    return _format_switch_range(value_range[0], value_range[1], english)


def _log_switch_combined_summary_text(summaries: list[dict[str, object]], english: bool = False) -> str:
    if not summaries:
        return ""
    if len(summaries) == 1:
        return _log_switch_summary_text(summaries[0], english)

    overall_groups: dict[str, list[dict[str, object]]] = defaultdict(list)
    for summary in summaries:
        overall_groups[str(summary["overall"])].append(summary)

    if len(overall_groups) == 1:
        overall = next(iter(overall_groups))
        return _log_switch_same_pattern_text(overall, summaries, english)
    return _log_switch_mixed_pattern_text(overall_groups, english)


def _log_switch_same_pattern_text(overall: str, summaries: list[dict[str, object]], english: bool = False) -> str:
    primary_range_parts = _log_switch_range_parts(summaries, _primary_log_switch_range_key(overall), english)
    risk_summaries = [summary for summary in summaries if summary.get("risk_range")]
    risk_range_parts = _log_switch_range_parts(risk_summaries, "risk_range", english)
    vietnamese_risk_range_parts = _log_switch_range_parts_by_range_first(risk_summaries) if not english else ""
    max_summary = max(summaries, key=lambda item: float(item["max"]))
    max_text = _format_switch_range(float(max_summary["max"]), float(max_summary["max"]), english)
    max_instance = str(max_summary["instance"])

    if english:
        subject = "database instances"
        if overall == "optimal":
            text = f"Overall, log switch frequency of the {subject} is optimal for most of the time, {primary_range_parts}."
        elif overall == "stable":
            text = f"Overall, log switch frequency of the {subject} is stable for most of the time, {primary_range_parts}."
        elif overall == "elevated":
            text = f"Overall, log switch frequency of the {subject} is quite high for most of the time, {primary_range_parts}."
        elif overall == "high":
            return f"Log switch frequency of the {subject} remains high for most of the time, {primary_range_parts}. Review online redo log size, redo-generating workload, and signs of overly frequent log switches."
        elif overall == "optimal_to_stable":
            text = f"Overall, log switch frequency of the {subject} remains optimal to stable for most of the time, {primary_range_parts}."
        elif overall == "frequent_elevated":
            text = f"Overall, log switch frequency of the {subject} frequently reaches quite high to high levels, {primary_range_parts}. Review redo log size and redo-generating workload."
        else:
            text = "Overall, log switch frequency of the database instances varies across time windows, without a clearly dominant level."
        if risk_range_parts and overall not in {"elevated", "high", "frequent_elevated"}:
            text += f" During higher-load periods, log switch frequency increased to {risk_range_parts} and should continue to be monitored."
        if float(max_summary["max"]) > 20:
            text += f" The highest recorded value is {max_text} at {max_instance}; periods above 20 times/hour should be checked if they repeat or persist."
        else:
            text += f" The highest recorded value is {max_text} at {max_instance}."
        return text

    if overall == "optimal":
        text = f"Nhìn chung, tần suất log switch của các instance trong phần lớn thời gian ở mức tối ưu, {primary_range_parts}."
    elif overall == "stable":
        text = f"Nhìn chung, tần suất log switch của các instance trong phần lớn thời gian ở mức ổn định, {primary_range_parts}."
    elif overall == "elevated":
        text = f"Nhìn chung, tần suất log switch của các instance trong phần lớn thời gian ở mức khá cao, {primary_range_parts}. Tần suất này nên được tiếp tục theo dõi và đối chiếu với kích thước redo log cũng như workload phát sinh redo."
    elif overall == "high":
        return f"Tần suất log switch của các instance duy trì ở mức cao trong phần lớn thời gian, {primary_range_parts}. Nên kiểm tra kích thước online redo log, workload phát sinh redo và các dấu hiệu log switch quá thường xuyên."
    elif overall == "optimal_to_stable":
        text = f"Nhìn chung, tần suất log switch của các instance trong phần lớn thời gian duy trì ở mức tối ưu đến ổn định, {primary_range_parts}."
    elif overall == "frequent_elevated":
        text = f"Nhìn chung, tần suất log switch của các instance ghi nhận mức khá cao đến cao xuất hiện thường xuyên, {primary_range_parts}. Nên kiểm tra kích thước redo log và workload phát sinh redo."
    else:
        text = "Nhìn chung, tần suất log switch của các instance có sự biến động giữa các khung giờ, không có mức nào chiếm ưu thế rõ."

    if risk_range_parts and overall not in {"elevated", "high", "frequent_elevated"}:
        text += f" Tại một số khung giờ tải cao, tần suất log switch tăng lên: {vietnamese_risk_range_parts}, thuộc mức khá cao và nên tiếp tục theo dõi."
    if float(max_summary["max"]) > 20:
        text += f" Mức cao nhất ghi nhận là {max_text} tại {max_instance}; các thời điểm vượt 20 lần/giờ nên được kiểm tra thêm nếu xuất hiện lặp lại hoặc kéo dài."
    else:
        text += f" Mức cao nhất ghi nhận là {max_text} tại {max_instance}."
    return text


def _log_switch_mixed_pattern_text(overall_groups: dict[str, list[dict[str, object]]], english: bool = False) -> str:
    parts = []
    for overall in ("optimal", "stable", "optimal_to_stable", "elevated", "high", "frequent_elevated", "mixed"):
        summaries = overall_groups.get(overall)
        if not summaries:
            continue
        range_parts = _log_switch_range_parts(summaries, _primary_log_switch_range_key(overall), english)
        names = _join_display_list([str(summary["instance"]) for summary in summaries], english)
        label = _log_switch_overall_label(overall, english)
        if english:
            parts.append(f"{names} is {label}, {range_parts}")
        else:
            parts.append(f"{names} ở mức {label}, {range_parts}")

    all_summaries = [summary for summaries in overall_groups.values() for summary in summaries]
    max_summary = max(all_summaries, key=lambda item: float(item["max"]))
    max_text = _format_switch_range(float(max_summary["max"]), float(max_summary["max"]), english)
    max_instance = str(max_summary["instance"])
    if english:
        text = f"Overall, log switch frequency differs between instances: {'; '.join(parts)}."
        text += f" The highest recorded value is {max_text} at {max_instance}."
    else:
        text = f"Nhìn chung, tần suất log switch có sự khác biệt giữa các instance: {'; '.join(parts)}."
        text += f" Mức cao nhất ghi nhận là {max_text} tại {max_instance}."
    return text


def _primary_log_switch_range_key(overall: str) -> str:
    return {
        "optimal": "optimal",
        "stable": "stable",
        "elevated": "elevated",
        "high": "high",
        "optimal_to_stable": "healthy_range",
        "frequent_elevated": "risk_range",
        "mixed": "healthy_range",
    }.get(overall, "healthy_range")


def _log_switch_range_parts(summaries: list[dict[str, object]], range_key: str, english: bool = False) -> str:
    if not summaries:
        return ""
    grouped: dict[str, list[str]] = defaultdict(list)
    for summary in summaries:
        value_range = _log_switch_summary_range(summary, range_key)
        if value_range is None:
            continue
        grouped[_format_switch_range_value(value_range, english)].append(str(summary["instance"]))
    if not grouped:
        return ""
    parts = []
    for range_text, instances in grouped.items():
        names = _join_display_list(instances, english)
        if len(instances) > 1:
            parts.append(f"{names} {'around' if english else 'dao động khoảng'} {range_text}")
        else:
            parts.append(f"{names}: {range_text}")
    return _join_display_list(parts, english)


def _log_switch_range_parts_by_range_first(summaries: list[dict[str, object]]) -> str:
    parts = []
    for summary in summaries:
        value_range = _log_switch_summary_range(summary, "risk_range")
        if value_range is None:
            continue
        parts.append(f"{_format_switch_range_value(value_range)} tại {summary['instance']}")
    return _join_display_list(parts)


def _log_switch_summary_range(summary: dict[str, object], range_key: str) -> tuple[float, float] | None:
    if range_key in {"healthy_range", "risk_range"}:
        return summary.get(range_key)  # type: ignore[return-value]
    ranges = summary["ranges"]
    assert isinstance(ranges, dict)
    return ranges.get(range_key)


def _join_display_list(items: list[str], english: bool = False) -> str:
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    separator = " and " if english else " và "
    return ", ".join(items[:-1]) + separator + items[-1]


def _log_switch_overall_label(overall: str, english: bool = False) -> str:
    labels = {
        "optimal": ("optimal for most of the time", "tối ưu trong phần lớn thời gian"),
        "stable": ("stable for most of the time", "ổn định trong phần lớn thời gian"),
        "elevated": ("quite high for most of the time", "khá cao trong phần lớn thời gian"),
        "high": ("high for most of the time", "cao trong phần lớn thời gian"),
        "optimal_to_stable": ("optimal to stable for most of the time", "tối ưu đến ổn định trong phần lớn thời gian"),
        "frequent_elevated": ("frequently quite high to high", "khá cao đến cao xuất hiện thường xuyên"),
        "mixed": ("mixed across time windows", "biến động giữa nhiều mức"),
    }
    english_label, vietnamese_label = labels[overall]
    return english_label if english else vietnamese_label


def _log_switch_summary_text(summary: dict[str, object], english: bool = False) -> str:
    instance = str(summary["instance"])
    overall = str(summary["overall"])
    maximum = float(summary["max"])
    ranges = summary["ranges"]
    healthy_range = summary["healthy_range"]
    risk_range = summary["risk_range"]
    ratios = summary["ratios"]
    assert isinstance(ranges, dict)
    assert isinstance(ratios, dict)

    optimal_range = _format_switch_range_value(ranges.get("optimal"), english)
    stable_range = _format_switch_range_value(ranges.get("stable"), english)
    elevated_range = _format_switch_range_value(ranges.get("elevated"), english)
    high_range = _format_switch_range_value(ranges.get("high"), english)
    healthy_range_text = _format_switch_range_value(healthy_range, english)
    risk_range_text = _format_switch_range_value(risk_range, english)
    max_text = _format_switch_range(maximum, maximum, english)

    if english:
        if overall == "optimal":
            text = f"Overall, log switch frequency of {instance} is optimal for most of the time, around {optimal_range}."
        elif overall == "stable":
            text = f"Overall, log switch frequency of {instance} is stable for most of the time, mainly around {stable_range}."
        elif overall == "elevated":
            text = f"Overall, log switch frequency of {instance} is quite high for most of the time, mainly around {elevated_range}."
        elif overall == "high":
            return f"Log switch frequency of {instance} remains high for most of the time, with recorded values around {high_range}. Review online redo log size, redo-generating workload, and signs of overly frequent log switches."
        elif overall == "optimal_to_stable":
            text = f"Overall, log switch frequency of {instance} remains optimal to stable for most of the time, mainly around {healthy_range_text}."
        elif overall == "frequent_elevated":
            text = f"Overall, log switch frequency of {instance} frequently reaches quite high to high levels, around {risk_range_text}."
        else:
            text = f"Overall, log switch frequency of {instance} varies across time windows, without a clearly dominant level."

        if maximum > 20 and overall != "high":
            text += f" However, some periods increased significantly, with the maximum recorded at {max_text}. Periods above 20 times/hour should be checked if they repeat or persist."
        elif ranges.get("elevated") and overall not in {"elevated", "frequent_elevated"}:
            text += f" Some higher-load periods reached {elevated_range} and should continue to be monitored."
        return text

    if overall == "optimal":
        text = f"Nhìn chung, tần suất log switch của {instance} trong phần lớn thời gian ở mức tối ưu, khoảng {optimal_range}."
    elif overall == "stable":
        text = f"Nhìn chung, tần suất log switch của {instance} trong phần lớn thời gian ở mức ổn định, chủ yếu dao động khoảng {stable_range}."
    elif overall == "elevated":
        text = f"Nhìn chung, tần suất log switch của {instance} trong phần lớn thời gian ở mức khá cao, chủ yếu dao động khoảng {elevated_range}. Tần suất này nên được tiếp tục theo dõi và đối chiếu với kích thước redo log cũng như workload phát sinh redo."
    elif overall == "high":
        return f"Tần suất log switch của {instance} duy trì ở mức cao trong phần lớn thời gian, với khoảng ghi nhận {high_range}. Nên kiểm tra kích thước online redo log, workload phát sinh redo và các dấu hiệu log switch quá thường xuyên."
    elif overall == "optimal_to_stable":
        text = f"Nhìn chung, tần suất log switch của {instance} trong phần lớn thời gian duy trì ở mức tối ưu đến ổn định, chủ yếu trong khoảng {healthy_range_text}."
    elif overall == "frequent_elevated":
        text = f"Nhìn chung, tần suất log switch của {instance} ghi nhận mức khá cao đến cao xuất hiện thường xuyên, trong khoảng {risk_range_text}. Nên kiểm tra kích thước redo log và workload phát sinh redo."
    else:
        dominant = "vùng tối ưu đến ổn định" if float(ratios.get("optimal", 0)) + float(ratios.get("stable", 0)) >= float(ratios.get("elevated", 0)) + float(ratios.get("high", 0)) else "vùng khá cao đến cao"
        text = f"Nhìn chung, tần suất log switch của {instance} có sự biến động giữa các khung giờ, chủ yếu tập trung trong {dominant}."

    if maximum > 20 and overall != "high":
        text += f" Tuy nhiên, một số thời điểm ghi nhận tần suất tăng cao, với mức tối đa {max_text}. Các thời điểm vượt 20 lần/giờ nên được kiểm tra thêm nếu xuất hiện lặp lại hoặc kéo dài."
    elif ranges.get("elevated") and overall not in {"elevated", "frequent_elevated"}:
        text += f" Một số khung giờ tải cao ghi nhận tần suất {elevated_range}, thuộc mức khá cao và nên tiếp tục theo dõi."
    return text


def _log_switch_instance_label(row: dict[str, str]) -> str:
    for key in ("INSTANCE_NUMBER", "INSTANCE_NAME", "INSTANCE", "INST_ID", "INSTANCE_ID"):
        value = row.get(key, "").strip()
        if value:
            return f"instance {value}"
    source_file = row.get("SOURCE_FILE", "")
    match = re.search(r"instance[_-](\d+)", source_file, flags=re.IGNORECASE)
    if match:
        return f"instance {match.group(1)}"
    return ""
