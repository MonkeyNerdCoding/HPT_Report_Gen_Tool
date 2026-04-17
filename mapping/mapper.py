from __future__ import annotations

from models import ExtractedContent, GenerationReport, MappingRule

from .content_registry import ContentRegistry


def resolve_mappings(
    rules: list[MappingRule],
    registry: ContentRegistry,
    report: GenerationReport,
) -> dict[str, tuple[MappingRule, ExtractedContent]]:
    resolved: dict[str, tuple[MappingRule, ExtractedContent]] = {}

    for rule in rules:
        # Dùng ContentRegistry để tìm nội dung phù hợp theo rule.
        # Truyền các tham số như type, source_key, source_file, section, table_index, variant
        matches = registry.find(
            rule.content_type,
            source_key=rule.source_key,
            source_file=rule.source_file,
            section=rule.section,
            index=rule.table_index,
            variant=rule.chart_variant,
        )

        if not matches:
            report.missing_content.append(
                f"{rule.placeholder}: no {rule.content_type} content for '{rule.source_key or rule.source_file or rule.section}'"
            )
            continue

        # Nếu tìm được nhiều kết quả: cố gắng tìm match chính xác bằng logical_key hoặc keys.
        # Nhưng nếu vẫn nhiều kết quả và rule không chỉ rõ table_index hay source_file,
        # báo cáo là ambiguous để người dùng cấu hình thêm (nguồn file hoặc index).
        if len(matches) > 1 and rule.table_index is None and not rule.source_file:
            exact = [item for item in matches if item.logical_key == rule.source_key or rule.source_key in item.keys]
            if len(exact) == 1:
                # Nếu chỉ có 1 kết quả chính xác, chọn nó
                matches = exact
            else:
                # Không thể phân biệt: để user điều chỉnh mapping (thêm source_file hoặc table_index)
                report.ambiguous.append(
                    f"{rule.placeholder}: {len(matches)} matches for '{rule.source_key}', add source_file or table_index"
                )
                continue

        # Lấy kết quả đầu tiên trong matches (sau khi đã xử lý ambiguous/exact)
        resolved[rule.placeholder] = (rule, matches[0])

    return resolved

