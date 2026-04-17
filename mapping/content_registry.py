from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from models import ExtractedContent
from utils.normalize import normalize_key, strip_chart_suffix


class ContentRegistry:
    def __init__(self, contents: list[ExtractedContent]) -> None:
        self.contents = contents
        self._by_key: dict[tuple[str, str], list[ExtractedContent]] = defaultdict(list)
        for content in contents:
            # Lặp qua các key mà content khai báo và lưu vào index theo (type, key)
            # normalize_key giúp chuẩn hoá key (viết thường, loại bỏ dấu, khoảng trắng,...)
            for key in content.keys:
                self._by_key[(content.content_type, normalize_key(key))].append(content)
                # Nếu là image (ảnh chart), cũng đăng ký dưới kiểu 'chart' để tìm thay thế linh hoạt
                if content.content_type == "image":
                    self._by_key[("chart", normalize_key(key))].append(content)

    def find(
        self,
        content_type: str,
        source_key: str = "",
        source_file: str = "",
        section: str = "",
        index: int | None = None,
        variant: str = "",
    ) -> list[ExtractedContent]:
        candidates = self.contents

        if source_file:
            # Nếu rule chỉ định source_file thì lọc ngay theo tên file (so sánh không phân biệt hoa thường)
            expected = Path(source_file).name.lower()
            candidates = [item for item in candidates if item.source_path.name.lower() == expected]

        if section:
            # Lọc theo section nếu rule có chỉ định (ví dụ: phần A, phần B của báo cáo)
            expected_section = section.lower()
            candidates = [item for item in candidates if item.section.lower() == expected_section]

        if source_key:
            # Tìm theo source_key: ưu tiên lookup nhanh từ _by_key
            key = normalize_key(source_key)
            key_candidates = list(self._by_key.get((content_type, key), []))
            # Nếu đang tìm chart nhưng không có entry 'chart', thử fallback sang 'image'
            if not key_candidates and content_type == "chart":
                key_candidates = list(self._by_key.get(("image", key), []))
            # Nếu vẫn rỗng thì dùng cơ chế tìm tương đối trên candidates hiện tại
            # Bao gồm so sánh normalize_key trên mỗi key của item, hoặc loại bỏ hậu tố chart khi cần
            if not key_candidates:
                key_candidates = [
                    item
                    for item in candidates
                    if key in {normalize_key(k) for k in item.keys}
                    or strip_chart_suffix(normalize_key(item.logical_key)) == key
                ]
            # Giữ lại những candidates khớp với key_candidates (áp dụng cả filter trước đó như source_file/section)
            candidates = [item for item in candidates if item in key_candidates]

        if content_type:
            # Lọc theo loại content: nếu rule yêu cầu 'chart' thì chấp nhận cả chart và image
            if content_type == "chart":
                candidates = [item for item in candidates if item.content_type in {"chart", "image"}]
            else:
                candidates = [item for item in candidates if item.content_type == content_type]

        if index is not None:
            # Nếu rule chỉ rõ index (ví dụ chọn dòng/biểu đồ thứ N), lọc theo thuộc tính index
            candidates = [item for item in candidates if item.index == index]

        if variant:
            # Nếu có variant cho chart (ví dụ: 'small', 'large'), chấp nhận cả '' (mặc định) hoặc đúng variant
            candidates = [item for item in candidates if getattr(item, "variant", "") in {"", variant}]

        return candidates

