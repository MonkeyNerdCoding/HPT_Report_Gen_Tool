from __future__ import annotations

import os
import re
import sys


def extract_sheet_name(filename: str) -> str:
    """
    Robust extract sheet name from filename.
    - Remove extension first.
    - Split by '-' and trim.
    - If last part is a long numeric timestamp (>=6 digits) -> take parts[-2].
    - Else if last part contains letters -> take last part.
    - Else (last part is a short number like '35' or other) -> scan from right to left
      to find the nearest part that contains letters (this is likely the info name).
    Returns sheet name trimmed to 31 chars.
    """
    name = os.path.splitext(filename)[0]
    parts = [p.strip() for p in name.split("-") if p.strip() != ""]

    _safe_print(f"\n📄 Đang xử lý file: {filename}")
    _safe_print(f"🔹 parts = {parts} (len={len(parts)})")

    if not parts:
        return name[:31]

    last = parts[-1]

    # Nếu phần cuối giống timestamp dài toàn số (ví dụ 202509111426542654)
    if re.fullmatch(r'\d{6,}', last):
        candidate = parts[-2] if len(parts) >= 2 else last
        reason = "last is long numeric timestamp -> use parts[-2]"
    else:
        # nếu phần cuối có chữ -> có thể chính là info
        if re.search(r'[A-Za-z]', last):
            candidate = last
            reason = "last contains letters -> use last"
        else:
            # last không phải timestamp nhưng là số (vd '35'/'75') -> tìm phần bên trái gần nhất có chữ
            candidate = None
            for p in reversed(parts[:-1]):
                if re.search(r'[A-Za-z]', p):
                    candidate = p
                    break
            if candidate:
                reason = "last is short number -> found nearest left part with letters"
            else:
                # fallback: dùng last nếu không tìm được phần có chữ
                candidate = last
                reason = "fallback -> use last"

    # nếu candidate vẫn là số thuần (vd '35'), cố gắng tìm phần có chữ từ phải sang trái toàn bộ parts
    if re.fullmatch(r'^\d+$', candidate):
        for p in reversed(parts):
            if re.search(r'[A-Za-z]', p):
                candidate = p
                reason = "candidate was numeric -> replaced by nearest part with letters"
                break

    sheet_name = candidate.strip()[:31]
    _safe_print(f"   ➡ detect -> '{sheet_name}'  ({reason})")
    return sheet_name


def _safe_print(message: str) -> None:
    try:
        print(message)
    except UnicodeEncodeError:
        encoding = sys.stdout.encoding or "utf-8"
        safe_message = message.encode(encoding, errors="backslashreplace").decode(encoding, errors="replace")
        print(safe_message)
