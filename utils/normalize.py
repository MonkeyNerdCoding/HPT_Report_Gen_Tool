import re


def normalize_key(value: str | None) -> str:
    """Convert report titles, placeholders, and filenames into stable lookup keys."""
    if not value:
        return ""

    value = value.strip().lower()
    value = re.sub(r"<|>", "", value)
    value = re.sub(r"\([^)]*\)", " ", value)
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("_")


def strip_chart_suffix(key: str) -> str:
    for suffix in ("_line_chart", "_pie_chart", "_bar_chart", "_chart"):
        if key.endswith(suffix):
            return key[: -len(suffix)]
    return key

