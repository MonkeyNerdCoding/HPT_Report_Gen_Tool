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


def content_key_aliases(value: str | None) -> set[str]:
    """Return lookup aliases for report keys that vary by date or RAC shape."""
    key = normalize_key(value)
    if not key:
        return set()

    aliases = {key, strip_chart_suffix(key)}
    changed = True
    while changed:
        changed = False
        for item in list(aliases):
            next_aliases = set()
            if "_between_" in item:
                next_aliases.add(item.split("_between_", 1)[0])
            normalized_days = re.sub(r"_for_\d+_days_of_history", "_for_days_of_history", item)
            if normalized_days != item:
                next_aliases.add(normalized_days)
            if "_for_days_of_history" in item:
                next_aliases.add(item.split("_for_days_of_history", 1)[0])
            normalized_instance = re.sub(r"_for_instance_\d+", "_for_instance", item)
            if normalized_instance != item:
                next_aliases.add(normalized_instance)
            if "for_cluster_for_9_days_of_history" in item:
                next_aliases.add(item.replace("for_cluster_for_9_days_of_history", "for_instance_1_for_9_days_of_history"))
            if "for_instance_1_for_9_days_of_history" in item:
                next_aliases.add(item.replace("for_instance_1_for_9_days_of_history", "for_cluster_for_9_days_of_history"))
            if not next_aliases.issubset(aliases):
                aliases.update(next_aliases)
                changed = True

    return {alias for alias in aliases if alias}

