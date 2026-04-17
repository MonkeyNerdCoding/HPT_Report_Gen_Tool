from pathlib import Path


def discover_html_files(input_path: str | Path) -> list[Path]:
    path = Path(input_path)
    if path.is_file():
        if path.suffix.lower() not in {".html", ".htm"}:
            raise ValueError(f"Input file is not HTML: {path}")
        return [path]

    if path.is_dir():
        return sorted(
            child
            for child in path.rglob("*")
            if child.is_file() and child.suffix.lower() in {".html", ".htm"}
        )

    raise FileNotFoundError(f"Input path does not exist: {path}")
