from __future__ import annotations

from pathlib import Path


def find_default_reference_docx(project_dir: Path) -> Path | None:
    roots = candidate_roots(project_dir)
    patterns = [
        "【2025修订版】1-1天津大学本科生毕业设计模板.docx",
        "*天津大学本科生毕业设计模板*.docx",
        "*本科生毕业设计模板*.docx",
    ]
    return first_existing_match(roots, patterns)


def find_default_csl(project_dir: Path) -> Path | None:
    roots = candidate_roots(project_dir)
    patterns = [
        "china-national-standard-gb-t-7714-2015-numeric.csl",
        "*gb*t*7714*2015*numeric*.csl",
        "*7714*.csl",
    ]
    return first_existing_match(roots, patterns)


def find_default_bibliography(project_dir: Path) -> Path | None:
    direct = project_dir / "reference.bib"
    if direct.is_file():
        return direct.resolve()
    matches = sorted(project_dir.glob("*.bib"))
    return matches[0].resolve() if matches else None


def candidate_roots(project_dir: Path) -> list[Path]:
    roots = [
        project_dir,
        project_dir.parent,
        Path.home() / "Desktop" / "编写latex",
    ]
    unique: list[Path] = []
    for root in roots:
        resolved = root.expanduser().resolve()
        if resolved not in unique and resolved.exists():
            unique.append(resolved)
    return unique


def first_existing_match(roots: list[Path], patterns: list[str]) -> Path | None:
    for root in roots:
        for pattern in patterns:
            direct = root / pattern
            if "*" not in pattern and direct.is_file():
                return direct.resolve()
            matches = sorted(root.glob(pattern))
            if matches:
                return matches[0].resolve()
    return None
