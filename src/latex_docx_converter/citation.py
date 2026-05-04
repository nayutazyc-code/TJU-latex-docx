from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re


@dataclass(frozen=True)
class CitationAudit:
    cited_keys: tuple[str, ...]
    bib_keys: tuple[str, ...]
    missing_keys: tuple[str, ...]

    @property
    def warnings(self) -> tuple[str, ...]:
        if not self.missing_keys:
            return ()
        missing = ", ".join(self.missing_keys)
        return (f"Missing bibliography entries for citation keys: {missing}",)


def audit_citations(tex_text: str, bib_path: Path | None) -> CitationAudit:
    cited_keys = tuple(sorted(extract_citation_keys(tex_text)))
    bib_keys = tuple(sorted(extract_bib_keys(bib_path))) if bib_path else ()
    missing_keys = tuple(key for key in cited_keys if key not in bib_keys)
    return CitationAudit(cited_keys=cited_keys, bib_keys=bib_keys, missing_keys=missing_keys)


def collect_project_tex_text(project_dir: Path) -> str:
    parts: list[str] = []
    for tex_file in sorted(project_dir.rglob("*.tex")):
        if any(part.startswith(".") for part in tex_file.relative_to(project_dir).parts):
            continue
        try:
            parts.append(tex_file.read_text(encoding="utf-8", errors="ignore"))
        except OSError:
            continue
    return "\n".join(parts)


def extract_citation_keys(tex_text: str) -> set[str]:
    stripped = strip_latex_comments(tex_text)
    keys: set[str] = set()
    pattern = re.compile(
        r"\\(?:cite|parencite|textcite|autocite|supercite|citep|citet)\*?"
        r"\s*(?:\[[^\]]*\]\s*)*\{([^{}]+)\}"
    )
    for match in pattern.finditer(stripped):
        for key in match.group(1).split(","):
            clean = key.strip()
            if clean:
                keys.add(clean)
    return keys


def extract_bib_keys(bib_path: Path | None) -> set[str]:
    if bib_path is None or not bib_path.is_file():
        return set()
    text = bib_path.read_text(encoding="utf-8", errors="ignore")
    return {match.group(1).strip() for match in re.finditer(r"@\w+\s*\{\s*([^,\s]+)", text)}


def strip_latex_comments(text: str) -> str:
    lines: list[str] = []
    for line in text.splitlines():
        escaped = False
        kept: list[str] = []
        for char in line:
            if char == "\\" and not escaped:
                escaped = True
                kept.append(char)
                continue
            if char == "%" and not escaped:
                break
            kept.append(char)
            escaped = False
        lines.append("".join(kept))
    return "\n".join(lines)
