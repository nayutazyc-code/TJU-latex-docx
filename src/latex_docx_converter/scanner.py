from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, order=True)
class TexCandidate:
    sort_key: tuple[int, int, str]
    path: Path
    score: int
    reason: str


def find_main_tex_candidates(project_dir: Path) -> list[Path]:
    candidates = score_tex_files(project_dir)
    return [candidate.path for candidate in candidates]


def score_tex_files(project_dir: Path) -> list[TexCandidate]:
    root = project_dir.expanduser().resolve()
    if not root.is_dir():
        return []

    candidates: list[TexCandidate] = []
    for tex_file in root.rglob("*.tex"):
        if _is_hidden(tex_file, root):
            continue
        text = read_preview(tex_file)
        score, reason = score_tex_content(tex_file, text)
        candidates.append(
            TexCandidate(
                sort_key=(-score, len(tex_file.parts), str(tex_file.relative_to(root)).lower()),
                path=tex_file,
                score=score,
                reason=reason,
            )
        )
    return sorted(candidates)


def score_tex_content(path: Path, text: str) -> tuple[int, str]:
    score = 0
    reasons: list[str] = []

    if "\\documentclass" in text:
        score += 60
        reasons.append("documentclass")
    if "\\begin{document}" in text:
        score += 60
        reasons.append("begin document")
    if "\\bibliography" in text or "\\addbibresource" in text:
        score += 10
        reasons.append("bibliography")
    if path.name.lower() in {"main.tex", "thesis.tex", "paper.tex", "article.tex"}:
        score += 15
        reasons.append("common main filename")

    if not reasons:
        reasons.append("tex file")
    return score, ", ".join(reasons)


def read_preview(path: Path, limit: int = 512_000) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore")[:limit]
    except OSError:
        return ""


def _is_hidden(path: Path, root: Path) -> bool:
    try:
        relative = path.relative_to(root)
    except ValueError:
        relative = path
    return any(part.startswith(".") for part in relative.parts)
