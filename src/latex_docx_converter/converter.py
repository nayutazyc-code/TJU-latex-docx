from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import os
from pathlib import Path
import subprocess
from typing import Iterable

from .citation import audit_citations, collect_project_tex_text
from .defaults import find_default_bibliography, find_default_csl, find_default_reference_docx
from .pandoc_manager import ensure_pandoc
from .tjuthesis import prepare_tjuthesis_input
from .word_postprocess import WordPostprocessProfile, postprocess_docx


@dataclass(frozen=True)
class ConversionConfig:
    project_dir: Path
    main_tex: Path
    output_docx: Path
    reference_docx: Path | None = None
    bibliography: Path | None = None
    csl: Path | None = None


@dataclass(frozen=True)
class ConversionResult:
    output_docx: Path
    log_path: Path
    command: tuple[str, ...]
    stdout: str
    stderr: str
    warnings: tuple[str, ...] = ()


class ConversionError(RuntimeError):
    """Raised when Pandoc cannot complete a conversion."""

    def __init__(self, message: str, log_path: Path | None = None) -> None:
        super().__init__(message)
        self.log_path = log_path


def convert_project(config: ConversionConfig) -> ConversionResult:
    normalized = normalize_config(config)
    pandoc = ensure_pandoc()

    normalized.output_docx.parent.mkdir(parents=True, exist_ok=True)
    log_path = make_log_path(normalized.output_docx)
    citation_audit = audit_citations(collect_project_tex_text(normalized.project_dir), normalized.bibliography)

    started_at = datetime.now().isoformat(timespec="seconds")
    command: list[str] = []
    postprocess_warnings: tuple[str, ...] = ()
    postprocess_notes: tuple[str, ...] = ()
    with prepare_tjuthesis_input(normalized.project_dir, normalized.main_tex) as prepared:
        command = build_pandoc_command(
            normalized,
            pandoc,
            input_tex=prepared.main_tex,
            add_toc=prepared.add_toc,
        )
        try:
            process = subprocess.run(
                command,
                cwd=normalized.project_dir,
                text=True,
                capture_output=True,
                check=False,
            )
        except OSError as exc:
            write_log(
                log_path,
                command,
                started_at,
                stdout="",
                stderr=str(exc),
                returncode=None,
                notes=prepared.notes,
            )
            raise ConversionError(f"Pandoc failed to start: {exc}", log_path) from exc

        if process.returncode == 0 and normalized.output_docx.exists() and prepared.postprocess_docx:
            postprocess = postprocess_docx(
                normalized.output_docx,
                WordPostprocessProfile(reference_docx=normalized.reference_docx),
                citation_audit,
            )
            postprocess_warnings = postprocess.warnings
            postprocess_notes = postprocess.notes

        write_log(
            log_path,
            command,
            started_at,
            stdout=process.stdout,
            stderr=process.stderr,
            returncode=process.returncode,
            notes=(
                *prepared.notes,
                *postprocess_notes,
            ),
            warnings=(
                *citation_audit.warnings,
                *postprocess_warnings,
            ),
        )

    if process.returncode != 0:
        message = process.stderr.strip() or process.stdout.strip() or "Pandoc failed without output."
        raise ConversionError(message, log_path)

    if not normalized.output_docx.exists():
        raise ConversionError("Pandoc finished successfully, but the DOCX file was not created.", log_path)

    return ConversionResult(
        output_docx=normalized.output_docx,
        log_path=log_path,
        command=tuple(command),
        stdout=process.stdout,
        stderr=process.stderr,
        warnings=(
            *citation_audit.warnings,
            *postprocess_warnings,
        ),
    )


def normalize_config(config: ConversionConfig) -> ConversionConfig:
    project_dir = config.project_dir.expanduser().resolve()
    main_tex = resolve_against_project(config.main_tex, project_dir)
    output_docx = resolve_export_docx(config.output_docx, project_dir)
    reference_docx = resolve_optional(config.reference_docx, project_dir) or find_default_reference_docx(project_dir)
    bibliography = resolve_optional(config.bibliography, project_dir) or find_default_bibliography(project_dir)
    csl = resolve_optional(config.csl, project_dir) or find_default_csl(project_dir)

    if not project_dir.is_dir():
        raise ValueError(f"Project folder does not exist: {project_dir}")
    if not main_tex.is_file():
        raise ValueError(f"Main .tex file does not exist: {main_tex}")
    if main_tex.suffix.lower() != ".tex":
        raise ValueError("Main file must be a .tex file.")
    for label, path in (
        ("reference DOCX", reference_docx),
        ("bibliography", bibliography),
        ("CSL", csl),
    ):
        if path is not None and not path.is_file():
            raise ValueError(f"Selected {label} file does not exist: {path}")

    return ConversionConfig(
        project_dir=project_dir,
        main_tex=main_tex,
        output_docx=output_docx,
        reference_docx=reference_docx,
        bibliography=bibliography,
        csl=csl,
    )


def build_pandoc_command(
    config: ConversionConfig,
    pandoc_executable: str,
    input_tex: Path | None = None,
    add_toc: bool = False,
) -> list[str]:
    main_arg = relative_or_absolute(input_tex or config.main_tex, config.project_dir)
    output_arg = config.output_docx
    command = [
        pandoc_executable,
        str(main_arg),
        "-f",
        "latex",
        "-t",
        "docx",
        "-s",
        "--resource-path",
        resource_path(config.project_dir),
        "--citeproc",
        "-o",
        str(output_arg),
    ]

    if add_toc:
        command.append("--toc")
    append_optional(command, "--reference-doc", config.reference_docx)
    append_optional(command, "--bibliography", config.bibliography)
    append_optional(command, "--csl", config.csl)
    return command


def append_optional(command: list[str], option: str, path: Path | None) -> None:
    if path is not None:
        command.append(f"{option}={path}")


def resolve_against_project(path: Path, project_dir: Path) -> Path:
    expanded = path.expanduser()
    if expanded.is_absolute():
        return expanded.resolve()
    return (project_dir / expanded).resolve()


def resolve_optional(path: Path | None, project_dir: Path) -> Path | None:
    if path is None:
        return None
    return resolve_against_project(path, project_dir)


def relative_or_absolute(path: Path, base: Path) -> Path:
    try:
        return path.relative_to(base)
    except ValueError:
        return path


def make_log_path(output_docx: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return output_docx.parent / "logs" / f"{output_docx.stem}-pandoc-{timestamp}.log"


def write_log(
    log_path: Path,
    command: Iterable[str],
    started_at: str,
    stdout: str,
    stderr: str,
    returncode: int | None,
    notes: Iterable[str] | None = None,
    warnings: Iterable[str] | None = None,
) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    finished_at = datetime.now().isoformat(timespec="seconds")
    log_path.write_text(
        "\n".join(
            [
                "LaTeX DOCX Converter Log",
                f"Started: {started_at}",
                f"Finished: {finished_at}",
                f"Return code: {returncode if returncode is not None else 'not started'}",
                "",
                "Notes:",
                *list(notes or ()),
                "",
                "Warnings:",
                *list(warnings or ()),
                "",
                "Command:",
                " ".join(command),
                "",
                "STDOUT:",
                stdout or "",
                "",
                "STDERR:",
                stderr or "",
                "",
            ]
        ),
        encoding="utf-8",
    )


def resource_path(project_dir: Path) -> str:
    paths = [project_dir]
    figures = project_dir / "figures"
    if figures.is_dir():
        paths.append(figures)
    return os.pathsep.join(str(path) for path in paths)


def export_dir(project_dir: Path) -> Path:
    return project_dir / "docx导出"


def resolve_export_docx(output_docx: Path, project_dir: Path) -> Path:
    expanded = output_docx.expanduser()
    filename = expanded.name if expanded.name else f"{project_dir.name}.docx"
    if not filename.lower().endswith(".docx"):
        filename = f"{filename}.docx"
    return (export_dir(project_dir) / filename).resolve()
