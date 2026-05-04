from __future__ import annotations

from dataclasses import dataclass
import shutil
import threading


@dataclass(frozen=True)
class PandocStatus:
    available: bool
    executable: str | None
    version: str | None
    message: str


class PandocInstallError(RuntimeError):
    """Raised when Pandoc is not available and cannot be installed."""


_install_lock = threading.Lock()


def check_pandoc() -> PandocStatus:
    executable = shutil.which("pandoc")
    if executable:
        return PandocStatus(True, executable, _read_pandoc_version(), "Pandoc is available.")

    try:
        import pypandoc

        path = pypandoc.get_pandoc_path()
        if path:
            return PandocStatus(True, path, _read_pandoc_version(), "Pandoc bundled by pypandoc is available.")
    except (ImportError, OSError, RuntimeError):
        pass

    return PandocStatus(False, None, None, "Pandoc was not found.")


def ensure_pandoc() -> str:
    status = check_pandoc()
    if status.available and status.executable:
        return status.executable

    with _install_lock:
        status = check_pandoc()
        if status.available and status.executable:
            return status.executable

        try:
            import pypandoc
        except ImportError as exc:
            raise PandocInstallError(
                "pypandoc is not installed. Run: python -m pip install pypandoc"
            ) from exc

        try:
            _download_pandoc()
            path = pypandoc.get_pandoc_path()
        except Exception as exc:
            raise PandocInstallError(
                "Pandoc could not be downloaded automatically. "
                "Please install Pandoc manually from https://pandoc.org/installing.html"
            ) from exc

        if not path:
            raise PandocInstallError("Pandoc download finished, but no executable path was found.")
        return path


def _read_pandoc_version() -> str | None:
    try:
        import pypandoc

        return pypandoc.get_pandoc_version()
    except Exception:
        return None


def _download_pandoc() -> None:
    try:
        from pypandoc.pandoc_download import download_pandoc
    except ImportError:
        import pypandoc

        download_pandoc = pypandoc.download_pandoc
    download_pandoc()
