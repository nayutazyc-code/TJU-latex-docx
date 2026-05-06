from __future__ import annotations

from datetime import datetime
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .ai_review_bundle import build_ai_review_bundle
from .converter import ConversionConfig, ConversionError, convert_project, resolve_export_docx
from .defaults import find_default_bibliography, find_default_csl, find_default_reference_docx
from .docx_review import review_docx
from .pandoc_manager import PandocInstallError, check_pandoc, ensure_pandoc
from .scanner import find_main_tex_candidates


class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("LaTeX 导出 DOCX")
        self.geometry("900x680")
        self.minsize(760, 560)

        self.project_dir = tk.StringVar()
        self.main_tex = tk.StringVar()
        self.output_docx = tk.StringVar()
        self.reference_docx = tk.StringVar()
        self.bibliography = tk.StringVar()
        self.csl = tk.StringVar()
        self.pandoc_status = tk.StringVar(value="正在检查 Pandoc...")

        self._events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._is_busy = False
        self._candidate_paths: list[Path] = []

        self._build_ui()
        self.after(100, self._poll_events)
        self._install_pandoc_in_background()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=16)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(8, weight=1)

        title = ttk.Label(root, text="LaTeX 论文项目导出 DOCX", font=("", 18, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        self._add_path_row(root, 1, "项目文件夹", self.project_dir, self._choose_project_dir)

        ttk.Label(root, text="主 .tex 文件").grid(row=2, column=0, sticky="w", pady=6)
        self.main_combo = ttk.Combobox(root, textvariable=self.main_tex, values=[], state="normal")
        self.main_combo.grid(row=2, column=1, sticky="ew", padx=(10, 8), pady=6)
        ttk.Button(root, text="选择", command=self._choose_main_tex).grid(row=2, column=2, sticky="ew", pady=6)

        self._add_path_row(root, 3, "输出 DOCX", self.output_docx, self._choose_output_docx)
        self._add_path_row(root, 4, "参考模板 DOCX", self.reference_docx, self._choose_reference_docx)
        self._add_path_row(root, 5, "参考文献 .bib", self.bibliography, self._choose_bibliography)
        self._add_path_row(root, 6, "引用样式 .csl", self.csl, self._choose_csl)

        status_frame = ttk.Frame(root)
        status_frame.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.pandoc_status).grid(row=0, column=0, sticky="w")
        ttk.Button(status_frame, text="重新检测/下载 Pandoc", command=self._install_pandoc_in_background).grid(
            row=0, column=1, sticky="e"
        )

        log_frame = ttk.LabelFrame(root, text="转换日志", padding=8)
        log_frame.grid(row=8, column=0, columnspan=3, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=14, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        actions = ttk.Frame(root)
        actions.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        actions.columnconfigure(0, weight=1)
        ttk.Button(actions, text="清空日志", command=self._clear_log).grid(row=0, column=1, padx=(0, 8))
        self.review_button = ttk.Button(actions, text="检查现有 DOCX", command=self._review_existing_docx_in_background)
        self.review_button.grid(row=0, column=2, padx=(0, 8))
        self.ai_bundle_button = ttk.Button(actions, text="生成 AI 审稿包", command=self._ai_bundle_in_background)
        self.ai_bundle_button.grid(row=0, column=3, padx=(0, 8))
        self.convert_button = ttk.Button(actions, text="开始导出 DOCX", command=self._convert_in_background)
        self.convert_button.grid(row=0, column=4)

    def _add_path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(10, 8), pady=6)
        ttk.Button(parent, text="选择", command=command).grid(row=row, column=2, sticky="ew", pady=6)

    def _choose_project_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择 LaTeX 项目文件夹")
        if not selected:
            return
        self.project_dir.set(selected)
        self._log(f"项目文件夹：{selected}")
        project_dir = Path(selected)
        self._scan_candidates(project_dir)
        self._fill_default_optional_files(project_dir)

    def _choose_main_tex(self) -> None:
        initial = self.project_dir.get() or str(Path.home())
        selected = filedialog.askopenfilename(
            title="选择主 .tex 文件",
            initialdir=initial,
            filetypes=[("LaTeX files", "*.tex"), ("All files", "*.*")],
        )
        if selected:
            self.main_tex.set(selected)
            self._suggest_output(Path(selected))

    def _choose_output_docx(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="选择输出 DOCX",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if selected:
            self.output_docx.set(selected)

    def _choose_reference_docx(self) -> None:
        self._choose_optional_file(self.reference_docx, "选择参考模板 DOCX", [("Word documents", "*.docx")])

    def _choose_bibliography(self) -> None:
        self._choose_optional_file(self.bibliography, "选择参考文献 .bib", [("BibTeX files", "*.bib")])

    def _choose_csl(self) -> None:
        self._choose_optional_file(self.csl, "选择引用样式 .csl", [("CSL files", "*.csl")])

    def _choose_optional_file(
        self,
        variable: tk.StringVar,
        title: str,
        filetypes: list[tuple[str, str]],
    ) -> None:
        initial = self.project_dir.get() or str(Path.home())
        selected = filedialog.askopenfilename(
            title=title,
            initialdir=initial,
            filetypes=filetypes + [("All files", "*.*")],
        )
        if selected:
            variable.set(selected)

    def _scan_candidates(self, project_dir: Path) -> None:
        self._candidate_paths = find_main_tex_candidates(project_dir)
        values = [str(path) for path in self._candidate_paths]
        self.main_combo.configure(values=values)
        if values:
            self.main_tex.set(values[0])
            self._suggest_output(self._candidate_paths[0])
            self._log(f"已找到 {len(values)} 个 .tex 文件，已选择最可能的主文件。")
        else:
            self.main_tex.set("")
            self._log("没有找到 .tex 文件，请手动选择。")

    def _fill_default_optional_files(self, project_dir: Path) -> None:
        defaults = [
            (self.reference_docx, find_default_reference_docx(project_dir), "参考模板"),
            (self.bibliography, find_default_bibliography(project_dir), "参考文献"),
            (self.csl, find_default_csl(project_dir), "引用样式"),
        ]
        for variable, path, label in defaults:
            if variable.get().strip() or path is None:
                continue
            variable.set(str(path))
            self._log(f"已自动选择{label}：{path}")

    def _suggest_output(self, main_tex: Path) -> None:
        if self.output_docx.get():
            return
        project = Path(self.project_dir.get()) if self.project_dir.get() else main_tex.parent
        self.output_docx.set(str(resolve_export_docx(main_tex.with_suffix(".docx"), project, datetime.now())))

    def _install_pandoc_in_background(self) -> None:
        if self._is_busy:
            return
        self._set_busy(True)
        self.pandoc_status.set("正在检查 Pandoc，如未安装将自动下载...")
        thread = threading.Thread(target=self._install_pandoc_worker, daemon=True)
        thread.start()

    def _install_pandoc_worker(self) -> None:
        try:
            before = check_pandoc()
            if before.available:
                self._events.put(("pandoc", before.message))
                return
            path = ensure_pandoc()
            self._events.put(("pandoc", f"Pandoc 已就绪：{path}"))
        except PandocInstallError as exc:
            self._events.put(("pandoc_error", str(exc)))
        except Exception as exc:
            self._events.put(("pandoc_error", f"Pandoc 检查失败：{exc}"))

    def _convert_in_background(self) -> None:
        if self._is_busy:
            return
        try:
            config = self._read_config()
        except ValueError as exc:
            messagebox.showerror("配置不完整", str(exc))
            return

        self._set_busy(True)
        self._log("开始调用 Pandoc 转换...")
        thread = threading.Thread(target=self._convert_worker, args=(config,), daemon=True)
        thread.start()

    def _convert_worker(self, config: ConversionConfig) -> None:
        try:
            result = convert_project(config)
            self._events.put(("converted", result))
        except ConversionError as exc:
            message = str(exc)
            if exc.log_path is not None:
                message = f"{message}\n日志文件：{exc.log_path}"
            self._events.put(("convert_error", message))
        except (PandocInstallError, ValueError) as exc:
            self._events.put(("convert_error", str(exc)))
        except Exception as exc:
            self._events.put(("convert_error", f"未知错误：{exc}"))

    def _review_existing_docx_in_background(self) -> None:
        if self._is_busy:
            return
        selected = filedialog.askopenfilename(
            title="选择要检查的 DOCX",
            initialdir=self.project_dir.get() or str(Path.home()),
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if not selected:
            return
        self._set_busy(True)
        self._log(f"开始检查 DOCX：{selected}")
        thread = threading.Thread(target=self._review_existing_docx_worker, args=(Path(selected),), daemon=True)
        thread.start()

    def _review_existing_docx_worker(self, docx_path: Path) -> None:
        try:
            report = review_docx(docx_path, docx_path.parent / "review")
            self._events.put(("reviewed", report))
        except Exception as exc:
            self._events.put(("review_error", f"DOCX 检查失败：{exc}"))

    def _ai_bundle_in_background(self) -> None:
        if self._is_busy:
            return
        selected = filedialog.askopenfilename(
            title="选择要生成审稿包的 DOCX",
            initialdir=self.project_dir.get() or str(Path.home()),
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if not selected:
            return
        project_dir = Path(self.project_dir.get().strip()) if self.project_dir.get().strip() else None
        main_tex = Path(self.main_tex.get().strip()) if self.main_tex.get().strip() else None
        bibliography = self._optional_path(self.bibliography)
        csl = self._optional_path(self.csl)
        reference_docx = self._optional_path(self.reference_docx)
        self._set_busy(True)
        self._log(f"开始生成 AI 审稿包：{selected}")
        thread = threading.Thread(
            target=self._ai_bundle_worker,
            args=(Path(selected), project_dir, main_tex, bibliography, csl, reference_docx),
            daemon=True,
        )
        thread.start()

    def _ai_bundle_worker(
        self,
        docx_path: Path,
        project_dir: Path | None,
        main_tex: Path | None,
        bibliography: Path | None,
        csl: Path | None,
        reference_docx: Path | None,
    ) -> None:
        try:
            bundle = build_ai_review_bundle(
                output_docx=docx_path,
                project_dir=project_dir,
                main_tex=main_tex,
                bibliography=bibliography,
                csl=csl,
                reference_docx=reference_docx,
            )
            self._events.put(("ai_bundle", bundle))
        except Exception as exc:
            self._events.put(("ai_bundle_error", f"AI 审稿包生成失败：{exc}"))

    def _read_config(self) -> ConversionConfig:
        if not self.project_dir.get().strip():
            raise ValueError("请选择 LaTeX 项目文件夹。")
        if not self.main_tex.get().strip():
            raise ValueError("请选择主 .tex 文件。")
        if not self.output_docx.get().strip():
            raise ValueError("请选择输出 DOCX 文件。")

        return ConversionConfig(
            project_dir=Path(self.project_dir.get().strip()),
            main_tex=Path(self.main_tex.get().strip()),
            output_docx=Path(self.output_docx.get().strip()),
            reference_docx=self._optional_path(self.reference_docx),
            bibliography=self._optional_path(self.bibliography),
            csl=self._optional_path(self.csl),
        )

    def _optional_path(self, variable: tk.StringVar) -> Path | None:
        value = variable.get().strip()
        return Path(value) if value else None

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self._events.get_nowait()
                if event == "pandoc":
                    self.pandoc_status.set(str(payload))
                    self._log(str(payload))
                    self._set_busy(False)
                elif event == "pandoc_error":
                    self.pandoc_status.set("Pandoc 未就绪")
                    self._log(str(payload))
                    messagebox.showwarning("Pandoc 未就绪", str(payload))
                    self._set_busy(False)
                elif event == "converted":
                    result = payload
                    self._log(f"导出完成：{result.output_docx}")
                    self._log(f"日志文件：{result.log_path}")
                    if result.review_markdown_path is not None:
                        self._log(f"格式检查报告：{result.review_markdown_path}")
                        self._log(
                            f"格式检查摘要：发现 {result.review_error_count} 个严重问题，"
                            f"{result.review_warning_count} 个警告。"
                        )
                    if result.ai_review_bundle_dir is not None:
                        self._log(f"AI 审稿包：{result.ai_review_bundle_dir}")
                        self._log(f"AI 审稿 Prompt：{result.ai_review_prompt_path}")
                    for warning in result.warnings:
                        self._log(f"警告：{warning}")
                    message = f"DOCX 已生成：\n{result.output_docx}\n\n日志文件：\n{result.log_path}"
                    if result.review_markdown_path is not None:
                        message += (
                            f"\n\n格式检查报告：\n{result.review_markdown_path}"
                            f"\n\n发现 {result.review_error_count} 个严重问题，"
                            f"{result.review_warning_count} 个警告。"
                        )
                    if result.ai_review_bundle_dir is not None:
                        message += f"\n\nAI 审稿包：\n{result.ai_review_bundle_dir}"
                    if result.warnings:
                        message += "\n\n有警告，请查看转换日志。"
                    messagebox.showinfo("导出完成", message)
                    self._set_busy(False)
                elif event == "reviewed":
                    report = payload
                    self._log(f"格式检查完成：{report.markdown_path}")
                    self._log(f"格式检查摘要：发现 {report.error_count} 个严重问题，{report.warning_count} 个警告。")
                    messagebox.showinfo(
                        "检查完成",
                        f"格式检查报告：\n{report.markdown_path}\n\n"
                        f"发现 {report.error_count} 个严重问题，{report.warning_count} 个警告。",
                    )
                    self._set_busy(False)
                elif event == "review_error":
                    self._log(str(payload))
                    messagebox.showerror("检查失败", str(payload))
                    self._set_busy(False)
                elif event == "ai_bundle":
                    bundle = payload
                    self._log(f"AI 审稿包生成完成：{bundle.bundle_dir}")
                    self._log(f"可复制 Prompt：{bundle.prompt_path}")
                    messagebox.showinfo(
                        "审稿包生成完成",
                        f"AI 审稿包：\n{bundle.bundle_dir}\n\n可复制 Prompt：\n{bundle.prompt_path}",
                    )
                    self._set_busy(False)
                elif event == "ai_bundle_error":
                    self._log(str(payload))
                    messagebox.showerror("审稿包生成失败", str(payload))
                    self._set_busy(False)
                elif event == "convert_error":
                    self._log(str(payload))
                    messagebox.showerror("转换失败", str(payload))
                    self._set_busy(False)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _set_busy(self, busy: bool) -> None:
        self._is_busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        self.convert_button.configure(state=state)
        self.review_button.configure(state=state)
        self.ai_bundle_button.configure(state=state)

    def _log(self, message: str) -> None:
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", tk.END)


def main() -> None:
    app = ConverterApp()
    app.mainloop()
