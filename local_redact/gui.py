"""CustomTkinter 本机界面。复用 cli.main，不改脱敏逻辑。"""

from __future__ import annotations

import queue
import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from local_redact.cli import ROOT, main as cli_main

_DEFAULT_LEXICON = ROOT / "lexicon.txt"
_DEFAULT_OUTPUT = ROOT / "output"
_DEFAULT_MAP = ROOT / "data" / "mapping.json"

_LEXICON_HEADER = """# 一行一个要在全文、文件名和子目录名里替换的词：客户全称、简称、厂区地址、两三字的人名。
# 指定固定代号：原词 => 代号
# 不写代号时，工具会按本机密钥生成稳定代号（同一词每次相同）。
# 以 # 开头的是注释。不要把「机构」「报告」这种代号前缀写进来。
#
# 示例环保科技有限公司
# 文三路1号
# 张三 => 人员甲

"""

_AUTO_TYPES_TEXT = """不依赖词表也会自动处理：
· 栏目后的值：委托/受检/建设等单位，项目名称，地址，联系人与采样/编制/审核/签发人，电话
· 结果表人员列：采样/检测人员、采样人员、检测人员、监测人员、人员（表头整行可识别时按列打码；格内多人按行拆分）
· 编号类：项目/合同/许可证号、统一社会信用代码（邮编、报告编号/文号、点位编号、样品编号保留）
· 格式类：手机与固话、坐标、身份证、银行卡、邮箱、网址、内网 IP、密钥类字符串
· 一律删除：图片、公章、签名、二维码、水印、背景图、嵌入对象；页眉页脚同样处理
需写入词表：客户简称、两三字人名（不在上述栏位/人员列时）、无标签的页脚地址/英文机构名、裸域名等"""

_LEXICON_HELP = """写法：一行一个词。只要原词则自动生成稳定代号；需要固定写法时用「原词 => 代号」。
# 开头是注释。保存后写入上方路径；开始脱敏前若有未保存修改会先自动保存。"""


class _QueueWriter:
    def __init__(self, target: queue.Queue[str]):
        self._target = target

    def write(self, text: str) -> int:
        if text:
            self._target.put(text)
        return len(text)

    def flush(self) -> None:
        return None


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("local-redact")
        self.geometry("860x760")
        self.minsize(720, 640)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self._log_queue: queue.Queue[str] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._busy = False
        self._lex_dirty = False
        self._lex_loaded_text = ""

        self.src_var = ctk.StringVar()
        self.out_var = ctk.StringVar(value=str(_DEFAULT_OUTPUT))
        self.lex_var = ctk.StringVar(
            value=str(_DEFAULT_LEXICON) if _DEFAULT_LEXICON.exists() else str(_DEFAULT_LEXICON)
        )

        self._build()
        self._load_lexicon_into_editor(show_error=False)
        self.after(80, self._drain_log)

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        tabs = ctk.CTkTabview(self)
        tabs.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        run_tab = tabs.add("脱敏")
        lex_tab = tabs.add("词表")
        self._build_run_tab(run_tab)
        self._build_lex_tab(lex_tab)

    def _build_run_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(5, weight=1)

        tip = ctk.CTkLabel(
            tab,
            text="数据只在本机处理。原稿不会被修改；对照表含原文，不要和脱敏结果一起上传。",
            anchor="w",
            justify="left",
            wraplength=800,
        )
        tip.grid(row=0, column=0, sticky="ew", padx=8, pady=(4, 6))

        form = ctk.CTkFrame(tab, fg_color="transparent")
        form.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        form.grid_columnconfigure(1, weight=1)

        self._row(form, 0, "输入", self.src_var, self._pick_src_file, self._pick_src_dir)
        self._row(form, 1, "输出目录", self.out_var, self._pick_out_dir)
        self._row(form, 2, "词表文件", self.lex_var, self._pick_lexicon)

        help_box = ctk.CTkFrame(tab)
        help_box.grid(row=2, column=0, sticky="ew", padx=8, pady=(8, 4))
        help_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(help_box, text="可自动识别的类型", anchor="w", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 2)
        )
        ctk.CTkLabel(
            help_box,
            text=_AUTO_TYPES_TEXT,
            anchor="nw",
            justify="left",
            wraplength=780,
        ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))

        actions = ctk.CTkFrame(tab, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=8, pady=(8, 4))

        self.run_btn = ctk.CTkButton(actions, text="开始脱敏", width=120, command=self._start)
        self.run_btn.grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(actions, text="打开输出目录", width=110, command=self._open_output).grid(
            row=0, column=1, padx=4
        )
        ctk.CTkButton(actions, text="打开对照表", width=110, command=self._open_mapping).grid(
            row=0, column=2, padx=4
        )
        ctk.CTkButton(actions, text="清空日志", width=90, command=self._clear_log).grid(
            row=0, column=3, padx=4
        )

        self.status = ctk.CTkLabel(tab, text="就绪", anchor="w")
        self.status.grid(row=4, column=0, sticky="ew", padx=8, pady=(2, 4))

        self.log = ctk.CTkTextbox(tab, wrap="word", font=ctk.CTkFont(family="Consolas", size=13))
        self.log.grid(row=5, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.log.insert("end", "选择输入后点「开始脱敏」。词表可在「词表」页编辑保存。\n")
        self.log.configure(state="disabled")

    def _build_lex_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            tab,
            text=_LEXICON_HELP,
            anchor="w",
            justify="left",
            wraplength=800,
        ).grid(row=0, column=0, sticky="ew", padx=8, pady=(4, 6))

        path_row = ctk.CTkFrame(tab, fg_color="transparent")
        path_row.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        path_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(path_row, text="保存到", width=56, anchor="w").grid(row=0, column=0, padx=(0, 8))
        ctk.CTkEntry(path_row, textvariable=self.lex_var).grid(row=0, column=1, sticky="ew")
        ctk.CTkButton(path_row, text="浏览", width=72, command=self._pick_lexicon).grid(
            row=0, column=2, padx=(8, 0)
        )
        ctk.CTkButton(path_row, text="重新加载", width=88, command=self._reload_lexicon).grid(
            row=0, column=3, padx=(8, 0)
        )
        ctk.CTkButton(path_row, text="保存词表", width=88, command=self._save_lexicon).grid(
            row=0, column=4, padx=(8, 0)
        )

        self.lex_editor = ctk.CTkTextbox(tab, wrap="word", font=ctk.CTkFont(family="Consolas", size=13))
        self.lex_editor.grid(row=2, column=0, sticky="nsew", padx=8, pady=(4, 4))
        self.lex_editor.bind("<<Modified>>", self._on_lex_modified)

        bottom = ctk.CTkFrame(tab, fg_color="transparent")
        bottom.grid(row=3, column=0, sticky="ew", padx=8, pady=(4, 8))
        self.lex_status = ctk.CTkLabel(bottom, text="", anchor="w")
        self.lex_status.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(bottom, text="插入示例模板", width=120, command=self._insert_lexicon_template).grid(
            row=0, column=1, padx=(12, 0)
        )

    def _row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: ctk.StringVar,
        *pickers,
        clearable: bool = False,
    ) -> None:
        ctk.CTkLabel(parent, text=label, width=72, anchor="w").grid(
            row=row, column=0, sticky="w", pady=4, padx=(0, 8)
        )
        entry = ctk.CTkEntry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        col = 2
        captions = ("选文件", "选文件夹") if len(pickers) > 1 else ("浏览",)
        for caption, picker in zip(captions, pickers):
            ctk.CTkButton(parent, text=caption, width=88, command=picker).grid(
                row=row, column=col, padx=(8, 0), pady=4
            )
            col += 1
        if clearable:
            ctk.CTkButton(
                parent, text="清空", width=64, command=lambda: variable.set("")
            ).grid(row=row, column=col, padx=(8, 0), pady=4)

    def _pick_src_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Word 报告",
            filetypes=[
                ("Word 文档", "*.doc *.docx"),
                ("全部文件", "*.*"),
            ],
        )
        if path:
            self.src_var.set(path)

    def _pick_src_dir(self) -> None:
        path = filedialog.askdirectory(title="选择报告文件夹")
        if path:
            self.src_var.set(path)

    def _pick_out_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.out_var.set(path)

    def _pick_lexicon(self) -> None:
        path = filedialog.askopenfilename(
            title="选择词表",
            filetypes=[("文本", "*.txt"), ("全部文件", "*.*")],
            initialdir=str(ROOT),
        )
        if path:
            self.lex_var.set(path)
            self._load_lexicon_into_editor()

    def _on_lex_modified(self, _event=None) -> None:
        try:
            if self.lex_editor.edit_modified():
                current = self.lex_editor.get("1.0", "end-1c")
                self._lex_dirty = current != self._lex_loaded_text
                self._update_lex_status()
                self.lex_editor.edit_modified(False)
        except Exception:
            pass

    def _update_lex_status(self) -> None:
        path = self.lex_var.get().strip() or str(_DEFAULT_LEXICON)
        flag = "（有未保存修改）" if self._lex_dirty else ""
        self.lex_status.configure(text=f"{path} {flag}".strip())

    def _set_editor_text(self, text: str) -> None:
        self.lex_editor.delete("1.0", "end")
        self.lex_editor.insert("1.0", text)
        self._lex_loaded_text = text
        self._lex_dirty = False
        self.lex_editor.edit_modified(False)
        self._update_lex_status()

    def _load_lexicon_into_editor(self, show_error: bool = True) -> None:
        raw = self.lex_var.get().strip() or str(_DEFAULT_LEXICON)
        path = Path(raw)
        if not path.exists():
            self._set_editor_text(_LEXICON_HEADER)
            if show_error:
                self.lex_status.configure(text=f"文件尚不存在，已载入模板，保存后写入：{path}")
            else:
                self._update_lex_status()
            return
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as exc:
            if show_error:
                messagebox.showerror("错误", f"无法读取词表：\n{exc}")
            return
        self.lex_var.set(str(path))
        self._set_editor_text(text)

    def _reload_lexicon(self) -> None:
        if self._lex_dirty and not messagebox.askyesno("重新加载", "有未保存修改，确定放弃并重新加载？"):
            return
        self._load_lexicon_into_editor()

    def _insert_lexicon_template(self) -> None:
        current = self.lex_editor.get("1.0", "end-1c").strip()
        if current and not messagebox.askyesno("插入模板", "将用示例模板替换当前编辑内容，继续？"):
            return
        self._set_editor_text(_LEXICON_HEADER)
        self._lex_dirty = True
        self._update_lex_status()

    def _save_lexicon(self, quiet: bool = False) -> bool:
        raw = self.lex_var.get().strip() or str(_DEFAULT_LEXICON)
        path = Path(raw)
        text = self.lex_editor.get("1.0", "end-1c")
        if text and not text.endswith("\n"):
            text += "\n"
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("错误", f"保存词表失败：\n{exc}")
            return False
        self.lex_var.set(str(path))
        self._set_editor_text(path.read_text(encoding="utf-8"))
        if not quiet:
            messagebox.showinfo("词表", f"已保存到：\n{path}")
        self._append_log(f"词表已保存：{path}\n")
        return True

    def _ensure_lexicon_saved(self) -> bool:
        if not self._lex_dirty:
            return True
        return self._save_lexicon(quiet=True)

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _drain_log(self) -> None:
        try:
            while True:
                self._append_log(self._log_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(80, self._drain_log)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.run_btn.configure(state=state, text="处理中…" if busy else "开始脱敏")
        self.status.configure(text="正在脱敏…" if busy else "就绪")

    def _start(self) -> None:
        if self._busy:
            return
        src = self.src_var.get().strip()
        out = self.out_var.get().strip()
        if not src:
            messagebox.showwarning("提示", "请先选择输入文件或文件夹。")
            return
        if not out:
            messagebox.showwarning("提示", "请指定输出目录。")
            return
        src_path = Path(src)
        if not src_path.exists():
            messagebox.showerror("错误", f"找不到输入：\n{src}")
            return
        if not self._ensure_lexicon_saved():
            return

        argv = [src, "-o", out]
        lexicon = self.lex_var.get().strip()
        if lexicon:
            argv.extend(["--lexicon", lexicon])

        self._set_busy(True)
        self._append_log("\n—— 开始 ——\n")
        self._worker = threading.Thread(target=self._run_job, args=(argv,), daemon=True)
        self._worker.start()

    def _run_job(self, argv: list[str]) -> None:
        writer = _QueueWriter(self._log_queue)
        old_out, old_err = sys.stdout, sys.stderr
        code = 1
        try:
            sys.stdout = writer  # type: ignore[assignment]
            sys.stderr = writer  # type: ignore[assignment]
            code = cli_main(argv)
        except Exception:
            self._log_queue.put(traceback.format_exc())
            code = 1
        finally:
            sys.stdout = old_out
            sys.stderr = old_err
            self.after(0, lambda: self._finished(code))

    def _finished(self, code: int) -> None:
        self._set_busy(False)
        if code == 0:
            self.status.configure(text="完成")
            self._append_log("—— 完成 ——\n")
        else:
            self.status.configure(text="完成（有失败或错误）")
            self._append_log("—— 结束（有失败或错误）——\n")

    def _open_output(self) -> None:
        path = Path(self.out_var.get().strip() or str(_DEFAULT_OUTPUT))
        path.mkdir(parents=True, exist_ok=True)
        self._reveal(path)

    def _open_mapping(self) -> None:
        path = _DEFAULT_MAP
        if path.exists():
            self._reveal(path)
        else:
            path.parent.mkdir(parents=True, exist_ok=True)
            messagebox.showinfo(
                "对照表",
                f"对照表尚未生成。\n跑完脱敏后会出现在：\n{path}",
            )
            self._reveal(path.parent)

    @staticmethod
    def _reveal(path: Path) -> None:
        import os

        os.startfile(path)  # noqa: S606 — Windows 本机打开


def main() -> int:
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
