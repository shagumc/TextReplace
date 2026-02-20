import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from dataclasses import dataclass, asdict
from pathlib import Path
import difflib

APP_NAME = "Text Replace"
RULES_FILE = Path.home() / ".text_replace_rules.json"


@dataclass
class Rule:
    enabled: bool
    src: str
    dst: str


# -----------------------------
# Tooltip
# -----------------------------
class Tooltip:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.tip = None

    def show(self, x, y, text):
        self.hide()
        self.tip = tk.Toplevel(self.master)
        self.tip.wm_overrideredirect(True)
        self.tip.attributes("-topmost", True)
        self.tip.geometry(f"+{x+12}+{y+12}")
        label = ttk.Label(self.tip, text=text, relief="solid", borderwidth=1, padding=(8, 6))
        label.pack()

    def hide(self):
        if self.tip:
            try:
                self.tip.destroy()
            except Exception:
                pass
        self.tip = None


# -----------------------------
# Scrollable Frame (for rules)
# -----------------------------
class ScrollableFrame(ttk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vscroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vscroll.set)

        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vscroll.grid(row=0, column=1, sticky="ns")

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Mouse wheel (Windows/Mac)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)       # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)   # Linux
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _on_inner_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.inner_id, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")


# -----------------------------
# Rule Manager (Row widgets)
# -----------------------------
class RuleManager(tk.Toplevel):
    def __init__(self, master, rules: list[Rule], on_save):
        super().__init__(master)
        self.title("辞書（置換ルール）")
        self.geometry("980x520")
        self.minsize(900, 460)

        self.master = master
        self.rules = rules
        self.on_save = on_save

        self.transient(master)
        self.grab_set()
        self.lift()
        self.focus_force()

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(2, weight=1)
        header.columnconfigure(3, weight=1)

        ttk.Label(header, text="適用", width=6).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="置換前", width=40).grid(row=0, column=2, sticky="w", padx=(10, 0))
        ttk.Label(header, text="置換後", width=40).grid(row=0, column=3, sticky="w", padx=(10, 0))
        ttk.Label(header, text="操作", width=18).grid(row=0, column=4, sticky="e")

        self.sf = ScrollableFrame(outer)
        self.sf.grid(row=1, column=0, sticky="nsew")

        footer = ttk.Frame(outer)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)

        self.status = ttk.Label(footer, text="", foreground="gray")
        self.status.grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(footer)
        btns.grid(row=0, column=1, sticky="e")

        ttk.Button(btns, text="＋ 追加", command=self.add_row).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="保存", command=self.save).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="閉じる", command=self.close).pack(side="left")

        self.row_widgets = []
        self.render_rows()

    def close(self):
        self.commit_to_model()
        self.destroy()

    def render_rows(self):
        for child in self.sf.inner.winfo_children():
            child.destroy()
        self.row_widgets.clear()

        for idx, rule in enumerate(self.rules):
            self._create_row(idx, rule)

        if not self.rules:
            hint = ttk.Label(self.sf.inner, text="「＋ 追加」でルールを作成できます。", foreground="gray")
            hint.pack(anchor="w", padx=6, pady=10)

    def _create_row(self, idx: int, rule: Rule):
        row = ttk.Frame(self.sf.inner)
        row.pack(fill="x", pady=4)

        v_enabled = tk.BooleanVar(master=self, value=rule.enabled)
        v_src = tk.StringVar(master=self, value=rule.src)
        v_dst = tk.StringVar(master=self, value=rule.dst)

        cb = ttk.Checkbutton(row, variable=v_enabled)
        cb.pack(side="left", padx=(2, 6))

        e_src = tk.Entry(row, textvariable=v_src, bd=1, relief="solid")
        e_src.pack(side="left", fill="x", expand=True, padx=(6, 6))
        e_dst = tk.Entry(row, textvariable=v_dst, bd=1, relief="solid")
        e_dst.pack(side="left", fill="x", expand=True, padx=(6, 6))

        ops = ttk.Frame(row)
        ops.pack(side="right")

        ttk.Button(ops, text="↑", width=3, command=lambda i=idx: self.move_row(i, -1)).pack(side="left", padx=(0, 4))
        ttk.Button(ops, text="↓", width=3, command=lambda i=idx: self.move_row(i, +1)).pack(side="left", padx=(0, 8))
        ttk.Button(ops, text="削除", command=lambda i=idx: self.delete_row(i)).pack(side="left")

        self.row_widgets.append({
            "enabled": v_enabled,
            "src": v_src,
            "dst": v_dst,
            "e_src": e_src,
            "e_dst": e_dst,
        })

    def add_row(self):
        self.commit_to_model()
        self.rules.append(Rule(enabled=True, src="", dst=""))
        self.render_rows()
        if self.row_widgets:
            self.row_widgets[-1]["e_src"].focus_set()
        self.status.config(text="行を追加しました")
        self.lift()
        self.focus_force()

    def delete_row(self, idx: int):
        self.commit_to_model()
        if idx < 0 or idx >= len(self.rules):
            return
        r = self.rules[idx]
        ok = messagebox.askyesno(
            "削除確認",
            f"この行を削除しますか？\n\n置換前: {r.src}\n置換後: {r.dst}",
            parent=self
        )
        if not ok:
            self.lift()
            self.focus_force()
            return

        self.rules.pop(idx)
        self.render_rows()
        self.status.config(text="行を削除しました")
        self.lift()
        self.focus_force()

    def move_row(self, idx: int, direction: int):
        self.commit_to_model()
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.rules):
            return
        self.rules[idx], self.rules[new_idx] = self.rules[new_idx], self.rules[idx]
        self.render_rows()
        self.status.config(text="行の順番を変更しました")
        self.lift()
        self.focus_force()

    def commit_to_model(self):
        new_rules = []
        for rw in self.row_widgets:
            new_rules.append(Rule(
                enabled=bool(rw["enabled"].get()),
                src=rw["src"].get(),
                dst=rw["dst"].get(),
            ))
        self.rules[:] = new_rules

    def save(self):
        self.commit_to_model()
        self.on_save()
        self.status.config(text="保存しました")
        self.lift()
        self.focus_force()


# -----------------------------
# Main App
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1100x780")

        self.rules: list[Rule] = self.load_rules()
        self.tooltip = Tooltip(self)

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="辞書（ルール）", command=self.open_rules).pack(side="left")

        # ★追加：TXT読み込みボタン（入力欄を全削除→ファイル内容を表示）
        ttk.Button(top, text="TXTを読み込み", command=self.open_text_file).pack(side="left", padx=8)

        ttk.Button(top, text="置換実行", command=self.replace).pack(side="left", padx=8)

        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Button(top, text="出力をコピー", command=self.copy).pack(side="left")
        ttk.Button(top, text="TXTとして保存", command=self.save_output).pack(side="left", padx=8)

        self.status = ttk.Label(top, text="", foreground="gray")
        self.status.pack(side="right")
        self.update_status()

        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        in_frame = ttk.Labelframe(paned, text="入力（ここに貼り付け / TXT読み込み）", padding=8)
        out_frame = ttk.Labelframe(paned, text="出力（置換箇所は色付き / ホバーで置換前表示）", padding=8)

        paned.add(in_frame, weight=1)
        paned.add(out_frame, weight=1)

        self.input = tk.Text(in_frame, wrap="word", undo=True)
        self.input.pack(fill="both", expand=True)

        self.output = tk.Text(out_frame, wrap="word")
        self.output.pack(fill="both", expand=True)

        self.output.tag_configure("chg", background="#fff3b0")
        self.output.bind("<Motion>", self.on_hover)
        self.output.bind("<Leave>", lambda e: self.tooltip.hide())

        self.tag_map = {}
        self._last_hover_tag = None

        if not self.input.get("1.0", "end-1c").strip():
            self.input.insert("1.0", "ここに貼り付け → 置換実行\n")

    # ★追加：ファイル読み込み（入力欄を全削除して差し替え）
    def open_text_file(self):
        path = filedialog.askopenfilename(
            title="テキストファイルを選択",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            content = Path(path).read_text(encoding="utf-8")
        except UnicodeDecodeError:
            # Shift-JISなどの可能性を考慮（追加で安全に）
            try:
                content = Path(path).read_text(encoding="cp932")
            except Exception as e:
                messagebox.showerror("読み込みエラー", f"文字コードの判定に失敗しました:\n{e}")
                return
        except Exception as e:
            messagebox.showerror("読み込みエラー", str(e))
            return

        # 既存テキストを全削除 → 新規表示
        self.input.delete("1.0", "end")
        self.input.insert("1.0", content)
        self.status.config(text=f"読み込みました: {Path(path).name}")

    # --- persistence ---
    def load_rules(self) -> list[Rule]:
        if RULES_FILE.exists():
            try:
                data = json.loads(RULES_FILE.read_text(encoding="utf-8"))
                out = []
                for r in data:
                    out.append(Rule(
                        enabled=bool(r.get("enabled", True)),
                        src=str(r.get("src", "")),
                        dst=str(r.get("dst", "")),
                    ))
                return out
            except Exception:
                return []
        return []

    def save_rules(self):
        RULES_FILE.write_text(
            json.dumps([asdict(r) for r in self.rules], ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        self.update_status()

    def update_status(self):
        enabled = sum(1 for r in self.rules if r.enabled and r.src)
        total = len(self.rules)
        self.status.config(text=f"適用ルール: {enabled}/{total}（上から順に実行）")

    # --- UI actions ---
    def open_rules(self):
        RuleManager(self, self.rules, on_save=self.save_rules)

    def replace(self):
        src_text = self.input.get("1.0", "end-1c")
        enabled_rules = [r for r in self.rules if r.enabled and r.src != ""]

        if not enabled_rules:
            messagebox.showwarning("ルールなし", "適用するルールがありません。\n辞書（ルール）で追加・ONにしてください。")
            return

        # 置換は「上から順に1回ずつ」適用（循環しません）
        out = src_text
        for r in enabled_rules:
            out = out.replace(r.src, r.dst)

        self.output.delete("1.0", "end")
        self.output.insert("1.0", out)

        self.apply_diff_highlight(src_text, out)
        self.update_status()

    def copy(self):
        text = self.output.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.config(text="コピーしました")

    def save_output(self):
        text = self.output.get("1.0", "end-1c")
        path = filedialog.asksaveasfilename(
            title="TXTとして保存",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")]
        )
        if not path:
            return
        Path(path).write_text(text, encoding="utf-8")
        self.status.config(text=f"保存しました: {Path(path).name}")

    # --- highlight / hover ---
    def clear_highlight(self):
        for tag in list(self.tag_map.keys()):
            try:
                self.output.tag_delete(tag)
            except Exception:
                pass
        self.tag_map.clear()
        self._last_hover_tag = None
        self.tooltip.hide()

    def apply_diff_highlight(self, original: str, modified: str):
        self.clear_highlight()

        sm = difflib.SequenceMatcher(a=original, b=modified)
        k = 0
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op == "equal":
                continue

            if op in ("replace", "insert") and j1 != j2:
                k += 1
                tag = f"chg_{k}"
                self.output.tag_configure(tag, background="#fff3b0")
                self.output.tag_add(tag, f"1.0+{j1}c", f"1.0+{j2}c")

                before = original[i1:i2]
                disp = before if len(before) <= 160 else before[:160] + "…"
                self.tag_map[tag] = disp

    def on_hover(self, event):
        idx = self.output.index(f"@{event.x},{event.y}")
        tags = self.output.tag_names(idx)
        dyn = None
        for t in tags:
            if t.startswith("chg_"):
                dyn = t
                break

        if dyn and dyn in self.tag_map:
            if self._last_hover_tag != dyn:
                self._last_hover_tag = dyn
                self.tooltip.show(self.winfo_pointerx(), self.winfo_pointery(), f"置換前: {self.tag_map[dyn]}")
        else:
            self._last_hover_tag = None
            self.tooltip.hide()


if __name__ == "__main__":
    app = App()
    app.mainloop()

