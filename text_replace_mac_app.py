import sys
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from dataclasses import dataclass, asdict
from pathlib import Path
import difflib
from typing import List, Dict, Optional, Callable
import tkinter.font as tkfont

APP_NAME = "Text Replace"

OLD_RULES_FILE = Path.home() / ".text_replace_rules.json"
DICT_STORE_FILE = Path.home() / ".text_replace_dictionaries.json"
SETTINGS_FILE = Path.home() / ".text_replace_settings.json"


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
    """
    Rules list frame with ttk.Scrollbar (same look as your dictionary dialog).
    Wheel binding in RuleManager is handled by binding recursively to all widgets,
    so here we disable internal bind by default for the rules dialog.
    """
    def __init__(self, master, enable_wheel_bind: bool = True, **kwargs):
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

        self._enable_wheel_bind = enable_wheel_bind
        self._wheel_bound = False

        if self._enable_wheel_bind:
            for w in (self.canvas, self.inner):
                w.bind("<Enter>", self._bind_wheel, add="+")
                w.bind("<Leave>", self._unbind_wheel, add="+")

    def _on_inner_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.inner_id, width=event.width)

    def _bind_wheel(self, _event=None):
        if not self._enable_wheel_bind:
            return
        if self._wheel_bound:
            return
        self._wheel_bound = True
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.bind_all("<Button-4>", self._on_mousewheel_linux, add="+")
        self.bind_all("<Button-5>", self._on_mousewheel_linux, add="+")

    def _unbind_wheel(self, _event=None):
        if not self._enable_wheel_bind:
            return
        if not self._wheel_bound:
            return
        self._wheel_bound = False
        try:
            self.unbind_all("<MouseWheel>")
            self.unbind_all("<Button-4>")
            self.unbind_all("<Button-5>")
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            if not self.winfo_exists() or not self.canvas.winfo_exists():
                return
            delta = getattr(event, "delta", 0)
            if delta == 0:
                return
            units = int(-1 * (delta / 120))
            if units == 0:
                units = -1 if delta > 0 else 1
            self.canvas.yview_scroll(units, "units")
        except tk.TclError:
            return

    def _on_mousewheel_linux(self, event):
        try:
            if not self.winfo_exists() or not self.canvas.winfo_exists():
                return
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        except tk.TclError:
            return

    def destroy(self):
        try:
            self._unbind_wheel()
        except Exception:
            pass
        super().destroy()


# -----------------------------
# Dictionary Store (multiple dicts)
# -----------------------------
class DictStore:
    """
    DICT_STORE_FILE: {"version":1,"dicts":{"default":[{rule},...], "foo":[...]}}
    旧形式 OLD_RULES_FILE があれば初回に default として移行。
    """
    def __init__(self, path: Path):
        self.path = path
        self.dicts: Dict[str, List[Rule]] = {}

    def load(self):
        if not self.path.exists():
            migrated = self._try_migrate_from_old()
            if migrated:
                self.save()
                return
            self.dicts = {"default": []}
            self.save()
            return

        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            d = data.get("dicts", {})
            out: Dict[str, List[Rule]] = {}
            for name, rules_list in d.items():
                rules: List[Rule] = []
                for r in rules_list:
                    rules.append(Rule(
                        enabled=bool(r.get("enabled", True)),
                        src=str(r.get("src", "")),
                        dst=str(r.get("dst", "")),
                    ))
                out[str(name)] = rules
            if not out:
                out = {"default": []}
            self.dicts = out
        except Exception:
            self.dicts = {"default": []}

    def _try_migrate_from_old(self) -> bool:
        if not OLD_RULES_FILE.exists():
            return False
        try:
            data = json.loads(OLD_RULES_FILE.read_text(encoding="utf-8"))
            rules: List[Rule] = []
            for r in data:
                rules.append(Rule(
                    enabled=bool(r.get("enabled", True)),
                    src=str(r.get("src", "")),
                    dst=str(r.get("dst", "")),
                ))
            self.dicts = {"default": rules}
            return True
        except Exception:
            return False

    def save(self):
        payload = {
            "version": 1,
            "dicts": {
                name: [asdict(r) for r in rules]
                for name, rules in self.dicts.items()
            }
        }
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def names(self) -> List[str]:
        return sorted(self.dicts.keys(), key=lambda s: (s != "default", s.lower()))

    def get_rules(self, name: str) -> List[Rule]:
        if name not in self.dicts:
            self.dicts[name] = []
        return self.dicts[name]

    def create(self, name: str) -> bool:
        name = name.strip()
        if not name or name in self.dicts:
            return False
        self.dicts[name] = []
        return True

    def delete(self, name: str) -> bool:
        if name == "default":
            return False
        if name not in self.dicts:
            return False
        del self.dicts[name]
        if "default" not in self.dicts:
            self.dicts["default"] = []
        return True


# -----------------------------
# Apply picker popup (multi-select / stays open on click)
# -----------------------------
class ApplyPickerPopup(tk.Toplevel):
    def __init__(
        self,
        master: tk.Tk,
        anchor_widget: tk.Widget,
        names: List[str],
        vars_by_name: Dict[str, tk.BooleanVar],
        on_change: Callable[[], None],
        on_closed: Callable[[], None],
    ):
        super().__init__(master)
        self.master = master
        self.anchor_widget = anchor_widget
        self.names = names
        self.vars_by_name = vars_by_name
        self.on_change = on_change
        self.on_closed = on_closed

        self.overrideredirect(True)
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass

        if sys.platform == "darwin":
            try:
                self.tk.call("::tk::unsupported::MacWindowStyle", "style", self._w, "help", "noActivates")
            except Exception:
                pass

        self.update_idletasks()
        x = anchor_widget.winfo_rootx()
        y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height()
        self.geometry(f"+{x}+{y}")

        outer = ttk.Frame(self, padding=8, relief="solid", borderwidth=1)
        outer.pack(fill="both", expand=True)

        max_h = int(master.winfo_screenheight() * 0.45)
        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(list_frame, highlightthickness=0, height=1)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def _on_inner_configure(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            try:
                req = inner.winfo_reqheight() + 4
                canvas.configure(height=min(req, max_h))
            except Exception:
                pass

        def _on_canvas_configure(e):
            try:
                canvas.itemconfigure(inner_id, width=e.width)
            except Exception:
                pass

        inner.bind("<Configure>", _on_inner_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # ホイール（トラックパッド含む）
        self._wheel_bound = False

        def _on_wheel(event):
            try:
                delta = getattr(event, "delta", 0)
                if delta == 0:
                    return "break"
                units = int(-1 * (delta / 120))
                if units == 0:
                    units = -1 if delta > 0 else 1
                canvas.yview_scroll(units, "units")
            except Exception:
                pass
            return "break"

        def _on_wheel_linux(event):
            try:
                if event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")
            except Exception:
                pass
            return "break"

        def _bind_wheel(_e=None):
            if self._wheel_bound:
                return
            self._wheel_bound = True
            self.bind_all("<MouseWheel>", _on_wheel, add="+")
            self.bind_all("<Button-4>", _on_wheel_linux, add="+")
            self.bind_all("<Button-5>", _on_wheel_linux, add="+")

        def _unbind_wheel(_e=None):
            if not self._wheel_bound:
                return
            self._wheel_bound = False
            try:
                self.unbind_all("<MouseWheel>")
                self.unbind_all("<Button-4>")
                self.unbind_all("<Button-5>")
            except Exception:
                pass

        for w in (canvas, inner, list_frame, outer):
            w.bind("<Enter>", _bind_wheel, add="+")
            w.bind("<Leave>", _unbind_wheel, add="+")

        for name in names:
            v = vars_by_name[name]
            cb = ttk.Checkbutton(inner, text=name, variable=v, command=self._changed)
            cb.pack(anchor="w")

        sep = ttk.Separator(outer)
        sep.pack(fill="x", pady=(8, 6))

        btns = ttk.Frame(outer)
        btns.pack(fill="x")
        ttk.Button(btns, text="すべて解除", command=self._clear_all).pack(side="left")
        ttk.Button(btns, text="閉じる", command=self.close).pack(side="right")

        self._prev_bind_all = self.master.bind_all("<Button-1>")
        self.master.bind_all("<Button-1>", self._on_global_click, add="+")

        self.bind("<Escape>", lambda _e: self.close())
        self.protocol("WM_DELETE_WINDOW", self.close)

        self.after(1, _on_inner_configure)

    def _changed(self):
        self.on_change()

    def _clear_all(self):
        for v in self.vars_by_name.values():
            try:
                v.set(False)
            except Exception:
                pass
        self.on_change()

    def _on_global_click(self, event):
        w = event.widget
        if self._is_descendant_of(w, self):
            return
        if w == self.anchor_widget or self._is_descendant_of(w, self.anchor_widget):
            self.close()
            return
        self.close()

    @staticmethod
    def _is_descendant_of(widget, parent) -> bool:
        try:
            while widget is not None:
                if widget == parent:
                    return True
                widget = widget.master
        except Exception:
            pass
        return False

    def close(self):
        try:
            self.master.unbind_all("<Button-1>")
            if self._prev_bind_all:
                self.master.bind_all("<Button-1>", self._prev_bind_all)
        except Exception:
            pass

        try:
            self.destroy()
        except Exception:
            pass

        try:
            self.on_closed()
        except Exception:
            pass


# -----------------------------
# Line numbers
# -----------------------------
class LineNumberCanvas(tk.Canvas):
    def __init__(self, master, text_widget: tk.Text, **kwargs):
        super().__init__(master, highlightthickness=0, **kwargs)
        self.text = text_widget
        self._after_id = None

    def schedule_redraw(self):
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = self.after(10, self.redraw)

    def redraw(self):
        self._after_id = None
        self.delete("all")

        i = self.text.index("@0,0")
        while True:
            d = self.text.dlineinfo(i)
            if d is None:
                break
            y = d[1]
            line = i.split(".")[0]
            self.create_text(4, y, anchor="nw", text=line, fill="#666666")
            i = self.text.index(f"{i}+1line")


# -----------------------------
# Rule Manager (modeless + autosave + dict selector + singleton)
# -----------------------------
class RuleManager(tk.Toplevel):
    def __init__(
        self,
        master,
        store: DictStore,
        initial_dict_name: str,
        on_save_store: Callable[[], None],
        on_message: Callable[[str], None],
        on_saved_callback: Optional[Callable[[], None]] = None,
        on_closed: Optional[Callable[[], None]] = None,
    ):
        super().__init__(master)
        self.title("辞書（置換ルール）")

        self.minsize(520, 320)

        self.master = master
        self.store = store
        self.on_save_store = on_save_store
        self.on_message = on_message
        self.on_saved_callback = on_saved_callback
        self.on_closed = on_closed

        self._save_after_id = None
        self._switching = False

        init_name = (initial_dict_name or "default").strip() or "default"
        if init_name not in self.store.dicts:
            init_name = "default"
        self.dict_name_var = tk.StringVar(master=self, value=init_name)

        self.rules: List[Rule] = self.store.get_rules(init_name)

        try:
            self.attributes("-topmost", False)
        except Exception:
            pass

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        topbar = ttk.Frame(outer)
        topbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        topbar.columnconfigure(1, weight=1)

        ttk.Label(topbar, text="編集辞書:").grid(row=0, column=0, sticky="w")

        self.dict_combo = ttk.Combobox(
            topbar,
            textvariable=self.dict_name_var,
            values=self.store.names(),
            state="readonly",
            width=22,
        )
        self.dict_combo.grid(row=0, column=1, sticky="w", padx=(6, 0))
        self.dict_combo.bind("<<ComboboxSelected>>", self.on_dict_change)

        header = ttk.Frame(outer)
        header.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(2, weight=1)
        header.columnconfigure(3, weight=1)

        ttk.Label(header, text="適用", width=6).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="置換前").grid(row=0, column=2, sticky="w", padx=(10, 0))
        ttk.Label(header, text="置換後").grid(row=0, column=3, sticky="w", padx=(10, 0))
        ttk.Label(header, text="操作", width=16).grid(row=0, column=4, sticky="e")

        # ★ここでは wheel bind を ScrollableFrame に任せず、下の「全ウィジェットに直接bind」で確実化する
        self.sf = ScrollableFrame(outer, enable_wheel_bind=False)
        self.sf.grid(row=2, column=0, sticky="nsew")

        footer = ttk.Frame(outer)
        footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(footer, text="自動保存（入力欄から離れたら保存）", foreground="gray").grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(footer)
        btns.grid(row=0, column=1, sticky="e")

        ttk.Button(btns, text="＋ 追加", command=self.add_row).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="閉じる", command=self.close).pack(side="left")

        self.row_widgets = []
        self.render_rows()

        self.protocol("WM_DELETE_WINDOW", self.close)

        # ★macトラックパッドでも「どこでも」スクロールできるように（辞書ダイアログ配下へ直接bind）
        self._install_wheel_bind_recursive_widgets()

    def _on_rules_touchpad(self, event):
        """
        ★Tk 8.7+ などでトラックパッド2本指が <TouchpadScroll> になる環境用
        event.delta (%D) は「dx,dy を詰めた値」になるので、tk::PreciseScrollDeltas があればそれで解凍。
        """
        try:
            if not self.sf.winfo_exists() or not self.sf.canvas.winfo_exists():
                return "break"
        except Exception:
            return "break"
    
        d = getattr(event, "delta", 0)
        if d == 0:
            return "break"
    
        dx = 0
        dy = 0
    
        # Tk 8.7 の TIP 684: tk::PreciseScrollDeltas がある場合に dx,dy を取る
        try:
            # 戻り値は文字列のリストになる（例: "0 -3"）
            parts = self.tk.call("tk::PreciseScrollDeltas", d)
            # tk.call の戻りは tuple になる場合もあるので両対応
            if isinstance(parts, (tuple, list)) and len(parts) >= 2:
                dx = int(parts[0])
                dy = int(parts[1])
            else:
                s = str(parts)
                sp = s.split()
                if len(sp) >= 2:
                    dx = int(float(sp[0]))
                    dy = int(float(sp[1]))
        except Exception:
            # PreciseScrollDeltas が無い環境では dy 相当として雑に扱う
            dy = int(d)
    
        # dy が下方向なら +、上方向なら - にしたいので符号を調整
        # Canvas の yview_scroll は「+で下へ」なので dy の符号を反転
        units = 0
        if dy != 0:
            # 高頻度イベント対策：dyが小さいときは1行に丸める
            units = 1 if dy > 0 else -1
    
        if units != 0:
            try:
                self.sf.canvas.yview_scroll(units, "units")
            except Exception:
                pass
    
        return "break"

    # -----------------------------
    # ★ macトラックパッドでも「どこでも」スクロール（辞書ダイアログ専用）
    # -----------------------------
    def _on_rules_wheel(self, event):
        try:
            if not self.sf.winfo_exists() or not self.sf.canvas.winfo_exists():
                return "break"
        except Exception:
            return "break"

        delta = getattr(event, "delta", 0)
        if delta == 0:
            return "break"

        # トラックパッドは delta が小さいことがあるのでフォールバック
        units = int(-delta / 120)
        if units == 0:
            units = -1 if delta > 0 else 1

        try:
            self.sf.canvas.yview_scroll(units, "units")
        except Exception:
            pass
        return "break"

    def _on_rules_wheel_linux(self, event):
        try:
            if not self.sf.winfo_exists() or not self.sf.canvas.winfo_exists():
                return "break"
        except Exception:
            return "break"

        try:
            if event.num == 4:
                self.sf.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.sf.canvas.yview_scroll(1, "units")
        except Exception:
            pass
        return "break"

    WHEEL_TAG = "RuleManagerWheelTag"

    def _install_wheel_bind_recursive_widgets(self):
        """
        ★macトラックパッド対策：
        ttk/tk標準のMouseWheel処理に“勝つ”ため、bindtags先頭に自前タグを入れて最優先で奪う。
        """
        # タグ(=bindtag)に対するクラスバインド（最優先で呼ばれたら break する）
        try:
            self.bind_class(self.WHEEL_TAG, "<MouseWheel>", self._on_rules_wheel)
            self.bind_class(self.WHEEL_TAG, "<Shift-MouseWheel>", self._on_rules_wheel)  # 念のため
            self.bind_class(self.WHEEL_TAG, "<TouchpadScroll>", self._on_rules_touchpad)
            self.bind_class(self.WHEEL_TAG, "<Button-4>", self._on_rules_wheel_linux)
            self.bind_class(self.WHEEL_TAG, "<Button-5>", self._on_rules_wheel_linux)
        except Exception:
            pass
    
        def _apply(w: tk.Widget):
            try:
                tags = list(w.bindtags())
                if self.WHEEL_TAG not in tags:
                    # ★先頭に入れる（標準クラスバインドより前に発火させる）
                    tags.insert(0, self.WHEEL_TAG)
                    w.bindtags(tuple(tags))
            except Exception:
                pass
    
            for c in w.winfo_children():
                _apply(c)
    
        _apply(self)

    # ---- front control ----
    def bring_to_front_no_focus(self):
        try:
            self.lift()
        except Exception:
            pass
        try:
            self.attributes("-topmost", True)
            self.lift()
            self.after(10, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def focus_existing(self):
        try:
            self.deiconify()
        except Exception:
            pass
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def refresh_dict_names(self):
        names = self.store.names()
        self.dict_combo["values"] = names
        cur = (self.dict_name_var.get().strip() or "default")
        if cur not in names:
            self.dict_name_var.set("default")
            self._do_switch_dict("default")

    def on_dict_change(self, _event=None):
        name = (self.dict_name_var.get().strip() or "default")
        if name not in self.store.dicts:
            name = "default"
            self.dict_name_var.set("default")
        self._do_switch_dict(name)

    def _do_switch_dict(self, new_name: str):
        if self._switching:
            return
        self._switching = True
        try:
            self.perform_save()
            self.rules = self.store.get_rules(new_name)
            self.render_rows()
            self.on_message(f"編集中の辞書: {new_name}")
            if self.on_saved_callback:
                try:
                    self.on_saved_callback()
                except Exception:
                    pass
        finally:
            self._switching = False

    def close(self):
        self.perform_save()
        try:
            self.destroy()
        finally:
            if self.on_closed:
                try:
                    self.on_closed()
                except Exception:
                    pass

    def render_rows(self):
        for child in self.sf.inner.winfo_children():
            child.destroy()
        self.row_widgets.clear()

        for idx, rule in enumerate(self.rules):
            self._create_row(idx, rule)

        if not self.rules:
            hint = ttk.Label(self.sf.inner, text="「＋ 追加」でルールを作成できます。", foreground="gray")
            hint.pack(anchor="w", padx=6, pady=10)

        # ★行追加/再描画で新しく作られたEntry等にも再bind（macの取りこぼし防止）
        self._install_wheel_bind_recursive_widgets()

    def _create_row(self, idx: int, rule: Rule):
        row = ttk.Frame(self.sf.inner)
        row.pack(fill="x", pady=4)

        v_enabled = tk.BooleanVar(master=self, value=rule.enabled)
        v_src = tk.StringVar(master=self, value=rule.src)
        v_dst = tk.StringVar(master=self, value=rule.dst)

        cb = ttk.Checkbutton(row, variable=v_enabled, command=self.schedule_save)
        cb.pack(side="left", padx=(2, 6))

        e_src = tk.Entry(row, textvariable=v_src, bd=1, relief="solid")
        e_src.pack(side="left", fill="x", expand=True, padx=(6, 6))
        e_dst = tk.Entry(row, textvariable=v_dst, bd=1, relief="solid")
        e_dst.pack(side="left", fill="x", expand=True, padx=(6, 6))

        e_src.bind("<FocusOut>", lambda _e: self.schedule_save())
        e_dst.bind("<FocusOut>", lambda _e: self.schedule_save())
        e_src.bind("<Return>", lambda _e: self.schedule_save())
        e_dst.bind("<Return>", lambda _e: self.schedule_save())

        ops = ttk.Frame(row)
        ops.pack(side="right")

        ttk.Button(ops, text="↑", width=3, command=lambda i=idx: self.move_row(i, -1)).pack(side="left", padx=(0, 4))
        ttk.Button(ops, text="↓", width=3, command=lambda i=idx: self.move_row(i, +1)).pack(side="left", padx=(0, 8))
        ttk.Button(ops, text="削除", command=lambda i=idx: self.delete_row(i)).pack(side="left")

        self.row_widgets.append({
            "enabled": v_enabled,
            "src": v_src,
            "dst": v_dst,
        })

    def add_row(self):
        self.commit_to_model()
        self.rules.append(Rule(enabled=True, src="", dst=""))
        self.render_rows()
        self.schedule_save()
        self.on_message("行を追加しました（自動保存）")

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
            return
        self.rules.pop(idx)
        self.render_rows()
        self.perform_save()
        self.on_message("行を削除しました（自動保存）")

    def move_row(self, idx: int, direction: int):
        self.commit_to_model()
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.rules):
            return
        self.rules[idx], self.rules[new_idx] = self.rules[new_idx], self.rules[idx]
        self.render_rows()
        self.perform_save()
        self.on_message("行の順番を変更しました（自動保存）")

    def commit_to_model(self):
        new_rules: List[Rule] = []
        for rw in self.row_widgets:
            new_rules.append(Rule(
                enabled=bool(rw["enabled"].get()),
                src=rw["src"].get(),
                dst=rw["dst"].get(),
            ))
        self.rules[:] = new_rules

    def schedule_save(self):
        if self._switching:
            return
        if self._save_after_id is not None:
            try:
                self.after_cancel(self._save_after_id)
            except Exception:
                pass
        self._save_after_id = self.after(250, self.perform_save)

    def perform_save(self):
        self._save_after_id = None
        self.commit_to_model()
        try:
            self.on_save_store()
            self.on_message("保存しました")
            if self.on_saved_callback:
                try:
                    self.on_saved_callback()
                except Exception:
                    pass
        except Exception as e:
            self.on_message(f"保存に失敗: {e}")


def resource_path(rel: str) -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base / rel)


# -----------------------------
# Main App
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- Style: macOS だけ ttk で色を出す（Windows は tk.Button を使う） ---
        self._use_ttk_colored_button = (sys.platform == "darwin")

        if self._use_ttk_colored_button:
            self.style = ttk.Style(self)
            try:
                self.style.theme_use("clam")
            except Exception:
                pass

            self.style.configure(
                "Replace.TButton",
                padding=(10, 4),
                foreground="white",
                background="#2E7D32",
                borderwidth=1,
                focusthickness=0,
            )
            self.style.map(
                "Replace.TButton",
                background=[
                    ("active", "#1B5E20"),
                    ("pressed", "#144A19"),
                    ("disabled", "#A7A7A7"),
                ],
                foreground=[
                    ("disabled", "#F2F2F2"),
                ],
            )

        try:
            img = tk.PhotoImage(file=resource_path("appicon.png"))
            self.iconphoto(True, img)
        except Exception:
            pass

        self.title(APP_NAME)

        # settings
        self._last_save_dir: Optional[Path] = None
        self._load_settings()

        # window size
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()

        target_w = int(sw * 0.34)
        target_w = max(740, min(980, target_w))

        target_h = int(sh * 0.88)
        target_h = max(720, min(920, target_h))

        self.geometry(f"{target_w}x{target_h}")
        self.minsize(720, 700)

        self.store = DictStore(DICT_STORE_FILE)
        self.store.load()

        self.tooltip = Tooltip(self)

        names = self.store.names()
        self.edit_dict_name_var = tk.StringVar(value=names[0] if names else "default")

        self.apply_vars: Dict[str, tk.BooleanVar] = {}
        self.apply_button = None
        self._apply_popup: Optional[ApplyPickerPopup] = None

        self.message_var = tk.StringVar(value="")

        self.zoom_values = ["50%", "75%", "100%", "125%", "150%", "175%", "200%"]
        self.zoom_var = tk.StringVar(value="100%")
        self._base_font_size = 11
        self._text_font = tkfont.Font(family="TkDefaultFont", size=self._base_font_size)

        self._syncing = False

        self._in_hl_after_id = None
        self._in_hl_tag = "src_target"
        self._in_hl_color = "#ffd6e7"

        self._last_loaded_text_path: Optional[Path] = None
        self._rule_manager_dialog: Optional[RuleManager] = None

        self._replacing = False
        self._progress_win = None
        self._progress_bar = None

        # ---------- TOP MENUS ----------
        top = ttk.Frame(self, padding=(10, 8, 10, 6))
        top.pack(fill="x")

        row1 = ttk.Frame(top)
        row1.pack(fill="x")
        row2 = ttk.Frame(top)
        row2.pack(fill="x", pady=(6, 0))
        row3 = ttk.Frame(top)
        row3.pack(fill="x", pady=(6, 0))

        ttk.Label(row1, text="使用辞書:").pack(side="left")

        self.apply_button = ttk.Menubutton(row1, text="選択…")
        self.apply_button.pack(side="left", padx=(6, 8))
        self.apply_button.bind("<Button-1>", self.toggle_apply_picker, add="+")

        if self._use_ttk_colored_button:
            self.replace_btn = ttk.Button(
                row1,
                text="置換実行",
                command=self.replace,
                style="Replace.TButton",
            )
        else:
            self.replace_btn = tk.Button(
                row1,
                text="置換実行",
                command=self.replace,
                bg="#2E7D32",
                fg="white",
                activebackground="#1B5E20",
                activeforeground="white",
                relief="solid",
                bd=1,
                padx=10,
                pady=4,
            )
        self.replace_btn.pack(side="left", padx=(0, 10))

        ttk.Button(row1, text="TXT読み込み", command=self.open_text_file).pack(side="left", padx=(0, 8))
        ttk.Button(row1, text="出力コピー", command=self.copy).pack(side="left", padx=(0, 8))
        ttk.Button(row1, text="TXTとして保存", command=self.save_output).pack(side="left")

        ttk.Label(row2, text="編集辞書:").pack(side="left")

        self.edit_combo = ttk.Combobox(
            row2,
            textvariable=self.edit_dict_name_var,
            values=self.store.names(),
            state="readonly",
            width=18
        )
        self.edit_combo.pack(side="left", padx=(6, 10))
        self.edit_combo.bind("<<ComboboxSelected>>", self.on_edit_dict_change)

        ttk.Button(row2, text="辞書新規", command=self.create_dictionary).pack(side="left", padx=(0, 6))
        ttk.Button(row2, text="辞書削除", command=self.delete_dictionary).pack(side="left", padx=(0, 10))
        ttk.Button(row2, text="辞書（ルール）", command=self.open_rules).pack(side="left")

        ttk.Label(row3, text="表示:").pack(side="left")
        zoom = ttk.Combobox(row3, textvariable=self.zoom_var, values=self.zoom_values, state="readonly", width=8)
        zoom.pack(side="left", padx=(6, 10))
        zoom.bind("<<ComboboxSelected>>", self.on_zoom_change)

        msg = ttk.Label(top, textvariable=self.message_var, foreground="gray")
        msg.pack(anchor="w", pady=(6, 0))

        self.bind("<Button-1>", self._on_main_interaction, add="+")
        self.bind("<FocusIn>", self._on_main_interaction, add="+")

        self.build_apply_menu(initial_select_edit=True)

        # ---------- TEXT AREAS ----------
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        in_frame = ttk.Labelframe(paned, text="入力（貼り付け / TXT読み込み）", padding=8)
        out_frame = ttk.Labelframe(paned, text="出力（置換箇所は色付き / ホバーで置換前表示）", padding=8)

        paned.add(in_frame, weight=1)
        paned.add(out_frame, weight=1)

        # --- 入力（gridでスクロールバー列を必ず確保） ---
        in_container = ttk.Frame(in_frame)
        in_container.pack(fill="both", expand=True)
        in_container.columnconfigure(1, weight=1)
        in_container.rowconfigure(0, weight=1)

        self.input = tk.Text(in_container, wrap="char", undo=True, font=self._text_font)
        self.input_ln = LineNumberCanvas(in_container, self.input, width=44, bg="#f3f3f3")
        self.in_vsb = ttk.Scrollbar(in_container, orient="vertical", command=lambda *a: self._scroll_both(*a))
        self.input.configure(yscrollcommand=self._on_input_yscroll)

        self.input_ln.grid(row=0, column=0, sticky="ns", padx=(0, 6))
        self.input.grid(row=0, column=1, sticky="nsew")
        self.in_vsb.grid(row=0, column=2, sticky="ns")

        self.input.tag_configure(self._in_hl_tag, background=self._in_hl_color)
        self.input.bind("<KeyRelease>", lambda _e: self.schedule_input_highlight(), add="+")
        self.input.bind("<ButtonRelease-1>", lambda _e: self.input_ln.schedule_redraw(), add="+")
        self.input.bind("<Configure>", lambda _e: self.input_ln.schedule_redraw(), add="+")

        # --- 出力（gridでスクロールバー列を必ず確保） ---
        out_container = ttk.Frame(out_frame)
        out_container.pack(fill="both", expand=True)
        out_container.columnconfigure(1, weight=1)
        out_container.rowconfigure(0, weight=1)

        self.output = tk.Text(out_container, wrap="char", font=self._text_font)
        self.output_ln = LineNumberCanvas(out_container, self.output, width=44, bg="#f3f3f3")
        self.out_vsb = ttk.Scrollbar(out_container, orient="vertical", command=lambda *a: self._scroll_both(*a))
        self.output.configure(yscrollcommand=self._on_output_yscroll)

        self.output_ln.grid(row=0, column=0, sticky="ns", padx=(0, 6))
        self.output.grid(row=0, column=1, sticky="nsew")
        self.out_vsb.grid(row=0, column=2, sticky="ns")

        self.output.bind("<Motion>", self.on_hover)
        self.output.bind("<Leave>", lambda _e: self.tooltip.hide())
        self.output.bind("<ButtonRelease-1>", lambda _e: self.output_ln.schedule_redraw(), add="+")
        self.output.bind("<Configure>", lambda _e: self.output_ln.schedule_redraw(), add="+")

        self._bind_sync_wheel(self.input)
        self._bind_sync_wheel(self.output)

        self.tag_map: Dict[str, str] = {}
        self._last_hover_tag: Optional[str] = None

        if not self.input.get("1.0", "end-1c").strip():
            self.input.insert("1.0", "ここに貼り付け → 置換実行\n")

        self.set_message("")
        self.apply_zoom("100%")

        self.input_ln.redraw()
        self.output_ln.redraw()
        self.schedule_input_highlight()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # -----------------------------
    # Settings (persist)
    # -----------------------------
    def _load_settings(self):
        try:
            if not SETTINGS_FILE.exists():
                return
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            last_dir = data.get("last_save_dir")
            if last_dir:
                p = Path(str(last_dir))
                if p.exists() and p.is_dir():
                    self._last_save_dir = p
        except Exception:
            pass

    def _save_settings(self):
        try:
            payload = {
                "version": 1,
                "last_save_dir": str(self._last_save_dir) if self._last_save_dir else "",
            }
            SETTINGS_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def on_close(self):
        try:
            if self._rule_manager_dialog is not None and self._rule_manager_dialog.winfo_exists():
                self._rule_manager_dialog.close()
        except Exception:
            pass
        self._save_settings()
        try:
            self.destroy()
        except Exception:
            pass

    def _on_main_interaction(self, _event=None):
        if self._rule_manager_dialog is not None and self._rule_manager_dialog.winfo_exists():
            try:
                self.after(1, self._rule_manager_dialog.bring_to_front_no_focus)
            except Exception:
                pass

    def set_message(self, text: str):
        self.message_var.set(text)

    def on_zoom_change(self, _event=None):
        self.apply_zoom(self.zoom_var.get())
        self.input_ln.schedule_redraw()
        self.output_ln.schedule_redraw()
        self.schedule_input_highlight()

    def apply_zoom(self, percent_text: str):
        try:
            p = int(percent_text.replace("%", "").strip())
        except Exception:
            p = 100
        p = max(50, min(400, p))
        scale = p / 100.0
        size = max(8, int(round(self._base_font_size * scale)))
        self._text_font.configure(size=size)

    # --- progress / lock ---
    def _set_edit_lock(self, locked: bool):
        state = "disabled" if locked else "normal"
        try:
            self.input.config(state=state)
        except Exception:
            pass
        try:
            self.output.config(state=state)
        except Exception:
            pass

    def _show_progress(self, message="置換中..."):
        if self._progress_win is not None and self._progress_win.winfo_exists():
            try:
                self._progress_win.title(message)
            except Exception:
                pass
            return

        win = tk.Toplevel(self)
        win.title(message)
        win.resizable(False, False)
        win.geometry("320x90")

        try:
            win.transient(self)
        except Exception:
            pass

        win.protocol("WM_DELETE_WINDOW", lambda: None)

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text=message).pack(anchor="w")

        bar = ttk.Progressbar(frm, mode="indeterminate")
        bar.pack(fill="x", pady=(10, 0))
        bar.start(12)

        self._progress_win = win
        self._progress_bar = bar

        win.update_idletasks()

        try:
            self.update_idletasks()
            x = self.winfo_rootx() + (self.winfo_width() // 2) - 160
            y = self.winfo_rooty() + (self.winfo_height() // 2) - 45
            win.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _hide_progress(self):
        try:
            if self._progress_bar is not None:
                self._progress_bar.stop()
        except Exception:
            pass
        self._progress_bar = None

        try:
            if self._progress_win is not None and self._progress_win.winfo_exists():
                self._progress_win.destroy()
        except Exception:
            pass
        self._progress_win = None

    # --- sync scroll ---
    def _scroll_both(self, *args):
        if self._syncing:
            return
        self._syncing = True
        try:
            self.input.yview(*args)
            self.output.yview(*args)
            self.input_ln.schedule_redraw()
            self.output_ln.schedule_redraw()
        finally:
            self._syncing = False

    def _on_input_yscroll(self, first, last):
        self.in_vsb.set(first, last)
        if self._syncing:
            self.input_ln.schedule_redraw()
            return
        self._syncing = True
        try:
            self.output.yview_moveto(first)
            self.out_vsb.set(first, last)
            self.input_ln.schedule_redraw()
            self.output_ln.schedule_redraw()
        finally:
            self._syncing = False

    def _on_output_yscroll(self, first, last):
        self.out_vsb.set(first, last)
        if self._syncing:
            self.output_ln.schedule_redraw()
            return
        self._syncing = True
        try:
            self.input.yview_moveto(first)
            self.in_vsb.set(first, last)
            self.input_ln.schedule_redraw()
            self.output_ln.schedule_redraw()
        finally:
            self._syncing = False

    def _bind_sync_wheel(self, widget: tk.Text):
        widget.bind("<MouseWheel>", self._on_wheel, add="+")
        widget.bind("<Button-4>", lambda _e: self._on_wheel_linux(-1), add="+")
        widget.bind("<Button-5>", lambda _e: self._on_wheel_linux(+1), add="+")

    def _on_wheel(self, event):
        delta = getattr(event, "delta", 0)
        if delta == 0:
            return "break"
        units = int(-1 * (delta / 120))
        if units == 0:
            units = -1 if delta > 0 else 1
        self._syncing = True
        try:
            self.input.yview_scroll(units, "units")
            self.output.yview_scroll(units, "units")
            self.input_ln.schedule_redraw()
            self.output_ln.schedule_redraw()
        finally:
            self._syncing = False
        return "break"

    def _on_wheel_linux(self, direction: int):
        self._syncing = True
        try:
            self.input.yview_scroll(direction, "units")
            self.output.yview_scroll(direction, "units")
            self.input_ln.schedule_redraw()
            self.output_ln.schedule_redraw()
        finally:
            self._syncing = False
        return "break"

    # --- dict selection ---
    def current_edit_dict_name(self) -> str:
        return (self.edit_dict_name_var.get().strip() or "default")

    def build_apply_menu(self, initial_select_edit: bool = False):
        self.apply_vars.clear()
        names = self.store.names()
        edit = self.current_edit_dict_name()
        for name in names:
            self.apply_vars[name] = tk.BooleanVar(value=False)
        if initial_select_edit and edit in self.apply_vars:
            self.apply_vars[edit].set(True)
        self.on_apply_selection_change()

    def toggle_apply_picker(self, event=None):
        if self._apply_popup is not None:
            try:
                self._apply_popup.close()
            except Exception:
                pass
            self._apply_popup = None
            return "break"

        names = self.store.names()
        for n in names:
            if n not in self.apply_vars:
                self.apply_vars[n] = tk.BooleanVar(value=False)

        def _closed():
            self._apply_popup = None

        self._apply_popup = ApplyPickerPopup(
            master=self,
            anchor_widget=self.apply_button,
            names=names,
            vars_by_name=self.apply_vars,
            on_change=self.on_apply_selection_change,
            on_closed=_closed,
        )
        return "break"

    def selected_apply_dicts(self) -> List[str]:
        names = self.store.names()
        selected = [n for n in names if self.apply_vars.get(n) and self.apply_vars[n].get()]
        if not selected:
            edit = self.current_edit_dict_name()
            if edit in self.apply_vars:
                self.apply_vars[edit].set(True)
                selected = [edit]
        return selected

    def on_apply_selection_change(self):
        selected = self.selected_apply_dicts()
        if len(selected) == 1:
            self.apply_button.config(text=selected[0])
        else:
            self.apply_button.config(text=f"{len(selected)}個選択")
        self.schedule_input_highlight()

    def on_edit_dict_change(self, _event=None):
        if not any(v.get() for v in self.apply_vars.values()):
            ed = self.current_edit_dict_name()
            if ed in self.apply_vars:
                self.apply_vars[ed].set(True)
        self.on_apply_selection_change()

        if self._rule_manager_dialog is not None and self._rule_manager_dialog.winfo_exists():
            try:
                self._rule_manager_dialog.dict_name_var.set(self.current_edit_dict_name())
                self._rule_manager_dialog.on_dict_change()
            except Exception:
                pass

    def create_dictionary(self):
        name = simpledialog.askstring("辞書を新規作成", "辞書名を入力してください", parent=self)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if not self.store.create(name):
            messagebox.showwarning("作成できません", "その辞書名は既に存在するか、無効な名前です。", parent=self)
            return
        self.store.save()
        self.edit_combo["values"] = self.store.names()
        self.edit_dict_name_var.set(name)
        self.build_apply_menu(initial_select_edit=True)
        self.set_message(f"辞書を作成しました: {name}")
        self.schedule_input_highlight()

    def delete_dictionary(self):
        name = self.current_edit_dict_name()
        if name == "default":
            messagebox.showinfo("削除できません", "default 辞書は削除できません。", parent=self)
            return
        ok = messagebox.askyesno("削除確認", f"辞書「{name}」を削除しますか？\n（中のルールも消えます）", parent=self)
        if not ok:
            return
        if not self.store.delete(name):
            messagebox.showwarning("削除できません", "削除に失敗しました。", parent=self)
            return
        self.store.save()
        self.edit_combo["values"] = self.store.names()
        self.edit_dict_name_var.set("default")
        self.build_apply_menu(initial_select_edit=True)
        self.set_message(f"辞書を削除しました: {name}")
        self.schedule_input_highlight()

    # --- input highlight ---
    def schedule_input_highlight(self):
        if self._in_hl_after_id is not None:
            try:
                self.after_cancel(self._in_hl_after_id)
            except Exception:
                pass
        self._in_hl_after_id = self.after(120, self.refresh_input_highlight)

    def refresh_input_highlight(self):
        self._in_hl_after_id = None
        try:
            self.input.tag_remove(self._in_hl_tag, "1.0", "end")
        except Exception:
            pass

        dict_names = self.selected_apply_dicts()
        src_list: List[str] = []
        for dn in dict_names:
            for r in self.store.get_rules(dn):
                if r.enabled and r.src:
                    src_list.append(r.src)

        seen = set()
        uniq_src = []
        for s in src_list:
            if s in seen:
                continue
            seen.add(s)
            uniq_src.append(s)

        if not uniq_src:
            return

        uniq_src.sort(key=len, reverse=True)

        for needle in uniq_src:
            start = "1.0"
            while True:
                pos = self.input.search(needle, start, stopindex="end")
                if not pos:
                    break
                end = f"{pos}+{len(needle)}c"
                self.input.tag_add(self._in_hl_tag, pos, end)
                start = end

        self.input_ln.schedule_redraw()

    # --- file load ---
    def open_text_file(self):
        path = filedialog.askopenfilename(
            title="テキストファイルを選択",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")]
        )
        if not path:
            return
        p = Path(path)
        try:
            content = p.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            try:
                content = p.read_text(encoding="cp932")
            except Exception as e:
                messagebox.showerror("読み込みエラー", f"文字コードの判定に失敗しました:\n{e}", parent=self)
                return
        except Exception as e:
            messagebox.showerror("読み込みエラー", str(e), parent=self)
            return

        self.input.delete("1.0", "end")
        self.input.insert("1.0", content)

        self._last_loaded_text_path = p

        self.set_message(f"読み込みました: {p.name}")
        self.input_ln.redraw()
        self.output_ln.redraw()
        self.schedule_input_highlight()

    def save_rules(self):
        self.store.save()

    # --- rules dialog (singleton) ---
    def open_rules(self):
        if self._rule_manager_dialog is not None and self._rule_manager_dialog.winfo_exists():
            self._rule_manager_dialog.focus_existing()
            try:
                self._rule_manager_dialog.dict_name_var.set(self.current_edit_dict_name())
                self._rule_manager_dialog.on_dict_change()
            except Exception:
                pass
            return

        def _closed():
            self._rule_manager_dialog = None

        self._rule_manager_dialog = RuleManager(
            self,
            store=self.store,
            initial_dict_name=self.current_edit_dict_name(),
            on_save_store=self.save_rules,
            on_message=self.set_message,
            on_saved_callback=self.schedule_input_highlight,
            on_closed=_closed,
        )

        try:
            self.update_idletasks()
            self._rule_manager_dialog.update_idletasks()

            main_x = self.winfo_rootx()
            main_y = self.winfo_rooty()
            main_w = self.winfo_width()
            main_h = self.winfo_height()
            main_bottom = main_y + main_h

            out_h = self.output.winfo_height()
            req_w = self._rule_manager_dialog.winfo_reqwidth()
            req_h = self._rule_manager_dialog.winfo_reqheight()

            Y_OFFSET = 30
            w = max(520, main_w, req_w)
            h = max(320, out_h, req_h)

            x = main_x
            y = main_bottom - h - Y_OFFSET

            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(0, min(x, sw - w))
            y = max(0, min(y, sh - h))

            self._rule_manager_dialog.geometry(f"{w}x{h}+{x}+{y}")
            self._rule_manager_dialog.update_idletasks()

            real_h = self._rule_manager_dialog.winfo_height()
            y2 = main_bottom - real_h - Y_OFFSET
            y2 = max(0, min(y2, sh - real_h))

            self._rule_manager_dialog.geometry(f"{w}x{real_h}+{x}+{y2}")
        except Exception:
            pass

        self._rule_manager_dialog.focus_existing()
        try:
            self._rule_manager_dialog.bring_to_front_no_focus()
        except Exception:
            pass

    # --- replace with progress + lock ---
    def replace(self):
        if self._replacing:
            return
        self._replacing = True

        try:
            self.replace_btn.config(state="disabled")
        except Exception:
            pass

        self._set_edit_lock(True)
        self._show_progress("置換中...")
        self.after(10, self._do_replace_impl)

    def _do_replace_impl(self):
        try:
            src_text = self.input.get("1.0", "end-1c")

            dict_names = self.selected_apply_dicts()
            enabled_rules: List[Rule] = []
            for dn in dict_names:
                enabled_rules.extend([r for r in self.store.get_rules(dn) if r.enabled and r.src != ""])

            if not enabled_rules:
                messagebox.showwarning("ルールなし", "適用するルールがありません。\n辞書（ルール）で追加・ONにしてください。", parent=self)
                return

            out = src_text
            for r in enabled_rules:
                out = out.replace(r.src, r.dst)

            self.output.config(state="normal")
            self.output.delete("1.0", "end")
            self.output.insert("1.0", out)
            self.output.config(state="disabled")

            self.apply_diff_highlight(src_text, out)
            self.set_message("置換しました")

            self.input_ln.redraw()
            self.output_ln.redraw()
            self.schedule_input_highlight()

        except Exception as e:
            messagebox.showerror("エラー", f"置換中にエラーが発生しました:\n{e}", parent=self)

        finally:
            self._hide_progress()
            self._replacing = False
            try:
                self.replace_btn.config(state="normal")
            except Exception:
                pass
            self._set_edit_lock(False)

    # --- output actions ---
    def copy(self):
        text = self.output.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(text)
        self.set_message("コピーしました")

    def save_output(self):
        text = self.output.get("1.0", "end-1c")

        initialdir = None
        initialfile = None

        if self._last_save_dir and self._last_save_dir.exists():
            initialdir = str(self._last_save_dir)
        elif self._last_loaded_text_path and self._last_loaded_text_path.exists():
            initialdir = str(self._last_loaded_text_path.parent)

        if self._last_loaded_text_path and self._last_loaded_text_path.exists():
            stem = self._last_loaded_text_path.stem
            initialfile = f"{stem}（TXT変換後）.txt"

        path = filedialog.asksaveasfilename(
            title="TXTとして保存",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")],
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if not path:
            return
        p = Path(path)
        p.write_text(text, encoding="utf-8")

        try:
            self._last_save_dir = p.parent
            self._save_settings()
        except Exception:
            pass

        self.set_message(f"保存しました: {p.name}")

    # --- output highlight / hover ---
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
    App().mainloop()
