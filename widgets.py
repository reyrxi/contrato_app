import re
import tkinter as tk

from constants import (
    ACCENT, ACCENT_L, BG, BORDER, BORDER2, CARD,
    DANGER, F10, F10B, F11B, F8, F9, F9B, SUCCESS, TEXT, TEXT2, TEXT3, WHITE,
)
from utils import fmt_cpf, fmt_cep, fmt_fone, validar_cpf


class ScrollableFrame(tk.Frame):
    def __init__(self, parent, **kwargs):
        kwargs.setdefault("bg", BG)
        super().__init__(parent, **kwargs)

        canvas = tk.Canvas(self, bg=BG, bd=0, highlightthickness=0)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)

        self.inner = tk.Frame(canvas, bg=BG)
        self.inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for widget in (canvas, self.inner):
            widget.bind(
                "<MouseWheel>",
                lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
            )


class Section(tk.Frame):
    def __init__(self, parent, title="", icon="", **kwargs):
        super().__init__(
            parent,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1,
            **kwargs,
        )
        if title:
            header = tk.Frame(self, bg=CARD)
            header.pack(fill="x", padx=20, pady=(16, 0))
            tk.Frame(header, bg=ACCENT, width=3, height=18).pack(side="left", padx=(0, 10))
            if icon:
                tk.Label(
                    header, text=icon, font=("Segoe UI", 12), bg=CARD, fg=ACCENT
                ).pack(side="left", padx=(0, 6))
            tk.Label(header, text=title, font=F11B, bg=CARD, fg=TEXT).pack(side="left")
            tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=20, pady=(10, 0))


class LabeledEntry(tk.Frame):
    def __init__(
        self,
        parent,
        label,
        var,
        width=28,
        validator=None,
        formatter=None,
        on_valid=None,
        hint="",
        **kwargs,
    ):
        super().__init__(parent, bg=CARD, **kwargs)
        self.var = var
        self.validator = validator
        self.formatter = formatter
        self.on_valid = on_valid
        self._skip = False

        top = tk.Frame(self, bg=CARD)
        top.pack(fill="x")
        tk.Label(top, text=label, font=F9B, bg=CARD, fg=TEXT2, anchor="w").pack(side="left")
        if hint:
            tk.Label(top, text=hint, font=F8, bg=CARD, fg=TEXT3).pack(side="left", padx=(4, 0))

        self._border = tk.Frame(self, bg=BORDER2, highlightthickness=0)
        self._border.pack(fill="x", pady=(3, 0))
        inner = tk.Frame(self._border, bg=WHITE)
        inner.pack(fill="x", padx=1, pady=1)

        self.entry = tk.Entry(
            inner,
            textvariable=var,
            font=F10,
            bg=WHITE,
            fg=TEXT,
            bd=0,
            relief="flat",
            width=width,
            insertbackground=ACCENT,
        )
        self.entry.pack(side="left", padx=10, pady=7, fill="x", expand=True)

        self._icon = tk.Label(inner, text="", font=F10, bg=WHITE, width=2)
        self._icon.pack(side="right", padx=6)

        self.entry.bind("<FocusIn>", self._on_focus_in)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        var.trace_add("write", self._on_change)

    def _on_focus_in(self, _=None):
        self._border.configure(bg=ACCENT)

    def _on_focus_out(self, _=None):
        valid = self.validator(self.var.get()) if self.validator else True
        self._border.configure(bg=SUCCESS if valid and self.var.get() else BORDER2)

    def _on_change(self, *_):
        if self._skip:
            return
        raw = self.var.get()
        if self.formatter:
            formatted = self.formatter(raw)
            if formatted != raw:
                self._skip = True
                cursor = self.entry.index("insert")
                self.var.set(formatted)
                try:
                    self.entry.icursor(min(cursor, len(formatted)))
                except tk.TclError:
                    pass
                self._skip = False

        if self.validator:
            value = self.var.get()
            valid = self.validator(value)
            if not value:
                self._icon.configure(text="", fg=TEXT3)
                self._border.configure(bg=BORDER2)
            elif valid:
                self._icon.configure(text="✓", fg=SUCCESS)
                self._border.configure(bg=SUCCESS)
                if self.on_valid:
                    self.on_valid(value)
            else:
                self._icon.configure(text="✗", fg=DANGER)
                self._border.configure(bg=BORDER2)


class Btn(tk.Frame):
    def __init__(self, parent, text, cmd, color=ACCENT, light=False, small=False, **kwargs):
        bg = ACCENT_L if light else color
        fg = ACCENT if light else WHITE
        super().__init__(parent, bg=bg, cursor="hand2", **kwargs)
        self._color = bg

        label = tk.Label(
            self,
            text=text,
            font=F9B if small else F10B,
            bg=bg,
            fg=fg,
            padx=14,
            pady=5 if small else 8,
        )
        label.pack()

        for widget in (self, label):
            widget.bind("<Button-1>", lambda e: cmd())
            widget.bind("<Enter>", lambda e, lbl=label: self._on_hover(lbl, True))
            widget.bind("<Leave>", lambda e, lbl=label: self._on_hover(lbl, False))
        self._lbl = label

    def _on_hover(self, label, hovering):
        h = self._color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        if hovering:
            new_color = f"#{max(0, r - 20):02x}{max(0, g - 20):02x}{max(0, b - 20):02x}"
        else:
            new_color = self._color
        self.configure(bg=new_color)
        label.configure(bg=new_color)


def build_fields(parent, rows, sv_dict, app):
    grid = tk.Frame(parent, bg=CARD)
    grid.pack(fill="x", padx=20, pady=(12, 16))
    grid.columnconfigure(0, weight=1, minsize=200)
    grid.columnconfigure(1, weight=1, minsize=200)

    for i, row in enumerate(rows):
        label, key, width, ftype = row[:4]
        hint = row[4] if len(row) > 4 else ""
        col = i % 2
        r = i // 2

        if key not in sv_dict:
            sv_dict[key] = tk.StringVar()
        var = sv_dict[key]

        kwargs = dict(label=label, var=var, width=width, hint=hint)

        if ftype == "cpf":
            kwargs["validator"] = validar_cpf
            kwargs["formatter"] = lambda v: fmt_cpf(v) if len(re.sub(r"\D", "", v)) == 11 else v
        elif ftype == "cep":
            kwargs["validator"] = lambda v: len(re.sub(r"\D", "", v)) == 8
            kwargs["formatter"] = lambda v: fmt_cep(v) if len(re.sub(r"\D", "", v)) == 8 else v
            kwargs["on_valid"] = lambda v, a=app: a._lookup_cep(v, "cont")
        elif ftype == "cep_resp":
            kwargs["validator"] = lambda v: len(re.sub(r"\D", "", v)) == 8
            kwargs["formatter"] = lambda v: fmt_cep(v) if len(re.sub(r"\D", "", v)) == 8 else v
            kwargs["on_valid"] = lambda v, a=app: a._lookup_cep(v, "resp")
        elif ftype == "fone":
            kwargs["formatter"] = (
                lambda v: fmt_fone(v) if len(re.sub(r"\D", "", v)) in (10, 11) else v
            )
            kwargs["validator"] = lambda v: len(re.sub(r"\D", "", v)) in (10, 11)

        entry = LabeledEntry(grid, **kwargs)
        entry.grid(
            row=r,
            column=col,
            sticky="ew",
            padx=(0, 12) if col == 0 else 0,
            pady=6,
        )

    return grid
