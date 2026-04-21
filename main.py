import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

import core.data_parser as data_parser
import core.transfer_curve as tc_module
import core.output_curve as oc_module
import utils.excel_exporter as excel_exporter


# ══════════════════════════════════════════════════════════════════
#   스타일 상수 (라이트 테마)
# ══════════════════════════════════════════════════════════════════
BG_DARK       = "#f0f0f0"
BG_PANEL      = "#ffffff"
BG_CARD       = "#f8f8f8"
BG_ACCENT     = "#e4e4e4"
FG_TEXT       = "#1e1e1e"
FG_MUTED      = "#777777"
ACCENT_BLUE   = "#1565c0"
ACCENT_GREEN  = "#2e7d32"
ACCENT_RED    = "#c62828"
FONT_TITLE    = ("Segoe UI", 15, "bold")
FONT_LABEL    = ("Segoe UI", 10)
FONT_SMALL    = ("Segoe UI", 9)
FONT_MONO     = ("Consolas", 9)
BTN_STYLE_PRI = {
    "bg": ACCENT_BLUE, "fg": "#ffffff",
    "activebackground": "#0d47a1", "activeforeground": "#ffffff",
    "relief": "flat", "bd": 0, "padx": 18, "pady": 8,
    "font": ("Segoe UI", 10, "bold"), "cursor": "hand2",
}
BTN_STYLE_SEC = {
    "bg": BG_ACCENT, "fg": FG_TEXT,
    "activebackground": "#d0d0d0", "activeforeground": FG_TEXT,
    "relief": "flat", "bd": 0, "padx": 14, "pady": 7,
    "font": ("Segoe UI", 10), "cursor": "hand2",
}
BTN_STYLE_GRN = {
    "bg": "#388e3c", "fg": "#ffffff",
    "activebackground": "#2e7d32", "activeforeground": "#ffffff",
    "relief": "flat", "bd": 0, "padx": 18, "pady": 8,
    "font": ("Segoe UI", 10, "bold"), "cursor": "hand2",
}

_CLICK_THRESHOLD_PX = 15  # 우클릭 삭제 허용 반경 (픽셀)


# ══════════════════════════════════════════════════════════════════
#   공통 유틸
# ══════════════════════════════════════════════════════════════════

def _labeled_entry(parent, label_text, default="", width=40):
    frm = tk.Frame(parent, bg=BG_CARD)
    tk.Label(frm, text=label_text, bg=BG_CARD, fg=FG_MUTED,
             font=FONT_SMALL, anchor="w").pack(anchor="w", padx=2)
    ent = tk.Entry(frm, font=FONT_MONO, width=width,
                   bg="#ffffff", fg=FG_TEXT, insertbackground=FG_TEXT,
                   relief="solid", bd=1)
    ent.insert(0, default)
    ent.pack(fill="x", padx=2)
    return frm, ent


def _scrollable_text(parent, height=6):
    frm = tk.Frame(parent, bg=BG_CARD)
    sb = tk.Scrollbar(frm, orient="vertical")
    txt = tk.Text(frm, height=height, font=FONT_SMALL,
                  bg="#ffffff", fg="#333333", insertbackground=FG_TEXT,
                  relief="solid", bd=1, yscrollcommand=sb.set)
    sb.config(command=txt.yview)
    sb.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    return frm, txt


# ══════════════════════════════════════════════════════════════════
#   탭 기본 클래스
# ══════════════════════════════════════════════════════════════════

class _CurveTab(tk.Frame):
    MODE_LABEL   = ""
    PARAM_LABEL  = ""
    X_LABEL      = ""
    DEFAULT_LOG  = True

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG_DARK, **kw)
        self._filepath     = None
        self._ref_filepath = None
        self._fig          = None
        self._ax           = None
        self._grouped      = None
        self._ref_grouped  = None
        self._canvas       = None
        self._toolbar      = None
        self._hover_annot  = None
        self._log_var      = tk.BooleanVar(value=self.DEFAULT_LOG)
        self._status_var   = tk.StringVar(value="파일을 선택해 주세요.")
        self._excluded     = {}

        self._xlim_min = tk.StringVar(value="")
        self._xlim_max = tk.StringVar(value="")
        self._ylim_min = tk.StringVar(value="")
        self._ylim_max = tk.StringVar(value="")

        self._build_ui()

    # ── UI 빌드 ─────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, weight=0, minsize=320)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left_outer = tk.Frame(self, bg=BG_PANEL, width=320)
        left_outer.grid(row=0, column=0, sticky="nsew")
        left_outer.grid_propagate(False)
        left_outer.rowconfigure(0, weight=1)
        left_outer.rowconfigure(1, weight=0)
        left_outer.columnconfigure(0, weight=1)

        left_vsb = tk.Scrollbar(left_outer, orient="vertical",
                                bg=BG_ACCENT, troughcolor=BG_PANEL, relief="flat")
        left_vsb.grid(row=0, column=1, sticky="ns")

        self._left_canvas = tk.Canvas(
            left_outer, bg=BG_PANEL, highlightthickness=0,
            yscrollcommand=left_vsb.set
        )
        self._left_canvas.grid(row=0, column=0, sticky="nsew")
        left_vsb.config(command=self._left_canvas.yview)

        left = tk.Frame(self._left_canvas, bg=BG_PANEL)
        self._left_window = self._left_canvas.create_window(
            (0, 0), window=left, anchor="nw"
        )
        left.bind("<Configure>", self._on_left_configure)
        self._left_canvas.bind("<Configure>", self._on_left_canvas_configure)
        self._left_canvas.bind("<MouseWheel>", self._on_mousewheel)
        left.bind("<MouseWheel>", self._on_mousewheel)

        self._build_left(left)

        sb_frame = tk.Frame(left_outer, bg=BG_ACCENT, pady=6)
        sb_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        tk.Label(sb_frame, textvariable=self._status_var, font=FONT_SMALL,
                 bg=BG_ACCENT, fg=FG_MUTED, wraplength=280,
                 justify="left").pack(padx=12, anchor="w")

        right = tk.Frame(self, bg=BG_DARK)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        self._build_right(right)

    def _build_left(self, parent):
        parent.columnconfigure(0, weight=1)

        hdr = tk.Frame(parent, bg=BG_PANEL, pady=12)
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text=self.MODE_LABEL, font=FONT_TITLE,
                 bg=BG_PANEL, fg=ACCENT_BLUE).pack(padx=16)
        tk.Label(hdr, text="SmartSpice 시뮬레이션 결과 분석기",
                 font=FONT_SMALL, bg=BG_PANEL, fg=FG_MUTED).pack(padx=16)

        cards = tk.Frame(parent, bg=BG_PANEL)
        cards.grid(row=1, column=0, sticky="ew", padx=10, pady=8)
        cards.columnconfigure(0, weight=1)

        # 파일 선택 카드
        fc = self._make_card(cards, "📂  입력 파일", row=0)
        self._file_label = tk.Label(fc, text="선택된 주 파일 없음", font=FONT_SMALL,
                                    bg=BG_CARD, fg=FG_MUTED, wraplength=260,
                                    justify="left", anchor="w")
        self._file_label.pack(fill="x", padx=8, pady=(0, 2))
        tk.Button(fc, text="메인 데이터 선택",
                  command=self._select_file, **BTN_STYLE_PRI).pack(fill="x", padx=8, pady=2)

        self._ref_label = tk.Label(fc, text="선택된 비교 데이터 없음", font=FONT_SMALL,
                                    bg=BG_CARD, fg=FG_MUTED, wraplength=260,
                                    justify="left", anchor="w")
        self._ref_label.pack(fill="x", padx=8, pady=(4, 2))
        tk.Button(fc, text="비교 데이터(Ref) 선택", command=self._select_ref_file,
                  **BTN_STYLE_SEC).pack(fill="x", padx=8, pady=(2, 4))

        # 옵션 카드
        oc = self._make_card(cards, "⚙️  옵션 및 축 범위", row=1)

        lf = tk.Frame(oc, bg=BG_CARD)
        lf.pack(fill="x", padx=8, pady=2)
        tk.Checkbutton(lf, text="Y축 로그 스케일 (Log Scale)",
                       variable=self._log_var,
                       bg=BG_CARD, fg=FG_TEXT, selectcolor=BG_ACCENT,
                       activebackground=BG_CARD, activeforeground=FG_TEXT,
                       font=FONT_LABEL, command=self._on_option_change,
                       relief="flat").pack(anchor="w")

        ax_frm = tk.Frame(oc, bg=BG_CARD)
        ax_frm.pack(fill="x", padx=8, pady=(4, 6))

        tk.Label(ax_frm, text="X축 범위 (Min - Max):", bg=BG_CARD, fg=FG_MUTED,
                 font=FONT_SMALL).grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Entry(ax_frm, textvariable=self._xlim_min, width=10, bg="#ffffff",
                 fg=FG_TEXT, relief="solid", bd=1).grid(row=1, column=0, padx=(0, 4), pady=2)
        tk.Label(ax_frm, text="~", bg=BG_CARD, fg=FG_TEXT).grid(row=1, column=1)
        tk.Entry(ax_frm, textvariable=self._xlim_max, width=10, bg="#ffffff",
                 fg=FG_TEXT, relief="solid", bd=1).grid(row=1, column=2, padx=(4, 0), pady=2)

        tk.Label(ax_frm, text="Y축 범위 (Min - Max):", bg=BG_CARD, fg=FG_MUTED,
                 font=FONT_SMALL).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
        tk.Entry(ax_frm, textvariable=self._ylim_min, width=10, bg="#ffffff",
                 fg=FG_TEXT, relief="solid", bd=1).grid(row=3, column=0, padx=(0, 4), pady=2)
        tk.Label(ax_frm, text="~", bg=BG_CARD, fg=FG_TEXT).grid(row=3, column=1)
        tk.Entry(ax_frm, textvariable=self._ylim_max, width=10, bg="#ffffff",
                 fg=FG_TEXT, relief="solid", bd=1).grid(row=3, column=2, padx=(4, 0), pady=2)

        tk.Button(oc, text="🔄 차트 새로 그리기 (적용)", command=self._on_option_change,
                  bg=BG_ACCENT, fg=FG_TEXT, activebackground="#d0d0d0",
                  activeforeground=FG_TEXT, relief="flat", bd=0, pady=5,
                  font=("Segoe UI", 9), cursor="hand2").pack(fill="x", padx=8, pady=(2, 4))

        # 이상치 삭제 카드
        outlier_card = self._make_card(cards, "🗑️  이상치 삭제", row=2)
        tk.Label(outlier_card,
                 text="파라미터 선택 후 목록에서 행을 클릭하여 선택\n(Ctrl+클릭 다중 선택, 차트 우클릭도 가능)",
                 font=FONT_SMALL, bg=BG_CARD, fg=FG_MUTED,
                 justify="left", anchor="w", wraplength=260).pack(fill="x", padx=8, pady=(0, 4))

        sel_frm = tk.Frame(outlier_card, bg=BG_CARD)
        sel_frm.pack(fill="x", padx=8, pady=2)
        sel_frm.columnconfigure(1, weight=1)

        tk.Label(sel_frm, text="파라미터 값:", bg=BG_CARD, fg=FG_MUTED,
                 font=FONT_SMALL).grid(row=0, column=0, sticky="w", padx=(0, 4))
        self._outlier_param_var = tk.StringVar()
        self._outlier_param_combo = ttk.Combobox(sel_frm, textvariable=self._outlier_param_var,
                                                  state="readonly", width=14)
        self._outlier_param_combo.grid(row=0, column=1, sticky="ew", pady=2)
        self._outlier_param_combo.bind("<<ComboboxSelected>>", self._update_outlier_data_preview)

        # 데이터 목록 Treeview
        preview_outer = tk.Frame(outlier_card, bg=BG_CARD)
        preview_outer.pack(fill="x", padx=8, pady=(6, 2))

        tree_frame = tk.Frame(preview_outer, bg=BG_CARD)
        tree_frame.pack(fill="x")

        preview_scroll = tk.Scrollbar(tree_frame, orient="vertical")
        self._preview_tree = ttk.Treeview(
            tree_frame,
            columns=("idx", "x", "y", "status"),
            show="headings",
            height=6,
            selectmode="extended",
            yscrollcommand=preview_scroll.set,
        )
        self._preview_tree.heading("idx", text="Idx")
        self._preview_tree.heading("x", text="X")
        self._preview_tree.heading("y", text="Y")
        self._preview_tree.heading("status", text="상태")
        self._preview_tree.column("idx", width=35, anchor="center", stretch=False)
        self._preview_tree.column("x", width=80, anchor="e", stretch=True)
        self._preview_tree.column("y", width=90, anchor="e", stretch=True)
        self._preview_tree.column("status", width=40, anchor="center", stretch=False)
        preview_scroll.config(command=self._preview_tree.yview)
        preview_scroll.pack(side="right", fill="y")
        self._preview_tree.pack(side="left", fill="x", expand=True)

        btn_frm = tk.Frame(outlier_card, bg=BG_CARD)
        btn_frm.pack(fill="x", padx=8, pady=(4, 2))
        tk.Button(btn_frm, text="✂️ 선택 삭제", command=self._remove_outlier,
                  bg=ACCENT_RED, fg="white", activebackground="#b71c1c",
                  relief="flat", bd=0, padx=10, pady=5,
                  font=("Segoe UI", 9, "bold"), cursor="hand2").pack(side="left", padx=(0, 4))
        tk.Button(btn_frm, text="↩ 전체 복원", command=self._reset_outliers,
                  **BTN_STYLE_SEC).pack(side="left")

        self._outlier_info_lbl = tk.Label(outlier_card, text="제외된 포인트: 없음",
                                          font=FONT_SMALL, bg=BG_CARD, fg=FG_MUTED,
                                          wraplength=260, justify="left", anchor="w")
        self._outlier_info_lbl.pack(fill="x", padx=8, pady=(2, 6))

        # 저장 경로 카드
        sc = self._make_card(cards, "💾  저장 설정", row=3)
        save_frm, self._save_entry = _labeled_entry(sc, "저장 파일명", "result.xlsx", width=28)
        save_frm.pack(fill="x", padx=8, pady=(2, 4))
        tk.Button(sc, text="저장 경로 선택", command=self._pick_save_dir,
                  **BTN_STYLE_SEC).pack(fill="x", padx=8, pady=(2, 4))
        self._save_dir_label = tk.Label(sc, text=os.path.expanduser("~"),
                                         font=FONT_SMALL, bg=BG_CARD,
                                         fg=FG_MUTED, wraplength=260,
                                         justify="left", anchor="w")
        self._save_dir_label.pack(fill="x", padx=8, pady=(0, 4))
        self._save_dir = os.path.expanduser("~")

        # 실행 카드
        rc = self._make_card(cards, "▶  실행", row=4)
        tk.Button(rc, text="📊  분석 실행",
                  command=self._run_analysis, **BTN_STYLE_PRI).pack(fill="x", padx=8, pady=4)
        tk.Button(rc, text="📥  엑셀로 저장",
                  command=self._export_excel, **BTN_STYLE_GRN).pack(fill="x", padx=8, pady=4)

        # 로그 카드
        lc = self._make_card(cards, "📋  로그", row=5)
        log_frm, self._log_text = _scrollable_text(lc, height=4)
        log_frm.pack(fill="both", expand=True, padx=8, pady=4)

    def _build_right(self, parent):
        hbar = tk.Frame(parent, bg=BG_ACCENT, height=40)
        hbar.grid(row=0, column=0, sticky="ew")
        hbar.grid_propagate(False)
        self._chart_title = tk.Label(hbar, text="차트 미리보기",
                                      font=("Segoe UI", 11, "bold"),
                                      bg=BG_ACCENT, fg=FG_TEXT)
        self._chart_title.pack(side="left", padx=16, pady=8)

        self._param_label = tk.Label(hbar, text="", font=FONT_SMALL,
                                      bg=BG_ACCENT, fg=ACCENT_GREEN)
        self._param_label.pack(side="right", padx=16)

        self._chart_frame = tk.Frame(parent, bg=BG_DARK)
        self._chart_frame.grid(row=1, column=0, sticky="nsew")
        self._chart_frame.rowconfigure(0, weight=1)
        self._chart_frame.columnconfigure(0, weight=1)

        tk.Label(self._chart_frame,
                 text="파일을 선택하고\n'분석 실행'을 눌러주세요.",
                 font=("Segoe UI", 14), bg=BG_DARK, fg=FG_MUTED).place(
                 relx=0.5, rely=0.5, anchor="center")

    def _make_card(self, parent, title, row):
        frm = tk.Frame(parent, bg=BG_CARD, bd=0, padx=4, pady=6,
                       relief="solid", highlightthickness=1,
                       highlightbackground=BG_ACCENT)
        frm.grid(row=row, column=0, sticky="ew", pady=5)
        frm.columnconfigure(0, weight=1)
        tk.Label(frm, text=title, font=("Segoe UI", 10, "bold"),
                 bg=BG_CARD, fg=ACCENT_BLUE, anchor="w").pack(fill="x", padx=8, pady=(4, 2))
        tk.Frame(frm, height=1, bg=BG_ACCENT).pack(fill="x", padx=8, pady=(0, 4))
        return frm

    # ── 스크롤 ──────────────────────────────────────────────────

    def _on_left_configure(self, event):
        self._left_canvas.configure(scrollregion=self._left_canvas.bbox("all"))

    def _on_left_canvas_configure(self, event):
        self._left_canvas.itemconfig(self._left_window, width=event.width)

    def _on_mousewheel(self, event):
        if event.delta:
            self._left_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    # ── 이상치 삭제 ──────────────────────────────────────────────

    def _update_outlier_combo(self):
        if self._grouped is None:
            return
        keys = sorted(self._grouped.keys())
        values = [self._format_val(k) for k in keys]
        self._outlier_param_combo["values"] = values
        if values:
            self._outlier_param_combo.set(values[0])
            self._update_outlier_data_preview()

    def _update_outlier_data_preview(self, event=None):
        for item in self._preview_tree.get_children():
            self._preview_tree.delete(item)

        param_key = self._get_selected_param_key()
        if param_key is None or self._grouped is None:
            return

        data = self._grouped[param_key]
        excluded = self._excluded.get(param_key, set())
        for i, (x, y) in enumerate(zip(data["x"], data["y"])):
            status = "제외" if i in excluded else ""
            tag = "excluded" if i in excluded else "active"
            self._preview_tree.insert(
                "", "end", iid=str(i),
                values=(i, f"{x:.4g}", f"{y:.3e}", status),
                tags=(tag,),
            )

    def _get_selected_param_key(self):
        if self._grouped is None:
            return None
        val_str = self._outlier_param_var.get().strip()
        for k in self._grouped:
            if self._format_val(k) == val_str:
                return k
        return None

    def _remove_outlier(self):
        param_key = self._get_selected_param_key()
        if param_key is None:
            messagebox.showwarning("선택 오류", "먼저 분석을 실행하고 파라미터를 선택하세요.")
            return
        selected_iids = self._preview_tree.selection()
        if not selected_iids:
            messagebox.showwarning("선택 없음", "삭제할 데이터를 목록에서 선택해주세요.\n(클릭 또는 Ctrl+클릭으로 다중 선택)")
            return

        already_excluded = self._excluded.get(param_key, set())
        new_indices = {int(iid) for iid in selected_iids} - already_excluded
        if not new_indices:
            messagebox.showinfo("알림", "선택한 항목이 이미 제외되어 있습니다.")
            return

        self._excluded.setdefault(param_key, set()).update(new_indices)
        self._log(f"파라미터 {self._format_val(param_key)}V → 인덱스 {sorted(new_indices)} 제외")
        self._update_outlier_label()
        self._update_outlier_data_preview()
        self._draw_chart()

    def _reset_outliers(self):
        self._excluded.clear()
        self._outlier_info_lbl.config(text="제외된 포인트: 없음")
        self._log("이상치 제외 목록 초기화")
        self._update_outlier_data_preview()
        if self._grouped is not None:
            self._draw_chart()

    def _update_outlier_label(self):
        if not self._excluded:
            self._outlier_info_lbl.config(text="제외된 포인트: 없음")
            return
        parts = [f"{self._format_val(k)}V→{sorted(v)}"
                 for k, v in sorted(self._excluded.items()) if v]
        self._outlier_info_lbl.config(text="제외: " + "  ".join(parts))

    def _get_filtered_grouped(self):
        if not self._excluded or self._grouped is None:
            return self._grouped
        filtered = {}
        for k, data in self._grouped.items():
            if k in self._excluded and self._excluded[k]:
                mask = [i for i in range(len(data["x"])) if i not in self._excluded[k]]
                filtered[k] = {"x": data["x"][mask], "y": data["y"][mask]}
            else:
                filtered[k] = data
        return filtered

    def _filtered_to_original_idx(self, param_key, filtered_idx: int) -> int:
        excluded = self._excluded.get(param_key, set())
        count = 0
        for orig_idx in range(len(self._grouped[param_key]["x"])):
            if orig_idx not in excluded:
                if count == filtered_idx:
                    return orig_idx
                count += 1
        return -1

    # ── 축 범위 파싱 / 자동 입력 ─────────────────────────────────

    def _parse_limits(self):
        def _get_val(s):
            v = s.get().strip()
            if not v:
                return None
            try:
                return float(v)
            except ValueError:
                return None

        x_min, x_max = _get_val(self._xlim_min), _get_val(self._xlim_max)
        y_min, y_max = _get_val(self._ylim_min), _get_val(self._ylim_max)
        xlim = (x_min, x_max) if (x_min is not None and x_max is not None) else None
        ylim = (y_min, y_max) if (y_min is not None and y_max is not None) else None
        return xlim, ylim

    def _auto_fill_axis_limits(self):
        if not self._grouped:
            return
        all_x = np.concatenate([d["x"] for d in self._grouped.values()])
        all_y_raw = np.concatenate([d["y"] for d in self._grouped.values()])

        self._xlim_min.set(f"{all_x.min():.4g}")
        self._xlim_max.set(f"{all_x.max():.4g}")

        if self._log_var.get():
            pos_y = np.abs(all_y_raw)
            pos_y = pos_y[pos_y > 0]
            if len(pos_y):
                self._ylim_min.set(f"{pos_y.min():.3e}")
                self._ylim_max.set(f"{pos_y.max():.3e}")
        else:
            self._ylim_min.set(f"{all_y_raw.min():.4g}")
            self._ylim_max.set(f"{all_y_raw.max():.4g}")

    # ── 파일 선택 / 저장 ─────────────────────────────────────────

    def _select_file(self):
        fp = filedialog.askopenfilename(
            title="메인 데이터 선택",
            filetypes=[("Excel & CSV Data", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if fp:
            self._filepath = fp
            fname = os.path.basename(fp)
            self._file_label.config(text=f"(Main) {fname}", fg=ACCENT_GREEN)
            self._log(f"메인 파일 선택됨: {fname}")
            self._status_var.set("✅ 파일 로드 준비 완료")
            base = os.path.splitext(fname)[0]
            self._save_entry.delete(0, "end")
            self._save_entry.insert(0, f"{base}_result.xlsx")
            self._save_dir = os.path.dirname(fp)
            self._save_dir_label.config(text=self._save_dir)

    def _select_ref_file(self):
        fp = filedialog.askopenfilename(
            title="비교 데이터(Reference) 선택",
            filetypes=[("Excel & CSV Data", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if fp:
            self._ref_filepath = fp
            fname = os.path.basename(fp)
            self._ref_label.config(text=f"(Ref) {fname}", fg=ACCENT_BLUE)
            self._log(f"비교 데이터 파일 선택됨: {fname}")

    def _pick_save_dir(self):
        d = filedialog.askdirectory(title="저장 폴더 선택")
        if d:
            self._save_dir = d
            self._save_dir_label.config(text=d)

    def _on_option_change(self):
        if self._grouped is not None:
            self._draw_chart()

    # ── 분석 실행 ────────────────────────────────────────────────

    def _run_analysis(self):
        if not self._filepath:
            messagebox.showwarning("파일 없음", "메인 데이터 파일을 먼저 선택해주세요.")
            return
        self._excluded.clear()
        self._log("분석 시작...")
        self._status_var.set("⏳ 데이터 파싱 중...")
        threading.Thread(target=self._analysis_thread, daemon=True).start()

    def _analysis_thread(self):
        try:
            self._grouped = self._parse_data(self._filepath)
            self._ref_grouped = self._parse_data(self._ref_filepath) if self._ref_filepath else None

            n_params = len(self._grouped)
            if n_params == 0:
                raise ValueError("유효한 그룹(파라미터) 데이터를 찾지 못했습니다.")

            param_list = ", ".join(f"{self._format_val(v)}V" for v in sorted(self._grouped.keys()))
            self.after(0, lambda: self._log(
                f"파싱 완료: {self.PARAM_LABEL} {n_params}개 감지\n  → {param_list}"
            ))
            if self._ref_grouped:
                self.after(0, lambda: self._log(
                    f"비교 데이터 파싱 완료: {len(self._ref_grouped)}개 파라미터"
                ))
            self.after(0, lambda: self._status_var.set(f"✅ {self.PARAM_LABEL} {n_params}개 감지됨"))
            self.after(0, lambda: self._param_label.config(
                text=f"{self.PARAM_LABEL} 값: {param_list}"
            ))
            self.after(0, self._auto_fill_axis_limits)
            self.after(0, self._update_outlier_combo)
            self.after(0, self._draw_chart)
        except Exception as e:
            self.after(0, lambda cur_e=e: self._log(f"❌ 오류: {cur_e}"))
            self.after(0, lambda: self._status_var.set("❌ 오류 발생"))
            self.after(0, lambda cur_e=e: messagebox.showerror("파싱 오류", str(cur_e)))

    # ── 차트 렌더링 ──────────────────────────────────────────────

    def _draw_chart(self):
        if self._grouped is None:
            return

        for w in self._chart_frame.winfo_children():
            w.destroy()

        log_scale = self._log_var.get()
        xlim, ylim = self._parse_limits()
        display_grouped = self._get_filtered_grouped()

        self._fig = self._create_figure(display_grouped, log_scale, xlim, ylim)
        self._ax = self._fig.axes[0]

        self._hover_annot = self._ax.annotate(
            "", xy=(0, 0), xytext=(12, 12), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=ACCENT_BLUE, alpha=0.92),
            fontsize=8.5, color=FG_TEXT, annotation_clip=False,
            visible=False, zorder=10,
        )

        self._canvas = FigureCanvasTkAgg(self._fig, master=self._chart_frame)
        canvas_widget = self._canvas.get_tk_widget()
        canvas_widget.configure(bg=BG_DARK)
        canvas_widget.pack(fill="both", expand=True)

        toolbar_frame = tk.Frame(self._chart_frame, bg=BG_PANEL)
        toolbar_frame.pack(fill="x", side="bottom")
        self._toolbar = NavigationToolbar2Tk(self._canvas, toolbar_frame)
        self._toolbar.config(background=BG_PANEL)
        self._toolbar.update()

        self._canvas.mpl_connect('motion_notify_event', self._on_chart_hover)
        self._canvas.mpl_connect('button_press_event', self._on_chart_click)
        self._canvas.draw()

        msg = f"차트 생성 완료 (로그: {'ON' if log_scale else 'OFF'})"
        if xlim or ylim:
            msg += " [축 범위 적용]"
        if self._excluded:
            msg += f" [이상치 {sum(len(v) for v in self._excluded.values())}개 제외]"
        self._log(msg)

    def _on_chart_hover(self, event):
        if self._hover_annot is None:
            return
        if event.inaxes and event.xdata is not None and event.ydata is not None:
            x, y = event.xdata, event.ydata
            y_fmt = f"{y:.3e}" if self._log_var.get() else f"{y:.4g}"
            self._hover_annot.set_text(f"x: {x:.4g}\ny: {y_fmt}")
            self._hover_annot.xy = (x, y)
            self._hover_annot.set_visible(True)
        else:
            self._hover_annot.set_visible(False)
        if self._canvas:
            self._canvas.draw_idle()

    def _on_chart_click(self, event):
        if event.button != 3 or not event.inaxes:
            return
        if event.xdata is None or event.ydata is None:
            return
        self._remove_nearest_point(event.xdata, event.ydata)

    def _remove_nearest_point(self, click_x: float, click_y: float):
        if self._grouped is None or self._ax is None:
            return

        log_scale = self._log_var.get()
        try:
            click_disp = self._ax.transData.transform([click_x, click_y])
        except Exception:
            return

        best_dist = float('inf')
        best_param = None
        best_filtered_idx = None

        for param_key, data in self._get_filtered_grouped().items():
            y_arr = np.abs(data["y"]) if log_scale else data["y"]
            for i, (x, y) in enumerate(zip(data["x"], y_arr)):
                if log_scale and y <= 0:
                    continue
                try:
                    pt_disp = self._ax.transData.transform([x, y])
                    dist = np.hypot(pt_disp[0] - click_disp[0], pt_disp[1] - click_disp[1])
                    if dist < best_dist:
                        best_dist = dist
                        best_param = param_key
                        best_filtered_idx = i
                except Exception:
                    continue

        if best_dist > _CLICK_THRESHOLD_PX or best_param is None:
            return

        orig_idx = self._filtered_to_original_idx(best_param, best_filtered_idx)
        if orig_idx < 0:
            return

        self._excluded.setdefault(best_param, set()).add(orig_idx)
        self._log(f"우클릭 제거: 파라미터 {self._format_val(best_param)}V → 인덱스 {orig_idx}")
        self._update_outlier_label()
        self._update_outlier_data_preview()
        self._draw_chart()

    # ── 엑셀 저장 ────────────────────────────────────────────────

    def _export_excel(self):
        if self._grouped is None or self._fig is None:
            messagebox.showwarning("분석 필요", "먼저 '분석 실행'을 눌러주세요.")
            return

        fname = self._save_entry.get().strip()
        if not fname.endswith(".xlsx"):
            fname += ".xlsx"
        save_path = os.path.join(self._save_dir, fname)

        try:
            log_scale = self._log_var.get()
            xlim, ylim = self._parse_limits()
            self._do_export(self._grouped, self._fig, save_path, log_scale, xlim, ylim)
            self._log(f"✅ 저장 완료: {save_path}")
            self._status_var.set("✅ 엑셀 저장 완료")
            if messagebox.askyesno("저장 완료",
                                    f"파일이 저장되었습니다.\n{save_path}\n\n지금 파일을 여시겠습니까?"):
                if sys.platform == "win32":
                    os.startfile(save_path)
                else:
                    os.system(f"open '{save_path}'")
        except Exception as e:
            self._log(f"❌ 저장 오류: {e}")
            messagebox.showerror("저장 오류", str(e))

    # ── 서브클래스에서 구현 ─────────────────────────────────────

    def _parse_data(self, filepath: str) -> dict:
        raise NotImplementedError

    def _create_figure(self, grouped: dict, log_scale: bool, xlim: tuple, ylim: tuple):
        raise NotImplementedError

    def _do_export(self, grouped, fig, save_path, log_scale, xlim, ylim):
        raise NotImplementedError

    # ── 내부 유틸 ───────────────────────────────────────────────

    def _log(self, msg: str):
        self._log_text.config(state="normal")
        self._log_text.insert("end", msg + "\n")
        self._log_text.see("end")
        self._log_text.config(state="disabled")

    @staticmethod
    def _format_val(v: float) -> str:
        if v == int(v):
            return str(int(v))
        return f"{v:.3g}"


# ══════════════════════════════════════════════════════════════════
#   Transfer Curve 탭
# ══════════════════════════════════════════════════════════════════

class TransferCurveTab(_CurveTab):
    MODE_LABEL  = "📈  Transfer Curve"
    PARAM_LABEL = "Vd"
    X_LABEL     = "Vg"
    DEFAULT_LOG = True

    def _parse_data(self, filepath):
        return data_parser.parse_transfer_curve(filepath)

    def _create_figure(self, grouped, log_scale, xlim, ylim):
        plt.close("all")
        return tc_module.create_transfer_figure(
            grouped, log_scale=log_scale,
            title="TFT Transfer Curve  (Vgs - Id)",
            xlim=xlim, ylim=ylim,
            ref_grouped=self._ref_grouped,
        )

    def _do_export(self, grouped, fig, save_path, log_scale, xlim, ylim):
        excel_exporter.export_transfer_curve(grouped, fig, save_path, log_scale, xlim, ylim)


# ══════════════════════════════════════════════════════════════════
#   Output Curve 탭
# ══════════════════════════════════════════════════════════════════

class OutputCurveTab(_CurveTab):
    MODE_LABEL  = "📉  Output Curve"
    PARAM_LABEL = "Vg"
    X_LABEL     = "Vd"
    DEFAULT_LOG = False

    def _parse_data(self, filepath):
        return data_parser.parse_output_curve(filepath)

    def _create_figure(self, grouped, log_scale, xlim, ylim):
        plt.close("all")
        return oc_module.create_output_figure(
            grouped, log_scale=log_scale,
            title="TFT Output Curve  (Vd - Id)",
            xlim=xlim, ylim=ylim,
            ref_grouped=self._ref_grouped,
        )

    def _do_export(self, grouped, fig, save_path, log_scale, xlim, ylim):
        excel_exporter.export_output_curve(grouped, fig, save_path, log_scale, xlim, ylim)


# ══════════════════════════════════════════════════════════════════
#   메인 앱
# ══════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("tr-curve-plotter  ─  SmartSpice 시뮬레이션 결과 분석")
        self.geometry("1400x820")
        self.minsize(960, 640)
        self.configure(bg=BG_DARK)

        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        self._apply_ttk_style()
        self._build_tabs()
        self._build_contact_bar()

    def _apply_ttk_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview",
                        background=BG_CARD,
                        foreground=FG_TEXT,
                        rowheight=22,
                        fieldbackground=BG_CARD,
                        borderwidth=1,
                        font=("Consolas", 8))
        style.configure("Treeview.Heading",
                        background=BG_ACCENT,
                        foreground=FG_TEXT,
                        relief="flat",
                        font=("Segoe UI", 8, "bold"))
        style.map("Treeview",
                  background=[("selected", ACCENT_BLUE)],
                  foreground=[("selected", "#ffffff")])
        style.configure("TCombobox",
                        fieldbackground="#ffffff",
                        background=BG_ACCENT,
                        foreground=FG_TEXT)

    def _build_tabs(self):
        tab_bar = tk.Frame(self, bg=BG_PANEL, height=48)
        tab_bar.pack(fill="x", side="top")
        tab_bar.pack_propagate(False)

        # 하단 구분선
        tk.Frame(self, bg=BG_ACCENT, height=1).pack(fill="x", side="top")

        content = tk.Frame(self, bg=BG_DARK)
        content.pack(fill="both", expand=True)

        self._tc_tab = TransferCurveTab(content)
        self._oc_tab = OutputCurveTab(content)
        self._tc_tab.place(relx=0, rely=0, relwidth=1, relheight=1)
        self._oc_tab.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._btn_tc = tk.Button(
            tab_bar, text="  📈  Transfer Curve  ",
            font=("Segoe UI", 11), relief="flat", bd=0, padx=20,
            cursor="hand2", command=lambda: self._switch_tab(self._tc_tab),
        )
        self._btn_oc = tk.Button(
            tab_bar, text="  📉  Output Curve  ",
            font=("Segoe UI", 11), relief="flat", bd=0, padx=20,
            cursor="hand2", command=lambda: self._switch_tab(self._oc_tab),
        )
        self._btn_tc.pack(side="left", ipady=12)
        self._btn_oc.pack(side="left", ipady=12)

        self._switch_tab(self._tc_tab)

    def _switch_tab(self, tab):
        tab.lift()
        if tab is self._tc_tab:
            self._btn_tc.config(bg=BG_ACCENT, fg=FG_TEXT)
            self._btn_oc.config(bg=BG_PANEL, fg=FG_MUTED)
        else:
            self._btn_tc.config(bg=BG_PANEL, fg=FG_MUTED)
            self._btn_oc.config(bg=BG_ACCENT, fg=FG_TEXT)

    def _build_contact_bar(self):
        bar = tk.Frame(self, bg=BG_ACCENT, pady=3)
        bar.pack(side="bottom", fill="x")
        tk.Label(bar, text="문의사항: dhooonk@lgdisplay.com",
                 font=("Segoe UI", 8), bg=BG_ACCENT, fg=FG_MUTED).pack(side="right", padx=12)


if __name__ == "__main__":
    app = App()
    app.mainloop()
