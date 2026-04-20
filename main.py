"""
main.py
-------
TFT Curve Analyzer - 메인 GUI 어플리케이션 (tkinter)
Transfer Curve (Vgs-Id) 및 Output Curve (Vd-Id) 분석을 지원.
"""

import os
import sys
import textwrap
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
#   스타일 상수
# ══════════════════════════════════════════════════════════════════
BG_DARK       = "#12121e"
BG_PANEL      = "#1e1e2e"
BG_CARD       = "#252535"
BG_ACCENT     = "#2e2e45"
FG_TEXT       = "#e0e0f0"
FG_MUTED      = "#9090b0"
ACCENT_BLUE   = "#4a9eff"
ACCENT_GREEN  = "#3de0a0"
ACCENT_RED    = "#ff5c6a"
FONT_TITLE    = ("Segoe UI", 15, "bold")
FONT_LABEL    = ("Segoe UI", 10)
FONT_SMALL    = ("Segoe UI", 9)
FONT_MONO     = ("Consolas", 9)
BTN_STYLE_PRI = {
    "bg": ACCENT_BLUE, "fg": "#ffffff",
    "activebackground": "#3080d0", "activeforeground": "#ffffff",
    "relief": "flat", "bd": 0, "padx": 18, "pady": 8,
    "font": ("Segoe UI", 10, "bold"), "cursor": "hand2",
}
BTN_STYLE_SEC = {
    "bg": BG_ACCENT, "fg": FG_TEXT,
    "activebackground": "#3a3a55", "activeforeground": FG_TEXT,
    "relief": "flat", "bd": 0, "padx": 14, "pady": 7,
    "font": ("Segoe UI", 10), "cursor": "hand2",
}
BTN_STYLE_GRN = {
    "bg": "#1a7a5a", "fg": "#ffffff",
    "activebackground": "#13604a", "activeforeground": "#ffffff",
    "relief": "flat", "bd": 0, "padx": 18, "pady": 8,
    "font": ("Segoe UI", 10, "bold"), "cursor": "hand2",
}


# ══════════════════════════════════════════════════════════════════
#   공통 유틸
# ══════════════════════════════════════════════════════════════════

def _labeled_entry(parent, label_text, default="", width=40):
    """레이블 + 입력창 쌍을 생성하여 (frame, entry) 반환. 배치는 호출자가 담당."""
    frm = tk.Frame(parent, bg=BG_CARD)
    tk.Label(frm, text=label_text, bg=BG_CARD, fg=FG_MUTED,
             font=FONT_SMALL, anchor="w").pack(anchor="w", padx=2)
    ent = tk.Entry(frm, font=FONT_MONO, width=width,
                   bg=BG_ACCENT, fg=FG_TEXT, insertbackground=FG_TEXT,
                   relief="flat", bd=4)
    ent.insert(0, default)
    ent.pack(fill="x", padx=2)
    return frm, ent


def _status_badge(parent, text="READY", color=ACCENT_GREEN):
    lbl = tk.Label(parent, text=f"  {text}  ", font=("Segoe UI", 9, "bold"),
                   bg=color, fg="#ffffff", relief="flat", bd=0)
    return lbl


def _scrollable_text(parent, height=6):
    frm = tk.Frame(parent, bg=BG_CARD)
    sb = tk.Scrollbar(frm, orient="vertical")
    txt = tk.Text(frm, height=height, font=FONT_SMALL,
                  bg="#0d0d1a", fg="#a0d0ff", insertbackground=FG_TEXT,
                  relief="flat", bd=4, yscrollcommand=sb.set)
    sb.config(command=txt.yview)
    sb.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    return frm, txt


# ══════════════════════════════════════════════════════════════════
#   탭 기본 클래스
# ══════════════════════════════════════════════════════════════════

class _CurveTab(tk.Frame):
    """
    Transfer Curve 및 Output Curve 탭에서 공통으로 사용되는 UI 레이아웃 컨테이너입니다.
    좌측에는 데이터 선택, 옵션 제어, 저장 경로 등을 입력하는 컨트롤 패널을 가지며,
    우측에는 Matplotlib 결과물이 렌더링되는 차트 캔버스를 가집니다.
    공통된 기능을 상속(Inheritance)받기 위한 기반 클래스(Base class)입니다.
    """

    MODE_LABEL   = ""   # 탭 제목 라벨 (서브클래스에서 오버라이드)
    PARAM_LABEL  = ""   # "Vd" or "Vg"
    X_LABEL      = ""   # "Vg" or "Vd"
    DEFAULT_LOG  = True

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG_DARK, **kw)
        self._filepath  = None
        self._ref_filepath = None
        self._fig       = None
        self._grouped   = None
        self._ref_grouped = None
        self._canvas    = None
        self._toolbar   = None
        self._log_var   = tk.BooleanVar(value=self.DEFAULT_LOG)
        self._status_var = tk.StringVar(value="파일을 선택해 주세요.")
        # 이상치 삭제를 위한 마스크 저장 {param_key: set of excluded indices}
        self._excluded = {}

        # 축 범위 변수
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

        # 좌측 외부 컨테이너 (고정 너비)
        left_outer = tk.Frame(self, bg=BG_PANEL, width=320)
        left_outer.grid(row=0, column=0, sticky="nsew")
        left_outer.grid_propagate(False)
        left_outer.rowconfigure(0, weight=1)   # 스크롤 영역이 확장
        left_outer.rowconfigure(1, weight=0)   # 상태 바는 고정
        left_outer.columnconfigure(0, weight=1)

        # 스크롤바
        left_vsb = tk.Scrollbar(left_outer, orient="vertical",
                                bg=BG_ACCENT, troughcolor=BG_PANEL, relief="flat")
        left_vsb.grid(row=0, column=1, sticky="ns")

        # 스크롤 가능한 Canvas
        self._left_canvas = tk.Canvas(
            left_outer, bg=BG_PANEL, highlightthickness=0,
            yscrollcommand=left_vsb.set
        )
        self._left_canvas.grid(row=0, column=0, sticky="nsew")
        left_vsb.config(command=self._left_canvas.yview)

        # Canvas 내부에 실제 콘텐츠 프레임
        left = tk.Frame(self._left_canvas, bg=BG_PANEL)
        self._left_window = self._left_canvas.create_window(
            (0, 0), window=left, anchor="nw"
        )
        left.bind("<Configure>", self._on_left_configure)
        self._left_canvas.bind("<Configure>", self._on_left_canvas_configure)
        # 마우스 휠은 left_canvas / left 프레임 위에서만 동작
        self._left_canvas.bind("<MouseWheel>", self._on_mousewheel)
        left.bind("<MouseWheel>", self._on_mousewheel)

        self._build_left(left)

        # 상태 바 – Canvas 스크롤 밖, left_outer row=1에 고정 배치
        sb_frame = tk.Frame(left_outer, bg="#0d0d20", pady=6)
        sb_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        tk.Label(sb_frame, textvariable=self._status_var, font=FONT_SMALL,
                 bg="#0d0d20", fg=FG_MUTED, wraplength=280,
                 justify="left").pack(padx=12, anchor="w")

        # 우측 차트 패널
        right = tk.Frame(self, bg=BG_DARK)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        self._build_right(right)

    def _build_left(self, parent):
        """Canvas 내부 스크롤 콘텐츠 프레임을 구성합니다.
        상태 바는 _build_ui에서 Canvas 외부(left_outer)에 별도 배치됩니다."""
        parent.columnconfigure(0, weight=1)

        # ── 제목 ──
        hdr = tk.Frame(parent, bg="#0d0d20", pady=12)
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text=self.MODE_LABEL, font=FONT_TITLE,
                 bg="#0d0d20", fg=ACCENT_BLUE).pack(padx=16)
        tk.Label(hdr, text="SmartSpice 시뮬레이션 결과 분석기",
                 font=FONT_SMALL, bg="#0d0d20", fg=FG_MUTED).pack(padx=16)

        # ── 카드 컨테이너 ──
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
                  command=self._select_file, **BTN_STYLE_PRI).pack(fill="x",
                  padx=8, pady=2)

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

        # 축 범위 그리드 프레임
        ax_frm = tk.Frame(oc, bg=BG_CARD)
        ax_frm.pack(fill="x", padx=8, pady=(4, 6))

        tk.Label(ax_frm, text="X축 범위 (Min - Max):", bg=BG_CARD, fg=FG_MUTED, font=FONT_SMALL).grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Entry(ax_frm, textvariable=self._xlim_min, width=10, bg=BG_ACCENT, fg=FG_TEXT, relief="flat", bd=3).grid(row=1, column=0, padx=(0,4), pady=2)
        tk.Label(ax_frm, text="~", bg=BG_CARD, fg=FG_TEXT).grid(row=1, column=1)
        tk.Entry(ax_frm, textvariable=self._xlim_max, width=10, bg=BG_ACCENT, fg=FG_TEXT, relief="flat", bd=3).grid(row=1, column=2, padx=(4,0), pady=2)

        tk.Label(ax_frm, text="Y축 범위 (Min - Max):", bg=BG_CARD, fg=FG_MUTED, font=FONT_SMALL).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6,0))
        tk.Entry(ax_frm, textvariable=self._ylim_min, width=10, bg=BG_ACCENT, fg=FG_TEXT, relief="flat", bd=3).grid(row=3, column=0, padx=(0,4), pady=2)
        tk.Label(ax_frm, text="~", bg=BG_CARD, fg=FG_TEXT).grid(row=3, column=1)
        tk.Entry(ax_frm, textvariable=self._ylim_max, width=10, bg=BG_ACCENT, fg=FG_TEXT, relief="flat", bd=3).grid(row=3, column=2, padx=(4,0), pady=2)

        tk.Button(oc, text="🔄 차트 새로 그리기 (적용)", command=self._on_option_change,
                  bg="#444466", fg="white", activebackground="#555577", relief="flat", bd=0, pady=4).pack(fill="x", padx=8, pady=(2,4))

        # 이상치 삭제 카드
        outlier_card = self._make_card(cards, "🗑️  이상치 삭제", row=2)
        info_lbl = tk.Label(outlier_card,
                            text="파라미터와 인덱스(0-base)를 입력하여\n해당 데이터 포인트를 제거합니다.",
                            font=FONT_SMALL, bg=BG_CARD, fg=FG_MUTED,
                            justify="left", anchor="w", wraplength=260)
        info_lbl.pack(fill="x", padx=8, pady=(0, 4))

        sel_frm = tk.Frame(outlier_card, bg=BG_CARD)
        sel_frm.pack(fill="x", padx=8, pady=2)
        sel_frm.columnconfigure(1, weight=1)

        tk.Label(sel_frm, text="파라미터 값:", bg=BG_CARD, fg=FG_MUTED, font=FONT_SMALL).grid(row=0, column=0, sticky="w", padx=(0,4))
        self._outlier_param_var = tk.StringVar()
        self._outlier_param_combo = ttk.Combobox(sel_frm, textvariable=self._outlier_param_var,
                                                  state="readonly", width=14)
        self._outlier_param_combo.grid(row=0, column=1, sticky="ew", pady=2)

        tk.Label(sel_frm, text="데이터 인덱스:", bg=BG_CARD, fg=FG_MUTED, font=FONT_SMALL).grid(row=1, column=0, sticky="w", padx=(0,4))
        self._outlier_idx_entry = tk.Entry(sel_frm, width=14, bg=BG_ACCENT, fg=FG_TEXT,
                                           insertbackground=FG_TEXT, relief="flat", bd=3,
                                           font=FONT_MONO)
        self._outlier_idx_entry.grid(row=1, column=1, sticky="ew", pady=2)
        tk.Label(sel_frm, text="  (쉼표 구분 다중 입력 가능)",
                 bg=BG_CARD, fg=FG_MUTED, font=("Segoe UI", 8)).grid(row=2, column=0, columnspan=2, sticky="w")

        btn_frm = tk.Frame(outlier_card, bg=BG_CARD)
        btn_frm.pack(fill="x", padx=8, pady=(4, 2))
        tk.Button(btn_frm, text="✂️ 선택 삭제", command=self._remove_outlier,
                  bg=ACCENT_RED, fg="white", activebackground="#cc3a47",
                  relief="flat", bd=0, padx=10, pady=5,
                  font=("Segoe UI", 9, "bold"), cursor="hand2").pack(side="left", padx=(0,4))
        tk.Button(btn_frm, text="↩ 전체 복원", command=self._reset_outliers,
                  **BTN_STYLE_SEC).pack(side="left")

        # 현재 제외된 포인트 표시 레이블
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
                  command=self._run_analysis, **BTN_STYLE_PRI).pack(
                  fill="x", padx=8, pady=4)
        tk.Button(rc, text="📥  엑셀로 저장",
                  command=self._export_excel, **BTN_STYLE_GRN).pack(
                  fill="x", padx=8, pady=4)

        # 로그 카드
        lc = self._make_card(cards, "📋  로그", row=5)
        log_frm, self._log_text = _scrollable_text(lc, height=4)
        log_frm.pack(fill="both", expand=True, padx=8, pady=4)

        # ▶ 상태 바는 _build_ui 내 left_outer에 고정 배치됨 (여기서는 생략)

    def _build_right(self, parent):
        # 차트 제목 바
        hbar = tk.Frame(parent, bg=BG_ACCENT, height=40)
        hbar.grid(row=0, column=0, sticky="ew")
        hbar.grid_propagate(False)
        self._chart_title = tk.Label(hbar, text="차트 미리보기",
                                      font=("Segoe UI", 11, "bold"),
                                      bg=BG_ACCENT, fg=FG_TEXT)
        self._chart_title.pack(side="left", padx=16, pady=8)

        # 파라미터 정보 레이블
        self._param_label = tk.Label(hbar, text="",
                                      font=FONT_SMALL,
                                      bg=BG_ACCENT, fg=ACCENT_GREEN)
        self._param_label.pack(side="right", padx=16)

        # 차트 영역
        self._chart_frame = tk.Frame(parent, bg=BG_DARK)
        self._chart_frame.grid(row=1, column=0, sticky="nsew")
        self._chart_frame.rowconfigure(0, weight=1)
        self._chart_frame.columnconfigure(0, weight=1)

        # 초기 플레이스홀더
        tk.Label(self._chart_frame,
                 text="파일을 선택하고\n'분석 실행'을 눌러주세요.",
                 font=("Segoe UI", 14), bg=BG_DARK, fg=FG_MUTED).place(
                 relx=0.5, rely=0.5, anchor="center")

    def _make_card(self, parent, title, row):
        frm = tk.Frame(parent, bg=BG_CARD, bd=0, padx=4, pady=6)
        frm.grid(row=row, column=0, sticky="ew", pady=5)
        frm.columnconfigure(0, weight=1)
        tk.Label(frm, text=title, font=("Segoe UI", 10, "bold"),
                 bg=BG_CARD, fg=ACCENT_BLUE, anchor="w").pack(
                 fill="x", padx=8, pady=(4, 2))
        sep = tk.Frame(frm, height=1, bg=BG_ACCENT)
        sep.pack(fill="x", padx=8, pady=(0, 4))
        return frm

    # ── 스크롤 이벤트 핸들러 ─────────────────────────────────────

    def _on_left_configure(self, event):
        """내부 콘텐츠 크기가 변할 때 Canvas의 scrollregion 갱신."""
        self._left_canvas.configure(
            scrollregion=self._left_canvas.bbox("all")
        )

    def _on_left_canvas_configure(self, event):
        """Canvas 크기가 변할 때 내부 윈도우 너비 동기화."""
        self._left_canvas.itemconfig(self._left_window, width=event.width)

    def _on_mousewheel(self, event):
        """마우스 휠 이벤트로 좌측 패널 스크롤 (macOS / Windows 공용)."""
        # macOS: event.delta는 양수면 위로, 음수면 아래로 스크롤
        # Windows/Linux: 120 단위로 들어오므로 나누기 처리
        if event.delta:
            amount = -1 if event.delta > 0 else 1
            self._left_canvas.yview_scroll(amount, "units")

    # ── 이상치 삭제 이벤트 핸들러 ────────────────────────────────

    def _update_outlier_combo(self):
        """그룹 데이터 파싱 후 Combobox 옵션을 갱신."""
        if self._grouped is None:
            return
        keys = sorted(self._grouped.keys())
        values = [self._format_val(k) for k in keys]
        self._outlier_param_combo["values"] = values
        if values:
            self._outlier_param_combo.set(values[0])

    def _get_selected_param_key(self):
        """Combobox에서 선택된 파라미터 값(float)을 반환."""
        if self._grouped is None:
            return None
        val_str = self._outlier_param_var.get().strip()
        for k in self._grouped:
            if self._format_val(k) == val_str:
                return k
        return None

    def _remove_outlier(self):
        """선택된 파라미터의 지정 인덱스 데이터를 제외 집합에 추가하고 차트 재렌더링."""
        param_key = self._get_selected_param_key()
        if param_key is None:
            messagebox.showwarning("선택 오류", "먼저 분석을 실행하고 파라미터를 선택하세요.")
            return
        idx_str = self._outlier_idx_entry.get().strip()
        if not idx_str:
            messagebox.showwarning("인덱스 없음", "삭제할 데이터 인덱스를 입력하세요.\n예: 0, 3, 5")
            return
        try:
            indices = {int(s.strip()) for s in idx_str.split(",") if s.strip()}
        except ValueError:
            messagebox.showerror("입력 오류", "인덱스는 정수여야 합니다. 예: 0, 3, 5")
            return

        data_len = len(self._grouped[param_key]["x"])
        invalid = [i for i in indices if i < 0 or i >= data_len]
        if invalid:
            messagebox.showerror("범위 오류",
                f"유효 범위(0~{data_len-1}) 벗어난 인덱스: {invalid}")
            return

        if param_key not in self._excluded:
            self._excluded[param_key] = set()
        self._excluded[param_key].update(indices)
        self._log(f"파라미터 {self._format_val(param_key)}V → 인덱스 {sorted(indices)} 제외")
        self._update_outlier_label()
        self._draw_chart()

    def _reset_outliers(self):
        """모든 이상치 제외 정보를 초기화하고 차트 재렌더링."""
        self._excluded.clear()
        self._outlier_info_lbl.config(text="제외된 포인트: 없음")
        self._log("이상치 제외 목록 초기화")
        if self._grouped is not None:
            self._draw_chart()

    def _update_outlier_label(self):
        """제외된 데이터 포인트 요약 텍스트 갱신."""
        if not self._excluded:
            self._outlier_info_lbl.config(text="제외된 포인트: 없음")
            return
        parts = []
        for k, idxs in sorted(self._excluded.items()):
            if idxs:
                parts.append(f"{self._format_val(k)}V→{sorted(idxs)}")
        self._outlier_info_lbl.config(text="제외: " + "  ".join(parts))

    def _get_filtered_grouped(self):
        """_excluded 마스크를 적용한 필터링된 grouped 딕셔너리 반환."""
        if not self._excluded or self._grouped is None:
            return self._grouped
        filtered = {}
        for k, data in self._grouped.items():
            if k in self._excluded and self._excluded[k]:
                mask = [i for i in range(len(data["x"])) if i not in self._excluded[k]]
                filtered[k] = {
                    "x": data["x"][mask],
                    "y": data["y"][mask],
                }
            else:
                filtered[k] = data
        return filtered

    # ── 이벤트 핸들러 ────────────────────────────────────────────

    def _parse_limits(self):
        """UI에 입력된 문자열을 파싱해서 xlim, ylim 튜플 반환."""
        def _get_val(s):
            v = s.get().strip()
            if not v: return None
            try: return float(v)
            except ValueError: return None
            
        x_min, x_max = _get_val(self._xlim_min), _get_val(self._xlim_max)
        y_min, y_max = _get_val(self._ylim_min), _get_val(self._ylim_max)
        
        xlim = (x_min, x_max) if (x_min is not None and x_max is not None) else None
        ylim = (y_min, y_max) if (y_min is not None and y_max is not None) else None
        
        return xlim, ylim

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
            self._status_var.set(f"✅ 파일 로드 준비 완료")
            # 파일명 기반으로 저장 파일명 자동 설정
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

    def _run_analysis(self):
        """
        '분석 실행' 버튼 클릭 이벤트 핸들러입니다.
        데이터 파싱 및 차트 렌더링 시 응용 프로그램 UI가 멈추는(Freezing) 현상을
        효과적으로 방지하기 위해 분석 태스크를 별도 스레드(Thread)로 넘깁니다.
        """
        if not self._filepath:
            messagebox.showwarning("파일 없음", "메인 데이터 파일을 먼저 선택해주세요.")
            return
        # 새 분석 시 이상치 제외 초기화
        self._excluded.clear()
        self._log("분석 시작...")
        self._status_var.set("⏳ 데이터 파싱 중...")
        # 별도 데몬 스레드에서 백그라운드 작업 수행
        threading.Thread(target=self._analysis_thread, daemon=True).start()

    def _analysis_thread(self):
        """
        분석 스레드의 핵심 로직입니다.
        데이터 파서(Parser)를 호출하여 엑셀/CSV 모델 데이터를 읽어들이고,
        비교 대상 파일(Reference)이 존재할 경우 동시에 로드한 뒤,
        Tkinter의 UI 메인 루프에 안전하게 통신(self.after 활용)하여 차트를 생성합니다.
        """
        try:
            # 1. 메인 파일 파싱
            self._grouped = self._parse_data(self._filepath)

            # 2. 비교 대상(Reference) 파일 파싱
            if self._ref_filepath:
                self._ref_grouped = self._parse_data(self._ref_filepath)
            else:
                self._ref_grouped = None

            n_params = len(self._grouped)
            if n_params == 0:
                raise ValueError("유효한 그룹(파라미터) 데이터를 찾지 못했습니다. 데이터 포맷이나 스케일을 확인하세요.")

            param_list = ", ".join(
                f"{self._format_val(v)}V" for v in sorted(self._grouped.keys())
            )
            self.after(0, lambda: self._log(
                f"파싱 완료: {self.PARAM_LABEL} {n_params}개 감지\n  → {param_list}"
            ))

            if self._ref_grouped:
                self.after(0, lambda: self._log(f"비교 데이터 파싱 완료: {len(self._ref_grouped)}개 파라미터 감지"))

            self.after(0, lambda: self._status_var.set(
                f"✅ {self.PARAM_LABEL} {n_params}개 감지됨"
            ))
            self.after(0, lambda: self._param_label.config(
                text=f"{self.PARAM_LABEL} 값: {param_list}"
            ))
            # 이상치 Combobox 갱신 및 차트 그리기
            self.after(0, self._update_outlier_combo)
            self.after(0, self._draw_chart)
        except Exception as e:
            self.after(0, lambda cur_e=e: self._log(f"❌ 오류: {cur_e}"))
            self.after(0, lambda: self._status_var.set(f"❌ 오류 발생"))
            self.after(0, lambda cur_e=e: messagebox.showerror("파싱 오류", str(cur_e)))

    def _draw_chart(self):
        if self._grouped is None:
            return

        # 기존 차트 제거
        for w in self._chart_frame.winfo_children():
            w.destroy()

        log_scale = self._log_var.get()
        xlim, ylim = self._parse_limits()

        # 이상치 필터 적용된 데이터로 차트 생성
        display_grouped = self._get_filtered_grouped()

        self._fig = self._create_figure(display_grouped, log_scale, xlim, ylim)

        self._canvas = FigureCanvasTkAgg(self._fig, master=self._chart_frame)
        canvas_widget = self._canvas.get_tk_widget()
        canvas_widget.configure(bg=BG_DARK)
        canvas_widget.pack(fill="both", expand=True)

        toolbar_frame = tk.Frame(self._chart_frame, bg="#1a1a2e")
        toolbar_frame.pack(fill="x", side="bottom")
        self._toolbar = NavigationToolbar2Tk(self._canvas, toolbar_frame)
        self._toolbar.config(background="#1a1a2e")
        self._toolbar.update()
        self._canvas.draw()

        msg = f"차트 생성 완료 (로그: {'ON' if log_scale else 'OFF'})"
        if xlim or ylim:
            msg += f" [축 범위 적용]"
        if self._excluded:
            excluded_cnt = sum(len(v) for v in self._excluded.values())
            msg += f" [이상치 {excluded_cnt}개 제외]"
        self._log(msg)

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
                os.startfile(save_path) if sys.platform == "win32" else os.system(f"open '{save_path}'")
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
    MODE_LABEL  = "🔵  Transfer Curve"
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
            ref_grouped=self._ref_grouped
        )

    def _do_export(self, grouped, fig, save_path, log_scale, xlim, ylim):
        excel_exporter.export_transfer_curve(grouped, fig, save_path, log_scale, xlim, ylim)


# ══════════════════════════════════════════════════════════════════
#   Output Curve 탭
# ══════════════════════════════════════════════════════════════════

class OutputCurveTab(_CurveTab):
    MODE_LABEL  = "🟢  Output Curve"
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
            ref_grouped=self._ref_grouped
        )

    def _do_export(self, grouped, fig, save_path, log_scale, xlim, ylim):
        excel_exporter.export_output_curve(grouped, fig, save_path, log_scale, xlim, ylim)


# ══════════════════════════════════════════════════════════════════
#   메인 앱
# ══════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TFT Curve Analyzer  ─  SmartSpice 시뮬레이션 결과 분석")
        self.geometry("1400x800")
        self.minsize(960, 620)
        self.configure(bg=BG_DARK)

        # 아이콘 없으면 무시
        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        self._build_menu()
        self._build_tabs()

    def _build_menu(self):
        menubar = tk.Menu(self, bg=BG_PANEL, fg=FG_TEXT,
                          activebackground=ACCENT_BLUE, activeforeground="white",
                          relief="flat", bd=0)
        file_menu = tk.Menu(menubar, tearoff=False,
                            bg=BG_PANEL, fg=FG_TEXT,
                            activebackground=ACCENT_BLUE, activeforeground="white")
        file_menu.add_command(label="종료", command=self.destroy)
        menubar.add_cascade(label="파일", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=False,
                            bg=BG_PANEL, fg=FG_TEXT,
                            activebackground=ACCENT_BLUE, activeforeground="white")
        help_menu.add_command(label="사용 방법", command=self._show_help)
        help_menu.add_command(label="정보", command=self._show_about)
        menubar.add_cascade(label="도움말", menu=help_menu)
        self.config(menu=menubar)

    def _build_tabs(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook",
                         background=BG_DARK, borderwidth=0, padding=0)
        style.configure("TNotebook.Tab",
                         background=BG_PANEL,
                         foreground=FG_MUTED,
                         font=("Segoe UI", 11),
                         padding=[20, 10],
                         borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", BG_ACCENT)],
                  foreground=[("selected", FG_TEXT)])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self._tc_tab = TransferCurveTab(nb)
        self._oc_tab = OutputCurveTab(nb)

        nb.add(self._tc_tab, text="  📈  Transfer Curve  ")
        nb.add(self._oc_tab, text="  📉  Output Curve  ")

    def _show_help(self):
        win = tk.Toplevel(self)
        win.title("사용 방법")
        win.configure(bg=BG_PANEL)
        win.geometry("640x540")
        win.resizable(False, False)
        help_text = """
        TFT Curve Analyzer  사용 방법
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        
        📈 Transfer Curve 탭
          • 입력 데이터 열 순서: A(무시), B=Vg, C=Vd, D=Id
          • Vd 값별로 Vgs-Id 곡선이 그려집니다.
          • 데이터 시작 행은 첫 숫자 등장 행부터 자동 감지됩니다.
          • Y축 로그 스케일 권장 (Id 값이 매우 작은 경우)
        
        📉 Output Curve 탭
          • 입력 데이터 열 순서: A(무시), B=Vd, C=Vg, D=Id
          • Vg 값별로 Vd-Id 곡선이 그려집니다.
          • 데이터 시작 행은 첫 숫자 등장 행부터 자동 감지됩니다.
        
        🤝 비교 분석 기능 (R-squared)
          • '메인 데이터'와 함께 '비교 데이터(Ref)'를 추가 업로드하세요.
          • 동일 조건(동일한 Vd 혹은 Vg) 데이터 곡선이 점선으로 오버레이 표시됩니다.
          • 두 곡선의 일치 수준 지표(R²)가 범례에 포함되어 표시됩니다.
          
        💾 결과 저장
          • '엑셀로 저장' 버튼으로 결과 파일 생성
          • Raw Data 시트: 파라미터별 정렬된 데이터
          • Chart 시트: 차트 이미지 고해상도 포함
        
        ⚠️  주의사항
          • 엑셀 파일(.xls, .xlsx) 및 CSV 파일(.csv) 형식 모두 지원합니다.
          • Vd/Vg 조건이 5개 미만인 이상치(Noise)는 렌더링에서 자동 제외됩니다.
        """
        txt = tk.Text(win, bg=BG_PANEL, fg=FG_TEXT, font=("Segoe UI", 10),
                      relief="flat", bd=8, wrap="word")
        txt.insert("1.0", textwrap.dedent(help_text).strip())
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True, padx=16, pady=16)
        tk.Button(win, text="닫기", command=win.destroy,
                  **BTN_STYLE_SEC).pack(pady=8)

    def _show_about(self):
        messagebox.showinfo(
            "정보",
            "TFT Curve Analyzer v1.1\n\n"
            "SmartSpice TFT 시뮬레이션 데이터를\n"
            "Transfer Curve / Output Curve로 분석 및 레퍼런스(R²) 비교하는 도구입니다.\n\n"
            "제작: 2026"
        )


# ══════════════════════════════════════════════════════════════════
#   진입점
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()

