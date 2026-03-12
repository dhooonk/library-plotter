"""
main.py
-------
TFT Curve Analyzer - 메인 GUI 어플리케이션 (tkinter)
Transfer Curve (Vgs-Id) 및 Output Curve (Vd-Id) 분석을 지원.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

import data_parser
import transfer_curve as tc_module
import output_curve as oc_module
import excel_exporter


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
    """Transfer/Output 탭의 공통 레이아웃."""

    MODE_LABEL   = ""   # 서브클래스에서 오버라이드
    PARAM_LABEL  = ""   # "Vd" or "Vg"
    X_LABEL      = ""   # "Vg" or "Vd"
    DEFAULT_LOG  = True

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG_DARK, **kw)
        self._filepath  = None
        self._fig       = None
        self._grouped   = None
        self._canvas    = None
        self._toolbar   = None
        self._log_var   = tk.BooleanVar(value=self.DEFAULT_LOG)
        self._status_var = tk.StringVar(value="파일을 선택해 주세요.")
        self._build_ui()

    # ── UI 빌드 ─────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, weight=0, minsize=300)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # 좌측 패널
        left = tk.Frame(self, bg=BG_PANEL, width=300)
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_propagate(False)
        self._build_left(left)

        # 우측 차트 패널
        right = tk.Frame(self, bg=BG_DARK)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        self._build_right(right)

    def _build_left(self, parent):
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
        cards.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        cards.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        # 파일 선택 카드
        fc = self._make_card(cards, "📂  입력 파일", row=0)
        self._file_label = tk.Label(fc, text="선택된 파일 없음", font=FONT_SMALL,
                                    bg=BG_CARD, fg=FG_MUTED, wraplength=240,
                                    justify="left", anchor="w")
        self._file_label.pack(fill="x", padx=8, pady=(0, 4))
        tk.Button(fc, text="엑셀 파일 선택 (.xlsx / .xls)",
                  command=self._select_file, **BTN_STYLE_PRI).pack(fill="x",
                  padx=8, pady=4)

        # 옵션 카드
        oc = self._make_card(cards, "⚙️  옵션", row=1)
        lf = tk.Frame(oc, bg=BG_CARD)
        lf.pack(fill="x", padx=8, pady=4)
        tk.Checkbutton(lf, text="Y축 로그 스케일 (Log Scale)",
                       variable=self._log_var,
                       bg=BG_CARD, fg=FG_TEXT, selectcolor=BG_ACCENT,
                       activebackground=BG_CARD, activeforeground=FG_TEXT,
                       font=FONT_LABEL, command=self._on_option_change,
                       relief="flat").pack(anchor="w")

        # 저장 경로 카드
        sc = self._make_card(cards, "💾  저장 설정", row=2)
        save_frm, self._save_entry = _labeled_entry(sc, "저장 파일명", "result.xlsx", width=28)
        save_frm.pack(fill="x", padx=8, pady=(2, 4))
        tk.Button(sc, text="저장 경로 선택", command=self._pick_save_dir,
                  **BTN_STYLE_SEC).pack(fill="x", padx=8, pady=(2, 4))
        self._save_dir_label = tk.Label(sc, text=os.path.expanduser("~"),
                                         font=FONT_SMALL, bg=BG_CARD,
                                         fg=FG_MUTED, wraplength=240,
                                         justify="left", anchor="w")
        self._save_dir_label.pack(fill="x", padx=8, pady=(0, 4))
        self._save_dir = os.path.expanduser("~")

        # 실행 카드
        rc = self._make_card(cards, "▶  실행", row=3)
        tk.Button(rc, text="📊  분석 실행",
                  command=self._run_analysis, **BTN_STYLE_PRI).pack(
                  fill="x", padx=8, pady=4)
        tk.Button(rc, text="📥  엑셀로 저장",
                  command=self._export_excel, **BTN_STYLE_GRN).pack(
                  fill="x", padx=8, pady=4)

        # 로그 카드
        lc = self._make_card(cards, "📋  로그", row=4)
        log_frm, self._log_text = _scrollable_text(lc, height=7)
        log_frm.pack(fill="both", expand=True, padx=8, pady=4)

        # 상태 바
        sb = tk.Frame(parent, bg="#0d0d20", pady=6)
        sb.grid(row=2, column=0, sticky="ew")
        tk.Label(sb, textvariable=self._status_var, font=FONT_SMALL,
                 bg="#0d0d20", fg=FG_MUTED, wraplength=260,
                 justify="left").pack(padx=12, anchor="w")

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
                 text="엑셀 파일을 선택하고\n'분석 실행'을 눌러주세요.",
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

    # ── 이벤트 핸들러 ────────────────────────────────────────────

    def _select_file(self):
        fp = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if fp:
            self._filepath = fp
            fname = os.path.basename(fp)
            self._file_label.config(text=fname, fg=ACCENT_GREEN)
            self._log(f"파일 선택됨: {fname}")
            self._status_var.set(f"✅ 파일 로드 준비 완료")
            # 파일명 기반으로 저장 파일명 자동 설정
            base = os.path.splitext(fname)[0]
            self._save_entry.delete(0, "end")
            self._save_entry.insert(0, f"{base}_result.xlsx")
            self._save_dir = os.path.dirname(fp)
            self._save_dir_label.config(text=self._save_dir)

    def _pick_save_dir(self):
        d = filedialog.askdirectory(title="저장 폴더 선택")
        if d:
            self._save_dir = d
            self._save_dir_label.config(text=d)

    def _on_option_change(self):
        if self._grouped is not None:
            self._draw_chart()

    def _run_analysis(self):
        if not self._filepath:
            messagebox.showwarning("파일 없음", "엑셀 파일을 먼저 선택해주세요.")
            return
        self._log("분석 시작...")
        self._status_var.set("⏳ 데이터 파싱 중...")
        # 별도 스레드에서 실행 (UI 블로킹 방지)
        threading.Thread(target=self._analysis_thread, daemon=True).start()

    def _analysis_thread(self):
        try:
            self._grouped = self._parse_data(self._filepath)
            n_params = len(self._grouped)
            param_list = ", ".join(
                f"{self._format_val(v)}V" for v in sorted(self._grouped.keys())
            )
            self.after(0, lambda: self._log(
                f"파싱 완료: {self.PARAM_LABEL} {n_params}개 감지\n  → {param_list}"
            ))
            self.after(0, lambda: self._status_var.set(
                f"✅ {self.PARAM_LABEL} {n_params}개 감지됨"
            ))
            self.after(0, lambda: self._param_label.config(
                text=f"{self.PARAM_LABEL} 값: {param_list}"
            ))
            self.after(0, self._draw_chart)
        except Exception as e:
            self.after(0, lambda: self._log(f"❌ 오류: {e}"))
            self.after(0, lambda: self._status_var.set(f"❌ 오류 발생"))
            self.after(0, lambda: messagebox.showerror("파싱 오류", str(e)))

    def _draw_chart(self):
        if self._grouped is None:
            return

        # 기존 차트 제거
        for w in self._chart_frame.winfo_children():
            w.destroy()

        log_scale = self._log_var.get()
        self._fig = self._create_figure(self._grouped, log_scale)

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

        self._log(f"차트 생성 완료 (로그 스케일: {'ON' if log_scale else 'OFF'})")

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
            self._do_export(self._grouped, self._fig, save_path, log_scale)
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

    def _create_figure(self, grouped: dict, log_scale: bool):
        raise NotImplementedError

    def _do_export(self, grouped, fig, save_path, log_scale):
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

    def _create_figure(self, grouped, log_scale):
        plt.close("all")
        return tc_module.create_transfer_figure(
            grouped, log_scale=log_scale,
            title="TFT Transfer Curve  (Vgs - Id)",
        )

    def _do_export(self, grouped, fig, save_path, log_scale):
        excel_exporter.export_transfer_curve(grouped, fig, save_path, log_scale)


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

    def _create_figure(self, grouped, log_scale):
        plt.close("all")
        return oc_module.create_output_figure(
            grouped, log_scale=log_scale,
            title="TFT Output Curve  (Vd - Id)",
        )

    def _do_export(self, grouped, fig, save_path, log_scale):
        excel_exporter.export_output_curve(grouped, fig, save_path, log_scale)


# ══════════════════════════════════════════════════════════════════
#   메인 앱
# ══════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TFT Curve Analyzer  ─  SmartSpice 시뮬레이션 결과 분석")
        self.geometry("1280x780")
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
                  foreground=[("selected", FG_TEXT)],
                  expand=[("selected", [1, 1, 1, 0])])

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
        win.geometry("560x480")
        win.resizable(False, False)
        help_text = """
TFT Curve Analyzer  사용 방법
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📈 Transfer Curve 탭
  • 입력 엑셀 열 순서: A(무시), B=Vg, C=Vd, D=Id
  • Vd 값별로 Vgs-Id 곡선이 그려집니다.
  • 데이터 시작 행은 자동 감지됩니다.
  • Y축 로그 스케일 권장 (Id 값이 매우 작은 경우)

📉 Output Curve 탭
  • 입력 엑셀 열 순서: A(무시), B=Vd, C=Vg, D=Id
  • Vg 값별로 Vd-Id 곡선이 그려집니다.
  • 데이터 시작 행은 자동 감지됩니다.

💾 결과 저장
  • '엑셀로 저장' 버튼으로 결과 파일 생성
  • Raw Data 시트: 파라미터별 정렬된 데이터
  • Chart 시트: 차트 이미지 포함

⚠️  주의사항
  • 엑셀 파일의 B, C, D 열에 숫자 데이터가 있어야 합니다.
  • .xlsx 및 .xls 형식 모두 지원합니다.
  • Vd/Vg 고유 값은 소수점 허용오차로 자동 구분됩니다.
"""
        txt = tk.Text(win, bg=BG_PANEL, fg=FG_TEXT, font=("Segoe UI", 10),
                      relief="flat", bd=8, wrap="word")
        txt.insert("1.0", help_text)
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True, padx=16, pady=16)
        tk.Button(win, text="닫기", command=win.destroy,
                  **BTN_STYLE_SEC).pack(pady=8)

    def _show_about(self):
        messagebox.showinfo(
            "정보",
            "TFT Curve Analyzer v1.0\n\n"
            "SmartSpice TFT 시뮬레이션 데이터를\n"
            "Transfer Curve / Output Curve로 분석하는 도구입니다.\n\n"
            "제작: 2026"
        )


# ══════════════════════════════════════════════════════════════════
#   진입점
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
