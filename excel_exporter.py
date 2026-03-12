"""
excel_exporter.py
-----------------
분석 결과를 엑셀 파일로 저장하는 모듈.
- Sheet 1: Raw 데이터 (파라미터별 정렬)
- Sheet 2: 차트 이미지 (matplotlib 차트 → PNG → 엑셀 삽입)
"""

import io
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")  # GUI 없이 이미지 렌더링
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from matplotlib.figure import Figure


# ────────────────────────── 스타일 상수 ──────────────────────────
_HEADER_FILL    = PatternFill("solid", fgColor="1F3864")  # 진한 파랑
_PARAM_FILL     = PatternFill("solid", fgColor="2E4057")  # 미디엄 파랑
_ALT_FILL       = PatternFill("solid", fgColor="F2F7FF")  # 연한 파랑 (짝수 행)
_HEADER_FONT    = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_PARAM_FONT     = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_DATA_FONT      = Font(name="Calibri", size=10)
_THIN_BORDER    = Border(
    left=Side(style="thin", color="C0C0C0"),
    right=Side(style="thin", color="C0C0C0"),
    top=Side(style="thin", color="C0C0C0"),
    bottom=Side(style="thin", color="C0C0C0"),
)
_CENTER         = Alignment(horizontal="center", vertical="center")
_NUM_FMT_SCI    = "0.000E+00"
_NUM_FMT_VOLT   = "0.000"


def _apply_cell(ws, row, col, value, fill=None, font=None, alignment=None,
                border=None, number_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format
    return cell


def _figure_to_xl_image(fig: Figure, width_px: int = 680, height_px: int = 480) -> XLImage:
    """matplotlib Figure를 openpyxl Image 객체로 변환."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight",
                facecolor=fig.get_facecolor(), dpi=130)
    buf.seek(0)
    img = XLImage(buf)
    # 셀 크기에 맞게 조정 (픽셀 단위)
    img.width = width_px
    img.height = height_px
    return img


import output_curve as oc_module
import transfer_curve as tc_module

# ───────────────────────── Transfer Curve 내보내기 ─────────────────────────

def export_transfer_curve(
    grouped_data: dict,
    fig: Figure,
    save_path: str,
    log_scale: bool = True,
    xlim: tuple = None,
    ylim: tuple = None
):
    """
    Transfer Curve 결과를 엑셀로 저장.
    """
    wb = Workbook()

    # ── Sheet 1: Raw 데이터 ──────────────────────────────────────────────
    ws_data = wb.active
    ws_data.title = "Raw Data"

    sorted_vds = sorted(grouped_data.keys())

    # 헤더 행 (Vd별 Vg / Id 쌍)
    ws_data.row_dimensions[1].height = 20
    ws_data.row_dimensions[2].height = 18

    col = 1
    vd_col_starts = {}
    for vd in sorted_vds:
        vd_col_starts[vd] = col
        vd_label = f"Vd = {_fmt(vd)} V"
        # 병합 헤더
        ws_data.merge_cells(
            start_row=1, start_column=col,
            end_row=1, end_column=col + 1
        )
        _apply_cell(ws_data, 1, col, vd_label,
                    fill=_HEADER_FILL, font=_HEADER_FONT,
                    alignment=_CENTER, border=_THIN_BORDER)
        # 서브 헤더
        _apply_cell(ws_data, 2, col, "Vgs (V)",
                    fill=_PARAM_FILL, font=_PARAM_FONT,
                    alignment=_CENTER, border=_THIN_BORDER)
        _apply_cell(ws_data, 2, col + 1, "Id (A)",
                    fill=_PARAM_FILL, font=_PARAM_FONT,
                    alignment=_CENTER, border=_THIN_BORDER)
        ws_data.column_dimensions[get_column_letter(col)].width = 14
        ws_data.column_dimensions[get_column_letter(col + 1)].width = 16
        col += 2

    # 데이터 행
    max_len = max(len(d["x"]) for d in grouped_data.values()) if grouped_data else 0
    for row_i in range(max_len):
        excel_row = row_i + 3
        is_odd = (row_i % 2 == 0)
        for vd in sorted_vds:
            c = vd_col_starts[vd]
            x_arr = grouped_data[vd]["x"]
            y_arr = grouped_data[vd]["y"]
            row_fill = None if is_odd else _ALT_FILL
            if row_i < len(x_arr):
                _apply_cell(ws_data, excel_row, c, float(x_arr[row_i]),
                            fill=row_fill, font=_DATA_FONT,
                            alignment=_CENTER, border=_THIN_BORDER,
                            number_format=_NUM_FMT_VOLT)
                _apply_cell(ws_data, excel_row, c + 1, float(y_arr[row_i]),
                            fill=row_fill, font=_DATA_FONT,
                            alignment=_CENTER, border=_THIN_BORDER,
                            number_format=_NUM_FMT_SCI)
            else:
                _apply_cell(ws_data, excel_row, c, None,
                            fill=row_fill, border=_THIN_BORDER)
                _apply_cell(ws_data, excel_row, c + 1, None,
                            fill=row_fill, border=_THIN_BORDER)

    # 틀 고정
    ws_data.freeze_panes = "A3"

    # ── Sheet 2: 차트 ────────────────────────────────────────────────────
    ws_chart = wb.create_sheet("Transfer Curve Chart")
    ws_chart.sheet_view.showGridLines = False
    ws_chart.column_dimensions["A"].width = 1

    title_font = Font(name="Calibri", bold=True, size=13, color="1F3864")
    _apply_cell(ws_chart, 1, 2, "TFT Transfer Curve (Vgs-Id)",
                font=title_font, alignment=_CENTER)
    _apply_cell(ws_chart, 2, 2,
                f"Y축: {'로그 스케일 (Log)' if log_scale else '선형 스케일 (Linear)'}",
                font=Font(name="Calibri", italic=True, size=10, color="555555"),
                alignment=_CENTER)

    # UI의 fig를 참조하면 빈 화면이 될 수 있으므로, 독립적으로 새로 랜더링
    export_fig = tc_module.create_transfer_figure(
        grouped_data, log_scale=log_scale,
        title="TFT Transfer Curve  (Vgs - Id)",
        xlim=xlim, ylim=ylim
    )
    xl_img = _figure_to_xl_image(export_fig, width_px=700, height_px=500)
    plt.close(export_fig)
    ws_chart.add_image(xl_img, "B4")

    if not save_path.endswith(".xlsx"):
        save_path += ".xlsx"
    wb.save(save_path)


# ───────────────────────── Output Curve 내보내기 ─────────────────────────

def export_output_curve(
    grouped_data: dict,
    fig: Figure,
    save_path: str,
    log_scale: bool = False,
    xlim: tuple = None,
    ylim: tuple = None
):
    """
    Output Curve 결과를 엑셀로 저장.
    """
    wb = Workbook()

    ws_data = wb.active
    ws_data.title = "Raw Data"

    sorted_vgs = sorted(grouped_data.keys())

    ws_data.row_dimensions[1].height = 20
    ws_data.row_dimensions[2].height = 18

    col = 1
    vg_col_starts = {}
    for vg in sorted_vgs:
        vg_col_starts[vg] = col
        vg_label = f"Vg = {_fmt(vg)} V"
        ws_data.merge_cells(
            start_row=1, start_column=col,
            end_row=1, end_column=col + 1
        )
        _apply_cell(ws_data, 1, col, vg_label,
                    fill=_HEADER_FILL, font=_HEADER_FONT,
                    alignment=_CENTER, border=_THIN_BORDER)
        _apply_cell(ws_data, 2, col, "Vds (V)",
                    fill=_PARAM_FILL, font=_PARAM_FONT,
                    alignment=_CENTER, border=_THIN_BORDER)
        _apply_cell(ws_data, 2, col + 1, "Id (A)",
                    fill=_PARAM_FILL, font=_PARAM_FONT,
                    alignment=_CENTER, border=_THIN_BORDER)
        ws_data.column_dimensions[get_column_letter(col)].width = 14
        ws_data.column_dimensions[get_column_letter(col + 1)].width = 16
        col += 2

    max_len = max(len(d["x"]) for d in grouped_data.values()) if grouped_data else 0
    for row_i in range(max_len):
        excel_row = row_i + 3
        is_odd = (row_i % 2 == 0)
        for vg in sorted_vgs:
            c = vg_col_starts[vg]
            x_arr = grouped_data[vg]["x"]
            y_arr = grouped_data[vg]["y"]
            row_fill = None if is_odd else _ALT_FILL
            if row_i < len(x_arr):
                _apply_cell(ws_data, excel_row, c, float(x_arr[row_i]),
                            fill=row_fill, font=_DATA_FONT,
                            alignment=_CENTER, border=_THIN_BORDER,
                            number_format=_NUM_FMT_VOLT)
                _apply_cell(ws_data, excel_row, c + 1, float(y_arr[row_i]),
                            fill=row_fill, font=_DATA_FONT,
                            alignment=_CENTER, border=_THIN_BORDER,
                            number_format=_NUM_FMT_SCI)
            else:
                _apply_cell(ws_data, excel_row, c, None,
                            fill=row_fill, border=_THIN_BORDER)
                _apply_cell(ws_data, excel_row, c + 1, None,
                            fill=row_fill, border=_THIN_BORDER)

    ws_data.freeze_panes = "A3"

    ws_chart = wb.create_sheet("Output Curve Chart")
    ws_chart.sheet_view.showGridLines = False
    ws_chart.column_dimensions["A"].width = 1

    title_font = Font(name="Calibri", bold=True, size=13, color="1F3864")
    _apply_cell(ws_chart, 1, 2, "TFT Output Curve (Vd-Id)",
                font=title_font, alignment=_CENTER)
    _apply_cell(ws_chart, 2, 2,
                f"Y축: {'로그 스케일 (Log)' if log_scale else '선형 스케일 (Linear)'}",
                font=Font(name="Calibri", italic=True, size=10, color="555555"),
                alignment=_CENTER)

    # 출력 전용 figure 생성으로 Blank 방지
    export_fig = oc_module.create_output_figure(
        grouped_data, log_scale=log_scale,
        title="TFT Output Curve  (Vd - Id)",
        xlim=xlim, ylim=ylim
    )
    xl_img = _figure_to_xl_image(export_fig, width_px=700, height_px=500)
    plt.close(export_fig)
    ws_chart.add_image(xl_img, "B4")

    if not save_path.endswith(".xlsx"):
        save_path += ".xlsx"
    wb.save(save_path)


def _fmt(v: float) -> str:
    if v == int(v):
        return str(int(v))
    return f"{v:.3g}"

