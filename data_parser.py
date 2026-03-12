"""
data_parser.py
--------------
엑셀 파일을 읽고 TFT 시뮬레이션 데이터를 파싱하는 모듈.
- Transfer Curve: B=Vg, C=Vd, D=Id
- Output Curve:   B=Vd, C=Vg, D=Id

행 자동 감지: 헤더 위치가 파일마다 다르므로,
숫자 데이터가 4열 연속으로 나타나는 첫 번째 행을 데이터 시작점으로 결정.
"""

import pandas as pd
import numpy as np


def _detect_data_start_row(filepath: str) -> int:
    """
    엑셀 파일에서 숫자 데이터가 시작하는 행 번호(0-indexed)를 자동 감지.
    각 행을 순서대로 읽어, B~D 열에 해당하는 값이 숫자인 첫 행을 반환.
    """
    # 헤더 없이 전부 읽기
    raw = pd.read_excel(filepath, header=None, engine="openpyxl")

    for row_idx in range(len(raw)):
        row = raw.iloc[row_idx]
        # 최소 4열이 있어야 함 (A, B, C, D)
        if len(row) < 4:
            continue
        try:
            # B(index 1), C(index 2), D(index 3) 가 모두 숫자인지 확인
            float(row[1])
            float(row[2])
            float(row[3])
            return row_idx
        except (ValueError, TypeError):
            continue

    raise ValueError(
        "데이터 시작 행을 찾을 수 없습니다. "
        "B, C, D 열에 숫자 데이터가 포함된 행이 있는지 확인하세요."
    )


def _group_by_param(df: pd.DataFrame, param_col: str, x_col: str, y_col: str) -> dict:
    """
    param_col 값(예: Vd)을 기준으로 데이터를 그룹화하고,
    각 그룹에서 x_col, y_col 값을 numpy 배열로 반환.

    부동소수점 오차를 허용하여 고유 param 값을 식별.
    반환: {param_value: {'x': np.array, 'y': np.array}, ...}
    """
    result = {}
    param_vals = df[param_col].values
    x_vals = df[x_col].values
    y_vals = df[y_col].values

    # 고유 파라미터 값 수집 (float 반올림으로 비교)
    unique_params = []
    seen = set()
    for v in param_vals:
        rounded = round(float(v), 6)
        if rounded not in seen:
            seen.add(rounded)
            unique_params.append(float(v))
    unique_params.sort()

    for p in unique_params:
        mask = np.abs(param_vals.astype(float) - p) < 1e-9
        x = x_vals[mask].astype(float)
        y = y_vals[mask].astype(float)
        # x 기준으로 정렬
        sort_idx = np.argsort(x)
        result[p] = {"x": x[sort_idx], "y": y[sort_idx]}

    return result


def parse_transfer_curve(filepath: str) -> dict:
    """
    Transfer Curve 데이터 파싱.
    열 레이아웃: A(무의미), B=Vg, C=Vd, D=Id
    반환: {vd_value: {'x': Vg 배열, 'y': Id 배열}, ...}
    """
    start_row = _detect_data_start_row(filepath)
    raw = pd.read_excel(filepath, header=None, skiprows=start_row, engine="openpyxl")

    # 컬럼명 매핑 (0=A, 1=B=Vg, 2=C=Vd, 3=D=Id)
    raw.columns = range(len(raw.columns))
    df = raw[[1, 2, 3]].copy()
    df.columns = ["Vg", "Vd", "Id"]
    df = df.dropna()
    df = df[pd.to_numeric(df["Vg"], errors="coerce").notna()]
    df = df[pd.to_numeric(df["Vd"], errors="coerce").notna()]
    df = df[pd.to_numeric(df["Id"], errors="coerce").notna()]
    df = df.astype({"Vg": float, "Vd": float, "Id": float})

    # Vd 기준으로 그룹화, x=Vg, y=Id
    grouped = _group_by_param(df, param_col="Vd", x_col="Vg", y_col="Id")
    return grouped


def parse_output_curve(filepath: str) -> dict:
    """
    Output Curve 데이터 파싱.
    열 레이아웃: A(무의미), B=Vd, C=Vg, D=Id
    반환: {vg_value: {'x': Vd 배열, 'y': Id 배열}, ...}
    """
    start_row = _detect_data_start_row(filepath)
    raw = pd.read_excel(filepath, header=None, skiprows=start_row, engine="openpyxl")

    raw.columns = range(len(raw.columns))
    df = raw[[1, 2, 3]].copy()
    df.columns = ["Vd", "Vg", "Id"]
    df = df.dropna()
    df = df[pd.to_numeric(df["Vd"], errors="coerce").notna()]
    df = df[pd.to_numeric(df["Vg"], errors="coerce").notna()]
    df = df[pd.to_numeric(df["Id"], errors="coerce").notna()]
    df = df.astype({"Vd": float, "Vg": float, "Id": float})

    # Vg 기준으로 그룹화, x=Vd, y=Id
    grouped = _group_by_param(df, param_col="Vg", x_col="Vd", y_col="Id")
    return grouped
