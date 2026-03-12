"""
data_parser.py
--------------
엑셀/CSV 파일을 읽고 TFT 시뮬레이션 데이터를 파싱하는 모듈.
- Transfer Curve: B=Vg, C=Vd, D=Id
- Output Curve:   B=Vd, C=Vg, D=Id

행 자동 감지: 헤더 위치가 파일마다 다르므로,
숫자 데이터가 4열 연속으로 나타나는 첫 번째 행을 데이터 시작점으로 결정하되,
알파벳 등 이상한 문자(1d 등)가 있는 행은 건너뜁니다.
"""

import pandas as pd
import numpy as np


def _is_valid_number(val) -> bool:
    """주어진 값이 순수한 숫자인지 검증 (과학적 표기법 허용, 이상한 문자 불허)"""
    if pd.isna(val):
        return False
    s = str(val).strip()
    if not s:
        return False
    # 과학적 표기법(예: 1.23E-10)의 e/E, 음수 부호 -, 소수점 . 외의 알파벳이 있으면 False
    for char in s:
        if char.isalpha() and char.lower() != 'e':
            return False
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def _detect_data_start_row(filepath: str, is_csv: bool = False) -> int:
    """
    엑셀/CSV 파일에서 순수 숫자 데이터가 시작하는 행 번호(0-indexed)를 자동 감지.
    각 행을 읽어, B~D 열에 해당하는 값이 순수한 숫자인 첫 행을 반환.
    """
    if is_csv:
        raw = pd.read_csv(filepath, header=None)
    else:
        raw = pd.read_excel(filepath, header=None, engine="openpyxl")

    for row_idx in range(len(raw)):
        row = raw.iloc[row_idx]
        if len(row) < 4:
            continue
        
        # B(1), C(2), D(3) 값이 모두 유효한 숫자인지 확인
        if _is_valid_number(row[1]) and _is_valid_number(row[2]) and _is_valid_number(row[3]):
            return row_idx

    raise ValueError(
        "데이터 시작 행을 찾을 수 없습니다. "
        "B, C, D 열에 순수 숫자 데이터가 포함된 행이 있는지 확인하세요."
    )


def _group_by_param(df: pd.DataFrame, param_col: str, x_col: str, y_col: str, min_points: int = 5) -> dict:
    """
    param_col 값(예: Vd)을 기준으로 데이터를 그룹화.
    각 그룹에서 x_col, y_col 값을 numpy 배열로 반환.
    고립된 쓰레기 값(예: 포인트 5개 미만)은 노이즈로 간주하고 무시.
    """
    result = {}
    param_vals = df[param_col].values
    x_vals = df[x_col].values
    y_vals = df[y_col].values

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
        
        # 쓰레기 값/노이즈 조건 (데이터가 너무 적음) 필터링
        if len(x) < min_points:
            continue
            
        sort_idx = np.argsort(x)
        result[p] = {"x": x[sort_idx], "y": y[sort_idx]}

    return result


def parse_transfer_curve(filepath: str) -> dict:
    """Transfer Curve 데이터 파싱"""
    is_csv = str(filepath).lower().endswith('.csv')
    start_row = _detect_data_start_row(filepath, is_csv)

    if is_csv:
        raw = pd.read_csv(filepath, header=None, skiprows=start_row)
    else:
        raw = pd.read_excel(filepath, header=None, skiprows=start_row, engine="openpyxl")

    raw.columns = range(len(raw.columns))
    df = raw[[1, 2, 3]].copy()
    df.columns = ["Vg", "Vd", "Id"]
    
    # 숫자형 변환 (문자열 등은 NaN 처리 후 drop)
    df["Vg"] = pd.to_numeric(df["Vg"], errors="coerce")
    df["Vd"] = pd.to_numeric(df["Vd"], errors="coerce")
    df["Id"] = pd.to_numeric(df["Id"], errors="coerce")
    df = df.dropna()
    
    grouped = _group_by_param(df, param_col="Vd", x_col="Vg", y_col="Id")
    return grouped


def parse_output_curve(filepath: str) -> dict:
    """Output Curve 데이터 파싱"""
    is_csv = str(filepath).lower().endswith('.csv')
    start_row = _detect_data_start_row(filepath, is_csv)

    if is_csv:
        raw = pd.read_csv(filepath, header=None, skiprows=start_row)
    else:
        raw = pd.read_excel(filepath, header=None, skiprows=start_row, engine="openpyxl")

    raw.columns = range(len(raw.columns))
    df = raw[[1, 2, 3]].copy()
    df.columns = ["Vd", "Vg", "Id"]
    
    df["Vd"] = pd.to_numeric(df["Vd"], errors="coerce")
    df["Vg"] = pd.to_numeric(df["Vg"], errors="coerce")
    df["Id"] = pd.to_numeric(df["Id"], errors="coerce")
    df = df.dropna()
    
    grouped = _group_by_param(df, param_col="Vg", x_col="Vd", y_col="Id")
    return grouped

