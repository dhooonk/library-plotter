import pandas as pd
import numpy as np

MIN_DATA_POINTS = 5
PARAM_TOLERANCE = 1e-9


def _is_valid_number(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip()
    if not s:
        return False
    # 과학적 표기법의 'e'/'E'와 부호·소수점을 제외한 알파벳이 있으면 숫자가 아님
    for char in s:
        if char.isalpha() and char.lower() != 'e':
            return False
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def _is_data_row(row) -> bool:
    if not (_is_valid_number(row[1]) and _is_valid_number(row[2]) and _is_valid_number(row[3])):
        return False
    a_val = row[0]
    if pd.isna(a_val):
        return True
    a_str = str(a_val).strip()
    if a_str == "":
        return True
    # SmartSpice 포맷: 실제 데이터 행의 A열 값은 반드시 'DataValue'
    return a_str == "DataValue"


def _detect_data_start_row(filepath: str, is_csv: bool = False) -> int:
    if is_csv:
        raw = pd.read_csv(filepath, header=None)
    else:
        raw = pd.read_excel(filepath, header=None, engine="openpyxl")

    for row_idx in range(len(raw)):
        row = raw.iloc[row_idx]
        if len(row) < 4:
            continue
        if _is_data_row(row):
            return row_idx

    raise ValueError(
        "데이터 시작 행을 찾을 수 없습니다. "
        "B, C, D 열에 순수 숫자 데이터가 포함되고, "
        "A 열에 'DataValue' 값이 있는 행이 있는지 확인하세요."
    )


def _load_raw_dataframe(filepath: str, is_csv: bool, start_row: int) -> pd.DataFrame:
    if is_csv:
        raw = pd.read_csv(filepath, header=None, skiprows=start_row)
    else:
        raw = pd.read_excel(filepath, header=None, skiprows=start_row, engine="openpyxl")
    raw.columns = range(len(raw.columns))
    valid_mask = raw.apply(
        lambda row: _is_data_row(row) if len(row) >= 4 else False, axis=1
    )
    return raw[valid_mask].reset_index(drop=True)


def _group_by_param(df: pd.DataFrame, param_col: str, x_col: str, y_col: str) -> dict:
    result = {}
    param_vals = df[param_col].values
    x_vals = df[x_col].values
    y_vals = df[y_col].values

    seen = set()
    unique_params = []
    for v in param_vals:
        # 부동소수점 오차 보정: 소수점 6자리 반올림으로 동일 파라미터 묶음
        rounded = round(float(v), 6)
        if rounded not in seen:
            seen.add(rounded)
            unique_params.append(float(v))
    unique_params.sort()

    for p in unique_params:
        mask = np.abs(param_vals.astype(float) - p) < PARAM_TOLERANCE
        x = x_vals[mask].astype(float)
        y = y_vals[mask].astype(float)
        if len(x) < MIN_DATA_POINTS:
            continue
        sort_idx = np.argsort(x)
        result[p] = {"x": x[sort_idx], "y": y[sort_idx]}

    return result


def parse_transfer_curve(filepath: str) -> dict:
    """Transfer Curve 파싱. Returns {Vd: {x: Vg[], y: Id[]}}."""
    is_csv = filepath.lower().endswith('.csv')
    start_row = _detect_data_start_row(filepath, is_csv)
    raw = _load_raw_dataframe(filepath, is_csv, start_row)

    df = raw[[1, 2, 3]].copy()
    df.columns = ["Vg", "Vd", "Id"]
    df = df.apply(pd.to_numeric, errors="coerce").dropna()
    return _group_by_param(df, param_col="Vd", x_col="Vg", y_col="Id")


def parse_output_curve(filepath: str) -> dict:
    """Output Curve 파싱. Returns {Vg: {x: Vd[], y: Id[]}}."""
    is_csv = filepath.lower().endswith('.csv')
    start_row = _detect_data_start_row(filepath, is_csv)
    raw = _load_raw_dataframe(filepath, is_csv, start_row)

    df = raw[[1, 2, 3]].copy()
    df.columns = ["Vd", "Vg", "Id"]
    df = df.apply(pd.to_numeric, errors="coerce").dropna()
    return _group_by_param(df, param_col="Vg", x_col="Vd", y_col="Id")
