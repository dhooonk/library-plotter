"""
data_parser.py
--------------
엑셀(.xlsx, .xls) 및 CSV(.csv) 파일을 읽고 TFT 시뮬레이션 데이터를 파싱하는 모듈입니다.
TFT 측정/시뮬레이션 데이터 형식은 아래와 같이 열이 구성되는 것을 전제로 합니다.
- Transfer Curve: B=Vg(게이트 전압), C=Vd(드레인 전압), D=Id(드레인 전류)
- Output Curve:   B=Vd(드레인 전압), C=Vg(게이트 전압), D=Id(드레인 전류)

주요 특징:
1. 행 자동 감지 기능: 파일 상단에 존재하는 불필요한 텍스트 행을 건너뛰고,
   실제 숫자 데이터가 나타나는 첫 번째 행을 탐색합니다.
   (의미 없는 '1d 5' 등의 텍스트 데이터는 무시)
2. 이상치 필터링(Noise filtering): 조건을 만족하는 데이터의 개수가 특정 값(예: 5개)
   미만인 경우 노이즈나 측정 오류로 간주하여 분석에서 제외합니다.
"""

import pandas as pd
import numpy as np


def _is_valid_number(val) -> bool:
    """
    주어진 값이 순수한 숫자로 변환 가능한지 검증하는 함수입니다.
    
    Args:
        val: 검증할 값 (int, float, str 등 다양한 타입 입력 가능)
        
    Returns:
        bool: 순수 숫자로 변환 가능하면 True, 그 외 문자열(알파벳 혼재 등)이면 False.
        
    설명:
        과학적 표기법(예: 1.23E-10의 'e' 또는 'E')과 음수 부호 '-', 
        소수점('.')을 제외한 다른 알파벳 문자가 섞여있으면 False를 반환하여
        엉뚱한 문자열 열이 값으로 매핑되는 것을 방지합니다.
    """
    if pd.isna(val):
        return False
    
    # 공백을 제거한 문자열 형태로 변환
    s = str(val).strip()
    if not s:
        return False
        
    # 과학적 표기법에 쓰일 수 있는 'e', 'E', '-', '.' 제외 외의 문자가 없어야 함
    for char in s:
        if char.isalpha() and char.lower() != 'e':
            return False
            
    # float 형변환 테스트로 최종 유효성 검증
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def _is_data_row(row) -> bool:
    """
    주어진 행이 유효한 TFT 데이터 행인지 검증합니다.
    
    데이터 행 판단 기준:
    1. B열(인덱스 1), C열(인덱스 2), D열(인덱스 3) 모두 순수 숫자여야 합니다.
    2. A열(인덱스 0)에 문자열 값이 존재하는 경우, 반드시 'DataValue'여야만
       데이터 행으로 인정합니다. A열이 비어 있거나 NaN인 경우는 조건 1만 만족하면 됩니다.
    
    이 검증을 통해 헤더 섹션이나 파라미터 정의 섹션 중
    우연히 B~D열에 숫자가 포함된 행이 데이터로 잘못 파싱되는 것을 방지합니다.
    
    Args:
        row: pandas Series (DataFrame의 한 행)
        
    Returns:
        bool: 유효한 데이터 행이면 True, 그렇지 않으면 False
    """
    # 조건 1: B, C, D열이 모두 유효한 숫자인지 확인
    if not (_is_valid_number(row[1]) and _is_valid_number(row[2]) and _is_valid_number(row[3])):
        return False
    
    # 조건 2: A열(인덱스 0) 값 확인
    a_val = row[0]
    if pd.isna(a_val):
        # A열이 비어 있으면 숫자 조건만으로 데이터 행으로 인정
        return True
    
    a_str = str(a_val).strip()
    if a_str == "":
        # A열이 빈 문자열이면 데이터 행으로 인정
        return True
    
    # A열에 값이 있으면 반드시 'DataValue'여야 데이터 행으로 인정
    return a_str == "DataValue"


def _detect_data_start_row(filepath: str, is_csv: bool = False) -> int:
    """
    엑셀/CSV 파일에서 실제 숫자 데이터가 시작되는 행(row)의 인덱스를 자동 감지합니다.
    
    Args:
        filepath (str): 불러올 파일 경로
        is_csv (bool): 파일 형식이 CSV인지 여부
        
    Returns:
        int: 데이터가 시작되는 0-based 행 인덱스
        
    Raises:
        ValueError: 유효한 데이터 행(B~D열 숫자 + A열 DataValue 조건)을 찾지 못한 경우
    """
    # 헤더 없이 원시 데이터를 전부 읽어옵니다.
    if is_csv:
        raw = pd.read_csv(filepath, header=None)
    else:
        raw = pd.read_excel(filepath, header=None, engine="openpyxl")

    # 모든 행을 위에서부터 차례대로 탐색
    for row_idx in range(len(raw)):
        row = raw.iloc[row_idx]
        
        # 행의 열이 최소 4개(A, B, C, D)가 되지 않으면 스킵
        if len(row) < 4:
            continue
        
        # B,C,D열 숫자 조건 + A열 DataValue 조건을 모두 통과하면 데이터 시작 행으로 반환
        if _is_data_row(row):
            return row_idx

    # 끝까지 찾지 못했을 때 예외 발생
    raise ValueError(
        "데이터 시작 행을 찾을 수 없습니다. "
        "B, C, D 열에 순수 숫자 데이터가 포함되고, "
        "A 열에 'DataValue' 값이 있는 행이 있는지 확인하세요."
    )


def _group_by_param(df: pd.DataFrame, param_col: str, x_col: str, y_col: str, min_points: int = 5) -> dict:
    """
    지정된 파라미터 컬럼(예: Vd 또는 Vg)의 값을 기준으로 데이터를 묶습니다(Grouping).
    
    Args:
        df (pd.DataFrame): 결측치 제거 등 전처리가 완료된 원본 데이터
        param_col (str): 그룹의 기준이 될 컬럼명
        x_col (str): X축으로 쓸 데이터 컬럼명
        y_col (str): Y축으로 쓸 데이터 컬럼명
        min_points (int): 파라미터 그룹으로 인정하기 위한 최소 데이터 개수 
                          (이 개수 미만이면 이상값, 쓰레기 값으로 판단하여 버림)
                          
    Returns:
        dict: 파라미터 고유값 (float)을 Key로 가지고, 해당 파라미터에 대한
              "x", "y" 넘파이 배열 딕셔너리를 포함하는 결과 딕셔너리.
    """
    result = {}
    
    # 판다스 Series를 고속 처리가 가능한 NumPy 배열로 추출
    param_vals = df[param_col].values
    x_vals = df[x_col].values
    y_vals = df[y_col].values

    # 고유한 파라미터 값(Vd 혹은 Vg)을 부동소수점 오차를 무시하며 찾습니다.
    unique_params = []
    seen = set()
    for v in param_vals:
        # 소수점 6자리까지 반올림을 수행하여 파라미터의 미세한 오차 보정
        rounded = round(float(v), 6)
        if rounded not in seen:
            seen.add(rounded)
            unique_params.append(float(v))
            
    # 고유 파라미터 오름차순 정렬 (차트 범례 등 표시 순서 일관성 보장)
    unique_params.sort()

    for p in unique_params:
        # 허용 오차(1e-9) 내에서 일치하는 파라미터 인덱스에 대한 boolean 마스크 생성
        mask = np.abs(param_vals.astype(float) - p) < 1e-9
        
        # 마스크를 통해 해당 파라미터인 x, y 값만 추출
        x = x_vals[mask].astype(float)
        y = y_vals[mask].astype(float)
        
        # 쓰레기 값 / 노이즈 조건 (데이터가 너무 적음) 필터링
        # 예를 들어 Vg 101V처럼 비정상적인 분포가 단 한 개만 있거나 할 때 차트에서 배제
        if len(x) < min_points:
            continue
            
        # x축 기준 오름차순으로 데이터 순서를 정렬 (차트 선이 꼬이는 것 방지)
        sort_idx = np.argsort(x)
        result[p] = {"x": x[sort_idx], "y": y[sort_idx]}

    return result


def parse_transfer_curve(filepath: str) -> dict:
    """
    Transfer Curve(게이트 전압-드레인 전류 특성) 데이터를 파싱합니다.
    
    Args:
        filepath (str): 엑셀 또는 CSV 파일 경로
        
    Returns:
        dict: Vd를 키(key)로 하는 X(Vg), Y(Id) 쌍의 그룹 딕셔너리
    """
    is_csv = str(filepath).lower().endswith('.csv')
    
    # 1. 실제 데이터가 존재하는 첫 줄(start_row)을 감지
    start_row = _detect_data_start_row(filepath, is_csv)

    # 2. 감지된 줄 위로는 모조리 스킵하고 파일 로드
    if is_csv:
        raw = pd.read_csv(filepath, header=None, skiprows=start_row)
    else:
        raw = pd.read_excel(filepath, header=None, skiprows=start_row, engine="openpyxl")

    # 3. 열 이름을 무명 인덱스로 초기화
    raw.columns = range(len(raw.columns))
    
    # 4. A열 DataValue 조건으로 유효한 데이터 행만 필터링
    #    (이으로 새 섹션 더미 행, 매개변수 정의 행 등이 데이터로 잘못 파싱되는 것을 방지)
    valid_mask = raw.apply(
        lambda row: _is_data_row(row) if len(row) >= 4 else False, axis=1
    )
    raw = raw[valid_mask].reset_index(drop=True)
    
    # 5. 1~3번째 인덱스 열을 추출
    df = raw[[1, 2, 3]].copy()
    
    # 파싱될 임시 헤더명 설정 (Transfer Curve: B=Vg, C=Vd, D=Id)
    df.columns = ["Vg", "Vd", "Id"]
    
    # 6. 각 열을 강제로 숫자형(Numeric)으로 변환, 문자 혼입 등 변환 실패 시 NaN으로 처리
    df["Vg"] = pd.to_numeric(df["Vg"], errors="coerce")
    df["Vd"] = pd.to_numeric(df["Vd"], errors="coerce")
    df["Id"] = pd.to_numeric(df["Id"], errors="coerce")
    
    # NaN이 하나라도 존재하는 열(행)은 버림
    df = df.dropna()
    
    # 7. Vd 파라미터를 기준으로 그룹화된 딕셔너리 반환
    grouped = _group_by_param(df, param_col="Vd", x_col="Vg", y_col="Id")
    return grouped



def parse_output_curve(filepath: str) -> dict:
    """
    Output Curve(드레인 전압-드레인 전류 특성) 데이터를 파싱합니다.
    
    Args:
        filepath (str): 엑셀 또는 CSV 파일 경로
        
    Returns:
        dict: Vg를 키(key)로 하는 X(Vd), Y(Id) 쌍의 그룹 딕셔너리
    """
    is_csv = str(filepath).lower().endswith('.csv')
    
    # 1. 실제 데이터가 존재하는 첫 줄(start_row)을 감지
    start_row = _detect_data_start_row(filepath, is_csv)

    # 2. 감지된 줄 위로는 모조리 스킵하고 파일 로드
    if is_csv:
        raw = pd.read_csv(filepath, header=None, skiprows=start_row)
    else:
        raw = pd.read_excel(filepath, header=None, skiprows=start_row, engine="openpyxl")

    # 3. 열 이름을 무명 인덱스로 초기화
    raw.columns = range(len(raw.columns))
    
    # 4. A열 DataValue 조건으로 유효한 데이터 행만 필터링
    #    (이으로 새 섹션 더미 행, 매개변수 정의 행 등이 데이터로 잘못 파싱되는 것을 방지)
    valid_mask = raw.apply(
        lambda row: _is_data_row(row) if len(row) >= 4 else False, axis=1
    )
    raw = raw[valid_mask].reset_index(drop=True)
    
    # 5. Dataframe 열 추출
    df = raw[[1, 2, 3]].copy()
    
    # 파싱될 임시 헤더명 설정 (Output Curve: B=Vd, C=Vg, D=Id)
    df.columns = ["Vd", "Vg", "Id"]
    
    # 6. 숫자형 변환 및 결측치 제거
    df["Vd"] = pd.to_numeric(df["Vd"], errors="coerce")
    df["Vg"] = pd.to_numeric(df["Vg"], errors="coerce")
    df["Id"] = pd.to_numeric(df["Id"], errors="coerce")
    df = df.dropna()
    
    # 7. Vg 파라미터를 기준으로 그룹화된 딕셔너리 반환
    grouped = _group_by_param(df, param_col="Vg", x_col="Vd", y_col="Id")
    return grouped



