# Lib. Plotter

> SmartSpice TFT 시뮬레이션 데이터를 자동으로 시각화·비교 분석하는 Python GUI 도구

[![버전](https://img.shields.io/badge/version-1.1.0-blue)](CHANGELOG.md)
[![라이선스](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.8+-yellow)](https://www.python.org)

---

## 목차

- [개요](#개요)
- [개발배경](#개발배경)
- [목적](#목적)
- [구현사항](#구현사항)
- [프로젝트 구조](#프로젝트-구조)
- [요구사항](#요구사항)
- [실행방법](#실행방법)
- [주요 기능](#주요-기능)
- [기술 스택](#기술-스택)
- [버전 히스토리](#버전-히스토리)
- [문의](#문의)
- [라이센스](#라이센스)

---

## 개요

TFT(박막 트랜지스터) 시뮬레이션 결과 파일(`.xlsx`, `.xls`, `.csv`)을 읽어  
**Transfer Curve(Vgs-Id)** 와 **Output Curve(Vds-Id)** 를 자동으로 플로팅하고,  
레퍼런스 데이터와 R² 비교 분석 후 결과를 엑셀 보고서로 내보냅니다.

## 개발배경

SmartSpice 시뮬레이션 결과를 매번 수작업으로 정리하고 그래프를 그리는 데 많은 시간이 소요되었습니다.  
파일 포맷이 일정하지 않아 헤더 위치나 데이터 시작 행이 달라지는 문제도 반복적으로 발생하였고,  
시뮬레이션 결과와 측정 데이터의 일치도를 정량적으로 비교할 수 있는 전용 도구가 필요했습니다.

## 목적

- 반복적인 TFT 특성 곡선 시각화 작업을 자동화
- 시뮬레이션 데이터와 측정 레퍼런스 간 R² 기반 정량 비교 지원
- 이상치 제거, 축 범위 조정 등 그래프 품질 관리 기능 제공
- 결과를 엑셀 보고서로 즉시 내보내는 워크플로 구축

## 구현사항

- 파일 내 헤더/데이터 시작 행 자동 감지 (SmartSpice `DataValue` 포맷 포함)
- Vd / Vg 파라미터별 곡선 자동 그룹화 및 오름차순 정렬
- 데이터 포인트 수 부족 그룹 자동 제외 (노이즈 필터링)
- 레퍼런스 곡선 선형 보간 후 R² 점수 산출 및 범례 표시
- 마우스 호버 시 x, y 좌표 실시간 어노테이션 표시
- 차트 우클릭으로 이상치 데이터 포인트 즉시 제거
- 분석 실행 후 실제 데이터 범위로 축 Min/Max 자동 입력
- 이상치 삭제 패널에 Idx·X·Y 미리보기 테이블 제공
- Raw Data + 차트 이미지를 포함한 엑셀 보고서 자동 생성

## 프로젝트 구조

```
tr-fitting-xls/
├── main.py              # GUI 진입점 (tkinter)
├── core/
│   ├── data_parser.py   # 파일 파싱 및 파라미터 그룹화
│   ├── transfer_curve.py# Transfer Curve 렌더링
│   └── output_curve.py  # Output Curve 렌더링
├── utils/
│   └── excel_exporter.py# 엑셀 보고서 생성
├── tests/               # 테스트 코드
├── docs/                # 문서
├── requirements.txt
└── README.md
```

## 요구사항

- Python 3.8 이상
- 아래 패키지 (상세 버전은 `requirements.txt` 참조)

| 패키지 | 용도 |
|--------|------|
| pandas | 데이터 파싱 및 가공 |
| numpy | 배열 연산 |
| matplotlib | 차트 렌더링 |
| openpyxl | 엑셀 보고서 생성 |
| scipy | 선형 보간 |
| scikit-learn | R² 점수 계산 |

## 실행방법

```bash
# 1. 가상환경 생성 및 의존성 설치
python -m venv venv
source venv/bin/activate      # macOS/Linux
# venv\Scripts\activate       # Windows

pip install -r requirements.txt

# 2. 앱 실행
python main.py
```

**사용 순서:**
1. `메인 데이터 선택` — 분석할 `.xlsx` / `.csv` 파일 선택
2. (선택) `비교 데이터(Ref) 선택` — R² 비교용 레퍼런스 파일 선택
3. `분석 실행` — 데이터 파싱 및 차트 렌더링
4. 필요 시 축 범위 조정 또는 이상치 제거 후 `🔄 차트 새로 그리기`
5. `엑셀로 저장` — Raw Data + 차트 포함 보고서 저장

## 주요 기능

- **자동 데이터 감지**: 헤더 위치와 무관하게 숫자 데이터 시작 행 자동 탐색
- **파라미터별 곡선 플로팅**: Vd / Vg 값 기준으로 곡선 자동 분리·정렬
- **레퍼런스 비교 (R²)**: 메인·레퍼런스 곡선 오버레이 및 일치도 점수 표시
- **호버 좌표 표시**: 마우스 위치의 x, y 값을 차트에 실시간 표시
- **우클릭 이상치 제거**: 차트 위에서 우클릭으로 데이터 포인트 즉시 삭제
- **축 범위 자동 입력**: 분석 실행 시 실제 데이터 범위로 Min/Max 자동 채움
- **엑셀 보고서 내보내기**: Raw Data 시트 + 고화질 차트 이미지 시트 생성

## 기술 스택

- **GUI**: Python tkinter, matplotlib (TkAgg 백엔드)
- **데이터 처리**: pandas, numpy
- **분석**: scipy (선형 보간), scikit-learn (R²)
- **보고서**: openpyxl

## 버전 히스토리

### v1.1.0
- 마우스 호버 x,y 좌표 어노테이션 추가
- 차트 우클릭 이상치 제거 기능 추가
- 분석 실행 후 축 범위 자동 입력
- 이상치 패널 Idx·X·Y 데이터 미리보기 테이블 추가
- `core/` · `utils/` 패키지 구조로 재편 (클린코드 가이드 적용)
- CSV 파일 지원 추가

### v1.0.0
- Transfer Curve / Output Curve 시각화
- 레퍼런스 R² 비교 분석
- 엑셀 보고서 내보내기

## 문의

dhooonk@lgdisplay.com

## 라이센스

이 프로젝트는 [MIT License](LICENSE) 하에 배포됩니다.
