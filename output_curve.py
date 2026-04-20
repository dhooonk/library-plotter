"""
output_curve.py
---------------
Output Curve (Vd-Id) 차트를 생성하는 모듈입니다.
Vg(게이트 전압)별로 그룹화된 데이터를 matplotlib으로 플로팅하며,
비교 대상 데이터(Ref)가 있을 경우 동일 파라미터에 대한 곡선을 병치하고 
$R^2$ (R-squared) 일치율 점수를 산출하여 시각화합니다.
"""

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.figure import Figure
from scipy.interpolate import interp1d
from sklearn.metrics import r2_score

from transfer_curve import _COLORS, _format_val, _style_axes


def create_output_figure(
    grouped_data: dict,
    log_scale: bool = False,
    title: str = "TFT Output Curve",
    figsize: tuple = (8, 6),
    xlim: tuple = None,
    ylim: tuple = None,
    ref_grouped: dict = None
) -> Figure:
    """
    Output Curve를 나타내는 matplotlib Figure 객체를 생성합니다.
    
    Args:
        grouped_data (dict): Vg를 키로 가지며 x(Vd), y(Id) 넘파이 배열을 지닌 기본 파싱 데이터
        log_scale (bool): Y축을 로그 스케일로 표시할지 여부 (기본값 False)
        title (str): 차트 제목
        figsize (tuple): 생성할 Figure의 가로, 세로 크기 설정 (기본 8x6)
        xlim (tuple): X축 최솟값/최댓값 (min, max). None이면 Auto 스케일
        ylim (tuple): Y축 최솟값/최댓값 (min, max). None이면 Auto 스케일
        ref_grouped (dict): 비교 대상(Ref) 데이터 딕셔너리. None이면 사용 안 함.
        
    Returns:
        Figure: 렌더링이 완료된 matplotlib Figure 객체
    """
    # 1. Figure 및 Subplot(Axes) 생성
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    
    # 2. 다크 배경 테마 적용
    fig.patch.set_facecolor("#1e1e2e")
    ax.set_facecolor("#2a2a3e")

    # 3. 데이터 순회하며 플로팅 (Vg 값이 작은 순서대로 정렬)
    for i, (vg, data) in enumerate(sorted(grouped_data.items())):
        color = _COLORS[i % len(_COLORS)]
        x = data["x"]
        y = data["y"]

        # 범례에 표시할 기본 Vg 레이블 텍스트
        label_base = f"Vg = {_format_val(vg)} V"
        r2_str = ""
        
        # ── 4. 비교 대상 데이터 플로팅 및 R-squared 점수 계산 ──
        if ref_grouped and vg in ref_grouped:
            ref_x = ref_grouped[vg]["x"]
            ref_y = ref_grouped[vg]["y"]
            
            # 메인 X축 범위에 맞추어 Ref 데이터 Y축 값을 선형 보간(Linear Interpolation) 적용
            try:
                f_interp = interp1d(ref_x, ref_y, kind='linear', bounds_error=False, fill_value=np.nan)
                ref_y_interp = f_interp(x)
                
                # NaN 값 및 (로그 스케일 시) 0 이하인 값을 제외한 유효 마스크 확보
                if log_scale:
                    valid_mask = ~np.isnan(ref_y_interp) & (np.abs(y) > 0)
                else:
                    valid_mask = ~np.isnan(ref_y_interp)
                
                # 유효 데이터 구간이 2포인트 이상일 때 $R^2$ 점수 도출
                if np.sum(valid_mask) > 1:
                    r2_val = r2_score(y[valid_mask], ref_y_interp[valid_mask])
                    r2_str = f" [R²: {r2_val:.3f}]"

                # Ref 데이터 오버레이(점선 형태의 반투명 두께 적용)
                if log_scale:
                    ref_y_abs = np.abs(ref_y)
                    ref_mask = ref_y_abs > 0
                    ax.semilogy(ref_x[ref_mask], ref_y_abs[ref_mask], color=color, linewidth=2,
                                linestyle="--", alpha=0.5, label=f"Ref Vg = {_format_val(vg)}")
                else:
                    ax.plot(ref_x, ref_y, color=color, linewidth=2, linestyle="--", 
                            alpha=0.5, label=f"Ref Vg = {_format_val(vg)}")
            except Exception:
                # 보간 에러 발생 시 부드럽게 무시하고 진행
                pass


        # ── 5. 메인 데이터 플로팅 ──
        if log_scale:
            # Output Curve에서 로그 스케일 시 전류의 절댓값 표기를 허용 (마이너스 제외)
            y_abs = np.abs(y)
            mask = y_abs > 0
            ax.semilogy(x[mask], y_abs[mask], color=color, linewidth=2,
                        label=label_base + r2_str, marker="o",
                        markersize=3, markeredgewidth=0)
        else:
            # 일반 선형 스케일 렌더
            ax.plot(x, y, color=color, linewidth=2,
                    label=label_base + r2_str, marker="o",
                    markersize=3, markeredgewidth=0)

    # 6. 축 스타일 및 커스텀 범위 지정
    _style_axes(ax, log_scale)

    if xlim is not None:
        ax.set_xlim(xlim)
    if ylim is not None:
        ax.set_ylim(ylim)

    ax.set_xlabel("Vds (V)", color="white", fontsize=12, labelpad=8)
    ax.set_ylabel("|Id| (A)" if log_scale else "Id (A)", color="white", fontsize=12, labelpad=8)
    ax.set_title(title, color="white", fontsize=14, fontweight="bold", pad=12)

    # 7. 범례 박스 스타일 적용
    legend = ax.legend(
        loc="best", fontsize=9, framealpha=0.3,
        facecolor="#1e1e2e", edgecolor="#555577", labelcolor="white"
    )

    plt.tight_layout(pad=1.5)
    return fig


