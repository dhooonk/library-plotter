"""
transfer_curve.py
-----------------
Transfer Curve (Vgs-Id) 차트를 생성하는 모듈.
Vd별로 그룹화된 데이터를 matplotlib으로 플롯.
"""

import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
from matplotlib.figure import Figure
from scipy.interpolate import interp1d
from sklearn.metrics import r2_score


# 컬러 팔레트 (최대 20개 Vd 값 지원)
_COLORS = [
    "#E63946", "#2196F3", "#2ecc71", "#FF9800", "#9C27B0",
    "#00BCD4", "#FF5722", "#607D8B", "#8BC34A", "#F44336",
    "#3F51B5", "#009688", "#CDDC39", "#795548", "#FFC107",
    "#673AB7", "#03A9F4", "#4CAF50", "#FF5252", "#1DE9B6",
]


def create_transfer_figure(
    grouped_data: dict,
    log_scale: bool = True,
    title: str = "TFT Transfer Curve",
    figsize: tuple = (8, 6),
    xlim: tuple = None,
    ylim: tuple = None,
    ref_grouped: dict = None
) -> Figure:
    """
    Transfer Curve Figure 생성 (비교 대상 및 R² 계산 포함).
    """
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    fig.patch.set_facecolor("#1e1e2e")
    ax.set_facecolor("#2a2a3e")

    for i, (vd, data) in enumerate(sorted(grouped_data.items())):
        color = _COLORS[i % len(_COLORS)]
        x = data["x"]
        y = np.abs(data["y"])  # Id는 절댓값 사용 (음수 가능)

        label_base = f"Vd = {_format_val(vd)} V"
        r2_str = ""
        
        # ── 비교 대상 데이터가 있을 경우 R² 계산 플로팅 ──
        if ref_grouped and vd in ref_grouped:
            ref_x = ref_grouped[vd]["x"]
            ref_y = np.abs(ref_grouped[vd]["y"])
            
            # 메인 x에 맞추어 Ref y 보간
            # 범위 벗어나는 값은 외삽하지 않고 nan 처리하거나 무시할 수 있으므로 
            # bounds_error=False, fill_value=np.nan 적용
            try:
                f_interp = interp1d(ref_x, ref_y, kind='linear', bounds_error=False, fill_value=np.nan)
                ref_y_interp = f_interp(x)
                
                # nan 없는 유효 범위 내에서만 R2 계산
                valid_mask = ~np.isnan(ref_y_interp) & (y > 0) if log_scale else ~np.isnan(ref_y_interp)
                if np.sum(valid_mask) > 1:
                    r2_val = r2_score(y[valid_mask], ref_y_interp[valid_mask])
                    r2_str = f" [R²: {r2_val:.3f}]"
                    
                # Ref 데이터 오버레이(점선/연한 색 등)
                if log_scale:
                    ref_mask = ref_y > 0
                    ax.semilogy(ref_x[ref_mask], ref_y[ref_mask], color=color, linewidth=2,
                                linestyle="--", alpha=0.5, label=f"Ref Vd = {_format_val(vd)}")
                else:
                    ax.plot(ref_x, ref_y, color=color, linewidth=2, linestyle="--", 
                            alpha=0.5, label=f"Ref Vd = {_format_val(vd)}")
            except Exception:
                pass

        if log_scale:
            # 0 이하 값 제거 (로그 스케일 호환)
            mask = y > 0
            ax.semilogy(x[mask], y[mask], color=color, linewidth=2,
                        label=label_base + r2_str, marker="o",
                        markersize=3, markeredgewidth=0)
        else:
            ax.plot(x, y, color=color, linewidth=2,
                    label=label_base + r2_str, marker="o",
                    markersize=3, markeredgewidth=0)

    # 축 스타일
    _style_axes(ax, log_scale)

    if xlim is not None:
        ax.set_xlim(xlim)
    if ylim is not None:
        ax.set_ylim(ylim)

    ax.set_xlabel("Vgs (V)", color="white", fontsize=12, labelpad=8)
    ax.set_ylabel("|Id| (A)" if log_scale else "Id (A)", color="white", fontsize=12, labelpad=8)
    ax.set_title(title, color="white", fontsize=14, fontweight="bold", pad=12)

    # 범례
    legend = ax.legend(
        loc="best", fontsize=9, framealpha=0.3,
        facecolor="#1e1e2e", edgecolor="#555577", labelcolor="white"
    )

    plt.tight_layout(pad=1.5)
    return fig


def _format_val(v: float) -> str:
    """소수점 불필요한 경우 정수로 표시."""
    if v == int(v):
        return str(int(v))
    return f"{v:.3g}"


def _style_axes(ax, log_scale: bool):
    """다크 테마 축 스타일 적용."""
    ax.tick_params(colors="white", labelsize=9)
    ax.spines["bottom"].set_color("#555577")
    ax.spines["left"].set_color("#555577")
    ax.spines["top"].set_color("#555577")
    ax.spines["right"].set_color("#555577")
    ax.grid(True, which="major", linestyle="--", alpha=0.3, color="#aaaacc")
    if log_scale:
        ax.grid(True, which="minor", linestyle=":", alpha=0.15, color="#aaaacc")

