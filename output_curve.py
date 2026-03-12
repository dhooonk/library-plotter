"""
output_curve.py
---------------
Output Curve (Vd-Id) 차트를 생성하는 모듈.
Vg별로 그룹화된 데이터를 matplotlib으로 플롯.
"""

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.figure import Figure


# 컬러 팔레트
_COLORS = [
    "#E63946", "#2196F3", "#2ecc71", "#FF9800", "#9C27B0",
    "#00BCD4", "#FF5722", "#607D8B", "#8BC34A", "#F44336",
    "#3F51B5", "#009688", "#CDDC39", "#795548", "#FFC107",
    "#673AB7", "#03A9F4", "#4CAF50", "#FF5252", "#1DE9B6",
]


def create_output_figure(
    grouped_data: dict,
    log_scale: bool = False,
    title: str = "TFT Output Curve",
    figsize: tuple = (8, 6),
) -> Figure:
    """
    Output Curve Figure 생성.

    Parameters
    ----------
    grouped_data : dict
        {vg_value: {'x': Vd 배열, 'y': Id 배열}, ...}
    log_scale : bool
        True이면 Y축 로그 스케일
    title : str
        차트 제목
    figsize : tuple
        matplotlib 그림 크기

    Returns
    -------
    matplotlib.figure.Figure
    """
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    fig.patch.set_facecolor("#1e1e2e")
    ax.set_facecolor("#2a2a3e")

    for i, (vg, data) in enumerate(sorted(grouped_data.items())):
        color = _COLORS[i % len(_COLORS)]
        x = data["x"]
        y = data["y"]

        if log_scale:
            y_abs = np.abs(y)
            mask = y_abs > 0
            ax.semilogy(x[mask], y_abs[mask], color=color, linewidth=2,
                        label=f"Vg = {_format_val(vg)} V", marker="o",
                        markersize=3, markeredgewidth=0)
        else:
            ax.plot(x, y, color=color, linewidth=2,
                    label=f"Vg = {_format_val(vg)} V", marker="o",
                    markersize=3, markeredgewidth=0)

    _style_axes(ax, log_scale)

    ax.set_xlabel("Vds (V)", color="white", fontsize=12, labelpad=8)
    ax.set_ylabel("|Id| (A)" if log_scale else "Id (A)", color="white", fontsize=12, labelpad=8)
    ax.set_title(title, color="white", fontsize=14, fontweight="bold", pad=12)

    legend = ax.legend(
        loc="best", fontsize=9, framealpha=0.3,
        facecolor="#1e1e2e", edgecolor="#555577", labelcolor="white"
    )

    plt.tight_layout(pad=1.5)
    return fig


def _format_val(v: float) -> str:
    if v == int(v):
        return str(int(v))
    return f"{v:.3g}"


def _style_axes(ax, log_scale: bool):
    ax.tick_params(colors="white", labelsize=9)
    ax.spines["bottom"].set_color("#555577")
    ax.spines["left"].set_color("#555577")
    ax.spines["top"].set_color("#555577")
    ax.spines["right"].set_color("#555577")
    ax.grid(True, which="major", linestyle="--", alpha=0.3, color="#aaaacc")
    if log_scale:
        ax.grid(True, which="minor", linestyle=":", alpha=0.15, color="#aaaacc")
