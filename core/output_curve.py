import matplotlib.pyplot as plt
import numpy as np
from matplotlib.figure import Figure
from scipy.interpolate import interp1d
from sklearn.metrics import r2_score

from core.transfer_curve import COLORS, _DARK_BG, _AXES_BG, format_val, _style_axes, _plot_main_curve


def create_output_figure(
    grouped_data: dict,
    log_scale: bool = False,
    title: str = "TFT Output Curve",
    figsize: tuple = (8, 6),
    xlim: tuple = None,
    ylim: tuple = None,
    ref_grouped: dict = None,
) -> Figure:
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    fig.patch.set_facecolor(_DARK_BG)
    ax.set_facecolor(_AXES_BG)

    for i, (vg, data) in enumerate(sorted(grouped_data.items())):
        color = COLORS[i % len(COLORS)]
        x = data["x"]
        y = data["y"]
        label = f"Vg = {format_val(vg)} V"
        r2_str = ""

        if ref_grouped and vg in ref_grouped:
            r2_str = _plot_ref_curve(
                ax, ref_grouped[vg], x, y, color, log_scale, f"Ref Vg = {format_val(vg)}"
            )

        _plot_main_curve(ax, x, np.abs(y) if log_scale else y, color, log_scale, label + r2_str)

    _style_axes(ax, log_scale)
    if xlim:
        ax.set_xlim(xlim)
    if ylim:
        ax.set_ylim(ylim)

    ax.set_xlabel("Vds (V)", color="white", fontsize=12, labelpad=8)
    ax.set_ylabel("|Id| (A)" if log_scale else "Id (A)", color="white", fontsize=12, labelpad=8)
    ax.set_title(title, color="white", fontsize=14, fontweight="bold", pad=12)
    ax.legend(loc="best", fontsize=9, framealpha=0.3,
              facecolor=_DARK_BG, edgecolor="#555577", labelcolor="white")
    plt.tight_layout(pad=1.5)
    return fig


def _plot_ref_curve(ax, ref_data: dict, main_x, main_y, color: str,
                    log_scale: bool, label: str) -> str:
    """Ref 곡선을 점선으로 오버레이하고 R² 값 문자열 반환. 보간 실패 시 빈 문자열."""
    try:
        f_interp = interp1d(ref_data["x"], ref_data["y"],
                            kind='linear', bounds_error=False, fill_value=np.nan)
        ref_y_interp = f_interp(main_x)
        valid_mask = (
            (~np.isnan(ref_y_interp) & (np.abs(main_y) > 0)) if log_scale
            else ~np.isnan(ref_y_interp)
        )
        r2_str = ""
        if np.sum(valid_mask) > 1:
            r2_val = r2_score(main_y[valid_mask], ref_y_interp[valid_mask])
            r2_str = f" [R²: {r2_val:.3f}]"

        if log_scale:
            ref_y_abs = np.abs(ref_data["y"])
            mask = ref_y_abs > 0
            ax.semilogy(ref_data["x"][mask], ref_y_abs[mask], color=color, linewidth=2,
                        linestyle="--", alpha=0.5, label=label)
        else:
            ax.plot(ref_data["x"], ref_data["y"], color=color, linewidth=2,
                    linestyle="--", alpha=0.5, label=label)
        return r2_str
    except Exception:
        return ""
