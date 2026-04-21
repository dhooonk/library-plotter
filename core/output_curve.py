import matplotlib.pyplot as plt
import numpy as np
from matplotlib.figure import Figure
from scipy.interpolate import interp1d
from sklearn.metrics import r2_score

from core.transfer_curve import (
    COLORS, _CHART_BG, _AXES_BG, _TEXT_CLR, _SPINE_CLR,
    format_val, _style_axes, _plot_main_curve,
)


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
    fig.patch.set_facecolor(_CHART_BG)
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

    ax.set_xlabel("Vds (V)", color=_TEXT_CLR, fontsize=12, labelpad=8)
    ax.set_ylabel("|Id| (A)" if log_scale else "Id (A)", color=_TEXT_CLR, fontsize=12, labelpad=8)
    ax.set_title(title, color=_TEXT_CLR, fontsize=14, fontweight="bold", pad=12)
    ax.legend(loc="best", fontsize=9, framealpha=0.8,
              facecolor="#ffffff", edgecolor=_SPINE_CLR, labelcolor=_TEXT_CLR)
    plt.tight_layout(pad=1.5)
    return fig


def _plot_ref_curve(ax, ref_data: dict, main_x, main_y, color: str,
                    log_scale: bool, label: str) -> str:
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
                        linestyle="--", alpha=0.6, label=label)
        else:
            ax.plot(ref_data["x"], ref_data["y"], color=color, linewidth=2,
                    linestyle="--", alpha=0.6, label=label)
        return r2_str
    except Exception:
        return ""
