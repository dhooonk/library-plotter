import matplotlib.pyplot as plt
import numpy as np
from matplotlib.figure import Figure
from scipy.interpolate import interp1d
from sklearn.metrics import r2_score

COLORS = [
    "#E63946", "#2196F3", "#2ecc71", "#FF9800", "#9C27B0",
    "#00BCD4", "#FF5722", "#607D8B", "#8BC34A", "#F44336",
    "#3F51B5", "#009688", "#CDDC39", "#795548", "#FFC107",
    "#673AB7", "#03A9F4", "#4CAF50", "#FF5252", "#1DE9B6",
]

_CHART_BG = "#ffffff"
_AXES_BG  = "#fafafa"
_TEXT_CLR = "#1e1e1e"
_GRID_CLR = "#cccccc"
_SPINE_CLR = "#bbbbbb"


def create_transfer_figure(
    grouped_data: dict,
    log_scale: bool = True,
    title: str = "TFT Transfer Curve",
    figsize: tuple = (8, 6),
    xlim: tuple = None,
    ylim: tuple = None,
    ref_grouped: dict = None,
) -> Figure:
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    fig.patch.set_facecolor(_CHART_BG)
    ax.set_facecolor(_AXES_BG)

    for i, (vd, data) in enumerate(sorted(grouped_data.items())):
        color = COLORS[i % len(COLORS)]
        x = data["x"]
        y = np.abs(data["y"])
        label = f"Vd = {format_val(vd)} V"
        r2_str = ""

        if ref_grouped and vd in ref_grouped:
            r2_str = _plot_ref_curve(
                ax, ref_grouped[vd], x, y, color, log_scale, f"Ref Vd = {format_val(vd)}"
            )

        _plot_main_curve(ax, x, y, color, log_scale, label + r2_str)

    _style_axes(ax, log_scale)
    if xlim:
        ax.set_xlim(xlim)
    if ylim:
        ax.set_ylim(ylim)

    ax.set_xlabel("Vgs (V)", color=_TEXT_CLR, fontsize=12, labelpad=8)
    ax.set_ylabel("|Id| (A)" if log_scale else "Id (A)", color=_TEXT_CLR, fontsize=12, labelpad=8)
    ax.set_title(title, color=_TEXT_CLR, fontsize=14, fontweight="bold", pad=12)
    ax.legend(loc="best", fontsize=9, framealpha=0.8,
              facecolor="#ffffff", edgecolor=_SPINE_CLR, labelcolor=_TEXT_CLR)
    plt.tight_layout(pad=1.5)
    return fig


def _plot_ref_curve(ax, ref_data: dict, main_x, main_y, color: str,
                    log_scale: bool, label: str) -> str:
    try:
        f_interp = interp1d(ref_data["x"], np.abs(ref_data["y"]),
                            kind='linear', bounds_error=False, fill_value=np.nan)
        ref_y_interp = f_interp(main_x)
        valid_mask = (
            (~np.isnan(ref_y_interp) & (main_y > 0)) if log_scale
            else ~np.isnan(ref_y_interp)
        )
        r2_str = ""
        if np.sum(valid_mask) > 1:
            r2_val = r2_score(main_y[valid_mask], ref_y_interp[valid_mask])
            r2_str = f" [R²: {r2_val:.3f}]"

        ref_y = np.abs(ref_data["y"])
        if log_scale:
            mask = ref_y > 0
            ax.semilogy(ref_data["x"][mask], ref_y[mask], color=color, linewidth=2,
                        linestyle="--", alpha=0.6, label=label)
        else:
            ax.plot(ref_data["x"], ref_y, color=color, linewidth=2,
                    linestyle="--", alpha=0.6, label=label)
        return r2_str
    except Exception:
        return ""


def _plot_main_curve(ax, x, y, color: str, log_scale: bool, label: str):
    if log_scale:
        mask = y > 0
        ax.semilogy(x[mask], y[mask], color=color, linewidth=2, label=label,
                    marker="o", markersize=3, markeredgewidth=0)
    else:
        ax.plot(x, y, color=color, linewidth=2, label=label,
                marker="o", markersize=3, markeredgewidth=0)


def format_val(v: float) -> str:
    if v == int(v):
        return str(int(v))
    return f"{v:.3g}"


def _style_axes(ax, log_scale: bool):
    ax.tick_params(colors=_TEXT_CLR, labelsize=9)
    for spine in ax.spines.values():
        spine.set_color(_SPINE_CLR)
    ax.grid(True, which="major", linestyle="--", alpha=0.5, color=_GRID_CLR)
    if log_scale:
        ax.grid(True, which="minor", linestyle=":", alpha=0.25, color=_GRID_CLR)
