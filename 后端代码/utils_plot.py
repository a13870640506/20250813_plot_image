# -*- coding: utf-8 -*-
"""
utils_plot.py — 通用绘图内核（时程 & 滞回）+ 统一风格
说明：
  - 这是 MVP 内置的轻量绘图实现。后续你可无缝替换为项目中已有的
    plot_excel_timeseries.py / plot_excel_hysteresis.py 的函数。
"""

import io
import base64
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, AutoMinorLocator
from scipy.signal import savgol_filter
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime
from cycler import cycler
from matplotlib.offsetbox import AnchoredText

"""
========== 通用风格 ==========
统一设置字体/导出以保证中文与数字正常显示：
- 优先使用系统可用的中文无衬线字体；
- 关闭 Unicode 负号问题；
- PDF/PS/SVG 保持文本为可编辑文字（嵌入 TrueType），避免乱码/空白。
"""

# 常见中文字体回退列表（跨 Windows/macOS/Linux）
_FONT_FALLBACK = [
    "Microsoft YaHei", "SimHei", "SimSun",  # Windows 常见
    "PingFang SC", "Heiti SC",                # macOS 常见
    "Noto Sans CJK SC", "WenQuanYi Micro Hei", "Source Han Sans CN",  # Linux 常见
    "Arial Unicode MS", "DejaVu Sans"         # 通用补充
]

_CYCLE_COLORS = [
    "#1f77b4", "#d62728", "#2ca02c", "#9467bd", "#ff7f0e",
    "#17becf", "#7f7f7f", "#bcbd22", "#e377c2", "#8c564b",
]

mpl.rcParams.update({
    # 字体与文本
    "font.family": "sans-serif",
    "font.sans-serif": _FONT_FALLBACK,
    "axes.unicode_minus": False,
    "axes.titlepad": 10,

    # 网格与坐标轴
    "axes.grid": True,
    "grid.linestyle": "--",
    "grid.alpha": 0.5,
    "grid.linewidth": 0.8,
    "grid.color": "#bfbfbf",
    "axes.linewidth": 1.5,

    # 尺寸（可被具体函数覆盖）
    "figure.dpi": 100,
    "axes.titlesize": 22,
    "axes.labelsize": 18,
    "xtick.labelsize": 16,
    "ytick.labelsize": 16,
    "legend.fontsize": 18,

    # 导出文本保持可编辑/避免乱码
    "pdf.fonttype": 42,
    "ps.fonttype": 42,
    "svg.fonttype": "none",
    "savefig.bbox": "tight",
    # 颜色循环
    "axes.prop_cycle": cycler(color=_CYCLE_COLORS),
})

def _cm2inch(w_cm, h_cm):
    return w_cm / 2.54, h_cm / 2.54

def _apply_ticks(ax, x_major=None, y_major=None, show_minor=True):
    if x_major is not None and x_major > 0:
        ax.xaxis.set_major_locator(MultipleLocator(x_major))
    if y_major is not None and y_major > 0:
        ax.yaxis.set_major_locator(MultipleLocator(y_major))
    if show_minor:
        ax.xaxis.set_minor_locator(AutoMinorLocator())
        ax.yaxis.set_minor_locator(AutoMinorLocator())

def _apply_axes_style(ax, major_len: float = 10.0, minor_len: float = 6.0, tick_width: float = 1.2):
    """参考常用论文绘图风格，统一坐标轴与刻度样式。"""
    ax.tick_params(direction="in", which="major", top=True, right=True, length=major_len, width=tick_width)
    ax.tick_params(direction="in", which="minor", top=True, right=True, length=minor_len, width=tick_width)
    for spine in ax.spines.values():
        spine.set_linewidth(1.5)


def _format_indicator_rows(rows):
    return "\n".join(rows)


def _add_indicator_box(ax, rows, loc: str = "lower right"):
    if not rows:
        return
    txt = _format_indicator_rows(rows)
    box = AnchoredText(
        txt,
        loc=loc,
        prop=dict(size=9),
        frameon=True,
        bbox_to_anchor=(0., 0.),  # 与 loc 搭配使用，由 AnchoredText 处理
        bbox_transform=ax.transAxes,
        borderpad=0.6,
        pad=0.4,
    )
    box.patch.set_alpha(0.9)
    box.patch.set_facecolor("white")
    box.patch.set_edgecolor("#999999")
    ax.add_artist(box)

def _smooth_series(y, method=None, kwargs=None):
    if method is None or method == "none":
        return y
    kwargs = kwargs or {}
    y = np.asarray(y, dtype=float)
    if method.lower() == "savgol":
        wl = int(kwargs.get("window_length", 11))
        po = int(kwargs.get("polyorder", 3))
        wl = max(3, wl if wl % 2 == 1 else wl + 1)
        wl = min(wl, len(y) - (1 - len(y) % 2)) if len(y) > 3 else 3
        wl = max(3, wl if wl % 2 == 1 else wl - 1)
        if len(y) >= wl:
            try:
                return savgol_filter(y, wl, po, mode="interp")
            except Exception:
                return y
        return y
    elif method.lower() == "ma":  # 简单滑动平均
        k = int(kwargs.get("k", 5))
        k = max(1, k)
        if k == 1:
            return y
        kernel = np.ones(k) / k
        return np.convolve(y, kernel, mode="same")
    return y

def _encode_fig_png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    data = base64.b64encode(buf.getvalue()).decode("ascii")
    return "data:image/png;base64," + data

def _save_multi_formats(fig, save_dir, filename_base, formats, dpi=600):
    os.makedirs(save_dir, exist_ok=True)
    paths = []
    for ext in formats:
        path = os.path.join(save_dir, f"{filename_base}.{ext}")
        fig.savefig(path, dpi=dpi, bbox_inches="tight")
        paths.append(path)
    return paths

def _zip_files(filepaths, zip_out):
    with ZipFile(zip_out, "w", compression=ZIP_DEFLATED) as zf:
        for p in filepaths:
            zf.write(p, arcname=os.path.basename(p))
    return zip_out

# ========== 时程图 ==========
def plot_timeseries_preview(df, params):
    """
    必要参数：
      time_col: str
      series_cols: List[str]
    可选：
      labels, xlabel, ylabel, title, figsize_cm, dpi,
      x_major, y_major, show_minor_grid, zero_baseline,
      legend_loc, linewidth, smooth, smooth_kwargs, annotate_peaks
    """
    time_col = params.get("time_col")
    series_cols = params.get("series_cols", [])
    labels = params.get("labels") or series_cols
    figsize_cm = params.get("figsize_cm", [16, 9])
    dpi = int(params.get("dpi", 120))
    legend_loc = params.get("legend_loc", "upper right")
    linewidth = float(params.get("linewidth", 2.0))
    x_major = params.get("x_major")
    y_major = params.get("y_major")
    show_minor = bool(params.get("show_minor_grid", True))
    zero_baseline = bool(params.get("zero_baseline", True))
    smooth = params.get("smooth")  # None/"savgol"/"ma"
    smooth_kwargs = params.get("smooth_kwargs") or {}
    annotate_peaks = bool(params.get("annotate_peaks", False))
    title_pad = float(params.get("title_pad", mpl.rcParams.get("axes.titlepad", 10)))
    show_indicator = bool(params.get("show_indicator", True))

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = df[time_col].to_numpy()

    indicator_rows = []
    for i, col in enumerate(series_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        label = (labels[i] if i < len(labels) else col)
        ax.plot(x, y, lw=linewidth, label=label)

        if len(y) > 0:
            y_valid = y[~np.isnan(y)]
            if y_valid.size:
                y_max = float(np.nanmax(y_valid))
                y_min = float(np.nanmin(y_valid))
                y_peak = float(np.nanmax(np.abs(y_valid)))
                y_rms = float(np.sqrt(np.nanmean(y_valid ** 2)))
                indicator_rows.append(f"{label}: 峰值={y_peak:.4g} 最大={y_max:.4g} 最小={y_min:.4g} RMS={y_rms:.4g}")
                if annotate_peaks:
                    idx = int(np.nanargmax(np.abs(y)))
                    ax.annotate(f"|峰| {y[idx]:.4g}",
                                xy=(x[idx], y[idx]),
                                xytext=(x[idx], y[idx] * (1.08 if y[idx] >= 0 else 0.92)),
                                arrowprops=dict(arrowstyle="->", lw=1),
                                fontsize=9, ha="center")

    ax.set_xlabel(params.get("xlabel", "时间 (s)"))
    ax.set_ylabel(params.get("ylabel", "响应"))
    ax.set_title(params.get("title", "时程曲线"), pad=title_pad)
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_baseline:
        ax.axhline(0, color="k", lw=1, alpha=0.6)
    _apply_axes_style(ax)
    ax.legend(loc=legend_loc, frameon=True, framealpha=0.9, facecolor="white", edgecolor="#999")
    if show_indicator and indicator_rows:
        _add_indicator_box(ax, indicator_rows[:5], loc="lower right")
    fig.tight_layout()
    data_url = _encode_fig_png(fig)
    plt.close(fig)
    return {"ok": True, "preview_data_url": data_url}

def plot_timeseries_export(df, params):
    res_prev = plot_timeseries_preview(df, {**params, "dpi": params.get("dpi", 600)})
    # 为导出重绘（避免 base64 再解析），直接保存多格式
    time_col = params["time_col"]
    series_cols = params.get("series_cols", [])
    labels = params.get("labels") or series_cols
    figsize_cm = params.get("figsize_cm", [16, 9])
    dpi = int(params.get("dpi", 600))
    legend_loc = params.get("legend_loc", "upper right")
    linewidth = float(params.get("linewidth", 2.0))
    x_major = params.get("x_major")
    y_major = params.get("y_major")
    show_minor = bool(params.get("show_minor_grid", True))
    zero_baseline = bool(params.get("zero_baseline", True))
    smooth = params.get("smooth")
    smooth_kwargs = params.get("smooth_kwargs") or {}

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = df[time_col].to_numpy()

    for i, col in enumerate(series_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        ax.plot(x, y, lw=linewidth, label=(labels[i] if i < len(labels) else col))

    ax.set_xlabel(params.get("xlabel", "时间 (s)"))
    ax.set_ylabel(params.get("ylabel", "响应"))
    ax.set_title(params.get("title", "时程曲线"), pad=mpl.rcParams.get("axes.titlepad", 10))
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_baseline:
        ax.axhline(0, color="k", lw=1, alpha=0.6)
    _apply_axes_style(ax)
    ax.legend(loc=legend_loc, frameon=True, framealpha=0.9, facecolor="white", edgecolor="#999")
    fig.tight_layout()

    save_dir = params.get("save_dir", os.path.join("exports", datetime.now().strftime("%Y-%m-%d")))
    filename_base = params.get("filename_base", f"timeseries_{datetime.now().strftime('%H%M%S')}")
    formats = params.get("export_formats", ["png", "pdf", "svg"])
    paths = _save_multi_formats(fig, save_dir, filename_base, formats, dpi=dpi)
    plt.close(fig)

    # 生成 zip
    zip_path = os.path.join(save_dir, filename_base + ".zip")
    _zip_files(paths, zip_path)
    return {"ok": True, "files": paths, "zip": zip_path, "preview_data_url": res_prev["preview_data_url"]}

# ========== 滞回图 ==========
def plot_hysteresis_preview(df, params):
    """
    必要参数：
      x_col: str
      y_cols: List[str]
    可选：
      labels, xlabel, ylabel, title, figsize_cm, dpi,
      x_major, y_major, show_minor_grid, zero_axes, equal_aspect, legend_loc, linewidth, smooth, smooth_kwargs
    """
    x_col = params.get("x_col")
    y_cols = params.get("y_cols", [])
    labels = params.get("labels") or y_cols
    figsize_cm = params.get("figsize_cm", [16, 16])
    dpi = int(params.get("dpi", 120))
    legend_loc = params.get("legend_loc", "best")
    linewidth = float(params.get("linewidth", 2.0))
    x_major = params.get("x_major")
    y_major = params.get("y_major")
    show_minor = bool(params.get("show_minor_grid", True))
    zero_axes = bool(params.get("zero_axes", True))
    equal_aspect = bool(params.get("equal_aspect", False))
    smooth = params.get("smooth")
    smooth_kwargs = params.get("smooth_kwargs") or {}

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = pd.to_numeric(df[x_col], errors="coerce").to_numpy()

    for i, col in enumerate(y_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        ax.plot(x, y, lw=linewidth, label=(labels[i] if i < len(labels) else col))

    ax.set_xlabel(params.get("xlabel", "位移"))
    ax.set_ylabel(params.get("ylabel", "内力/反力"))
    ax.set_title(params.get("title", "滞回曲线"), pad=mpl.rcParams.get("axes.titlepad", 10))
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_axes:
        ax.axhline(0, color="k", lw=1, alpha=0.6)
        ax.axvline(0, color="k", lw=1, alpha=0.6)
    if equal_aspect:
        ax.set_aspect("equal", adjustable="box")
    _apply_axes_style(ax)
    ax.legend(loc=legend_loc, frameon=True, framealpha=0.9, facecolor="white", edgecolor="#999")
    fig.tight_layout()
    data_url = _encode_fig_png(fig)
    plt.close(fig)
    return {"ok": True, "preview_data_url": data_url}

def plot_hysteresis_export(df, params):
    res_prev = plot_hysteresis_preview(df, {**params, "dpi": params.get("dpi", 600)})
    x_col = params["x_col"]
    y_cols = params.get("y_cols", [])
    labels = params.get("labels") or y_cols
    figsize_cm = params.get("figsize_cm", [16, 16])
    dpi = int(params.get("dpi", 600))
    legend_loc = params.get("legend_loc", "best")
    linewidth = float(params.get("linewidth", 2.0))
    x_major = params.get("x_major")
    y_major = params.get("y_major")
    show_minor = bool(params.get("show_minor_grid", True))
    zero_axes = bool(params.get("zero_axes", True))
    equal_aspect = bool(params.get("equal_aspect", False))
    smooth = params.get("smooth")
    smooth_kwargs = params.get("smooth_kwargs") or {}

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = pd.to_numeric(df[x_col], errors="coerce").to_numpy()

    for i, col in enumerate(y_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        ax.plot(x, y, lw=linewidth, label=(labels[i] if i < len(labels) else col))

    ax.set_xlabel(params.get("xlabel", "位移"))
    ax.set_ylabel(params.get("ylabel", "内力/反力"))
    ax.set_title(params.get("title", "滞回曲线"), pad=mpl.rcParams.get("axes.titlepad", 10))
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_axes:
        ax.axhline(0, color="k", lw=1, alpha=0.6)
        ax.axvline(0, color="k", lw=1, alpha=0.6)
    if equal_aspect:
        ax.set_aspect("equal", adjustable="box")
    _apply_axes_style(ax)
    ax.legend(loc=legend_loc, frameon=True, framealpha=0.9, facecolor="white", edgecolor="#999")
    fig.tight_layout()

    save_dir = params.get("save_dir", os.path.join("exports", datetime.now().strftime("%Y-%m-%d")))
    filename_base = params.get("filename_base", f"hysteresis_{datetime.now().strftime('%H%M%S')}")
    formats = params.get("export_formats", ["png", "pdf", "svg"])
    paths = _save_multi_formats(fig, save_dir, filename_base, formats, dpi=dpi)
    plt.close(fig)

    zip_path = os.path.join(save_dir, filename_base + ".zip")
    _zip_files(paths, zip_path)
    return {"ok": True, "files": paths, "zip": zip_path, "preview_data_url": res_prev["preview_data_url"]}
