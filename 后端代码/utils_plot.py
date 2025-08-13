# -*- coding: utf-8 -*-
"""
utils_plot.py — 通用绘图内核（时程 & 滞回）+ 统一风格
说明：
  - 这是 MVP 内置的轻量绘图实现。后续你可无缝替换为项目中已有的
    plot_excel_timeseries.py / plot_excel_hysteresis.py 的函数。
  - 已修复峰值标注、指标框、刻度线样式，并添加坐标轴范围控制
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
    # 字体与文本（参照参考代码的设置）
    "font.family": "sans-serif",
    "font.sans-serif": _FONT_FALLBACK,
    "font.size": 10.5,  # 基础字体大小，与参考代码一致
    "axes.unicode_minus": False,
    "axes.linewidth": 1.0,  # 与参考代码一致

    # 网格与坐标轴（参照参考代码）
    "grid.linestyle": "--",
    "grid.linewidth": 0.8,
    "grid.alpha": 0.5,
    "figure.constrained_layout.use": True,
    "savefig.dpi": 300,

    # 让标注盒子观感更统一
    "patch.linewidth": 0.8,

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

def _normalize_lim(lim, dmin=None, dmax=None):
    """
    将 xlim/ylim 规格化为二元组 (low, high)。
    - lim = None -> None（不设置，走默认）
    - lim = (a, b) / [a, b]；任一端为 None 则用数据端点替代
    """
    if lim is None:
        return None
    if isinstance(lim, (list, tuple)) and len(lim) == 2:
        lo, hi = lim
        if lo is None:
            lo = dmin
        if hi is None:
            hi = dmax
        return (float(lo), float(hi))
    raise ValueError("xlim/ylim 应为长度为2的 (min, max) 或 None")


def _apply_ticks(ax, x_major=None, y_major=None, show_minor=True):
    # 主网格设置（与参考代码一致）
    ax.grid(True)
    if show_minor:
        ax.xaxis.set_minor_locator(AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(AutoMinorLocator(2))
        ax.grid(True, which="minor", alpha=0.25, linestyle="--", linewidth=0.6)
    if x_major is not None and x_major > 0:
        ax.xaxis.set_major_locator(MultipleLocator(x_major))
    if y_major is not None and y_major > 0:
        ax.yaxis.set_major_locator(MultipleLocator(y_major))

def _apply_axes_style(ax, major_len: float = 6.0, minor_len: float = 3.5, tick_width: float = 1.2):
    """统一坐标轴与刻度样式：刻度朝外、仅下/左、较短。"""
    # 仅在下/左显示刻度，方向朝外，符合参考图风格
    ax.tick_params(direction="out", which="major", top=False, right=False,
                   length=major_len, width=tick_width)
    ax.tick_params(direction="out", which="minor", top=False, right=False,
                   length=minor_len, width=tick_width)
    for spine in ax.spines.values():
        spine.set_linewidth(1.5)


# ================= 峰值标注（防重叠升级版） =================
def _bboxes_overlap(bb1, bb2, pad=2.0):
    """判断两个 bbox 是否重叠，pad 为像素级扩展边距"""
    x11, y11, x12, y12 = bb1.x0 - pad, bb1.y0 - pad, bb1.x1 + pad, bb1.y1 + pad
    x21, y21, x22, y22 = bb2.x0 - pad, bb2.y0 - pad, bb2.x1 + pad, bb2.y1 + pad
    return not (x12 < x21 or x22 < x11 or y12 < y21 or y22 < y11)


def _place_annot_no_overlap(ax, xy, text, color=None, fontsize=None,
                            candidates=None, used_bboxes=None):
    """
    尝试多组偏移(单位: offset points)放置注释，若与已放置注释 bbox 重叠则换下一组。
    成功返回(annot, bbox)，失败返回(None, None)。
    """
    if candidates is None:
        # 优先上下左右，再斜向，尽量离得近一些
        candidates = [(6, 10), (0, 14), (0, -18), (14, 0), (-18, 0),
                      (10, 12), (-12, 10), (12, -10), (-10, -12), (22, 6), (-22, 6)]
    if used_bboxes is None:
        used_bboxes = []

    # 先画一个点增强可读性
    ax.scatter([xy[0]], [xy[1]], s=28, zorder=5, color=color)

    # 逐个候选尝试
    for dx, dy in candidates:
        ann = ax.annotate(
            text, xy=xy, xycoords="data",
            xytext=(dx, dy), textcoords="offset points",
            ha="left", va="bottom",
            arrowprops=dict(arrowstyle="->", lw=0.8, alpha=0.9, color=color),
            bbox=dict(facecolor="white", alpha=0.85, edgecolor="#444", boxstyle="round,pad=0.25"),
            fontsize=fontsize if fontsize is not None else max(9, int(plt.rcParams["font.size"])) - 1,
            color=color
        )
        # 需要先 draw 才能拿到 bbox
        ann.figure.canvas.draw()
        bb = ann.get_window_extent(ann.figure.canvas.get_renderer())

        # 与已放置 bbox 做碰撞检测
        conflict = any(_bboxes_overlap(bb, ub) for ub in used_bboxes)
        if not conflict:
            used_bboxes.append(bb)
            return ann, bb
        # 冲突则移除，试下一个候选
        ann.remove()

    # 都冲突则保底放一个最初偏移
    ann = ax.annotate(
        text, xy=xy, xycoords="data",
        xytext=(6, 10), textcoords="offset points",
        ha="left", va="bottom",
        arrowprops=dict(arrowstyle="->", lw=0.8, alpha=0.9, color=color),
        bbox=dict(facecolor="white", alpha=0.85, edgecolor="#444", boxstyle="round,pad=0.25"),
        fontsize=fontsize if fontsize is not None else max(9, int(plt.rcParams["font.size"])) - 1,
        color=color
    )
    ann.figure.canvas.draw()
    bb = ann.get_window_extent(ann.figure.canvas.get_renderer())
    used_bboxes.append(bb)
    return ann, bb


def _annotate_absmax(ax, x, y, label=None, fmt="{:.1f}",
                     used_bboxes=None, line_color=None):
    """
    自动错位避免文字相互遮挡（基于bbox碰撞检测）。
    used_bboxes: list 存已放置文本的窗口坐标 bbox（像素）。
    """
    idx = int(np.nanargmax(np.abs(y)))
    xv, yv = x[idx], y[idx]
    s = (label + "：") if label else ""
    s += fmt.format(yv)
    _place_annot_no_overlap(
        ax, (xv, yv), s,
        color=line_color, used_bboxes=used_bboxes
    )


# _format_indicator_rows函数已删除


# ================= 插件：两曲线峰值对比（带边框的锚定文本框） =================
def _add_peak_metrics_box(ax, y_ref, y_cmp, loc="lower right"):
    """添加峰值对比框（参考参考代码的样式）"""
    def peak(a):
        return float(np.nanmax(np.abs(a)))

    p1, p2 = peak(y_ref), peak(y_cmp)
    peak_drop = np.nan
    if p1 > 1e-12:
        peak_drop = 100.0 * (1.0 - p2 / p1)
    text = (f"峰值参考 = {p1:.1f}\n"
            f"峰值对比 = {p2:.1f}\n"
            f"峰值降低 = {peak_drop:.1f}%")
    at = AnchoredText(
        text, loc={"lower right": 4, "lower left": 3, "upper right": 1, "upper left": 2}.get(loc, 4),
        prop=dict(size=max(9, int(plt.rcParams["font.size"])) - 1),
        frameon=True, borderpad=0.4, pad=0.3
    )
    # 统一的外观
    at.patch.set_alpha(0.9)
    at.patch.set_edgecolor("#444")
    at.patch.set_facecolor("white")
    ax.add_artist(at)


# 传统指标框功能已删除

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
      xlim, ylim, metrics_box
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
    xlim = params.get("xlim")
    ylim = params.get("ylim")
    metrics_box = bool(params.get("metrics_box", False))

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = df[time_col].to_numpy()

    plotted = []
    y_all_min, y_all_max = np.inf, -np.inf
    
    for i, col in enumerate(series_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        label = (labels[i] if i < len(labels) else col)
        line, = ax.plot(x, y, lw=linewidth, label=label)
        plotted.append((label, x, y, line))
        
        # 统计全局 y 范围
        if len(y) > 0:
            y_valid = y[~np.isnan(y)]
            if y_valid.size:
                y_all_min = min(y_all_min, float(np.nanmin(y_valid)))
                y_all_max = max(y_all_max, float(np.nanmax(y_valid)))

    ax.set_xlabel(params.get("xlabel", "时间 (s)"))
    ax.set_ylabel(params.get("ylabel", "响应"))
    ax.set_title(params.get("title", "时程曲线"), pad=title_pad)
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_baseline:
        ax.axhline(0, color="0.25", lw=1.0, alpha=0.9)
    _apply_axes_style(ax)
    
    # 横坐标紧贴图框：默认锁到数据端点
    ax.margins(x=0.0)
    try:
        default_xlim = (float(np.nanmin(x)), float(np.nanmax(x)))
        ax.set_xlim(*default_xlim)
    except Exception:
        default_xlim = None

    # 应用用户指定的坐标轴范围
    try:
        xlim_norm = _normalize_lim(xlim, dmin=default_xlim[0] if default_xlim else None,
                                   dmax=default_xlim[1] if default_xlim else None)
        if xlim_norm is not None:
            ax.set_xlim(*xlim_norm)
    except Exception:
        pass

    try:
        default_ylim = (y_all_min, y_all_max) if np.isfinite(y_all_min) and np.isfinite(y_all_max) else None
        ylim_norm = _normalize_lim(ylim, dmin=default_ylim[0] if default_ylim else None,
                                   dmax=default_ylim[1] if default_ylim else None)
        if ylim_norm is not None:
            ax.set_ylim(*ylim_norm)
    except Exception:
        pass

    # 峰值标注（基于bbox防重叠）
    if annotate_peaks:
        used_bboxes = []
        # 先 draw 一次，确保坐标系确定
        fig.canvas.draw()
        for (lab, x_data, y_data, line) in plotted:
            _annotate_absmax(ax, x_data, y_data, label=lab, used_bboxes=used_bboxes, line_color=line.get_color())

    # 峰值对比文本框（仅两条曲线生效）
    if metrics_box and len(plotted) >= 2:
        _add_peak_metrics_box(ax, plotted[0][2], plotted[1][2], loc="lower right")


    ax.legend(loc=legend_loc, framealpha=0.95, fancybox=True, edgecolor="#444")
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
    annotate_peaks = bool(params.get("annotate_peaks", False))
    title_pad = float(params.get("title_pad", mpl.rcParams.get("axes.titlepad", 10)))

    xlim = params.get("xlim")
    ylim = params.get("ylim")
    metrics_box = bool(params.get("metrics_box", False))

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = df[time_col].to_numpy()

    plotted = []

    y_all_min, y_all_max = np.inf, -np.inf

    for i, col in enumerate(series_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        label = (labels[i] if i < len(labels) else col)
        line, = ax.plot(x, y, lw=linewidth, label=label)
        plotted.append((label, x, y, line))
        
        # 统计全局 y 范围和指标信息
        if len(y) > 0:
            y_valid = y[~np.isnan(y)]
            if y_valid.size:
                y_all_min = min(y_all_min, float(np.nanmin(y_valid)))
                y_all_max = max(y_all_max, float(np.nanmax(y_valid)))
                y_max = float(np.nanmax(y_valid))
                y_min = float(np.nanmin(y_valid))
                y_peak = float(np.nanmax(np.abs(y_valid)))
                y_rms = float(np.sqrt(np.nanmean(y_valid ** 2)))


    ax.set_xlabel(params.get("xlabel", "时间 (s)"))
    ax.set_ylabel(params.get("ylabel", "响应"))
    ax.set_title(params.get("title", "时程曲线"), pad=title_pad)
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_baseline:
        ax.axhline(0, color="0.25", lw=1.0, alpha=0.9)
    _apply_axes_style(ax)
    
    # 横坐标紧贴图框
    ax.margins(x=0.0)
    try:
        default_xlim = (float(np.nanmin(x)), float(np.nanmax(x)))
        ax.set_xlim(*default_xlim)
    except Exception:
        default_xlim = None

    # 应用用户指定的坐标轴范围
    try:
        xlim_norm = _normalize_lim(xlim, dmin=default_xlim[0] if default_xlim else None,
                                   dmax=default_xlim[1] if default_xlim else None)
        if xlim_norm is not None:
            ax.set_xlim(*xlim_norm)
    except Exception:
        pass

    try:
        default_ylim = (y_all_min, y_all_max) if np.isfinite(y_all_min) and np.isfinite(y_all_max) else None
        ylim_norm = _normalize_lim(ylim, dmin=default_ylim[0] if default_ylim else None,
                                   dmax=default_ylim[1] if default_ylim else None)
        if ylim_norm is not None:
            ax.set_ylim(*ylim_norm)
    except Exception:
        pass

    # 峰值标注（基于bbox防重叠）
    if annotate_peaks:
        used_bboxes = []
        fig.canvas.draw()
        for (lab, x_data, y_data, line) in plotted:
            _annotate_absmax(ax, x_data, y_data, label=lab, used_bboxes=used_bboxes, line_color=line.get_color())

    # 峰值对比文本框（仅两条曲线生效）
    if metrics_box and len(plotted) >= 2:
        _add_peak_metrics_box(ax, plotted[0][2], plotted[1][2], loc="lower right")


    ax.legend(loc=legend_loc, framealpha=0.95, fancybox=True, edgecolor="#444")
    fig.tight_layout()

    save_dir = params.get("save_dir", os.path.join("exports", datetime.now().strftime("%Y-%m-%d")))
    filename_base = params.get("filename_base", f"timeseries_{datetime.now().strftime('%Y-%m-%d%H%M%S')}")
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
      xlim, ylim
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
    title_pad = float(params.get("title_pad", mpl.rcParams.get("axes.titlepad", 10)))
    xlim = params.get("xlim")
    ylim = params.get("ylim")

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = pd.to_numeric(df[x_col], errors="coerce").to_numpy()

    x_all_min, x_all_max = np.inf, -np.inf
    y_all_min, y_all_max = np.inf, -np.inf

    for i, col in enumerate(y_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        ax.plot(x, y, lw=linewidth, label=(labels[i] if i < len(labels) else col))
        
        # 统计范围
        if len(x) > 0:
            x_valid = x[~np.isnan(x)]
            if x_valid.size:
                x_all_min = min(x_all_min, float(np.nanmin(x_valid)))
                x_all_max = max(x_all_max, float(np.nanmax(x_valid)))
        if len(y) > 0:
            y_valid = y[~np.isnan(y)]
            if y_valid.size:
                y_all_min = min(y_all_min, float(np.nanmin(y_valid)))
                y_all_max = max(y_all_max, float(np.nanmax(y_valid)))

    ax.set_xlabel(params.get("xlabel", "位移"))
    ax.set_ylabel(params.get("ylabel", "内力/反力"))
    ax.set_title(params.get("title", "滞回曲线"), pad=title_pad)
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_axes:
        ax.axhline(0, color="0.25", lw=1.0, alpha=0.9)
        ax.axvline(0, color="0.25", lw=1.0, alpha=0.9)
    if equal_aspect:
        ax.set_aspect("equal", adjustable="box")
    _apply_axes_style(ax)

    # 应用用户指定的坐标轴范围
    try:
        default_xlim = (x_all_min, x_all_max) if np.isfinite(x_all_min) and np.isfinite(x_all_max) else None
        xlim_norm = _normalize_lim(xlim, dmin=default_xlim[0] if default_xlim else None,
                                   dmax=default_xlim[1] if default_xlim else None)
        if xlim_norm is not None:
            ax.set_xlim(*xlim_norm)
    except Exception:
        pass

    try:
        default_ylim = (y_all_min, y_all_max) if np.isfinite(y_all_min) and np.isfinite(y_all_max) else None
        ylim_norm = _normalize_lim(ylim, dmin=default_ylim[0] if default_ylim else None,
                                   dmax=default_ylim[1] if default_ylim else None)
        if ylim_norm is not None:
            ax.set_ylim(*ylim_norm)
    except Exception:
        pass

    ax.legend(loc=legend_loc, framealpha=0.95, fancybox=True, edgecolor="#444")
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
    title_pad = float(params.get("title_pad", mpl.rcParams.get("axes.titlepad", 10)))
    xlim = params.get("xlim")
    ylim = params.get("ylim")

    fig, ax = plt.subplots(figsize=_cm2inch(*figsize_cm), dpi=dpi)
    x = pd.to_numeric(df[x_col], errors="coerce").to_numpy()

    x_all_min, x_all_max = np.inf, -np.inf
    y_all_min, y_all_max = np.inf, -np.inf

    for i, col in enumerate(y_cols):
        y_raw = pd.to_numeric(df[col], errors="coerce").to_numpy()
        y = _smooth_series(y_raw, smooth, smooth_kwargs)
        ax.plot(x, y, lw=linewidth, label=(labels[i] if i < len(labels) else col))
        
        # 统计范围
        if len(x) > 0:
            x_valid = x[~np.isnan(x)]
            if x_valid.size:
                x_all_min = min(x_all_min, float(np.nanmin(x_valid)))
                x_all_max = max(x_all_max, float(np.nanmax(x_valid)))
        if len(y) > 0:
            y_valid = y[~np.isnan(y)]
            if y_valid.size:
                y_all_min = min(y_all_min, float(np.nanmin(y_valid)))
                y_all_max = max(y_all_max, float(np.nanmax(y_valid)))

    ax.set_xlabel(params.get("xlabel", "位移"))
    ax.set_ylabel(params.get("ylabel", "内力/反力"))
    ax.set_title(params.get("title", "滞回曲线"), pad=title_pad)
    _apply_ticks(ax, x_major, y_major, show_minor)
    if zero_axes:
        ax.axhline(0, color="0.25", lw=1.0, alpha=0.9)
        ax.axvline(0, color="0.25", lw=1.0, alpha=0.9)
    if equal_aspect:
        ax.set_aspect("equal", adjustable="box")
    _apply_axes_style(ax)

    # 应用用户指定的坐标轴范围
    try:
        default_xlim = (x_all_min, x_all_max) if np.isfinite(x_all_min) and np.isfinite(x_all_max) else None
        xlim_norm = _normalize_lim(xlim, dmin=default_xlim[0] if default_xlim else None,
                                   dmax=default_xlim[1] if default_xlim else None)
        if xlim_norm is not None:
            ax.set_xlim(*xlim_norm)
    except Exception:
        pass

    try:
        default_ylim = (y_all_min, y_all_max) if np.isfinite(y_all_min) and np.isfinite(y_all_max) else None
        ylim_norm = _normalize_lim(ylim, dmin=default_ylim[0] if default_ylim else None,
                                   dmax=default_ylim[1] if default_ylim else None)
        if ylim_norm is not None:
            ax.set_ylim(*ylim_norm)
    except Exception:
        pass

    ax.legend(loc=legend_loc, framealpha=0.95, fancybox=True, edgecolor="#444")
    fig.tight_layout()

    save_dir = params.get("save_dir", os.path.join("exports", datetime.now().strftime("%Y-%m-%d")))
    filename_base = params.get("filename_base", f"hysteresis_{datetime.now().strftime('%Y-%m-%d%H%M%S')}")
    formats = params.get("export_formats", ["png", "pdf", "svg"])
    paths = _save_multi_formats(fig, save_dir, filename_base, formats, dpi=dpi)
    plt.close(fig)

    zip_path = os.path.join(save_dir, filename_base + ".zip")
    _zip_files(paths, zip_path)
    return {"ok": True, "files": paths, "zip": zip_path, "preview_data_url": res_prev["preview_data_url"]}
