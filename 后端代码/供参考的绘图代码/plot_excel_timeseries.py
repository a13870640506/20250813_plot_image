# -*- coding: utf-8 -*-
"""
plot_excel_timeseries.py — Excel时程绘图（统一风格 + 可插拔增强）
改进点：
  1) 命令行示例每行含行内注释  （⚠ 现已移除命令行，仅保留通用函数）
  2) 支持一次导出多种格式：export_formats=["png","pdf","svg"]
     若既不传 outfile 也不传 export_formats，则默认导出 png+pdf
  3) 峰值标注自动错位防重叠（基于bbox碰撞检测，比原版更稳）
  4) 横坐标紧贴图框（取消左右留白）
  5) 右下角说明改为统一带边框的锚定文本框
  6) ✅ 新增 save_dir 与 filename_base 用于指定保存文件夹与文件名（不含扩展名）
  7) ✅ 新增 xlim / ylim 用于指定坐标轴范围（支持 None 表示自适应）
"""

import os
import numpy as np
import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, AutoMinorLocator
from matplotlib.offsetbox import AnchoredText


# ================= 统一风格 =================
def apply_project_style(font_family=("Microsoft YaHei", "SimHei", "DejaVu Sans"),
                        base_fontsize=10.5):
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": list(font_family),
        "font.size": base_fontsize,
        "axes.unicode_minus": False,
        "axes.linewidth": 1.0,
        "grid.linestyle": "--",
        "grid.linewidth": 0.8,
        "grid.alpha": 0.5,
        "figure.constrained_layout.use": True,
        "savefig.dpi": 300,  # 可被函数入参覆盖
        # 让标注盒子观感更统一
        "patch.linewidth": 0.8,
    })


# ================= 小工具 =================
def cm2inch(w_cm, h_cm):
    return (w_cm / 2.54, h_cm / 2.54)


def _smart_detect_time_col(df: pd.DataFrame, user_col: str | None):
    if user_col and user_col in df.columns:
        return user_col
    for c in ["时间", "time", "Time", "t", "T", "时间(s)", "时间（s）", "Time(s)"]:
        if c in df.columns:
            return c
    return df.columns[0]


def _to_list_or_none(s):
    if s is None:
        return None
    if isinstance(s, (list, tuple)):
        return list(s)
    parts = [x.strip() for x in str(s).split(",") if x.strip() != ""]
    return parts if parts else None


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


# 可选平滑（默认不用）
def smooth_signal(y, method=None, **kwargs):
    if method is None:
        return y
    if method == "ma":
        w = max(3, int(kwargs.get("window", 5)))
        if w % 2 == 0:
            w += 1
        k = np.ones(w) / w
        return np.convolve(y, k, mode="same")
    try:
        from scipy.signal import savgol_filter, butter, filtfilt
    except Exception:
        raise ImportError("使用 'savgol' 或 'butter' 需要安装 scipy：pip install scipy")
    if method == "savgol":
        wl = max(5, int(kwargs.get("window_length", 11)))
        if wl % 2 == 0:
            wl += 1
        po = int(kwargs.get("polyorder", 3))
        return savgol_filter(y, window_length=wl, polyorder=po, mode="interp")
    if method == "butter":
        cutoff = float(kwargs.get("cutoff", 2.0))
        fs = float(kwargs.get("fs", 100.0))
        order = int(kwargs.get("order", 2))
        b, a = butter(order, cutoff / (0.5 * fs), btype="low")
        return filtfilt(b, a, y)
    return y


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


def plugin_annotate_absmax(ax, x, y, label=None, fmt="{:.1f}",
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


# ================= 插件：两曲线峰值对比（带边框的锚定文本框） =================
def plugin_peak_metrics_box(ax, y_ref, y_cmp, loc="lower right"):
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


# ================= 主函数：从Excel绘图 =================
def plot_from_excel(
        excel_path: str,
        sheet_name: str | int | None = 0,
        time_col: str | None = None,
        series_cols: list[str] | None = None,
        labels: list[str] | None = None,
        figsize_cm=(14, 8),
        dpi=300,
        xlabel="时间（s）",
        ylabel="位移响应（mm）",
        title=None,
        title_pad=6,  # 新增：标题与图框的距离（pt）
        colors=None,
        linewidth=2.2,
        legend_loc="upper right",
        x_major: float | None = None,
        y_major: float | None = None,
        show_minor_grid=True,
        zero_baseline=True,
        annotate_peaks=False,
        metrics_box=False,
        smooth=None,
        smooth_kwargs=None,
        # ======= 新增：坐标轴范围 =======
        xlim: tuple | None = None,   # 例如 (0, 30) 或 (None, 30)
        ylim: tuple | None = None,   # 例如 (-50, 50) 或 (None, None)
        # ===== 保存相关参数（兼容旧版） =====
        outfile: str | None = None,            # 指定单一文件名（完整路径，含扩展名）
        export_formats: list[str] | None = None,  # 例如 ["png","pdf","svg"]
        save_dir: str | None = None,           # ✅ 指定保存文件夹
        filename_base: str | None = None,      # ✅ 指定文件名（不含扩展名）
        # ===== 其他 =====
        plugins: list | None = None,
        close_after_save=True,
):
    if smooth_kwargs is None:
        smooth_kwargs = {}

    # 读数据
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    tcol = _smart_detect_time_col(df, time_col)
    time = df[tcol].to_numpy(dtype=float)

    if series_cols is None:
        series_cols = [c for c in df.columns
                       if c != tcol and np.issubdtype(df[c].dtype, np.number)]
    if not series_cols:
        raise ValueError("未找到可用的曲线列，请检查 Excel。")

    if labels is None or len(labels) != len(series_cols):
        labels = series_cols

    if colors is None:
        colors = ["#1F77B4", "#D62728", "#FF7F0E", "#2CA02C", "#9467BD", "#8C564B"]

    fig, ax = plt.subplots(figsize=cm2inch(*figsize_cm), dpi=dpi)

    plotted = []
    y_all_min, y_all_max = np.inf, -np.inf

    for i, col in enumerate(series_cols):
        y = df[col].to_numpy(dtype=float)
        if smooth:
            y = smooth_signal(y, method=smooth, **smooth_kwargs)
        color = colors[i % len(colors)]
        line, = ax.plot(time, y, lw=linewidth, color=color, label=labels[i])
        plotted.append((labels[i], time, y, line, color))
        # 统计全局 y 范围，便于 ylim None 端点替代
        y_all_min = min(y_all_min, float(np.nanmin(y)))
        y_all_max = max(y_all_max, float(np.nanmax(y)))

    # 坐标轴与网格
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    if title:
        ax.set_title(title, pad=title_pad)  # 应用标题间距
    ax.grid(True)
    if show_minor_grid:
        ax.xaxis.set_minor_locator(AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(AutoMinorLocator(2))
        ax.grid(True, which="minor", alpha=0.25, linestyle="--", linewidth=0.6)
    if x_major:
        ax.xaxis.set_major_locator(MultipleLocator(x_major))
    if y_major:
        ax.yaxis.set_major_locator(MultipleLocator(y_major))
    if zero_baseline:
        ax.axhline(0.0, color="0.25", lw=1.0, alpha=0.9)

    # —— 让横坐标紧贴图框：默认锁到数据端点
    ax.margins(x=0.0)
    try:
        default_xlim = (float(np.nanmin(time)), float(np.nanmax(time)))
        ax.set_xlim(*default_xlim)
    except Exception:
        default_xlim = None

    # ====== 应用用户指定的坐标轴范围（若提供） ======
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
        for (lab, x, y, _, color) in plotted:
            plugin_annotate_absmax(ax, x, y, label=lab, used_bboxes=used_bboxes, line_color=color)

    # 峰值对比文本框（仅两条曲线生效；在一个带边框的框中）
    if metrics_box and len(plotted) >= 2:
        plugin_peak_metrics_box(ax, plotted[0][2], plotted[1][2], loc="lower right")

    # 额外插件
    if plugins:
        for fn in plugins:
            for (lab, x, y, _, _) in plotted:
                try:
                    fn(ax, x, y, lab)
                except TypeError:
                    fn(ax, x, y)

    leg = ax.legend(loc=legend_loc, framealpha=0.95, fancybox=True, edgecolor="#444")
    for t in leg.get_texts():
        t.set_fontsize(plt.rcParams["font.size"])

    # ===== 导出逻辑（兼容旧版 + 新增 save_dir/filename_base） =====
    saved = []

    if outfile:  # 1) 若显式指定了完整文件（含扩展名），优先使用
        outdir = os.path.dirname(os.path.abspath(outfile))
        if outdir:
            os.makedirs(outdir, exist_ok=True)
        fig.savefig(outfile, bbox_inches="tight")
        saved.append(outfile)

    else:
        # 2) 多格式导出
        if not export_formats:
            export_formats = ["png", "pdf"]  # 时程默认 png+pdf（与原版一致）

        # 2.1 组合保存目录
        if save_dir is None:
            save_dir = os.getcwd()  # 如未指定，保存到当前工作目录
        os.makedirs(save_dir, exist_ok=True)

        # 2.2 组合文件名
        if filename_base is None:
            base_name = os.path.splitext(os.path.basename(excel_path))[0] + "_plot"
        else:
            base_name = filename_base

        # 2.3 逐格式导出
        for ext in export_formats:
            path = os.path.join(save_dir, f"{base_name}.{ext.lower()}")
            fig.savefig(path, bbox_inches="tight")
            saved.append(path)

    if close_after_save:
        plt.close(fig)
    return saved, fig
