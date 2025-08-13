# -*- coding: utf-8 -*-
"""
plot_excel_hysteresis.py — Excel滞回曲线绘图（统一风格 + 多格式导出）

功能特性：
  1) 统一项目风格：中文字体、虚线网格、紧凑布局、次网格、坐标轴0基线等
  2) Excel 读取：支持指定位移列（x_col）与一个或多个力列（y_cols）
  3) 多格式导出：export_formats=["png","pdf","svg"]；若既不传 outfile 也不传 export_formats，则默认导出 png+pdf+svg
  4) 细节可配：主刻度步长、线宽、图例位置、DPI、图幅(cm)等
  5) 与现有项目脚本风格保持一致（Agg 后端、统一 rcParams、工具函数等）
  6) ✅ save_dir 与 filename_base 指定保存文件夹与文件名（不含扩展名）
  7) ✅ xlim / ylim 指定坐标轴范围（支持 None 表示自适应）
  8) ✅ 智能留白：当未指定坐标轴范围时，自动按数据范围添加边距（避免曲线贴边）

（⚠ 已移除命令行，仅保留通用函数）
"""

import os
import numpy as np
import pandas as pd
import matplotlib

# 与项目一致：非交互式后端
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, AutoMinorLocator


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
        "savefig.dpi": 300,  # 可被入参覆盖
        "patch.linewidth": 0.8,
    })


# ================= 小工具 =================
def cm2inch(w_cm, h_cm):
    return (w_cm / 2.54, h_cm / 2.54)


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
    - lim = None -> None（不设置，走默认/智能留白）
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


def _nice_auto_limits(dmin, dmax, pad_frac=0.06, symmetric_if_cross_zero=True):
    """
    根据数据范围返回“好看”的坐标范围（带边距）。
    - pad_frac: 相对边距比例（6% 比较合适）
    - 若数据同时跨越正负且 symmetric_if_cross_zero=True，则按零对称扩展。
    """
    if not (np.isfinite(dmin) and np.isfinite(dmax)):
        return None

    # 防止 dmin == dmax（平线）的情况
    if np.isclose(dmin, dmax):
        span = 1.0 if dmin == 0 else abs(dmin) * 0.2
        dmin, dmax = dmin - span, dmax + span

    if symmetric_if_cross_zero and (dmin < 0.0) and (dmax > 0.0):
        m = max(abs(dmin), abs(dmax))
        lo = -m * (1.0 + pad_frac)
        hi = +m * (1.0 + pad_frac)
        return (lo, hi)
    else:
        span = dmax - dmin
        pad = span * pad_frac
        return (dmin - pad, dmax + pad)


# ================= 主函数：从Excel绘制滞回曲线 =================
def plot_hysteresis_from_excel(
        excel_path: str,
        sheet_name: str | int | None = 0,
        x_col: str = "位移",
        y_cols: list[str] | None = None,
        labels: list[str] | None = None,
        figsize_cm=(14, 10),
        dpi=600,
        xlabel="位移（mm）",
        ylabel="滞回力（kN）",
        title="阻尼器滞回曲线",
        title_pad=6,  # 新增：标题与图框的距离（pt）
        colors=None,
        linewidth=1.8,
        legend_loc="upper right",
        x_major: float | None = None,
        y_major: float | None = None,
        show_minor_grid=True,
        zero_axes=True,
        equal_aspect=False,  # 若希望 x、y 比例尺一致，可设 True
        # ======= 坐标轴范围 =======
        xlim: tuple | None = None,
        ylim: tuple | None = None,
        # ===== 保存相关参数（兼容旧版） =====
        outfile: str | None = None,              # 指定单一文件名（完整路径，含扩展名）
        export_formats: list[str] | None = None,  # 例如 ["png","pdf","svg"]
        save_dir: str | None = None,             # ✅ 保存文件夹
        filename_base: str | None = None,        # ✅ 文件名（不含扩展名）
        # ===== 其他 =====
        close_after_save=True,
        # —— 智能留白控制（仅在对应轴未显式设置时生效）
        auto_pad_frac_x: float = 0.05,
        auto_pad_frac_y: float = 0.08,
        symmetric_if_cross_zero: bool = True,
):
    # 读数据
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    if x_col not in df.columns:
        raise ValueError(f"未找到位移列：{x_col}（Excel列名区分大小写）")

    # 自动收集力列
    if y_cols is None:
        y_cols = [c for c in df.columns if c != x_col and np.issubdtype(df[c].dtype, np.number)]
    if not y_cols:
        raise ValueError("未找到可用的力列，请通过 y_cols 指定，或检查 Excel。")

    # 图例名称
    if labels is None or len(labels) != len(y_cols):
        labels = y_cols

    # 颜色方案
    if colors is None:
        colors = ["#D62728", "#1F77B4", "#FF7F0E", "#2CA02C", "#9467BD", "#8C564B"]

    # 统一风格
    apply_project_style()

    # 数据
    x = df[x_col].to_numpy(dtype=float)

    # 汇总 y 的全局范围（用于智能留白/默认 ylim）
    y_all_min, y_all_max = np.inf, -np.inf

    # 绘图
    fig, ax = plt.subplots(figsize=cm2inch(*figsize_cm), dpi=dpi)

    plotted = []
    for i, col in enumerate(y_cols):
        y = df[col].to_numpy(dtype=float)
        color = colors[i % len(colors)]
        line, = ax.plot(x, y, lw=linewidth, color=color, label=labels[i])
        plotted.append((labels[i], x, y, line, color))
        y_all_min = min(y_all_min, float(np.nanmin(y)))
        y_all_max = max(y_all_max, float(np.nanmax(y)))

    # 坐标 & 网格
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
    if zero_axes:
        ax.axhline(0.0, color="0.25", lw=1.0, alpha=0.9)
        ax.axvline(0.0, color="0.25", lw=1.0, alpha=0.9)

    # 可选等比例坐标（用于严格比较环线形状）
    if equal_aspect:
        try:
            ax.set_aspect('equal', adjustable='datalim')
        except Exception:
            pass

    # ====== 坐标轴范围与智能留白 ======
    # 1) x 轴
    try:
        x_min = float(np.nanmin(x))
        x_max = float(np.nanmax(x))
    except Exception:
        x_min = x_max = None

    default_xlim = (x_min, x_max) if (x_min is not None and x_max is not None) else None
    xlim_norm = _normalize_lim(xlim, dmin=x_min, dmax=x_max)

    if xlim_norm is not None:
        ax.set_xlim(*xlim_norm)
    elif default_xlim is not None:
        lo, hi = _nice_auto_limits(x_min, x_max, pad_frac=auto_pad_frac_x,
                                   symmetric_if_cross_zero=symmetric_if_cross_zero)
        ax.set_xlim(lo, hi)

    # 2) y 轴
    default_ylim = (y_all_min, y_all_max) if np.isfinite(y_all_min) and np.isfinite(y_all_max) else None
    ylim_norm = _normalize_lim(ylim,
                               dmin=default_ylim[0] if default_ylim else None,
                               dmax=default_ylim[1] if default_ylim else None)

    if ylim_norm is not None:
        ax.set_ylim(*ylim_norm)
    elif default_ylim is not None:
        lo, hi = _nice_auto_limits(y_all_min, y_all_max, pad_frac=auto_pad_frac_y,
                                   symmetric_if_cross_zero=symmetric_if_cross_zero)
        ax.set_ylim(lo, hi)

    # 再用 margins 做一次轻微兜底，防止极端情况下仍有贴边感
    ax.margins(x=0.0, y=0.0)  # 不扩大百分比（范围已手动扩过），仅启用“非粘连”机制

    # 图例
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
            export_formats = ["png", "pdf", "svg"]  # 滞回默认 png+pdf+svg

        # 2.1 组合保存目录
        if save_dir is None:
            save_dir = os.getcwd()
        os.makedirs(save_dir, exist_ok=True)

        # 2.2 组合文件名
        if filename_base is None:
            base_name = os.path.splitext(os.path.basename(excel_path))[0] + "_hysteresis"
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
