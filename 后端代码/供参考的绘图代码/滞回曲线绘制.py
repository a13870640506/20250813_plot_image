# -*- coding: utf-8 -*-
# 纯 Python 方式调用，无需命令行

# 1) 从单文件工具中导入
from plot_excel_hysteresis import (
    apply_project_style,   # 一次性设置全局绘图风格（中文字体、虚线网格、紧凑布局等）
    plot_hysteresis_from_excel  # 从 Excel 读数据并绘滞回曲线
)

# 2) 可选：应用全局风格（建议放在程序开头，只调一次）
apply_project_style(
    # 这里可以自定义字体族/字号；不传就用默认（微软雅黑/黑体/DejaVu Sans，字号10.5）
    # font_family=("Microsoft YaHei", "SimHei", "DejaVu Sans"),
    # base_fontsize=10.5,
)

# 3) 设定 Excel 文件与列信息
excel_path = "2.xlsx"                # Excel 路径
sheet_name = "只有粘滞阻尼器"          # 工作表名；也可用 0 表示第一个表
x_col = "位移"                         # 横轴列名
y_cols = ["力"]                        # 纵轴列名，可多条曲线对比
labels = ["粘滞阻尼器"]                # 图例中文名（与上面列一一对应）

# 4) 直接绘图并导出（支持指定保存文件夹与文件名；✅ 现在支持 xlim/ylim）
saved_files, fig = plot_hysteresis_from_excel(
    excel_path=excel_path,      # Excel 文件路径
    sheet_name=sheet_name,      # 表名/索引
    x_col=x_col,                # 位移列
    y_cols=y_cols,              # 力列
    labels=labels,              # 图例标签
    figsize_cm=(14, 10),        # 图幅（厘米）
    dpi=600,                    # 清晰度（论文可用600~1200）
    xlabel="位移（mm）",            # x 轴名
    ylabel="滞回力（kN）",          # y 轴名
    title="粘滞阻尼器-滞回曲线",     # 图标题
    title_pad=10,  # 设置标题距离
    linewidth=1.8,              # 线宽
    legend_loc="lower right",   # 图例位置
    x_major=None,               # x 轴主刻度步长（可设数值）
    y_major=None,               # y 轴主刻度步长（可设数值）
    show_minor_grid=True,       # 显示次网格
    zero_axes=False,             # 显示零轴
    equal_aspect=False,         # 是否等比例坐标

    # ===== ✅ 新增：坐标轴范围（任一端可为 None 表示自适应） =====
    # xlim=(None, None),
    # ylim=(None, None),

    # ===== 统一的保存控制 =====
    save_dir="./只考虑粘滞阻尼器-滞回曲线",  # 指定保存文件夹（自动创建）
    filename_base="粘滞阻尼器-滞回曲线",  # 指定文件名（不含扩展名）
    export_formats=["png", "pdf", "svg"],  # 一次导出多种格式
    close_after_save=True       # 保存后关闭 Figure
)

print("已导出文件：", saved_files)

# 5) 如果你在交互式环境里（如 Jupyter）想立即展示：
# import matplotlib.pyplot as plt
# plt.show()
