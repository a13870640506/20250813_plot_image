# -*- coding: utf-8 -*-
# 纯 Python 方式调用，无需命令行

# 1) 从单文件工具中导入
from plot_excel_timeseries import (
    apply_project_style,  # 一次性设置全局绘图风格（中文字体、虚线网格、紧凑布局等）
    plot_from_excel  # 从 Excel 读数据并绘图，支持峰值标注、导出多格式等
)

# 2) 可选：应用全局风格（建议放在程序开头，只调一次）
apply_project_style(
    # 这里可以自定义字体族/字号；不传就用默认（微软雅黑/黑体/DejaVu Sans，字号10.5）
    # font_family=("Microsoft YaHei", "SimHei", "DejaVu Sans"),
    # base_fontsize=10.5,
)

# 3) 设定 Excel 文件与列信息
excel_path = "1.xlsx"  # 你的 Excel 路径
sheet_name = "Sheet3"  # 工作表名；也可用 0 表示第一个表
time_col = "时间"  # 时间列名；不写会自动取第一列为时间列
series_cols = ["设置粘滞阻尼器", "设置粘滞阻尼器+铅芯橡胶支座"]  # 需要绘制的两条曲线
labels = ["设置粘滞阻尼器", "设置粘滞阻尼器+铅芯橡胶支座"]  # 图例中文名（与上面列一一对应）

# 4) 直接绘图并导出（支持指定保存文件夹与文件名；✅ 现在支持 xlim/ylim）
saved_files, fig = plot_from_excel(
    excel_path=excel_path,  # Excel 文件路径
    sheet_name=sheet_name,  # 表名/索引
    time_col=time_col,  # 时间列
    series_cols=series_cols,  # 数据列
    labels=labels,  # 图例标签
    figsize_cm=(16, 9),  # 图幅（厘米）；默认(14,8)
    dpi=300,  # 清晰度（论文可用600~1200）
    xlabel="时间（s）",  # x 轴名
    ylabel="位移响应（mm）",  # y 轴名
    title="跨中Y方向时程曲线对比",  # 可填写标题或 None
    title_pad=10,  # 设置标题距离
    linewidth=1.5,  # 线宽
    legend_loc="upper right",  # 图例位置
    x_major=5,  # x 轴主刻度步长（秒）
    y_major=20,  # y 轴主刻度步长（mm）
    show_minor_grid=True,  # 显示次网格
    zero_baseline=True,  # 显示 y=0 基线
    annotate_peaks=True,  # 开启峰值标注（内置自动错位，避免文字重叠）
    metrics_box=True,  # 显示两曲线“峰值对比/峰值降低(%)”（已去掉 RMS）
    smooth=None,  # 可选: "ma" / "savgol" / "butter"
    smooth_kwargs=None,

    # ===== ✅ 新增：坐标轴范围（任一端可为 None 表示自适应） =====
    xlim=(0, 30),  # 例如仅看 0~30 s；可写 (None, 30) / (0, None)
    ylim=(-80, 80),  # 例如固定纵轴范围

    # ===== 统一的保存控制 =====
    save_dir="./两种抗震装置时程对比1",  # 指定保存文件夹（自动创建）
    filename_base="两种抗震装置-时程曲线对比",  # 指定文件名（不含扩展名）
    export_formats=["png", "pdf", "svg"],  # 一次导出多种格式
    close_after_save=True,  # 保存后关闭 Figure，脚本批量跑更省内存
)

print("已导出文件：", saved_files)

# 5) 如果你在交互式环境里（如 Jupyter）想立即展示：
# import matplotlib.pyplot as plt
# plt.show()
