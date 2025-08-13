# -*- coding: utf-8 -*-
import os
import uuid
import json
from datetime import datetime
from flask import Flask, request, jsonify, send_file, abort
from flask_cors import CORS
import pandas as pd

from utils_plot import (
    plot_timeseries_preview, plot_timeseries_export,
    plot_hysteresis_preview, plot_hysteresis_export
)

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TMP_DIR = os.path.join(BASE_DIR, "tmp")
UPLOAD_DIR = os.path.join(TMP_DIR, "uploads")
META_DIR = os.path.join(TMP_DIR, "meta")
EXPORT_ROOT = os.path.join(BASE_DIR, "..", "exports")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(META_DIR, exist_ok=True)
os.makedirs(EXPORT_ROOT, exist_ok=True)


def _excel_path(file_id):
    return os.path.join(UPLOAD_DIR, f"{file_id}.xlsx")


def _read_excel_head(path, sheet_name, nrows=5):
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, nrows=nrows, engine="openpyxl")
        return df
    except Exception as e:
        raise e


def _read_excel_full(path, sheet_name):
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
        return df
    except Exception as e:
        raise e


@app.route("/api/excel/upload", methods=["POST"])
def upload_excel():
    if "file" not in request.files:
        return jsonify({"ok": False, "msg": "缺少文件字段 file"}), 400

    f = request.files["file"]
    if f.filename == "":
        return jsonify({"ok": False, "msg": "未选择文件"}), 400
    if not (f.filename.lower().endswith(".xlsx") or f.filename.lower().endswith(".xls")):
        return jsonify({"ok": False, "msg": "仅支持 .xlsx / .xls"}), 400

    file_id = str(uuid.uuid4())
    save_path = _excel_path(file_id)
    f.save(save_path)

    try:
        xls = pd.ExcelFile(save_path, engine="openpyxl")
        sheets = xls.sheet_names
        sniff = {}
        for sh in sheets[:3]:  # 仅预览前3个 sheet，避免超大表卡顿
            head = _read_excel_head(save_path, sh, nrows=5)
            sniff[sh] = {
                "columns": list(map(str, head.columns)),
                "head": head.astype(str).values.tolist()
            }
        meta = {
            "file_id": file_id, "filename": f.filename,
            "sheets": sheets, "created_at": datetime.now().isoformat()
        }
        with open(os.path.join(META_DIR, f"{file_id}.json"), "w", encoding="utf-8") as fp:
            json.dump(meta, fp, ensure_ascii=False, indent=2)

        return jsonify({"ok": True, "file_id": file_id, "sheets": sheets, "sniff": sniff})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"Excel 解析失败: {e}"}), 500


@app.route("/api/excel/columns", methods=["GET"])
def excel_columns():
    file_id = request.args.get("file_id")
    sheet = request.args.get("sheet")
    if not file_id or not sheet:
        return jsonify({"ok": False, "msg": "缺少 file_id 或 sheet"}), 400

    path = _excel_path(file_id)
    if not os.path.exists(path):
        return jsonify({"ok": False, "msg": "文件不存在或已过期"}), 404
    try:
        head = _read_excel_head(path, sheet, nrows=20)
        cols = list(map(str, head.columns))
        # 简单推断：时间列候选（含 'time'/'时间'）
        time_candidates = [c for c in cols if "time" in c.lower() or "时间" in c]
        numeric_cols = []
        for c in cols:
            s = pd.to_numeric(head[c], errors="coerce")
            if s.notna().sum() >= max(3, len(head) * 0.6):
                numeric_cols.append(c)
        return jsonify({"ok": True, "columns": cols, "time_candidates": time_candidates, "numeric_cols": numeric_cols})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"读取列失败: {e}"}), 500


@app.route("/api/plot/preview", methods=["POST"])
def plot_preview():
    data = request.get_json(force=True, silent=True) or {}
    file_id = data.get("file_id")
    sheet = data.get("sheet")
    plot_type = data.get("plot_type")
    params = data.get("params") or {}

    if not file_id or not sheet or plot_type not in ("timeseries", "hysteresis"):
        return jsonify({"ok": False, "msg": "参数缺失或 plot_type 非法"}), 400

    path = _excel_path(file_id)
    if not os.path.exists(path):
        return jsonify({"ok": False, "msg": "文件不存在或已过期"}), 404

    try:
        df = _read_excel_full(path, sheet)
        if plot_type == "timeseries":
            if not params.get("time_col") or not params.get("series_cols"):
                return jsonify({"ok": False, "msg": "timeseries 需要 time_col 与 series_cols"}), 400
            out = plot_timeseries_preview(df, params)
        else:
            if not params.get("x_col") or not params.get("y_cols"):
                return jsonify({"ok": False, "msg": "hysteresis 需要 x_col 与 y_cols"}), 400
            out = plot_hysteresis_preview(df, params)
        return jsonify(out)
    except Exception as e:
        return jsonify({"ok": False, "msg": f"绘图预览失败: {e}"}), 500


@app.route("/api/plot/export", methods=["POST"])
def plot_export():
    data = request.get_json(force=True, silent=True) or {}
    file_id = data.get("file_id")
    sheet = data.get("sheet")
    plot_type = data.get("plot_type")
    params = data.get("params") or {}

    if not file_id or not sheet or plot_type not in ("timeseries", "hysteresis"):
        return jsonify({"ok": False, "msg": "参数缺失或 plot_type 非法"}), 400

    path = _excel_path(file_id)
    if not os.path.exists(path):
        return jsonify({"ok": False, "msg": "文件不存在或已过期"}), 404

    # 统一导出根目录
    save_dir = params.get("save_dir") or os.path.join(EXPORT_ROOT, datetime.now().strftime("%Y-%m-%d"))
    params["save_dir"] = save_dir

    try:
        df = _read_excel_full(path, sheet)
        if plot_type == "timeseries":
            if not params.get("time_col") or not params.get("series_cols"):
                return jsonify({"ok": False, "msg": "timeseries 需要 time_col 与 series_cols"}), 400
            out = plot_timeseries_export(df, params)
        else:
            if not params.get("x_col") or not params.get("y_cols"):
                return jsonify({"ok": False, "msg": "hysteresis 需要 x_col 与 y_cols"}), 400
            out = plot_hysteresis_export(df, params)
        # 转换为可下载 URL
        files = [f"/download?path={os.path.abspath(p)}" for p in out["files"]]
        zip_url = f"/download?path={os.path.abspath(out['zip'])}"
        return jsonify({"ok": True, "files": files, "zip": zip_url, "preview_data_url": out.get("preview_data_url")})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"导出失败: {e}"}), 500


@app.route("/download", methods=["GET"])
def download():
    path = request.args.get("path")
    if not path:
        return abort(400, "缺少 path")
    path = os.path.abspath(path)
    # 简单防护：限制在工程根目录内
    root = os.path.abspath(os.path.join(BASE_DIR, ".."))
    if not path.startswith(root) or not os.path.exists(path):
        return abort(404)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
