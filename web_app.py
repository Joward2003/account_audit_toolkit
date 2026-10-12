"""Account Audit Toolkit — Web 前端

轻量 Flask Web 应用：
- 上传 WP Excel 文件 + supporting 文件夹（zip）
- 一键运行审阅流水线
- 在线查看 / 下载审阅结论

启动:
  python web_app.py
  浏览器打开 http://localhost:5050
"""

from __future__ import annotations

import json
import os
import queue
import re
import shutil
import tempfile
import threading
import zipfile
from pathlib import Path
from datetime import datetime

from flask import Flask, Response, jsonify, render_template, request, send_file, stream_with_context

# 复用 main.py 的核心逻辑
from main import AuditConclusionGenerator
from pipeline import AccountAuditPipeline

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB

APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = APP_DIR / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

generator = AuditConclusionGenerator()
pipeline_lock = threading.Lock()


TECHNICAL_DISPLAY_REPLACEMENTS = [
    ("流水线", "检查流程"),
    ("supporting", "支持性文件"),
    ("Supporting", "支持性文件"),
    ("Sheet", "工作表"),
    ("sheet", "工作表"),
    ("SOD", "前后台管理员重合检查"),
    ("派生检查", "规则检查"),
    ("ZIP", "压缩包"),
    ("zip", "压缩包"),
    ("OCR触发/未触发原因", "截图核对原因"),
    ("OCR已匹配账号", "按表格位置已确认账号"),
    ("ocr_global_matched_accounts", "全文补充确认账号"),
    ("global_matched_accounts", "全文补充确认账号"),
    ("OCR缺失账号", "截图中未确认账号"),
    ("OCR明细", "截图核对明细"),
    ("OCR状态", "截图核对状态"),
    ("是否执行OCR", "是否核对截图"),
    ("OCR", "截图文字识别"),
    ("ocr", "截图文字识别"),
    ("ExpectedSetStatus", "账号清单状态"),
    ("ExpectedSetNote", "账号清单说明"),
    ("expected set", "待核对账号清单"),
    ("Expected set", "待核对账号清单"),
    ("expected accounts", "待核对账号"),
    ("expected", "待核对"),
    ("ParseMethod", "检查方式"),
    ("parse_method", "检查方式"),
    ("manual_upload", "支持性材料上传情况"),
    ("ManualUpload", "支持性材料上传情况"),
    ("手工上传标识", "支持性材料标识"),
    ("fallback", "补充核对"),
    ("兜底", "补充核对"),
    ("全量文本池", "截图全文"),
    ("结构化", "按表格位置"),
    ("结构", "表格位置"),
    ("scope", "检查范围"),
    ("Scope", "检查范围"),
    ("PENDING", "待确认"),
    ("REVIEW", "待确认"),
    ("UNKNOWN", "待确认"),
    ("BUILT", "已确认"),
    ("PASS", "通过"),
    ("FAIL", "未通过"),
    ("SKIP", "跳过"),
]

STATUS_DISPLAY_MAP = {
    "ACCOUNT_EXPORT_FILE_NOT_FOUND": "未找到账号审阅导出文件",
    "EXPORT_REQUIRED_COLUMNS_MISSING": "账号审阅导出缺少必要列",
    "EXPORT_EMPTY_WORKBOOK": "账号审阅导出为空",
    "EXPORT_READ_ERROR": "账号审阅导出读取失败",
    "NO_RESOLVED_FOLDER_PATH": "未找到支持性文件夹",
    "REVIEW_OBJECT_NOT_FOUND": "审阅导出中未找到对应审阅对象",
    "EXPORT_MISSING_REVIEW_OBJECT_COL": "账号审阅导出缺少审阅对象列",
    "EXPORT_MISSING_ACCOUNT_COL": "账号审阅导出缺少账号列",
    "ONLY_SKIPPED_ACCOUNTS": "未找到需要核对的有效账号",
    "EMPTY_AFTER_FILTER": "未找到需要核对的有效账号",
    "NO_REVIEW_OBJECT_RULE": "该检查项无需账号清单",
    "NOT_ACCOUNT_CHECK_NODE": "该检查项不按账号清单核对",
    "READY_FOR_EVIDENCE": "待核对支持性材料",
    "READY_FOR_SERVER_EVIDENCE_REVIEW": "待核对服务器支持性材料",
    "MISSING_EVIDENCE_BY_MANUAL_FLAG": "支持性材料未上传",
    "EVIDENCE_UPLOADED_BY_MANUAL_FLAG": "支持性材料已上传",
    "PENDING_DERIVED_CHECK": "待进行规则检查",
}


@app.route("/")
def index():
    return render_template("index.html")


def _friendly_message(message: object) -> str:
    """把页面展示文字中的内部术语转换成业务方更容易理解的说法。"""
    text = "" if message is None else str(message)
    if not text:
        return ""

    if text.startswith("N/A-Refer to "):
        target = text.removeprefix("N/A-Refer to ").strip()
        return f"不适用（引用 {target} 材料）"
    if text.startswith("Refer to "):
        target = text.removeprefix("Refer to ").strip()
        return f"引用 {target} 材料"
    if text == "N/A":
        return "不适用"
    if text.startswith("N/A-"):
        return "不适用（" + text.removeprefix("N/A-").strip() + "）"

    for raw, friendly in STATUS_DISPLAY_MAP.items():
        text = text.replace(raw, friendly)
    for raw, friendly in TECHNICAL_DISPLAY_REPLACEMENTS:
        text = text.replace(raw, friendly)

    text = re.sub(r"\bN/A\b", "不适用", text)
    text = re.sub(r"\bY\b", "是", text)
    text = re.sub(r"\bN\b", "否", text)
    text = re.sub(r"\bTRUE\b", "是", text, flags=re.IGNORECASE)
    text = re.sub(r"\bFALSE\b", "否", text, flags=re.IGNORECASE)
    text = text.replace("截图文字识别 截图", "截图")
    text = text.replace("截图文字识别截图", "截图")
    text = text.replace("无需截图文字识别", "无需核对截图")
    return text


def _friendly_column_name(column: object) -> str:
    mapping = {
        "System": "系统",
        "Category": "系统类别",
        "System Owner": "系统负责人",
        "N/A": "不适用",
    }
    name = "" if column is None else str(column)
    return mapping.get(name, _friendly_message(name))


def _friendly_rows_for_web(rows: list[dict]) -> list[dict]:
    display_rows = []
    for row in rows:
        display_rows.append({
            _friendly_column_name(key): _friendly_message(value)
            for key, value in row.items()
        })
    return display_rows


def _need_label_for_web(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if text in {"测试", "需要测试"}:
        return "是"
    if text == "不测试":
        return "否"
    if text == "引用其他材料":
        return "引用"
    return text or "待确认"


def _short_detail_for_web(reason: object) -> str:
    text = "" if reason is None else str(reason).strip()
    if not text:
        return ""

    compact_phrases = [
        ("已通过账号审阅导出验证", "系统导出验证"),
        ("该节点无需上传截图支持文件", "无需截图"),
        ("支持性材料标识=N", "无需截图"),
        ("账号审阅导出中无待核对账号", "无待核对账号"),
        ("无账号审阅导出数据", "无账号审阅导出"),
    ]
    for raw, short in compact_phrases:
        if raw in text:
            return short

    patterns = [
        r"已确认\d+个，未确认\d+个",
        r"截图已上传\(\d+张\)",
        r"只读\(\d+条\)",
        r"未出现在导出表\d+人",
        r"前后台管理员重合[:：][^；。]*",
        r"无前后台管理员重合",
        r"未找到[^；。]*",
        r"缺少[^；。]*",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(0)

    for sep in ("；", "。", ";"):
        if sep in text:
            text = text.split(sep, 1)[0].strip()
            break
    return text[:24] + ("..." if len(text) > 24 else "")


def _web_cell_for_result(row: dict) -> str:
    need = _need_label_for_web(row.get("是否需要测试", ""))
    status = str(row.get("测试结果", row.get("结论", ""))).strip() or "待确认"
    reason = _short_detail_for_web(row.get("结果说明", ""))

    if need == "是":
        text = f"是 - {status}"
    elif need == "否":
        text = "否 - 不执行"
    elif need == "引用":
        text = f"引用 - {status or '看其他材料'}"
    else:
        text = f"{need} - {status}"

    if reason and status not in {"不适用"}:
        text += f"（{reason}）"
    return text


def _build_web_result_rows(result_rows: list[dict]) -> list[dict]:
    """页面按系统横向展示各检查程序，单元格只保留“是否执行-简短结果”。"""
    grouped: dict[str, dict] = {}
    for row in result_rows:
        system = str(row.get("系统", row.get("System", ""))).strip()
        sid = str(row.get("SID", "")).strip()
        if not sid:
            continue

        item = grouped.setdefault(sid, {
            "系统": system,
            "SID": sid,
            "系统类别": row.get("系统类别", row.get("Category", "")),
        })
        node = str(row.get("检查节点", "")).strip()
        if node:
            item[node] = _web_cell_for_result(row)
    return list(grouped.values())


@app.route("/api/run", methods=["POST"])
def run_pipeline():
    """接收上传文件，并以 NDJSON 流持续返回真实执行进度。"""
    if "file" not in request.files:
        return jsonify({"error": "请上传审阅文件（Excel）"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "未选择审阅文件"}), 400
    if Path(file.filename).suffix.lower() not in {".xlsx", ".xls", ".csv"}:
        return jsonify({"error": "审阅文件仅支持 .xlsx / .xls / .csv"}), 400
    input_filename = Path(file.filename).name

    sheet_name = request.form.get("sheet", "Sheet1")

    # 处理 supporting 文件夹上传（zip）
    supporting_zip = request.files.get("supporting")
    tmp_supporting_dir = None
    zip_path = None

    if supporting_zip and supporting_zip.filename:
        if Path(supporting_zip.filename).suffix.lower() != ".zip":
            return jsonify({"error": "支持性文件仅支持 .zip 压缩包"}), 400
        tmp_supporting_dir = Path(tempfile.mkdtemp(prefix="audit_supporting_"))
        zip_path = tmp_supporting_dir / Path(supporting_zip.filename).name
        supporting_zip.save(str(zip_path))

    # 保存审阅文件
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = f"{ts}_{input_filename}"
    upload_path = UPLOAD_DIR / safe_name
    file.save(str(upload_path))

    event_queue: queue.Queue = queue.Queue()

    def emit_progress(event: dict) -> None:
        if "message" in event:
            event = {**event, "message": _friendly_message(event["message"])}
        event_queue.put({"type": "progress", **event})

    def worker() -> None:
        try:
            if pipeline_lock.locked():
                emit_progress({
                    "stage": "waiting",
                    "message": "已有任务正在执行，正在等待运行资源",
                })

            with pipeline_lock:
                if zip_path and tmp_supporting_dir:
                    emit_progress({
                        "stage": "unzip",
                        "message": "正在解压并检查支持性文件压缩包",
                    })
                    with zipfile.ZipFile(zip_path, "r") as zf:
                        _extract_zip_fix_encoding(zf, tmp_supporting_dir)
                    zip_path.unlink(missing_ok=True)

                emit_progress({
                    "stage": "read_input",
                    "message": f"正在读取审阅文件：{input_filename}",
                })
                wp_df = AccountAuditPipeline.read_file(
                    str(upload_path), sheet_name=sheet_name
                )
                system_count = len(wp_df)
                systems = [str(s) for s in wp_df["System"].tolist()]

                supporting_dir = (
                    _find_supporting_subdir(tmp_supporting_dir)
                    if tmp_supporting_dir
                    else None
                )
                result = generator.generate(
                    wp_df,
                    supporting_dir=supporting_dir,
                    progress_callback=emit_progress,
                )

                import pandas as pd
                conclusion_path = result.get("conclusion", "")
                rows_data = []
                if conclusion_path and Path(conclusion_path).exists():
                    try:
                        df = pd.read_excel(conclusion_path, sheet_name="测试结果")
                    except ValueError:
                        df = pd.read_excel(conclusion_path, sheet_name="审阅结论")
                    rows_data = df.fillna("").to_dict(orient="records")
                web_rows = _build_web_result_rows(rows_data)

                event_queue.put({
                    "type": "result",
                    "data": {
                        "success": True,
                        "systems": systems,
                        "system_count": system_count,
                        "conclusion_file": os.path.basename(conclusion_path),
                        "rows": _friendly_rows_for_web(web_rows),
                    },
                })
        except zipfile.BadZipFile:
            event_queue.put({
                "type": "error",
                "error": "支持性文件不是有效的压缩包，请重新打包",
            })
        except Exception as exc:
            event_queue.put({"type": "error", "error": str(exc)})
        finally:
            upload_path.unlink(missing_ok=True)
            if tmp_supporting_dir:
                shutil.rmtree(str(tmp_supporting_dir), ignore_errors=True)
            event_queue.put(None)

    def event_stream():
        threading.Thread(target=worker, daemon=True).start()
        while True:
            event = event_queue.get()
            if event is None:
                break
            yield json.dumps(event, ensure_ascii=False) + "\n"

    return Response(
        stream_with_context(event_stream()),
        mimetype="application/x-ndjson",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


def _extract_zip_fix_encoding(zf: zipfile.ZipFile, dest: Path) -> None:
    """
    解压 zip，修复中文文件名编码。

    macOS 打包的 zip 可能不用 UTF-8 flag，导致 Python zipfile
    用 CP437 解码中文文件名，变成乱码（如 账号 → Φ┤ªσÅ╖）。

    此函数逐个读取 zip entry 的原始文件名字节，
    尝试 UTF-8 → 若为乱码则用 CP437 回编码再 UTF-8 解码，
    然后用正确文件名写出。
    """
    for member in zf.infolist():
        # 获取原始文件名字节
        try:
            raw_name = member.filename.encode("cp437")
        except (UnicodeEncodeError, UnicodeDecodeError):
            raw_name = member.filename.encode("utf-8", errors="surrogateescape")

        # 尝试用 UTF-8 解码
        try:
            correct_name = raw_name.decode("utf-8")
        except UnicodeDecodeError:
            correct_name = member.filename  # 保持原样

        # 跳过 macOS 资源分支目录
        if correct_name.startswith("__MACOSX") or "/__MACOSX" in correct_name:
            continue
        # 跳过隐藏文件
        parts = correct_name.replace("\\", "/").split("/")
        if any(p.startswith("._") or p == "__MACOSX" for p in parts):
            continue

        target_path = (dest / correct_name).resolve()
        dest_resolved = dest.resolve()
        if target_path != dest_resolved and dest_resolved not in target_path.parents:
            raise ValueError(f"zip 包含不安全路径：{correct_name}")
        # 创建父目录
        target_path.parent.mkdir(parents=True, exist_ok=True)

        if member.is_dir():
            target_path.mkdir(parents=True, exist_ok=True)
        else:
            with zf.open(member) as src, open(target_path, "wb") as dst:
                shutil.copyfileobj(src, dst)


def _find_supporting_subdir(tmp_dir: Path) -> Path:
    """
    在解压目录中找 supporting 根目录。

    兼容两种情况：
    1. zip 根目录直接包含各系统文件夹 → 返回 tmp_dir
    2. zip 内有一层父目录（如 supporting/）→ 返回该父目录

    跳过 macOS 干扰文件（__MACOSX、._*）。
    """
    SKIP = {"__MACOSX"}

    def _is_valid_dir(p: Path) -> bool:
        """是否为有效的目录（非系统干扰文件）。"""
        return p.is_dir() and p.name not in SKIP and not p.name.startswith("._")

    current = tmp_dir
    # 只有一个包装目录且当前层没有业务文件时持续下钻；多系统目录则停在共同根目录。
    while True:
        has_business_file = any(
            p.is_file() and p.suffix.lower() in {".xlsx", ".xls", ".csv"}
            and not p.name.startswith("~$")
            for p in current.iterdir()
        )
        valid_children = [c for c in current.iterdir() if _is_valid_dir(c)]
        if has_business_file or len(valid_children) != 1:
            return current
        current = valid_children[0]


@app.route("/api/download")
def download_conclusion():
    """下载最终审阅结论 Excel。"""
    from main import CONCLUSION_FILE
    path = Path(CONCLUSION_FILE)
    if not path.exists():
        return jsonify({"error": "结论文件不存在，请先运行流水线"}), 404
    return send_file(
        str(path),
        as_attachment=True,
        download_name="account_audit_conclusion.xlsx",
    )


if __name__ == "__main__":
    print(f"\n{'='*50}")
    print("  Account Audit Toolkit — Web 前端")
    print(f"  打开浏览器访问: http://localhost:5000")
    print(f"{'='*50}\n")
    # 演示版仅允许本机访问，避免审阅文件通过局域网暴露。
    app.run(debug=False, host="127.0.0.1", port=5000, threaded=True)
