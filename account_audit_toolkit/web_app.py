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

import io
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_file

# 复用 main.py 的核心逻辑
from main import AuditConclusionGenerator
from pipeline import AccountAuditPipeline

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "account_audit_toolkit" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

generator = AuditConclusionGenerator()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/run", methods=["POST"])
def run_pipeline():
    """接收上传的 WP Excel + supporting zip，运行流水线，返回结果。"""
    if "file" not in request.files:
        return jsonify({"error": "请上传审阅文件（Excel）"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "未选择审阅文件"}), 400

    sheet_name = request.form.get("sheet", "Sheet1")

    # 处理 supporting 文件夹上传（zip）
    supporting_zip = request.files.get("supporting")
    tmp_supporting_dir = None

    if supporting_zip and supporting_zip.filename:
        tmp_supporting_dir = Path(tempfile.mkdtemp(prefix="audit_supporting_"))
        zip_path = tmp_supporting_dir / supporting_zip.filename
        supporting_zip.save(str(zip_path))

        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                _extract_zip_fix_encoding(zf, tmp_supporting_dir)
        except zipfile.BadZipFile:
            shutil.rmtree(str(tmp_supporting_dir), ignore_errors=True)
            tmp_supporting_dir = None
            return jsonify({"error": "supporting 文件不是有效的 zip 压缩包，请重新打包"}), 400
        except Exception as e:
            shutil.rmtree(str(tmp_supporting_dir), ignore_errors=True)
            tmp_supporting_dir = None
            return jsonify({"error": f"zip 解压失败: {e}"}), 400

        zip_path.unlink(missing_ok=True)

        # 验证 zip 内容结构
        subdir = _find_supporting_subdir(tmp_supporting_dir)
        has_opm = (subdir / "OPM").is_dir()
        has_yumc = (subdir / "YumcZoo").is_dir()
        if not has_opm and not has_yumc:
            shutil.rmtree(str(tmp_supporting_dir), ignore_errors=True)
            tmp_supporting_dir = None
            return jsonify({
                "error": "zip 中未找到 OPM 或 YumcZoo 文件夹。"
                         "请确保 zip 根目录包含 OPM/ 和/或 YumcZoo/ 子文件夹。"
            }), 400

    # 保存审阅文件
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = f"{ts}_{file.filename}"
    upload_path = UPLOAD_DIR / safe_name
    file.save(str(upload_path))

    try:
        # 读取 WP
        wp_df = AccountAuditPipeline.read_file(str(upload_path), sheet_name=sheet_name)
        system_count = len(wp_df)
        systems = [str(s) for s in wp_df["System"].tolist()]

        # 运行流水线（传入自定义 supporting 目录）
        if tmp_supporting_dir:
            supporting_dir = _find_supporting_subdir(tmp_supporting_dir)
            result = generator.generate(wp_df, supporting_dir=supporting_dir)
        else:
            result = generator.generate(wp_df)

        # 读取结论 Excel 内容
        import pandas as pd
        conclusion_path = result.get("conclusion", "")
        rows_data = []
        if conclusion_path and Path(conclusion_path).exists():
            df = pd.read_excel(conclusion_path, sheet_name="审阅结论")
            rows_data = df.fillna("").to_dict(orient="records")

        return jsonify({
            "success": True,
            "systems": systems,
            "system_count": system_count,
            "conclusion_file": os.path.basename(conclusion_path),
            "rows": rows_data,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

    finally:
        # 清理上传文件
        upload_path.unlink(missing_ok=True)
        if tmp_supporting_dir:
            shutil.rmtree(str(tmp_supporting_dir), ignore_errors=True)


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

        target_path = dest / correct_name
        # 创建父目录
        target_path.parent.mkdir(parents=True, exist_ok=True)

        if member.is_dir():
            target_path.mkdir(parents=True, exist_ok=True)
        else:
            with zf.open(member) as src, open(target_path, "wb") as dst:
                shutil.copyfileobj(src, dst)


def _find_supporting_subdir(tmp_dir: Path) -> Path:
    """在解压目录中找 supporting 子目录（兼容压缩时多一层父目录，跳过 macOS 干扰文件）。"""
    SKIP = {"__MACOSX"}

    # 如果解压后直接有 OPM / YumcZoo 子目录，就是它
    for child in tmp_dir.iterdir():
        if child.name in SKIP or child.name.startswith("._"):
            continue
        if child.is_dir() and child.name in {"OPM", "YumcZoo"}:
            return tmp_dir
    # 否则可能有额外的父目录，比如 supporting/OPM
    for child in tmp_dir.iterdir():
        if child.name in SKIP or child.name.startswith("._"):
            continue
        if child.is_dir():
            for sub in child.iterdir():
                if sub.name in SKIP or sub.name.startswith("._"):
                    continue
                if sub.is_dir() and sub.name in {"OPM", "YumcZoo"}:
                    return child
    return tmp_dir


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
    print(f"  打开浏览器访问: http://localhost:5050")
    print(f"{'='*50}\n")
    app.run(debug=True, host="0.0.0.0", port=5050)
