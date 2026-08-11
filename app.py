"""엑셀 문서 비교 웹 앱.

브라우저에서 두 개의 엑셀 파일을 업로드하면 시트별·셀 단위 차이를
색으로 하이라이트하여 보여준다.

실행:
    pip install -r requirements.txt
    python app.py
    브라우저에서 http://127.0.0.1:5000 접속
"""

from __future__ import annotations

import os
import tempfile

from flask import Flask, jsonify, render_template, request, send_file

from comparator import column_label, compare_files
from excel_diff import export_diff_xlsx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 파일당 최대 20MB

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


def _allowed(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/compare", methods=["POST"])
def compare():
    file_a = request.files.get("file_a")
    file_b = request.files.get("file_b")

    if not file_a or not file_b or not file_a.filename or not file_b.filename:
        return jsonify({"error": "두 개의 엑셀 파일을 모두 업로드해 주세요."}), 400

    for f in (file_a, file_b):
        if not _allowed(f.filename):
            return jsonify(
                {"error": f"지원하지 않는 파일 형식입니다: {f.filename} (.xlsx, .xls, .xlsm 만 가능)"}
            ), 400

    tmp_a = tmp_b = None
    try:
        tmp_a = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_b = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        file_a.save(tmp_a.name)
        file_b.save(tmp_b.name)
        tmp_a.close()
        tmp_b.close()

        result = compare_files(tmp_a.name, tmp_b.name)

        # 각 셀에 엑셀식 좌표(예: B3) 라벨을 추가한다.
        for sheet in result["sheets"]:
            for cell in sheet["cells"]:
                cell["ref"] = f"{column_label(cell['col'])}{cell['row'] + 1}"

        result["file_a_name"] = file_a.filename
        result["file_b_name"] = file_b.filename
        return jsonify(result)

    except Exception as exc:  # noqa: BLE001 — 사용자에게 오류 메시지를 전달
        return jsonify({"error": f"파일을 비교하는 중 오류가 발생했습니다: {exc}"}), 500

    finally:
        for tmp in (tmp_a, tmp_b):
            if tmp is not None and os.path.exists(tmp.name):
                os.unlink(tmp.name)


@app.route("/export", methods=["POST"])
def export():
    """두 파일을 비교해 '개발자 diff 뷰'로 색칠된 결과 엑셀을 내려준다."""
    file_a = request.files.get("file_a")
    file_b = request.files.get("file_b")

    if not file_a or not file_b or not file_a.filename or not file_b.filename:
        return jsonify({"error": "두 개의 엑셀 파일을 모두 업로드해 주세요."}), 400

    for f in (file_a, file_b):
        if not _allowed(f.filename):
            return jsonify(
                {"error": f"지원하지 않는 파일 형식입니다: {f.filename} (.xlsx, .xls, .xlsm 만 가능)"}
            ), 400

    tmp_a = tmp_b = tmp_out = None
    try:
        tmp_a = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_b = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_out = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        file_a.save(tmp_a.name)
        file_b.save(tmp_b.name)
        tmp_a.close()
        tmp_b.close()
        tmp_out.close()

        export_diff_xlsx(tmp_a.name, tmp_b.name, tmp_out.name)
        return send_file(
            tmp_out.name,
            as_attachment=True,
            download_name="비교결과.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:  # noqa: BLE001 — 사용자에게 오류 메시지를 전달
        return jsonify({"error": f"결과 엑셀을 만드는 중 오류가 발생했습니다: {exc}"}), 500

    finally:
        for tmp in (tmp_a, tmp_b):
            if tmp is not None and os.path.exists(tmp.name):
                os.unlink(tmp.name)
        # tmp_out 은 send_file 전송 후 OS 임시정리에 맡긴다(전송 중 삭제 방지).


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
