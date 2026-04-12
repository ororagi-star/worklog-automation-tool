from __future__ import annotations

import shutil
import uuid
from io import BytesIO
from datetime import date
from pathlib import Path

from flask import Flask, Response, flash, redirect, render_template_string, request, send_file, url_for
from werkzeug.utils import secure_filename

from automate_worklog import fill_worklog


BASE_DIR = Path(__file__).resolve().parent
TEMP_ROOT = BASE_DIR / ".tmp_uploads"

app = Flask(__name__)
app.secret_key = "worklog-local-secret"
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

PAGE = """
<!doctype html>
<html lang="ko">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>업무일지 자동 생성</title>
    <style>
      :root {
        font-family: Arial, "Malgun Gothic", sans-serif;
        color: #18212b;
        background: #f4f7f6;
      }

      * {
        box-sizing: border-box;
      }

      body {
        margin: 0;
      }

      main {
        width: min(980px, calc(100% - 32px));
        margin: 0 auto;
        padding: 40px 0;
      }

      h1 {
        margin: 0 0 10px;
        font-size: 34px;
        line-height: 1.25;
      }

      p {
        margin: 0;
        color: #586776;
        line-height: 1.6;
      }

      .layout {
        display: grid;
        grid-template-columns: minmax(0, 1fr) 280px;
        gap: 22px;
        margin-top: 24px;
        align-items: start;
      }

      form,
      aside {
        border: 1px solid #dce5e2;
        border-radius: 8px;
        background: #ffffff;
        padding: 22px;
      }

      label {
        display: grid;
        gap: 8px;
        margin-bottom: 18px;
        font-weight: 700;
      }

      input {
        width: 100%;
        border: 1px solid #c7d4d0;
        border-radius: 8px;
        padding: 12px;
        font: inherit;
      }

      input[type="file"] {
        background: #f8faf9;
      }

      button {
        width: 100%;
        border: 0;
        border-radius: 8px;
        padding: 13px 16px;
        background: #15806d;
        color: #ffffff;
        font: inherit;
        font-weight: 700;
        cursor: pointer;
      }

      button:hover {
        background: #0f6b5b;
      }

      .hint {
        margin-top: 8px;
        font-size: 13px;
      }

      .notice {
        margin: 16px 0 0;
        border-left: 4px solid #15806d;
        background: #edf7f4;
        padding: 12px;
        color: #26443d;
      }

      ul {
        margin: 12px 0 0;
        padding-left: 18px;
        color: #586776;
        line-height: 1.7;
      }

      @media (max-width: 760px) {
        .layout {
          grid-template-columns: 1fr;
        }

        h1 {
          font-size: 28px;
        }
      }
    </style>
  </head>
  <body>
    <main>
      <h1>업무일지 자동 생성</h1>
      <p>출결 파일과 업무일지 양식을 올리면 완성된 엑셀 파일이 바로 다운로드됩니다.</p>

      {% with messages = get_flashed_messages() %}
        {% if messages %}
          {% for message in messages %}
            <p class="notice">{{ message }}</p>
          {% endfor %}
        {% endif %}
      {% endwith %}

      <div class="layout">
        <form action="{{ url_for('generate') }}" method="post" enctype="multipart/form-data">
          <label>
            출결 파일 low.xlsx
            <input type="file" name="low_file" accept=".xlsx" required />
            <span class="hint">출결 현황 원본 파일을 선택하세요.</span>
          </label>

          <label>
            업무일지 양식 worklog_set.xlsx
            <input type="file" name="template_file" accept=".xlsx" required />
            <span class="hint">4월 시트가 들어있는 업무일지 양식을 선택하세요.</span>
          </label>

          <label>
            기준일
            <input type="date" name="target_date" />
            <span class="hint">비워두면 출석 데이터가 있는 마지막 날짜를 자동으로 사용합니다.</span>
          </label>

          <button type="submit">결과 엑셀 만들기</button>
        </form>

        <aside>
          <strong>처리 방식</strong>
          <ul>
            <li>업로드한 파일은 처리 중에만 사용됩니다.</li>
            <li>프로그램명 매칭 후 일계와 월계를 입력합니다.</li>
            <li>결과 파일은 바로 다운로드됩니다.</li>
          </ul>
        </aside>
      </div>
    </main>
  </body>
</html>
"""


def allowed_xlsx(file_storage) -> bool:
    filename = secure_filename(file_storage.filename or "")
    return filename.lower().endswith(".xlsx")


@app.get("/")
def index() -> str:
    return render_template_string(PAGE)


@app.post("/generate")
def generate() -> Response:
    low_file = request.files.get("low_file")
    template_file = request.files.get("template_file")
    target_date_text = (request.form.get("target_date") or "").strip()

    if not low_file or not template_file:
        flash("두 개의 엑셀 파일을 모두 선택해주세요.")
        return redirect(url_for("index"))

    if not allowed_xlsx(low_file) or not allowed_xlsx(template_file):
        flash("xlsx 파일만 업로드할 수 있습니다.")
        return redirect(url_for("index"))

    try:
        target_date = date.fromisoformat(target_date_text) if target_date_text else None
    except ValueError:
        flash("기준일 형식이 올바르지 않습니다.")
        return redirect(url_for("index"))

    TEMP_ROOT.mkdir(parents=True, exist_ok=True)

    temp_dir = TEMP_ROOT / uuid.uuid4().hex
    temp_dir.mkdir(parents=True, exist_ok=True)
    try:
        low_path = temp_dir / "low.xlsx"
        template_path = temp_dir / "worklog_set.xlsx"
        output_path = temp_dir / "worklog_result.xlsx"

        low_file.save(low_path)
        template_file.save(template_path)

        try:
            summary = fill_worklog(
                low_file=low_path,
                template_file=template_path,
                output_file=output_path,
                target_date=target_date,
            )
        except Exception as exc:
            flash(f"처리 중 오류가 발생했습니다: {exc}")
            return redirect(url_for("index"))

        download_name = f"worklog_result_{summary['target_date']}.xlsx"
        result_bytes = BytesIO(output_path.read_bytes())
        result_bytes.seek(0)
        return send_file(
            result_bytes,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
