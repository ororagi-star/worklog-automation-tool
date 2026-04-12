from __future__ import annotations

import shutil
import sys
import uuid
import webbrowser
from datetime import date
from pathlib import Path
from threading import Timer

from flask import Flask, Response, flash, redirect, render_template_string, request, url_for
from werkzeug.utils import secure_filename

from automate_worklog import fill_worklog


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
TEMP_ROOT = BASE_DIR / ".tmp_uploads"
DEFAULT_TEMPLATE_FILE = BASE_DIR / "worklog_set.xlsx"
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"

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
        grid-template-columns: minmax(0, 1fr) 300px;
        gap: 22px;
        margin-top: 24px;
        align-items: start;
      }

      form,
      aside,
      .result {
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
        color: #687786;
        font-size: 13px;
        font-weight: 400;
      }

      .notice {
        margin: 16px 0 0;
        border-left: 4px solid #15806d;
        background: #edf7f4;
        padding: 12px;
        color: #26443d;
      }

      .result {
        margin-top: 18px;
        border-color: #b7d9d1;
        background: #eef9f6;
      }

      .path {
        display: block;
        margin-top: 8px;
        overflow-wrap: anywhere;
        color: #153b33;
        font-family: Consolas, "Courier New", monospace;
        font-size: 14px;
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
      <p>출결 파일만 올리면 기본 업무일지 양식으로 결과 엑셀을 만들어 저장합니다.</p>

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
            <span class="hint">이 파일만 선택하면 바로 결과를 만들 수 있습니다.</span>
          </label>

          <label>
            수정 양식 업로드
            <input type="file" name="template_file" accept=".xlsx" />
            <span class="hint">비워두면 앱 폴더의 worklog_set.xlsx를 자동 사용합니다.</span>
          </label>

          <label>
            기준일
            <input type="date" name="target_date" />
            <span class="hint">비워두면 출석 데이터가 있는 마지막 날짜를 자동으로 사용합니다.</span>
          </label>

          <label>
            결과 저장 폴더
            <input type="text" name="output_dir" value="{{ default_output_dir }}" />
            <span class="hint">기본값은 앱 폴더 안의 output 폴더입니다.</span>
          </label>

          <button type="submit">결과 엑셀 만들기</button>
        </form>

        <aside>
          <strong>사용 흐름</strong>
          <ul>
            <li>보통은 low.xlsx만 올리면 됩니다.</li>
            <li>worklog_set.xlsx는 앱 폴더에 둡니다.</li>
            <li>결과는 지정한 폴더에 저장됩니다.</li>
          </ul>
        </aside>
      </div>

      {% if result %}
        <section class="result">
          <strong>생성 완료</strong>
          <p>기준일: {{ result.target_date }}</p>
          <p>결과 파일:</p>
          <span class="path">{{ result.output_file }}</span>
        </section>
      {% endif %}
    </main>
  </body>
</html>
"""


def allowed_xlsx(file_storage) -> bool:
    filename = secure_filename(file_storage.filename or "")
    return filename.lower().endswith(".xlsx")


def render_page(result: dict | None = None, output_dir: Path = DEFAULT_OUTPUT_DIR) -> str:
    return render_template_string(PAGE, default_output_dir=str(output_dir), result=result)


@app.get("/")
def index() -> str:
    return render_page()


@app.post("/generate")
def generate() -> Response | str:
    low_file = request.files.get("low_file")
    template_file = request.files.get("template_file")
    target_date_text = (request.form.get("target_date") or "").strip()
    output_dir_text = (request.form.get("output_dir") or "").strip()

    if not low_file:
        flash("출결 파일 low.xlsx를 선택해주세요.")
        return redirect(url_for("index"))

    has_template_upload = template_file is not None and bool(template_file.filename)
    if not allowed_xlsx(low_file) or (has_template_upload and not allowed_xlsx(template_file)):
        flash("xlsx 파일만 사용할 수 있습니다.")
        return redirect(url_for("index"))

    try:
        target_date = date.fromisoformat(target_date_text) if target_date_text else None
    except ValueError:
        flash("기준일 형식이 올바르지 않습니다.")
        return redirect(url_for("index"))

    output_dir = Path(output_dir_text) if output_dir_text else DEFAULT_OUTPUT_DIR
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        flash(f"결과 저장 폴더를 만들 수 없습니다: {exc}")
        return redirect(url_for("index"))

    TEMP_ROOT.mkdir(parents=True, exist_ok=True)
    temp_dir = TEMP_ROOT / uuid.uuid4().hex
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        low_path = temp_dir / "low.xlsx"
        low_file.save(low_path)

        if has_template_upload:
            template_path = temp_dir / "worklog_set.xlsx"
            template_file.save(template_path)
        else:
            template_path = DEFAULT_TEMPLATE_FILE
            if not template_path.exists():
                flash("기본 양식 worklog_set.xlsx가 앱 폴더에 없습니다.")
                return redirect(url_for("index"))

        temp_output_path = temp_dir / "worklog_result.xlsx"

        try:
            summary = fill_worklog(
                low_file=low_path,
                template_file=template_path,
                output_file=temp_output_path,
                target_date=target_date,
            )
        except Exception as exc:
            flash(f"처리 중 오류가 발생했습니다: {exc}")
            return redirect(url_for("index"))

        final_output_path = output_dir / f"worklog_result_{summary['target_date']}.xlsx"
        shutil.copy2(temp_output_path, final_output_path)

        return render_page(
            output_dir=output_dir,
            result={
                "target_date": summary["target_date"],
                "output_file": str(final_output_path),
            },
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def open_browser() -> None:
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    Timer(1.0, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
