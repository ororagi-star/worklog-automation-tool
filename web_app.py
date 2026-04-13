from __future__ import annotations

import shutil
import sys
import uuid
import webbrowser
from datetime import date
from pathlib import Path
from threading import Timer

from flask import Flask, Response, flash, jsonify, redirect, render_template_string, request, url_for
from werkzeug.utils import secure_filename

from automate_worklog import fill_worklog_dates


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
TEMP_ROOT = BASE_DIR / ".tmp_uploads"
DEFAULT_TEMPLATE_FILES = [
    ("worklog_set1", BASE_DIR / "worklog_set1.xlsx"),
    ("worklog_set2", BASE_DIR / "worklog_set2.xlsx"),
]
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"
VERSION_FILE = BASE_DIR / "version.txt"
APP_VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "dev"

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
        background: #f3f7f8;
      }

      * {
        box-sizing: border-box;
      }

      body {
        margin: 0;
      }

      main {
        width: min(1040px, calc(100% - 32px));
        margin: 0 auto;
        padding: 42px 0;
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
        grid-template-columns: minmax(0, 1fr) 320px;
        gap: 24px;
        margin-top: 26px;
        align-items: start;
      }

      form,
      aside,
      .result {
        border: 1px solid #dce5e2;
        border-radius: 8px;
        background: #ffffff;
        padding: 24px;
        box-shadow: 0 16px 45px rgba(31, 47, 62, 0.08);
      }

      label,
      .field-group {
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

      .calendar-panel {
        border: 1px solid #cfe0dc;
        border-radius: 8px;
        background: linear-gradient(180deg, #ffffff, #f6fbfa);
        padding: 14px;
      }

      .calendar-head {
        display: grid;
        grid-template-columns: 42px minmax(0, 1fr) 42px;
        gap: 8px;
        align-items: center;
        margin-bottom: 12px;
      }

      .calendar-title {
        text-align: center;
        color: #162832;
        font-size: 18px;
        font-weight: 800;
      }

      .calendar-nav,
      .calendar-day,
      .calendar-clear {
        border: 1px solid #c4d6d2;
        border-radius: 8px;
        background: #ffffff;
        color: #20323d;
      }

      .calendar-nav {
        min-height: 38px;
        padding: 0;
        font-size: 20px;
      }

      .calendar-grid,
      .calendar-weekdays {
        display: grid;
        grid-template-columns: repeat(7, minmax(0, 1fr));
        gap: 6px;
      }

      .calendar-weekdays {
        margin-bottom: 6px;
        color: #6a7780;
        font-size: 12px;
        font-weight: 800;
        text-align: center;
      }

      .calendar-day {
        aspect-ratio: 1;
        min-width: 0;
        padding: 0;
        font-size: 14px;
        font-weight: 700;
      }

      .calendar-day:hover {
        background: #eaf4f1;
      }

      .calendar-day.is-muted {
        color: #a2adb4;
        background: #f7f9fa;
      }

      .calendar-day.is-today {
        border-color: #14816e;
      }

      .calendar-day.is-selected {
        border-color: #176f62;
        background: #176f62;
        color: #ffffff;
      }

      .calendar-selected {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        min-height: 34px;
        margin-top: 12px;
      }

      .date-chip {
        border-radius: 8px;
        background: #e5f2ef;
        padding: 6px 9px;
        color: #17483f;
        font-size: 13px;
        font-weight: 700;
      }

      .calendar-actions {
        display: grid;
        grid-template-columns: 1fr;
        margin-top: 10px;
      }

      .calendar-clear {
        padding: 9px 12px;
      }

      .field-row {
        display: grid;
        grid-template-columns: minmax(0, 1fr) 132px;
        gap: 10px;
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

      .secondary {
        border: 1px solid #b7c9c5;
        background: #ffffff;
        color: #20323d;
      }

      .secondary:hover {
        background: #edf4f3;
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

      .version {
        display: inline-flex;
        margin-top: 12px;
        border: 1px solid #d5e3e0;
        border-radius: 8px;
        padding: 6px 10px;
        background: #ffffff;
        color: #566772;
        font-size: 13px;
        font-weight: 700;
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

        .field-row {
          grid-template-columns: 1fr;
        }
      }
    </style>
  </head>
  <body>
    <main>
      <h1>업무일지 자동 생성</h1>
      <p>출결 파일만 올리면 기본 업무일지 양식으로 결과 엑셀을 만들어 저장합니다.</p>
      <span class="version">버전 {{ app_version }}</span>

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

          <div class="field-group">
            <span>기준일</span>
            <input type="hidden" id="target-dates" name="target_dates" />
            <div class="calendar-panel">
              <div class="calendar-head">
                <button type="button" class="calendar-nav" id="calendar-prev" aria-label="이전 달">‹</button>
                <div class="calendar-title" id="calendar-title"></div>
                <button type="button" class="calendar-nav" id="calendar-next" aria-label="다음 달">›</button>
              </div>
              <div class="calendar-weekdays" aria-hidden="true">
                <span>일</span>
                <span>월</span>
                <span>화</span>
                <span>수</span>
                <span>목</span>
                <span>금</span>
                <span>토</span>
              </div>
              <div class="calendar-grid" id="calendar-grid"></div>
              <div class="calendar-selected" id="calendar-selected">
                <span class="date-chip">선택 없음</span>
              </div>
              <div class="calendar-actions">
                <button type="button" class="calendar-clear" id="calendar-clear">선택 초기화</button>
              </div>
            </div>
            <span class="hint">여러 날짜를 클릭해서 선택할 수 있습니다. 선택하지 않으면 출결 파일의 마지막 출석일을 사용합니다.</span>
          </div>

          <div class="field-group">
            <span>결과 저장 폴더</span>
            <div class="field-row">
              <input type="text" id="output-dir" name="output_dir" value="{{ default_output_dir }}" readonly />
              <button type="button" class="secondary" id="choose-output-dir">폴더 선택</button>
            </div>
            <span class="hint">선택하지 않으면 앱 폴더 안의 output 폴더에 저장합니다.</span>
          </div>

          <button type="submit">결과 엑셀 만들기</button>
        </form>

        <aside>
          <strong>사용 흐름</strong>
          <ul>
            <li>보통은 low.xlsx만 올리면 됩니다.</li>
            <li>양식 1, 2는 앱 폴더에서 자동으로 찾습니다.</li>
            <li>기준일은 달력에서 여러 날짜를 클릭해 고릅니다.</li>
            <li>결과는 양식별 파일 2개로 저장됩니다.</li>
          </ul>
        </aside>
      </div>

      {% if result %}
        <section class="result">
          <strong>생성 완료</strong>
          <p>기준일: {{ result.target_date }}</p>
          <p>결과 파일:</p>
          {% for output_file in result.output_files %}
            <span class="path">{{ output_file }}</span>
          {% endfor %}
        </section>
      {% endif %}
    </main>
    <script>
      const targetDates = document.querySelector("#target-dates");
      const calendarTitle = document.querySelector("#calendar-title");
      const calendarGrid = document.querySelector("#calendar-grid");
      const calendarSelected = document.querySelector("#calendar-selected");
      const calendarPrev = document.querySelector("#calendar-prev");
      const calendarNext = document.querySelector("#calendar-next");
      const calendarClear = document.querySelector("#calendar-clear");
      const outputDir = document.querySelector("#output-dir");
      const chooseOutputDir = document.querySelector("#choose-output-dir");

      const selectedDates = new Set();
      const today = new Date();
      let visibleYear = today.getFullYear();
      let visibleMonth = today.getMonth();

      function toDateKey(year, month, day) {
        const monthText = String(month + 1).padStart(2, "0");
        const dayText = String(day).padStart(2, "0");
        return `${year}-${monthText}-${dayText}`;
      }

      function renderSelectedDates() {
        const dates = Array.from(selectedDates).sort();
        targetDates.value = dates.join(",");
        calendarSelected.innerHTML = "";

        if (dates.length === 0) {
          const emptyChip = document.createElement("span");
          emptyChip.className = "date-chip";
          emptyChip.textContent = "선택 없음";
          calendarSelected.appendChild(emptyChip);
          return;
        }

        for (const dateKey of dates) {
          const chip = document.createElement("span");
          chip.className = "date-chip";
          chip.textContent = dateKey;
          calendarSelected.appendChild(chip);
        }
      }

      function renderCalendar() {
        calendarTitle.textContent = `${visibleYear}년 ${visibleMonth + 1}월`;
        calendarGrid.innerHTML = "";

        const firstDay = new Date(visibleYear, visibleMonth, 1);
        const startOffset = firstDay.getDay();
        const daysInMonth = new Date(visibleYear, visibleMonth + 1, 0).getDate();
        const prevMonthDays = new Date(visibleYear, visibleMonth, 0).getDate();
        const todayKey = toDateKey(today.getFullYear(), today.getMonth(), today.getDate());

        for (let index = 0; index < 42; index += 1) {
          const dayButton = document.createElement("button");
          dayButton.type = "button";
          dayButton.className = "calendar-day";

          let cellYear = visibleYear;
          let cellMonth = visibleMonth;
          let cellDay = index - startOffset + 1;

          if (cellDay <= 0) {
            cellMonth -= 1;
            if (cellMonth < 0) {
              cellMonth = 11;
              cellYear -= 1;
            }
            cellDay = prevMonthDays + cellDay;
            dayButton.classList.add("is-muted");
          } else if (cellDay > daysInMonth) {
            cellDay -= daysInMonth;
            cellMonth += 1;
            if (cellMonth > 11) {
              cellMonth = 0;
              cellYear += 1;
            }
            dayButton.classList.add("is-muted");
          }

          const dateKey = toDateKey(cellYear, cellMonth, cellDay);
          dayButton.textContent = cellDay;
          dayButton.dataset.date = dateKey;

          if (dateKey === todayKey) {
            dayButton.classList.add("is-today");
          }

          if (selectedDates.has(dateKey)) {
            dayButton.classList.add("is-selected");
          }

          dayButton.addEventListener("click", () => {
            if (selectedDates.has(dateKey)) {
              selectedDates.delete(dateKey);
            } else {
              selectedDates.add(dateKey);
            }
            renderSelectedDates();
            renderCalendar();
          });

          calendarGrid.appendChild(dayButton);
        }
      }

      calendarPrev.addEventListener("click", () => {
        visibleMonth -= 1;
        if (visibleMonth < 0) {
          visibleMonth = 11;
          visibleYear -= 1;
        }
        renderCalendar();
      });

      calendarNext.addEventListener("click", () => {
        visibleMonth += 1;
        if (visibleMonth > 11) {
          visibleMonth = 0;
          visibleYear += 1;
        }
        renderCalendar();
      });

      calendarClear.addEventListener("click", () => {
        selectedDates.clear();
        renderSelectedDates();
        renderCalendar();
      });

      renderSelectedDates();
      renderCalendar();

      chooseOutputDir.addEventListener("click", async () => {
        chooseOutputDir.disabled = true;
        chooseOutputDir.textContent = "선택 중";

        try {
          const response = await fetch("{{ url_for('select_output_dir') }}", { method: "POST" });
          const data = await response.json();

          if (data.path) {
            outputDir.value = data.path;
          } else if (data.error) {
            alert(data.error);
          }
        } catch (error) {
          alert("폴더 선택 창을 열 수 없습니다.");
        } finally {
          chooseOutputDir.disabled = false;
          chooseOutputDir.textContent = "폴더 선택";
        }
      });
    </script>
  </body>
</html>
"""


def allowed_xlsx(file_storage) -> bool:
    filename = secure_filename(file_storage.filename or "")
    return filename.lower().endswith(".xlsx")


def parse_target_dates(value: str) -> list[date] | None:
    if not value.strip():
        return None

    parts = [part.strip() for part in value.replace("\n", ",").split(",") if part.strip()]
    if not parts:
        return None

    return [date.fromisoformat(part) for part in parts]


def default_template_path(default_path: Path) -> Path:
    if not default_path.exists():
        raise FileNotFoundError(f"기본 양식 {default_path.name}가 앱 폴더에 없습니다.")

    return default_path


def render_page(result: dict | None = None, output_dir: Path = DEFAULT_OUTPUT_DIR) -> str:
    return render_template_string(
        PAGE,
        app_version=APP_VERSION,
        default_output_dir=str(output_dir),
        result=result,
    )


@app.get("/")
def index() -> str:
    return render_page()


@app.post("/select-output-dir")
def select_output_dir() -> Response:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected_dir = filedialog.askdirectory(
            initialdir=str(DEFAULT_OUTPUT_DIR.parent),
            title="결과 저장 폴더 선택",
        )
        root.destroy()
    except Exception as exc:
        return jsonify({"error": f"폴더 선택 창을 열 수 없습니다: {exc}"}), 500

    return jsonify({"path": selected_dir})


@app.post("/generate")
def generate() -> Response | str:
    low_file = request.files.get("low_file")
    target_dates_text = (request.form.get("target_dates") or "").strip()
    output_dir_text = (request.form.get("output_dir") or "").strip()

    if not low_file:
        flash("출결 파일 low.xlsx를 선택해주세요.")
        return redirect(url_for("index"))

    if not allowed_xlsx(low_file):
        flash("xlsx 파일만 사용할 수 있습니다.")
        return redirect(url_for("index"))

    try:
        target_dates = parse_target_dates(target_dates_text)
    except ValueError:
        flash("기준일 형식이 올바르지 않습니다. 예: 2026-04-09, 2026-04-10")
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

        try:
            templates = [
                (
                    "worklog_set1",
                    default_template_path(DEFAULT_TEMPLATE_FILES[0][1]),
                ),
                (
                    "worklog_set2",
                    default_template_path(DEFAULT_TEMPLATE_FILES[1][1]),
                ),
            ]

            output_files = []
            target_date_labels = []
            for template_label, template_path in templates:
                temp_output_path = temp_dir / f"{template_label}_result.xlsx"
                summary = fill_worklog_dates(
                    low_file=low_path,
                    template_file=template_path,
                    output_file=temp_output_path,
                    target_dates=target_dates,
                )

                date_part = "_".join(summary["target_dates"])
                final_output_path = output_dir / f"{template_label}_result_{date_part}.xlsx"
                shutil.copy2(temp_output_path, final_output_path)
                output_files.append(str(final_output_path))
                target_date_labels = summary["target_dates"]
        except Exception as exc:
            flash(f"처리 중 오류가 발생했습니다: {exc}")
            return redirect(url_for("index"))

        return render_page(
            output_dir=output_dir,
            result={
                "target_date": ", ".join(target_date_labels),
                "output_files": output_files,
            },
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def open_browser() -> None:
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    Timer(1.0, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
