from __future__ import annotations

import shutil
import sys
import tempfile
from datetime import date
from io import BytesIO
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook

from automate_worklog import fill_worklog_dates, find_date_columns, latest_populated_date


BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DEFAULT_TEMPLATE_FILES = [
    ("worklog_set1", BASE_DIR / "worklog_set1.xlsx"),
    ("worklog_set2", BASE_DIR / "worklog_set2.xlsx"),
]
VERSION_FILE = BASE_DIR / "version.txt"
APP_VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "dev"


st.set_page_config(
    page_title="업무일지 자동 생성",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      .stApp {
        background: #f5f7f8;
      }

      [data-testid="stHeader"] {
        background: rgba(245, 247, 248, 0.85);
      }

      .hero {
        border: 1px solid #d9e2e0;
        border-radius: 8px;
        background: #ffffff;
        padding: 24px;
        box-shadow: 0 16px 45px rgba(31, 47, 62, 0.08);
      }

      .hero h1 {
        margin: 0 0 8px;
        color: #17242e;
        font-size: 34px;
        line-height: 1.25;
      }

      .hero p {
        margin: 0;
        color: #586776;
        font-size: 16px;
        line-height: 1.6;
      }

      .version-pill {
        display: inline-flex;
        margin-top: 14px;
        border: 1px solid #d5e3e0;
        border-radius: 8px;
        padding: 6px 10px;
        background: #f7fbfa;
        color: #566772;
        font-size: 13px;
        font-weight: 700;
      }

      .date-chip {
        display: inline-flex;
        margin: 0 6px 6px 0;
        border-radius: 8px;
        background: #e5f2ef;
        padding: 6px 10px;
        color: #17483f;
        font-size: 13px;
        font-weight: 700;
      }

      div[data-testid="stMetric"] {
        border: 1px solid #d9e2e0;
        border-radius: 8px;
        background: #ffffff;
        padding: 14px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def uploaded_bytes(uploaded_file) -> bytes:
    data = uploaded_file.getvalue()
    if not data:
        raise ValueError("업로드한 파일이 비어 있습니다.")
    return data


def available_dates_from_low(low_data: bytes) -> tuple[list[date], date | None]:
    workbook = load_workbook(BytesIO(low_data), data_only=True)
    sheet = workbook.active
    date_columns = find_date_columns(sheet)
    available_dates = sorted(date_columns)
    default_date = latest_populated_date(sheet, date_columns) if available_dates else None
    return available_dates, default_date


def render_selected_dates(selected_dates: list[date]) -> None:
    if not selected_dates:
        st.caption("선택된 날짜가 없습니다. 생성 시 마지막 출석일을 자동으로 사용합니다.")
        return

    chips = "".join(f"<span class='date-chip'>{day.isoformat()}</span>" for day in selected_dates)
    st.markdown(chips, unsafe_allow_html=True)


def generate_files(low_data: bytes, selected_dates: list[date]) -> list[tuple[str, bytes]]:
    missing_templates = [template_path.name for _, template_path in DEFAULT_TEMPLATE_FILES if not template_path.exists()]
    if missing_templates:
        missing_text = ", ".join(missing_templates)
        raise FileNotFoundError(f"앱 폴더에 기본 양식 파일이 없습니다: {missing_text}")

    with tempfile.TemporaryDirectory() as temp_dir_text:
        temp_dir = Path(temp_dir_text)
        low_path = temp_dir / "low.xlsx"
        low_path.write_bytes(low_data)

        output_files: list[tuple[str, bytes]] = []
        target_dates = selected_dates or None

        for template_label, template_path in DEFAULT_TEMPLATE_FILES:
            temp_output_path = temp_dir / f"{template_label}_result.xlsx"
            summary = fill_worklog_dates(
                low_file=low_path,
                template_file=template_path,
                output_file=temp_output_path,
                target_dates=target_dates,
            )
            date_part = "_".join(summary["target_dates"])
            output_name = f"{template_label}_result_{date_part}.xlsx"
            output_files.append((output_name, temp_output_path.read_bytes()))

    return output_files


st.markdown(
    f"""
    <section class="hero">
      <h1>업무일지 자동 생성</h1>
      <p>출결 파일을 올리고 기준일을 고르면 업무일지 결과 엑셀을 바로 내려받을 수 있습니다.</p>
      <span class="version-pill">버전 {APP_VERSION}</span>
    </section>
    """,
    unsafe_allow_html=True,
)

left, right = st.columns([1.45, 1], gap="large")

with left:
    st.subheader("파일과 날짜")
    low_file = st.file_uploader("출결 파일 low.xlsx", type=["xlsx"])

    low_data: bytes | None = None
    selected_dates: list[date] = []

    if low_file is not None:
        try:
            low_data = uploaded_bytes(low_file)
            available_dates, default_date = available_dates_from_low(low_data)
        except Exception as exc:
            st.error(f"출결 파일을 읽을 수 없습니다: {exc}")
            available_dates = []
            default_date = None

        if available_dates:
            default_selection = [default_date] if default_date else []
            selected_dates = st.multiselect(
                "기준일",
                options=available_dates,
                default=default_selection,
                format_func=lambda day: day.isoformat(),
                help="여러 날짜를 선택할 수 있습니다. 비워두면 마지막 출석일을 사용합니다.",
            )
            render_selected_dates(selected_dates)
        else:
            st.info("출결 파일에서 날짜 컬럼을 찾으면 기준일 선택이 표시됩니다.")
    else:
        st.info("먼저 출결 파일을 올려주세요.")

with right:
    st.subheader("안내")
    st.markdown("**사용 흐름**")
    st.write("1. 출결 파일을 올립니다.")
    st.write("2. 기준일을 하나 이상 선택합니다.")
    st.write("3. 결과 엑셀 만들기를 누릅니다.")
    st.write("4. 생성된 파일 2개를 내려받습니다.")
    st.info("기본 양식은 웹앱에 포함된 worklog_set1.xlsx, worklog_set2.xlsx를 사용합니다.")

st.divider()

can_generate = low_data is not None
if st.button("결과 엑셀 만들기", type="primary", use_container_width=True, disabled=not can_generate):
    try:
        output_files = generate_files(
            low_data=low_data or b"",
            selected_dates=sorted(selected_dates),
        )
    except Exception as exc:
        st.error(f"처리 중 오류가 발생했습니다: {exc}")
    else:
        st.success("생성이 완료되었습니다.")
        cols = st.columns(2)
        for index, (output_name, output_data) in enumerate(output_files):
            with cols[index % 2]:
                st.metric(f"결과 파일 {index + 1}", output_name)
                st.download_button(
                    label="다운로드",
                    data=output_data,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
