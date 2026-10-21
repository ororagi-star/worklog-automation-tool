from __future__ import annotations

import hmac
import json
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

from course_catalog import add_category, delete_category, place_course, remove_course_everywhere, rename_category
from worklog_config import SESSION_RULES, check_source_ownership, resolve_sources, validate_config


PLACEMENTS = {"미지정": [], "업무일지 1": ["worklog_set1"], "업무일지 2": ["worklog_set2"],
              "업무일지 1·2": ["worklog_set1", "worklog_set2"]}


def placement(config, course_id):
    keys = [t["key"] for t in config["templates"] if any(i["course_id"] == course_id for i in t["items"])]
    return next(label for label, values in PLACEMENTS.items() if set(keys) == set(values))


def commit_change(store, config, revision, programs):
    check_source_ownership(config, programs)
    store.save(config, revision)
    st.session_state.settings_saved = True
    st.session_state.result_zip = None
    st.session_state.pop("editing_course_id", None)
    st.rerun()


def save_action(store, revision, programs, action):
    try:
        commit_change(store, action(), revision, programs)
    except ValueError as exc:
        st.error(str(exc))
    except Exception:
        st.error("저장하지 못했습니다. 연결을 확인하고 다시 시도해 주세요.")


def close_edit():
    st.session_state.pop("editing_course_id", None)


@st.dialog("수업 수정", width="large", on_dismiss=close_edit)
def edit_course(store, config, revision, programs, course):
    with st.form(f"edit_course_{course['id']}_{revision}"):
        cols = st.columns([2, 3, 2])
        category = cols[0].selectbox("카테고리", config["categories"],
            index=config["categories"].index(course["category"]), accept_new_options=True)
        name = cols[1].text_input("수업명", value=course["name"])
        current = placement(config, course["id"])
        destination = cols[2].selectbox("업무일지", list(PLACEMENTS), index=list(PLACEMENTS).index(current))
        with st.expander("추가 설정"):
            options = sorted(programs | set(course["sources"]))
            sources = st.multiselect("연결할 원본 수업", options, default=course["sources"])
            excluded = st.checkbox("출결 집계 제외", value=course["mode"] == "집계 제외")
            col1, col2 = st.columns(2)
            target_sessions = col1.number_input("목표 건수", min_value=0, step=1, value=course["target_sessions"])
            target_people = col2.number_input("목표 인원", min_value=0, step=1, value=course["target_people"])
            rule = st.selectbox("건수 기준", SESSION_RULES, index=SESSION_RULES.index(course["session_rule"]))
        save = st.form_submit_button("저장", type="primary")
        remove = st.form_submit_button("수업 삭제")
    if save:
        values = {**course, "category": category, "name": name, "sources": sources,
                  "mode": "집계 제외" if excluded else ("직접 연결" if sources else "자동 연결"),
                  "target_sessions": target_sessions, "target_people": target_people, "session_rule": rule}
        save_action(store, revision, programs, lambda: place_course(config, values, PLACEMENTS[destination]))
    if remove:
        save_action(store, revision, programs, lambda: remove_course_everywhere(config, course["id"]))


def render_settings(store, config, revision, updated_at, programs, admin_password):
    with st.expander("수업 관리", expanded=False):
        if getattr(store, "session_only", False):
            st.info("DB 연결 전입니다. 수정한 설정은 현재 접속 중에만 유지됩니다. 보관하려면 추가 설정에서 설정 백업을 내려받으세요.")
        saved = st.session_state.pop("settings_saved", False)
        if saved or "settings_snapshot" not in st.session_state:
            st.session_state.settings_snapshot = (config, revision, updated_at)
        current_revision = revision
        config, revision, updated_at = st.session_state.settings_snapshot
        if "courses" not in config:
            del st.session_state.settings_snapshot
            st.rerun()
        if current_revision != revision:
            st.warning("다른 창에서 내용이 바뀌었습니다. 추가 설정에서 새로 불러와 주세요.")
        if saved:
            st.success("저장했습니다.")
        editable = not store.database_url and not admin_password
        if admin_password:
            password = st.text_input("관리자 비밀번호", type="password", key="settings_password")
            editable = hmac.compare_digest(password.encode(), admin_password.encode())
        elif store.database_url:
            st.info("관리자 비밀번호가 설정되지 않아 조회만 가능합니다.")

        with st.form(f"add_course_{revision}"):
            cols = st.columns([2, 3, 2, 1], vertical_alignment="bottom")
            category = cols[0].selectbox("카테고리", config["categories"], accept_new_options=True,
                                       placeholder="선택하거나 새 이름 입력")
            name = cols[1].text_input("수업명", placeholder="예: 한글교실 초급")
            destination = cols[2].selectbox("업무일지", list(PLACEMENTS), index=1)
            add = cols[3].form_submit_button("추가", type="primary", disabled=not editable, use_container_width=True)
        if add and editable:
            values = {"category": category, "name": name, "order": len(config["courses"]) + 1}
            save_action(store, revision, programs, lambda: place_course(config, values, PLACEMENTS[destination]))

        with st.popover("카테고리 편집"):
            with st.form(f"category_add_{revision}"):
                new_category = st.text_input("새 카테고리")
                add_category_clicked = st.form_submit_button("카테고리 추가", disabled=not editable)
            if add_category_clicked and editable:
                save_action(store, revision, programs, lambda: add_category(config, new_category))
            if config["categories"]:
                selected = st.selectbox("변경할 카테고리", config["categories"])
                with st.form(f"category_edit_{selected}_{revision}"):
                    renamed = st.text_input("변경할 이름", value=selected)
                    rename = st.form_submit_button("이름 변경", disabled=not editable)
                    remove = st.form_submit_button("카테고리 삭제", disabled=not editable)
                if rename and editable:
                    save_action(store, revision, programs, lambda: rename_category(config, selected, renamed))
                if remove and editable:
                    save_action(store, revision, programs, lambda: delete_category(config, selected))

        search = st.text_input("수업 찾기", placeholder="카테고리 또는 수업명 검색", label_visibility="collapsed")
        courses = [course for course in config["courses"] if search.strip() in f"{course['category']} {course['name']}"]
        st.caption(f"등록된 수업 {len(courses)}개")
        headings = st.columns([2, 4, 2, 1])
        for col, label in zip(headings, ["카테고리", "수업명", "업무일지", ""]):
            col.write(label)
        for course in courses:
            cols = st.columns([2, 4, 2, 1], vertical_alignment="center")
            cols[0].write(course["category"])
            cols[1].write(course["name"])
            cols[2].caption(placement(config, course["id"]))
            if cols[3].button("수정", key=f"edit_{course['id']}", disabled=not editable):
                st.session_state.editing_course_id = course["id"]
        editing = next((course for course in config["courses"] if course["id"] == st.session_state.get("editing_course_id")), None)
        if editing and editable:
            edit_course(store, config, revision, programs, editing)
        if not courses:
            st.info("등록된 수업이 없습니다.")

        with st.expander("추가 설정"):
            updated = datetime.fromisoformat(updated_at).astimezone(ZoneInfo("Asia/Seoul"))
            location = "Neon / PostgreSQL" if store.database_url else ("브라우저 임시 설정" if getattr(store, "session_only", False) else "로컬 저장 · Neon 미연결")
            st.caption(f"{location} | 설정 {revision} | {updated:%Y-%m-%d %H:%M}")
            if st.button("저장된 설정 다시 불러오기"):
                del st.session_state.settings_snapshot
                st.rerun()
            if programs:
                linked = {name for t in config["templates"] for item in t["items"] if item["enabled"] for name in resolve_sources(item, programs)}
                unlinked = sorted(programs - linked)
                if unlinked:
                    st.warning("연결되지 않은 원본 수업: " + ", ".join(unlinked))
            st.download_button("설정 백업", json.dumps(config, ensure_ascii=False, indent=2).encode("utf-8"),
                file_name=f"worklog_settings_v{revision}.json", mime="application/json")
            backup = st.file_uploader("설정 백업 불러오기", type=["json"], disabled=not editable)
            if st.button("백업으로 설정 복원", disabled=not editable or backup is None) and editable:
                save_action(store, revision, programs, lambda: validate_config(json.loads(backup.getvalue())))
            previous = [row[0] for row in store.history() if row[0] != revision]
            if previous:
                restore_revision = st.selectbox("이전 설정 버전", previous)
                if st.button("이전 버전으로 복원", disabled=not editable) and editable:
                    try:
                        store.restore(restore_revision, revision)
                    except ValueError as exc:
                        st.error(str(exc))
                    except Exception:
                        st.error("복원하지 못했습니다. 저장소 연결을 확인해 주세요.")
                    else:
                        st.session_state.settings_saved = True
                        st.rerun()
