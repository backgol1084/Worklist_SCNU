import streamlit as st
from supabase import create_client, Client
import pandas as pd
from datetime import datetime, date, timedelta
import calendar
import openpyxl
from io import BytesIO

# 1. 페이지 기본 설정 (가장 처음에 위치해야 함)
st.set_page_config(
    page_title="개인 업무 진행 상황 관리",
    page_icon="✅",
    layout="wide"
)

# 2. Supabase 클라이언트 초기화
@st.cache_resource
def init_client():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase: Client = init_client()

# 3. 데이터베이스 연동 함수들
def load_data():
    """DB에서 전체 업무 리스트 불러오기"""
    try:
        response = supabase.table("tasks").select("*").order("id", desc=False).execute()
        return response.data
    except Exception as e:
        st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
        return []

def add_task_db(category, title, subtasks, start_date, due_date):
    """새로운 업무, 세부 체크리스트 및 일정 추가"""
    new_row = {
        "category": category,
        "title": title,
        "status": "대기",
        "subtasks": [{"name": item, "done": False, "memo": "", "start_date": str(start_date) if start_date else None, "end_date": str(due_date) if due_date else None} for item in subtasks],
        "start_date": str(start_date) if start_date else None,
        "due_date": str(due_date) if due_date else None,
        "end_date": None
    }
    supabase.table("tasks").insert(new_row).execute()

def update_task_db(task_id, subtasks, status, start_date=None, due_date=None, end_date=None):
    """업무 정보, 세부 업무, 상태 및 일정 업데이트"""
    update_data = {
        "subtasks": subtasks,
        "status": status
    }
    if start_date is not None:
        update_data["start_date"] = str(start_date) if start_date else None
    if due_date is not None:
        update_data["due_date"] = str(due_date) if due_date else None
    if end_date is not None:
        update_data["end_date"] = str(end_date) if end_date else None

    supabase.table("tasks").update(update_data).eq("id", task_id).execute()

def delete_task_db(task_id):
    """업무 삭제"""
    supabase.table("tasks").delete().eq("id", task_id).execute()

def _update_and_save(task_id, subtasks):
    """진행률 및 상태 자동 계산 후 DB 저장 헬퍼 함수"""
    total = len(subtasks)
    done = sum(1 for s in subtasks if s["done"])
    prog = (done / total) if total > 0 else 0
    
    tasks = load_data()
    curr = next((t for t in tasks if t["id"] == task_id), None)
    current_end_date = curr.get("end_date") if curr else None

    if prog == 1.0:
        status = "완료"
        if not current_end_date:
            current_end_date = str(date.today())
    elif prog > 0:
        status = "진행중"
        current_end_date = None
    else:
        status = "대기"
        current_end_date = None

    update_task_db(task_id, subtasks, status, end_date=current_end_date)

def _clear_widget_keys(task_id):
    """위젯 충돌 방지를 위해 해당 업무의 세션 상태 키 모두 초기화"""
    keys_to_del = [k for k in st.session_state.keys() if f"_{task_id}_" in k]
    for k in keys_to_del:
        if k != "editing_task_id" and k != "show_calendar_modal":
            try:
                del st.session_state[k]
            except:
                pass

if "editing_task_id" not in st.session_state:
    st.session_state.editing_task_id = None

if "show_calendar_modal" not in st.session_state:
    st.session_state.show_calendar_modal = False

# --- 커스텀 CSS ---
st.markdown("""
<style>
    textarea {
        resize: vertical !important;
        min-height: 100px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 앱 UI 레이아웃 시작 ---
st.title("✅ 개인 업무 진행 상황 관리 앱")
st.caption("나만의 간편한 업무 리스트 & 실시간 클라우드 DB 트래커")
st.markdown("---")

CATEGORY_OPTIONS = ["융합", "전공", "기초", "가치", "기타"]

# 4. 사이드바: 새 업무 등록 및 캘린더 버튼
with st.sidebar:
    st.header("➕ 새 업무 등록")
    with st.form("add_task_form", clear_on_submit=True):
        category = st.selectbox("업무 구분", CATEGORY_OPTIONS)
        task_title = st.text_input("업무 제목", placeholder="예: 3분기 실적 보고서 작성")
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            start_date = st.date_input("시작일", value=date.today())
        with col_d2:
            due_date = st.date_input("마감예정일", value=date.today())

        subtasks_text = st.text_area(
            "세부 업무 리스트 (줄바꿈으로 구분)",
            placeholder="자료 수집\n초안 작성\n팀장님 검토\n최종 제출"
        )
        
        submitted = st.form_submit_button("업무 추가하기", use_container_width=True)
        
        if submitted:
            if task_title.strip():
                with st.spinner("새 업무를 등록하는 중입니다..."):
                    sub_items = [s.strip() for s in subtasks_text.split("\n") if s.strip()]
                    add_task_db(category, task_title.strip(), sub_items, start_date, due_date)
                st.success("새 업무가 추가되었습니다!")
                st.rerun()
            else:
                st.warning("업무 제목을 입력해주세요.")

    st.markdown("---")
    st.subheader("📅 캘린더 일정 관리")
    if st.button("🗓️ 월간 캘린더 뷰 열기 / 토글", use_container_width=True, type="primary"):
        st.session_state.show_calendar_modal = not st.session_state.show_calendar_modal
        st.rerun()

# 데이터 불러오기
with st.spinner("데이터를 불러오는 중입니다..."):
    tasks = load_data()

# 5. 월간 캘린더 격자 뷰 영역 (사이드바 버튼 클릭 시 표시)
if st.session_state.show_calendar_modal:
    st.markdown("""
    <div style="padding: 20px; border-radius: 10px; border: 2px solid #2196F3; background-color: #e3f2fd; margin-bottom: 20px;">
        <h3 style="margin-top: 0; color: #0d47a1;">🗓️ 구글 캘린더 스타일 월간 격자 뷰어</h3>
        <p style="color: #333; font-size: 14px;">연도와 월을 선택하여 달력 형태의 격자에서 날짜별로 진행 중인 업무 일정을 한눈에 파악하세요.</p>
    </div>
    """, unsafe_allow_html=True)

    c_col1, c_col2 = st.columns(2)
    with c_col1:
        cal_year = st.number_input("조회 연도", min_value=2020, max_value=2035, value=date.today().year, step=1)
    with c_col2:
        cal_month = st.selectbox("조회 월 선택", list(range(1, 13)), index=date.today().month - 1)

    st.markdown(f"### 📌 {cal_year}년 {cal_month}월 캘린더")

    # 월간 캘린더 매트릭스 생성 (월요일 시작 기준)
    cal_matrix = calendar.monthcalendar(cal_year, cal_month)
    days_header = ["월", "화", "수", "목", "금", "토", "일"]

    # 요일 헤더 렌더링
    head_cols = st.columns(7)
    for idx, day_name in enumerate(days_header):
        head_cols[idx].markdown(f"<div style='text-align: center; font-weight: bold; color: {'#d32f2f' if idx==6 else ('#1976d2' if idx==5 else '#333')};'>{day_name}</div>", unsafe_allow_html=True)

    color_map = {
        "융합": "#e3f2fd", 
        "전공": "#e8f5e9", 
        "기초": "#fff3e0", 
        "가치": "#fce4ec", 
        "기타": "#f3e5f5"
    }

    # 날짜별 격자 렌더링
    for week in cal_matrix:
        week_cols = st.columns(7)
        for i, day in enumerate(week):
            with week_cols[i]:
                if day == 0:
                    st.markdown("<div style='min-height: 110px; background-color: #fcfcfc; border: 1px solid #eee; border-radius: 4px; padding: 6px;'></div>", unsafe_allow_html=True)
                else:
                    curr_date = date(cal_year, cal_month, day)
                    # 해당 날짜에 포함되는 업무 수집
                    day_tasks = []
                    for t in tasks:
                        s_str = t.get("start_date")
                        d_str = t.get("due_date")
                        if s_str and d_str:
                            try:
                                s_dt = datetime.strptime(s_str, "%Y-%m-%d").date()
                                d_dt = datetime.strptime(d_str, "%Y-%m-%d").date()
                                if s_dt <= curr_date <= d_dt:
                                    day_tasks.append(t)
                            except:
                                pass

                    badges_html = ""
                    for dt_task in day_tasks:
                        cat = dt_task.get("category", "기타")
                        bg_color = color_map.get(cat, "#f0f2f6")
                        badges_html += f"<div style='font-size: 11px; background-color: {bg_color}; padding: 2px 5px; margin-bottom: 3px; border-radius: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;' title='{dt_task['title']}'><b>[{cat}]</b> {dt_task['title']}</div>"

                    is_today = (curr_date == date.today())
                    border_style = "border: 2px solid #4CAF50;" if is_today else "border: 1px solid #ddd;"
                    bg_color_cell = "#f1f8e9" if is_today else "white"

                    st.markdown(f"""
                    <div style='min-height: 110px; {border_style} border-radius: 4px; padding: 6px; background-color: {bg_color_cell};'>
                        <div style='font-weight: bold; font-size: 13px; margin-bottom: 5px; color: {"#d32f2f" if i==6 else ("#1976d2" if i==5 else "#333")};'>{day}</div>
                        {badges_html}
                    </div>
                    """, unsafe_allow_html=True)

    st.write("")
    if st.button("❌ 캘린더 뷰 닫기", use_container_width=True):
        st.session_state.show_calendar_modal = False
        st.rerun()

    st.markdown("---")

# 6. 메인 화면: 탭 구성
tab1, tab2 = st.tabs(["📋 업무 진행 & 관리", "📊 한눈에 보기 & 엑셀 다운로드"])

with tab1:
    # --- 상단 편집 패널 ---
    if st.session_state.editing_task_id is not None:
        current_task = next((t for t in tasks if t["id"] == st.session_state.editing_task_id), None)
        if current_task:
            st.markdown(
                f"""
                <div style="padding: 20px; border-radius: 10px; border: 2px solid #4CAF50; background-color: #f1f8e9; margin-bottom: 20px;">
                    <h3 style="margin-top: 0; color: #2e7d32;">⚙️ 세부 업무, 일정 및 메모 편집 모드</h3>
                    <p style="font-size: 16px; font-weight: bold; margin-bottom: 5px;">[{current_task['category']}] {current_task['title']}</p>
                    <p style="color: #555; font-size: 13px;">전체 일정 수정, 세부 업무별 시작일/종료일 지정, 체크, 추진계획 메모 작성, 순서 변경(⬆️, ⬇️), 삭제(❌)를 자유롭게 진행하세요.</p>
                </div>
                """, 
                unsafe_allow_html=True
            )
            
            task_id = current_task["id"]
            
            col_ed1, col_ed2, col_ed3 = st.columns(3)
            with col_ed1:
                cur_start = datetime.strptime(current_task["start_date"], "%Y-%m-%d").date() if current_task.get("start_date") else date.today()
                new_start = st.date_input("전체 시작일 수정", value=cur_start, key=f"edit_start_{task_id}")
            with col_ed2:
                cur_due = datetime.strptime(current_task["due_date"], "%Y-%m-%d").date() if current_task.get("due_date") else date.today()
                new_due = st.date_input("전체 마감예정일 수정", value=cur_due, key=f"edit_due_{task_id}")
            with col_ed3:
                cur_end = datetime.strptime(current_task["end_date"], "%Y-%m-%d").date() if current_task.get("end_date") else None
                new_end = st.date_input("전체 종료일 (실제 완료일)", value=cur_end, key=f"edit_end_{task_id}", help="완료된 경우 자동 입력되거나 직접 수정할 수 있습니다.")

            if new_start != cur_start or new_due != cur_due or new_end != cur_end:
                update_task_db(task_id, current_task["subtasks"], current_task["status"], start_date=new_start, due_date=new_due, end_date=new_end)

            st.markdown("---")
            st.subheader("📝 세부 업무별 체크리스트, 일정 및 추진계획 메모 관리")

            subtasks = []
            default_sub_start = current_task.get("start_date") or str(date.today())
            default_sub_end = current_task.get("due_date") or str(date.today())

            for s in current_task["subtasks"]:
                if isinstance(s, str):
                    subtasks.append({"name": s, "done": False, "memo": "", "start_date": default_sub_start, "end_date": default_sub_end})
                else:
                    item_dict = dict(s)
                    if "memo" not in item_dict: item_dict["memo"] = ""
                    if "start_date" not in item_dict or not item_dict["start_date"]: item_dict["start_date"] = default_sub_start
                    if "end_date" not in item_dict or not item_dict["end_date"]: item_dict["end_date"] = default_sub_end
                    subtasks.append(item_dict)
            
            if not subtasks:
                st.info("등록된 세부 업무가 없습니다. 아래에서 추가해보세요.")
            
            for idx, sub in enumerate(subtasks):
                chk_key = f"edit_chk_{task_id}_{idx}"
                memo_key = f"edit_memo_{task_id}_{idx}"
                sub_start_key = f"edit_sub_start_{task_id}_{idx}"
                sub_end_key = f"edit_sub_end_{task_id}_{idx}"
                
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = sub["done"]
                if memo_key not in st.session_state:
                    st.session_state[memo_key] = sub["memo"]

                st.markdown(f"**[{idx+1}] 세부 항목 컨트롤**")
                
                rc1, rc2, rc3, rc4 = st.columns([5, 1, 1, 1])
                with rc1:
                    new_done = st.checkbox(sub["name"], key=chk_key)
                    if new_done != sub["done"]:
                        subtasks[idx]["done"] = new_done
                        _update_and_save(task_id, subtasks)
                with rc2:
                    if idx > 0:
                        if st.button("⬆️", key=f"edit_up_{task_id}_{idx}"):
                            subtasks[idx], subtasks[idx - 1] = subtasks[idx - 1], subtasks[idx]
                            _update_and_save(task_id, subtasks)
                            _clear_widget_keys(task_id)
                            st.rerun()
                with rc3:
                    if idx < len(subtasks) - 1:
                        if st.button("⬇️", key=f"edit_down_{task_id}_{idx}"):
                            subtasks[idx], subtasks[idx + 1] = subtasks[idx + 1], subtasks[idx]
                            _update_and_save(task_id, subtasks)
                            _clear_widget_keys(task_id)
                            st.rerun()
                with rc4:
                    if st.button("❌", key=f"edit_del_{task_id}_{idx}"):
                        subtasks.pop(idx)
                        _update_and_save(task_id, subtasks)
                        _clear_widget_keys(task_id)
                        st.rerun()

                sc1, sc2 = st.columns(2)
                with sc1:
                    cur_s_date = datetime.strptime(sub["start_date"], "%Y-%m-%d").date() if sub.get("start_date") else date.today()
                    new_sub_start = st.date_input("세부 시작일", value=cur_s_date, key=sub_start_key)
                with sc2:
                    cur_e_date = datetime.strptime(sub["end_date"], "%Y-%m-%d").date() if sub.get("end_date") else date.today()
                    new_sub_end = st.date_input("세부 종료일", value=cur_e_date, key=sub_end_key)

                if str(new_sub_start) != sub.get("start_date") or str(new_sub_end) != sub.get("end_date"):
                    subtasks[idx]["start_date"] = str(new_sub_start)
                    subtasks[idx]["end_date"] = str(new_sub_end)
                    update_task_db(task_id, subtasks, current_task["status"])

                new_memo = st.text_area("추진계획 메모", value=sub["memo"], key=memo_key, placeholder="이 세부 업무와 관련된 추진계획이나 메모를 여러 줄로 자유롭게 입력하세요")
                if new_memo != sub["memo"]:
                    subtasks[idx]["memo"] = new_memo
                    update_task_db(task_id, subtasks, current_task["status"])

                st.markdown("---")
            
            with st.form(key=f"edit_add_form_{task_id}", clear_on_submit=True):
                new_sub_name = st.text_input("새 세부 업무 추가", placeholder="세부 업무명 입력")
                
                acol1, acol2 = st.columns(2)
                with acol1:
                    new_sub_sdate = st.date_input("새 세부 시작일", value=date.today())
                with acol2:
                    new_sub_edate = st.date_input("새 세부 종료일", value=date.today())

                submitted_sub = st.form_submit_button("➕ 세부 업무 추가하기", use_container_width=True)
                if submitted_sub:
                    if new_sub_name.strip():
                        subtasks.append({
                            "name": new_sub_name.strip(),
                            "done": False,
                            "memo": "",
                            "start_date": str(new_sub_sdate),
                            "end_date": str(new_sub_edate)
                        })
                        _update_and_save(task_id, subtasks)
                        _clear_widget_keys(task_id)
                        st.success("세부 업무가 추가되었습니다!")
                        st.rerun()
                    else:
                        st.warning("내용을 입력해주세요.")

            st.write("")
            if st.button("📌 편집 완료 (닫기)", use_container_width=True, type="primary"):
                _clear_widget_keys(task_id)
                st.session_state.editing_task_id = None
                st.rerun()
            
            st.markdown("---")
            st.markdown("<br>", unsafe_allow_html=True)

    if not tasks:
        st.info("등록된 업무가 없습니다. 왼쪽 사이드바에서 새 업무를 추가해보세요!")
    else:
        # --- 필터 및 정렬 컨트롤 영역 (년, 월, 주, 일 별 보기) ---
        st.subheader("🔍 업무 필터 및 정렬 설정")
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            filter_options = ["전체"] + CATEGORY_OPTIONS
            selected_cat = st.selectbox("📂 업무구분 필터", filter_options, key="main_cat_filter")
        with f_col2:
            period_mode = st.selectbox("📅 기간 단위 보기", ["전체 기간", "연간 보기", "월간 보기", "주간 보기", "일간 보기"], key="main_period_mode")
        with f_col3:
            sort_option = st.selectbox(
                "🔄 정렬 기준",
                ["ID 순 (기본)", "시작일 순", "마감예정일 순", "종료일 순"],
                key="main_sort_option"
            )

        target_year = date.today().year
        target_month = date.today().month
        target_date = date.today()

        if period_mode == "연간 보기":
            y_col1 = st.columns(1)[0]
            with y_col1:
                target_year = st.number_input("조회 연도 선택", min_value=2020, max_value=2035, value=date.today().year, step=1)
        elif period_mode == "월간 보기":
            m_col1, m_col2 = st.columns(2)
            with m_col1:
                target_year = st.number_input("조회 연도", min_value=2020, max_value=2035, value=date.today().year, step=1)
            with m_col2:
                target_month = st.selectbox("조회 월 선택", list(range(1, 13)), index=date.today().month - 1)
        elif period_mode == "주간 보기" or period_mode == "일간 보기":
            d_col1 = st.columns(1)[0]
            with d_col1:
                target_date = st.date_input("기준 날짜 선택 (해당 주 또는 일자 조회)", value=date.today(), key="period_기준날짜")

        st.markdown("---")

        status_tab1, status_tab2, status_tab3 = st.tabs(["⏳ 대기 & 진행중", "✅ 완료", "📋 전체 보기"])

        def render_task_cards(target_statuses, tab_prefix):
            res = tasks if selected_cat == "전체" else [t for t in tasks if t.get("category") == selected_cat]
            res = [t for t in res if t["status"] in target_statuses]

            filtered_by_period = []
            for t in res:
                s_str = t.get("start_date")
                d_str = t.get("due_date")
                if not s_str or not d_str:
                    continue
                try:
                    s_dt = datetime.strptime(s_str, "%Y-%m-%d").date()
                    d_dt = datetime.strptime(d_str, "%Y-%m-%d").date()

                    if period_mode == "전체 기간":
                        filtered_by_period.append(t)
                    elif period_mode == "연간 보기":
                        if s_dt.year <= target_year <= d_dt.year:
                            filtered_by_period.append(t)
                    elif period_mode == "월간 보기":
                        last_day = calendar.monthrange(target_year, target_month)[1]
                        m_start = date(target_year, target_month, 1)
                        m_end = date(target_year, target_month, last_day)
                        if not (d_dt < m_start or s_dt > m_end):
                            filtered_by_period.append(t)
                    elif period_mode == "주간 보기":
                        start_of_week = target_date - timedelta(days=target_date.weekday())
                        end_of_week = start_of_week + timedelta(days=6)
                        if not (d_dt < start_of_week or s_dt > end_of_week):
                            filtered_by_period.append(t)
                    elif period_mode == "일간 보기":
                        if s_dt <= target_date <= d_dt:
                            filtered_by_period.append(t)
                except:
                    pass

            res = filtered_by_period

            if target_statuses == ["완료"]:
                res = sorted(res, key=lambda x: x.get("end_date") or "", reverse=True)
            else:
                if sort_option == "시작일 순":
                    res = sorted(res, key=lambda x: x.get("start_date") or "9999-12-31")
                elif sort_option == "마감예정일 순":
                    res = sorted(res, key=lambda x: x.get("due_date") or "9999-12-31")
                elif sort_option == "종료일 순":
                    res = sorted(res, key=lambda x: x.get("end_date") or "9999-12-31")
                else:
                    res = sorted(res, key=lambda x: x["id"])

            if not res:
                st.info("해당 조건 및 기간에 등록된 업무가 없습니다.")
                return

            for task in res:
                subtasks = task["subtasks"]
                task_id = task["id"]
                
                total_sub = len(subtasks)
                done_sub = sum(1 for s in subtasks if isinstance(s, dict) and s.get("done"))
                
                progress = (done_sub / total_sub) if total_sub > 0 else (1.0 if task["status"] == "완료" else 0.0)
                progress_percent = int(progress * 100)

                if total_sub > 0:
                    if progress_percent == 100:
                        new_status = "완료"
                        if not task.get("end_date"):
                            update_task_db(task_id, subtasks, new_status, end_date=str(date.today()))
                            task["end_date"] = str(date.today())
                    elif progress_percent > 0:
                        new_status = "진행중"
                    else:
                        new_status = "대기"
                else:
                    new_status = task["status"]

                if new_status != task["status"]:
                    task["status"] = new_status
                    update_task_db(task_id, subtasks, new_status)

                start_str = task.get("start_date") or "미지정"
                due_str = task.get("due_date") or "미지정"
                end_str = task.get("end_date") or "-"

                with st.container():
                    st.markdown(f"""
                    <div style="padding: 15px; border-radius: 8px; border: 1px solid #e0e0e0; margin-bottom: 5px; background-color: #ffffff;">
                        <span style="background-color: #f0f2f6; padding: 3px 8px; border-radius: 4px; font-size: 12px; font-weight: bold;">{task['category']}</span>
                        <span style="font-size: 16px; font-weight: bold; margin-left: 10px;">{task['title']}</span>
                        <span style="color: #666; font-size: 14px; margin-left: 10px;">({task['status']} - {progress_percent}%)</span>
                        <div style="margin-top: 8px; font-size: 13px; color: #555;">
                            📅 <b>전체 시작일:</b> {start_str} &nbsp;|&nbsp; ⏳ <b>전체 마감예정일:</b> {due_str} &nbsp;|&nbsp; ✅ <b>전체 종료일:</b> {end_str}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    col_bar, col_btn1, col_btn2 = st.columns([4, 1, 1])
                    with col_bar:
                        st.progress(progress)
                    with col_btn1:
                        if st.button("⚙️ 세부업무 & 일정 편집", key=f"edit_btn_{tab_prefix}_{task_id}", use_container_width=True):
                            st.session_state.editing_task_id = task_id
                            st.rerun()
                    with col_btn2:
                        if st.button("🗑️ 전체 삭제", key=f"del_btn_{tab_prefix}_{task_id}", use_container_width=True):
                            with st.spinner("삭제 중..."):
                                delete_task_db(task_id)
                            st.success("삭제되었습니다!")
                            st.rerun()
                    st.markdown("---")

        with status_tab1:
            render_task_cards(["대기", "진행중"], "ongoing")
            
        with status_tab2:
            render_task_cards(["완료"], "completed")
            
        with status_tab3:
            render_task_cards(["대기", "진행중", "완료"], "all")

with tab2:
    st.subheader("📌 전체 업무 통계 대시보드")
    if tasks:
        tot_count = len(tasks)
        completed_count = sum(1 for t in tasks if t["status"] == "완료")
        in_progress_count = sum(1 for t in tasks if t["status"] == "진행중")
        pending_count = sum(1 for t in tasks if t["status"] == "대기")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("전체 업무", f"{tot_count}개")
        m2.metric("진행 중", f"{in_progress_count}개")
        m3.metric("완료됨", f"{completed_count}개")
        m4.metric("대기 중", f"{pending_count}개")

        st.markdown("---")
        st.subheader("📂 업무 구분별 요약")
        cat_map = {cat: {"total": 0, "completed": 0} for cat in CATEGORY_OPTIONS}
        for t in tasks:
            cat = t.get("category", "기타")
            if cat not in cat_map:
                cat_map[cat] = {"total": 0, "completed": 0}
            cat_map[cat]["total"] += 1
            if t["status"] == "완료":
                cat_map[cat]["completed"] += 1

        for cat, val in cat_map.items():
            rate = int((val["completed"] / val["total"]) * 100) if val["total"] > 0 else 0
            st.write(f"**{cat}** ({val['completed']}/{val['total']} 완료)")
            st.progress(rate / 100)

        st.markdown("---")
        st.subheader("📥 전체 업무 리스트 엑셀 다운로드")
        st.caption("다운로드할 업무 상태를 체크박스로 선택하고, 세부 업무별 일정 및 셀 병합이 적용된 엑셀 파일을 다운로드하세요.")

        st.write("**다운로드할 업무 상태 선택**")
        chk_col1, chk_col2, chk_col3 = st.columns(3)
        with chk_col1:
            sel_pending = st.checkbox("대기", value=True)
        with chk_col2:
            sel_progress = st.checkbox("진행중", value=True)
        with chk_col3:
            sel_completed = st.checkbox("완료", value=True)

        selected_statuses = []
        if sel_pending: selected_statuses.append("대기")
        if sel_progress: selected_statuses.append("진행중")
        if sel_completed: selected_statuses.append("완료")

        filtered_export_tasks = [t for t in tasks if t.get("status") in selected_statuses]

        if not filtered_export_tasks:
            st.warning("선택된 상태의 업무가 없습니다. 적어도 하나의 상태를 체크해주세요.")
        else:
            excel_data = []
            for t in filtered_export_tasks:
                subtasks = t["subtasks"]
                total_sub = len(subtasks)
                done_sub = sum(1 for s in subtasks if isinstance(s, dict) and s.get("done"))
                progress_percent = int((done_sub / total_sub) * 100) if total_sub > 0 else (100 if t["status"] == "완료" else 0)
                
                if subtasks:
                    for s in subtasks:
                        if isinstance(s, dict):
                            sub_name = s.get('name', '')
                            sub_done = "완료" if s.get('done') else "진행중"
                            sub_memo = s.get('memo', '')
                            sub_sdate = s.get('start_date', '')
                            sub_edate = s.get('end_date', '')
                        else:
                            sub_name = str(s)
                            sub_done = "진행중"
                            sub_memo = ""
                            sub_sdate = ""
                            sub_edate = ""

                        excel_data.append({
                            "ID": t["id"],
                            "업무 구분": t["category"],
                            "업무 제목": t["title"],
                            "전체 상태": t["status"],
                            "전체 진행률(%)": f"{progress_percent}%",
                            "전체 시작일": t.get("start_date", ""),
                            "전체 마감예정일": t.get("due_date", ""),
                            "전체 종료일": t.get("end_date", ""),
                            "세부 업무명": sub_name,
                            "세부 완료 여부": sub_done,
                            "세부 시작일": sub_sdate,
                            "세부 종료일": sub_edate,
                            "추진계획 메모": sub_memo
                        })
                else:
                    excel_data.append({
                        "ID": t["id"],
                        "업무 구분": t["category"],
                        "업무 제목": t["title"],
                        "전체 상태": t["status"],
                        "전체 진행률(%)": f"{progress_percent}%",
                        "전체 시작일": t.get("start_date", ""),
                        "전체 마감예정일": t.get("due_date", ""),
                        "전체 종료일": t.get("end_date", ""),
                        "세부 업무명": "-",
                        "세부 완료 여부": "-",
                        "세부 시작일": "",
                        "세부 종료일": "",
                        "추진계획 메모": ""
                    })

            df = pd.DataFrame(excel_data)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="업무진행상황")

            output.seek(0)
            wb = openpyxl.load_workbook(output)
            ws = wb.active
            
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center", horizontal="center" if cell.column in [1, 4, 5, 6, 7, 8, 10, 11, 12] else "left")

            merge_cols = [1, 2, 3, 4, 5, 6, 7, 8]
            start_row = 2
            max_row = ws.max_row
            
            while start_row <= max_row:
                current_id = ws.cell(row=start_row, column=1).value
                end_row = start_row
                
                while end_row + 1 <= max_row and ws.cell(row=end_row + 1, column=1).value == current_id:
                    end_row += 1
                
                if end_row > start_row:
                    for col in merge_cols:
                        ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
                
                start_row = end_row + 1

            thin_border = openpyxl.styles.Border(
                left=openpyxl.styles.Side(style='thin', color='CCCCCC'),
                right=openpyxl.styles.Side(style='thin', color='CCCCCC'),
                top=openpyxl.styles.Side(style='thin', color='CCCCCC'),
                bottom=openpyxl.styles.Side(style='thin', color='CCCCCC')
            )
            
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = thin_border

            final_output = BytesIO()
            wb.save(final_output)
            excel_file = final_output.getvalue()

            today_str = datetime.now().strftime("%Y%m%d")
            st.download_button(
                label="📊 세부 일정 및 셀 병합된 엑셀 파일 다운로드 (.xlsx)",
                data=excel_file,
                file_name=f"업무진행상황_달력형반영_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    else:
        st.info("등록된 데이터가 없습니다.")
