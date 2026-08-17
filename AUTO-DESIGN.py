# -*- coding: utf-8 -*-
"""
DWG-BOM 정합성 검토 AI
BOM(재질/가공공법) ↔ 도면(PDF 텍스트) 자동 비교 MVP
UI 스타일: JOINT-AI-APP-6 계열 (다크 콘솔 테마, JetBrains Mono, 오렌지 포인트)
"""

import io
import re
import json
import difflib
from datetime import datetime

import streamlit as st
import pandas as pd

# 선택적 의존성 (없어도 앱은 동작해야 함 - MVP는 graceful degrade)
try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

try:
    from groq import Groq
    GROQ_OK = True
except ImportError:
    GROQ_OK = False


# =========================================================
# 0. 기본 설정
# =========================================================
st.set_page_config(
    page_title="DWG-BOM 정합성 검토 AI",
    page_icon="※",
    layout="wide",
    initial_sidebar_state="expanded",
)

OWNER_PASSWORD = st.secrets.get("OWNER_PASSWORD", "nt1234")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "")
TEMP_PWD_SHEET_ID = st.secrets.get("temp_pwd_sheet_id", "")

if "lang" not in st.session_state:
    st.session_state["lang"] = "ko"
if "authed" not in st.session_state:
    st.session_state["authed"] = False
if "material_map" not in st.session_state:
    # 표준 재질 코드 매핑 테이블 (별칭 -> 표준코드). 사이드바에서 편집 가능.
    st.session_state["material_map"] = pd.DataFrame(
        [
            {"표준코드": "SCM435", "별칭(콤마구분)": "SCM435, SCM435H, SCM 435"},
            {"표준코드": "SUS304", "별칭(콤마구분)": "SUS304, STS304, SUS 304, STS 304"},
            {"표준코드": "S45C", "별칭(콤마구분)": "S45C, S 45 C, SM45C"},
            {"표준코드": "SPCC", "별칭(콤마구분)": "SPCC, SPCC-SD"},
        ]
    )
if "review_history" not in st.session_state:
    st.session_state["review_history"] = []


# =========================================================
# 1. 다국어 라벨
# =========================================================
LANG_DICT = {
    "ko": {
        "app_title": "DWG-BOM 정합성 검토 AI",
        "app_sub": "도면(PDF) × BOM 재질/가공공법 자동 비교",
        "login_title": "접속 비밀번호를 입력하세요",
        "login_btn": "접속",
        "login_fail": "비밀번호가 올바르지 않습니다.",
        "tab_upload": "1. 문서 업로드",
        "tab_map": "2. 재질 표준코드표",
        "tab_result": "3. 비교 결과",
        "tab_history": "4. 검토 이력",
        "bom_upload": "BOM 파일 업로드 (xlsx / csv)",
        "dwg_upload": "도면 PDF 업로드 (여러 개 가능)",
        "col_map_header": "BOM 컬럼 매핑",
        "col_partno": "품번 컬럼",
        "col_material": "재질 컬럼",
        "col_process": "가공공법 컬럼",
        "run_btn": "비교 실행",
        "no_bom": "BOM 파일을 먼저 업로드하세요.",
        "no_dwg": "도면 PDF를 먼저 업로드하세요.",
        "material_map_desc": "재질 표기가 달라도 같은 재질로 인식할 별칭을 등록하세요. (예: SUS304 = STS304)",
        "save_map": "표준코드표 저장",
        "map_saved": "저장되었습니다.",
        "result_partno": "품번",
        "result_bom_mat": "BOM 재질",
        "result_dwg_mat": "도면 인식 재질",
        "result_match": "재질 일치",
        "result_process": "BOM 가공공법",
        "result_conf": "신뢰도",
        "result_comment": "비고",
        "badge_exact": "정확일치",
        "badge_std": "표준코드일치",
        "badge_low": "유사도낮음",
        "badge_none": "불일치/미검출",
        "no_result": "아직 비교 결과가 없습니다. 1번 탭에서 문서를 업로드하고 비교를 실행하세요.",
        "export_btn": "결과 엑셀 다운로드",
        "pdf_missing": "pdfplumber가 설치되어 있지 않아 PDF 텍스트 추출을 사용할 수 없습니다. requirements.txt를 확인하세요.",
        "llm_check": "애매한 재질 표기는 LLM으로 재확인",
        "history_empty": "저장된 검토 이력이 없습니다.",
        "sidebar_lang": "언어 / Language",
        "raw_text_expander": "도면에서 추출한 원문 보기 (검증용)",
        "manual_edit_note": "자동 인식이 틀렸다면 직접 수정할 수 있습니다.",
        "logout": "로그아웃",
    },
    "en": {
        "app_title": "DWG-BOM Consistency Checker AI",
        "app_sub": "Automated comparison of drawing (PDF) vs BOM material / process",
        "login_title": "Enter access password",
        "login_btn": "Sign in",
        "login_fail": "Incorrect password.",
        "tab_upload": "1. Upload Documents",
        "tab_map": "2. Material Code Map",
        "tab_result": "3. Comparison Result",
        "tab_history": "4. Review History",
        "bom_upload": "Upload BOM file (xlsx / csv)",
        "dwg_upload": "Upload drawing PDF(s)",
        "col_map_header": "BOM Column Mapping",
        "col_partno": "Part No. column",
        "col_material": "Material column",
        "col_process": "Process column",
        "run_btn": "Run Comparison",
        "no_bom": "Please upload a BOM file first.",
        "no_dwg": "Please upload drawing PDFs first.",
        "material_map_desc": "Register aliases that should be treated as the same material (e.g. SUS304 = STS304).",
        "save_map": "Save code map",
        "map_saved": "Saved.",
        "result_partno": "Part No.",
        "result_bom_mat": "BOM Material",
        "result_dwg_mat": "Drawing Material",
        "result_match": "Match",
        "result_process": "BOM Process",
        "result_conf": "Confidence",
        "result_comment": "Comment",
        "badge_exact": "Exact",
        "badge_std": "Std-Mapped",
        "badge_low": "Low-Sim",
        "badge_none": "Mismatch/Not found",
        "no_result": "No results yet. Upload documents in tab 1 and run comparison.",
        "export_btn": "Download result (Excel)",
        "pdf_missing": "pdfplumber is not installed, so PDF text extraction is unavailable.",
        "llm_check": "Re-check ambiguous materials with LLM",
        "history_empty": "No saved review history.",
        "sidebar_lang": "언어 / Language",
        "raw_text_expander": "View extracted drawing text (for verification)",
        "manual_edit_note": "You can manually correct results if auto-detection is wrong.",
        "logout": "Log out",
    },
}


def t(key: str) -> str:
    return LANG_DICT[st.session_state["lang"]].get(key, key)


# =========================================================
# 2. 다크 콘솔 테마 CSS (JOINT 계열과 통일)
# =========================================================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700;800&display=swap');

html, body, [class*="css"]  {
    font-family: 'JetBrains Mono', monospace;
}
.stApp {
    background-color: #0f0f0f;
    color: #e6e6e6;
}
h1, h2, h3, h4 {
    color: #f5f5f5 !important;
    font-family: 'JetBrains Mono', monospace;
}
section[data-testid="stSidebar"] {
    background-color: #161616;
    border-right: 1px solid #2a2a2a;
}
div.stButton > button {
    background-color: #ff9f1c;
    color: #0f0f0f;
    font-weight: 700;
    border: none;
    border-radius: 4px;
    padding: 0.5rem 1.2rem;
}
div.stButton > button:hover {
    background-color: #ffb347;
    color: #0f0f0f;
}
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 0.82rem;
    font-weight: 700;
}
.badge-green  { background-color: #1e5b32; color: #7CFFA0; border: 1px solid #2fae5c; }
.badge-yellow { background-color: #5b4d1e; color: #ffe27a; border: 1px solid #c9a52f; }
.badge-orange { background-color: #5b3a1e; color: #ffb877; border: 1px solid #d47a2f; }
.badge-red    { background-color: #5b1e1e; color: #ff8a8a; border: 1px solid #d43a3a; }

.card {
    background-color: #161616;
    border: 1px solid #2a2a2a;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.8rem;
}
hr { border-color: #2a2a2a; }
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# 3. Google Sheets 임시 비밀번호 저장 (선택적 - 없으면 세션 전용)
# =========================================================
def _get_ws():
    if not (GSPREAD_OK and TEMP_PWD_SHEET_ID and "gcp_service_account" in st.secrets):
        return None
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(TEMP_PWD_SHEET_ID)
        return sh.sheet1
    except Exception:
        return None


def _load_temp_pwds():
    ws = _get_ws()
    if ws is None:
        return st.session_state.get("_temp_pwds", [])
    try:
        vals = ws.col_values(1)
        return [v for v in vals if v]
    except Exception:
        return st.session_state.get("_temp_pwds", [])


def _save_temp_pwds(pwds):
    st.session_state["_temp_pwds"] = pwds
    ws = _get_ws()
    if ws is None:
        return
    try:
        ws.update([[p] for p in pwds])  # clear() 없이 바로 update (데이터 유실 방지)
    except Exception:
        pass


# =========================================================
# 4. 로그인 화면
# =========================================================
def login_screen():
    col1, col2 = st.columns([1, 1])
    with col1:
        st.selectbox(
            t("sidebar_lang"), ["ko", "en"],
            index=0 if st.session_state["lang"] == "ko" else 1,
            key="_lang_select",
            on_change=lambda: st.session_state.update(lang=st.session_state["_lang_select"]),
            label_visibility="collapsed",
        )

    st.markdown(f"## {t('app_title')}")
    st.caption(t("app_sub"))
    st.markdown("---")

    with st.form("login_form"):
        pwd = st.text_input(t("login_title"), type="password")
        submitted = st.form_submit_button(t("login_btn"))
        if submitted:
            valid_pwds = [OWNER_PASSWORD] + _load_temp_pwds()
            if pwd in valid_pwds:
                st.session_state["authed"] = True
                st.rerun()
            else:
                st.error(t("login_fail"))


# =========================================================
# 5. 재질 매칭 엔진
# =========================================================
def normalize(s: str) -> str:
    if not isinstance(s, str):
        return ""
    return re.sub(r"[\s\-_/]", "", s).upper()


def build_alias_lookup(map_df: pd.DataFrame) -> dict:
    """별칭(정규화) -> 표준코드"""
    lookup = {}
    for _, row in map_df.iterrows():
        std = str(row.get("표준코드", "")).strip()
        if not std:
            continue
        aliases = str(row.get("별칭(콤마구분)", "")).split(",")
        for a in aliases:
            a = a.strip()
            if a:
                lookup[normalize(a)] = std
    return lookup


def match_material(bom_mat: str, dwg_mat: str, alias_lookup: dict):
    """
    반환: (일치여부, 신뢰도등급, 코멘트)
    신뢰도등급: exact / std / low / none
    """
    if not dwg_mat:
        return False, "none", "도면에서 재질을 인식하지 못했습니다."

    n_bom, n_dwg = normalize(bom_mat), normalize(dwg_mat)

    if n_bom == n_dwg and n_bom != "":
        return True, "exact", "표기 완전 일치"

    std_bom = alias_lookup.get(n_bom)
    std_dwg = alias_lookup.get(n_dwg)
    if std_bom and std_dwg and std_bom == std_dwg:
        return True, "std", f"표준코드 매핑 일치 ({std_bom})"

    ratio = difflib.SequenceMatcher(None, n_bom, n_dwg).ratio()
    if ratio >= 0.6:
        return False, "low", f"유사도 {ratio*100:.0f}% - 수동 확인 필요"

    return False, "none", "재질 표기 불일치"


def llm_recheck(bom_mat: str, dwg_mat: str) -> str:
    """애매한 케이스만 LLM에게 판단시킴. 실패 시 빈 문자열 반환."""
    if not (GROQ_OK and GROQ_API_KEY):
        return ""
    try:
        client = Groq(api_key=GROQ_API_KEY)
        prompt = (
            f'다음 두 재질 표기가 같은 재질을 의미하는지 판단해줘.\n'
            f'A: "{bom_mat}"\nB: "{dwg_mat}"\n'
            f'"동일" 또는 "다름" 중 하나로만 답하고, 이유를 15자 이내로 덧붙여줘. '
            f'형식: 동일|이유  또는  다름|이유'
        )
        resp = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=50,
            temperature=0,
        )
        return resp.choices[0].message.content.strip()
    except Exception:
        return ""


# =========================================================
# 6. 도면 PDF 텍스트 추출 + 재질 인식
# =========================================================
MATERIAL_KEYWORDS = [
    r"재\s*질\s*[:：]?\s*([A-Za-z0-9\-\.]+)",
    r"MAT'?L\s*[:：]?\s*([A-Za-z0-9\-\.]+)",
    r"MATERIAL\s*[:：]?\s*([A-Za-z0-9\-\.]+)",
]
PARTNO_KEYWORDS = [
    r"품\s*번\s*[:：]?\s*([A-Za-z0-9\-]+)",
    r"PART\s*NO\.?\s*[:：]?\s*([A-Za-z0-9\-]+)",
    r"DWG\s*NO\.?\s*[:：]?\s*([A-Za-z0-9\-]+)",
]


def extract_pdf_text(file) -> str:
    if not PDF_OK:
        return ""
    text = ""
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
    except Exception:
        pass
    return text


def guess_field(text: str, patterns) -> str:
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


# =========================================================
# 7. 메인 화면
# =========================================================
def main_app():
    with st.sidebar:
        st.markdown(f"### {t('app_title')}")
        st.selectbox(
            t("sidebar_lang"), ["ko", "en"],
            index=0 if st.session_state["lang"] == "ko" else 1,
            key="_lang_select2",
            on_change=lambda: st.session_state.update(lang=st.session_state["_lang_select2"]),
        )
        st.markdown("---")
        if st.button(t("logout")):
            st.session_state["authed"] = False
            st.rerun()

    st.markdown(f"## {t('app_title')}")
    st.caption(t("app_sub"))

    tab1, tab2, tab3, tab4 = st.tabs(
        [t("tab_upload"), t("tab_map"), t("tab_result"), t("tab_history")]
    )

    # ---------------- Tab 1: 업로드 ----------------
    with tab1:
        if not PDF_OK:
            st.warning(t("pdf_missing"))

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"**{t('bom_upload')}**")
            bom_file = st.file_uploader(
                t("bom_upload"), type=["xlsx", "csv"], label_visibility="collapsed"
            )
        with col_b:
            st.markdown(f"**{t('dwg_upload')}**")
            dwg_files = st.file_uploader(
                t("dwg_upload"), type=["pdf"], accept_multiple_files=True,
                label_visibility="collapsed",
            )

        bom_df = None
        if bom_file is not None:
            try:
                if bom_file.name.lower().endswith(".csv"):
                    bom_df = pd.read_csv(bom_file)
                else:
                    bom_df = pd.read_excel(bom_file)
                st.dataframe(bom_df.head(20), use_container_width=True)
            except Exception as e:
                st.error(f"BOM 파일을 읽는 중 오류: {e}")

        if bom_df is not None:
            st.markdown(f"##### {t('col_map_header')}")
            c1, c2, c3 = st.columns(3)
            cols = list(bom_df.columns)
            with c1:
                partno_col = st.selectbox(t("col_partno"), cols, key="partno_col")
            with c2:
                material_col = st.selectbox(t("col_material"), cols, key="material_col")
            with c3:
                process_col = st.selectbox(t("col_process"), cols, key="process_col")

            use_llm = st.checkbox(t("llm_check"), value=bool(GROQ_API_KEY))

            if st.button(t("run_btn"), type="primary"):
                if not dwg_files:
                    st.warning(t("no_dwg"))
                else:
                    with st.spinner("..."):
                        results = run_comparison(
                            bom_df, partno_col, material_col, process_col,
                            dwg_files, use_llm,
                        )
                    st.session_state["last_result"] = results
                    st.session_state["review_history"].append(
                        {
                            "시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            "BOM파일": bom_file.name,
                            "도면수": len(dwg_files),
                            "검토건수": len(results),
                        }
                    )
                    st.success(f"완료: {len(results)}건 비교")
        else:
            st.info(t("no_bom"))

    # ---------------- Tab 2: 재질 표준코드표 ----------------
    with tab2:
        st.caption(t("material_map_desc"))
        edited = st.data_editor(
            st.session_state["material_map"],
            num_rows="dynamic",
            use_container_width=True,
            key="material_map_editor",
        )
        if st.button(t("save_map")):
            st.session_state["material_map"] = edited
            st.success(t("map_saved"))

    # ---------------- Tab 3: 비교 결과 ----------------
    with tab3:
        results = st.session_state.get("last_result")
        if not results:
            st.info(t("no_result"))
        else:
            render_results(results)

    # ---------------- Tab 4: 검토 이력 ----------------
    with tab4:
        hist = st.session_state["review_history"]
        if not hist:
            st.info(t("history_empty"))
        else:
            st.dataframe(pd.DataFrame(hist), use_container_width=True)


def run_comparison(bom_df, partno_col, material_col, process_col, dwg_files, use_llm):
    alias_lookup = build_alias_lookup(st.session_state["material_map"])

    # 도면들에서 품번 -> (재질, 원문) 추출
    dwg_index = {}
    for f in dwg_files:
        text = extract_pdf_text(f)
        pn = guess_field(text, PARTNO_KEYWORDS)
        mat = guess_field(text, MATERIAL_KEYWORDS)
        key = normalize(pn) if pn else normalize(f.name.rsplit(".", 1)[0])
        dwg_index[key] = {"material": mat, "raw_text": text, "filename": f.name}

    results = []
    for _, row in bom_df.iterrows():
        part_no = str(row.get(partno_col, "")).strip()
        bom_mat = str(row.get(material_col, "")).strip()
        process = str(row.get(process_col, "")).strip()

        dwg_entry = dwg_index.get(normalize(part_no), {})
        dwg_mat = dwg_entry.get("material", "")

        match, grade, comment = match_material(bom_mat, dwg_mat, alias_lookup)

        if grade == "low" and use_llm:
            llm_out = llm_recheck(bom_mat, dwg_mat)
            if llm_out.startswith("동일"):
                match, grade = True, "std"
                comment = "LLM 재확인: " + llm_out.split("|", 1)[-1].strip()
            elif llm_out.startswith("다름"):
                grade = "none"
                comment = "LLM 재확인: " + llm_out.split("|", 1)[-1].strip()

        results.append(
            {
                "part_no": part_no,
                "bom_material": bom_mat,
                "dwg_material": dwg_mat or "-",
                "match": match,
                "grade": grade,
                "process": process,
                "comment": comment,
                "raw_text": dwg_entry.get("raw_text", ""),
                "dwg_filename": dwg_entry.get("filename", ""),
            }
        )
    return results


BADGE_CLASS = {"exact": "badge-green", "std": "badge-green", "low": "badge-yellow", "none": "badge-red"}
BADGE_LABEL_KEY = {"exact": "badge_exact", "std": "badge_std", "low": "badge_low", "none": "badge_none"}


def render_results(results):
    n_total = len(results)
    n_ok = sum(1 for r in results if r["match"])
    n_low = sum(1 for r in results if r["grade"] == "low")
    n_bad = sum(1 for r in results if r["grade"] == "none")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("전체" if st.session_state["lang"] == "ko" else "Total", n_total)
    c2.metric("일치" if st.session_state["lang"] == "ko" else "Match", n_ok)
    c3.metric("확인필요" if st.session_state["lang"] == "ko" else "Review", n_low)
    c4.metric("불일치" if st.session_state["lang"] == "ko" else "Mismatch", n_bad)

    st.markdown("---")

    for r in results:
        badge_cls = BADGE_CLASS[r["grade"]]
        badge_label = t(BADGE_LABEL_KEY[r["grade"]])
        st.markdown(
            f"""
<div class="card">
  <b>{t('result_partno')}: {r['part_no']}</b>
  &nbsp;&nbsp;<span class="badge {badge_cls}">{badge_label}</span>
  <br><br>
  {t('result_bom_mat')}: <code>{r['bom_material']}</code>
  &nbsp;&nbsp;→&nbsp;&nbsp;
  {t('result_dwg_mat')}: <code>{r['dwg_material']}</code>
  <br>
  {t('result_process')}: {r['process']}
  <br>
  {t('result_comment')}: {r['comment']}
</div>
""",
            unsafe_allow_html=True,
        )
        if r["raw_text"]:
            with st.expander(f"{t('raw_text_expander')} — {r['dwg_filename']}"):
                st.text(r["raw_text"][:2000])

    st.markdown("---")
    export_df = pd.DataFrame(
        [
            {
                t("result_partno"): r["part_no"],
                t("result_bom_mat"): r["bom_material"],
                t("result_dwg_mat"): r["dwg_material"],
                t("result_match"): "O" if r["match"] else "X",
                t("result_process"): r["process"],
                t("result_conf"): t(BADGE_LABEL_KEY[r["grade"]]),
                t("result_comment"): r["comment"],
            }
            for r in results
        ]
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="result")
    st.download_button(
        t("export_btn"),
        data=buf.getvalue(),
        file_name=f"dwg_bom_check_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================================================
# 8. 엔트리 포인트
# =========================================================
if not st.session_state["authed"]:
    login_screen()
else:
    main_app()
