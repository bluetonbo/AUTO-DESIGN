# -*- coding: utf-8 -*-
"""
DWG-BOM 정합성 검토 AI (AUTO-DESIGN)
BOM(재질/가공공법) ↔ 도면(PDF 텍스트) 자동 비교
UI/인프라: JOINT-AI-APP-6 원본 기준으로 통일
  - 다크 콘솔 CSS (glass-card, Inter/JetBrains Mono, 탭/익스팬더 스타일)
  - 로그인 화면 (배지+타이틀 글래스카드, 언어선택, 임시비번 만료시스템)
  - Google Sheets 기반 임시 비밀번호 영구 저장 + 사이드바 관리 패널
  - 4단계 신뢰도 배지 (render_confidence_badge)
  - API 키 등 민감정보 마스킹 (_sanitize_secret_text)
"""

import io
import re
import json
import difflib
from datetime import datetime, timedelta

import streamlit as st
import pandas as pd

# 선택적 의존성 (없어도 앱은 동작해야 함 - graceful degrade)
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

try:
    from google import genai as gemini_genai
    from google.genai import types as gemini_types
    GEMINI_OK = True
except ImportError:
    GEMINI_OK = False


# =========================================================
# 0. 기본 설정
# =========================================================
st.set_page_config(
    layout="wide",
    page_title="AUTO-DESIGN AI - DWG-BOM Consistency Suite",
)

OWNER_PWD = st.secrets.get("OWNER_PASSWORD", "nt1234")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY", "")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")
_TEMP_PWD_WORKSHEET = "temp_pwd_store"
_DEFAULT_TEMP_PWD = "design1234"  # Sheets가 비어있을 때 최초 1회 자동 생성되는 기본 임시 비번 (7일)


# =========================================================
# 1. 다국어 사전
# =========================================================
LANG_DICT = {
    "KO": {
        "console": "데이터 컨트롤",
        "title": "DWG-BOM 정합성 검토 AI",
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
        "no_result": "아직 비교 결과가 없습니다. 1번 탭에서 문서를 업로드하고 비교를 실행하세요.",
        "export_btn": "결과 엑셀 다운로드",
        "pdf_missing": "pdfplumber가 설치되어 있지 않아 PDF 텍스트 추출을 사용할 수 없습니다. requirements.txt를 확인하세요.",
        "llm_check": "애매한 재질 표기는 LLM으로 재확인",
        "vision_check": "스캔 도면(이미지 PDF) Vision 인식 사용 (텍스트 없는 도면만 대상, 느림)",
        "vision_not_configured": "GEMINI_API_KEY가 설정되지 않아 Vision 인식이 비활성 상태입니다.",
        "vision_badge": "Vision 인식",
        "history_empty": "저장된 검토 이력이 없습니다.",
        "history_sheets_error": "검토 이력 저장소(Sheets) 연결 오류 — 이번 세션 이력만 임시로 표시됩니다.",
        "history_session_only": "Sheets 미설정 상태입니다. 이 목록은 현재 세션에서만 유지되며 새로고침 시 사라집니다.",
        "hist_time": "시각", "hist_bom_file": "BOM파일", "hist_dwg_count": "도면수",
        "hist_score": "신뢰도(%)", "hist_grade": "등급",
        "tab_cad": "5. 3D 물성치 비교",
        "cad_upload": "CAD 물성치 엑셀 업로드 (xlsx/csv)",
        "cad_col_partno": "품번 컬럼", "cad_col_weight": "무게 컬럼 (kg)",
        "cad_col_x": "가로(X) 컬럼 (mm)", "cad_col_y": "세로(Y) 컬럼 (mm)", "cad_col_z": "높이(Z) 컬럼 (mm)",
        "cad_col_none": "사용 안 함",
        "cad_compare_bom": "BOM 기준값과 비교",
        "cad_compare_prev": "이전 리비전과 비교 (Sheets 이력 기반)",
        "cad_compare_dwg": "도면 표기 치수와 비교 (1번 탭 실행 필요, 타이틀블록에 명시된 값만 인식)",
        "cad_no_dwg_run": "1번 탭에서 비교 실행을 먼저 하셔야 도면 치수를 사용할 수 있습니다.",
        "cad_bom_need_cols": "BOM 기준값 컬럼 매핑 (BOM에 해당 컬럼이 없으면 '사용 안 함')",
        "cad_tolerance": "허용오차 (%)",
        "cad_run_btn": "물성치 비교 실행",
        "cad_no_file": "CAD 물성치 파일을 먼저 업로드하세요.",
        "cad_no_bom": "BOM 기준값과 비교하려면 1번 탭에서 BOM을 먼저 업로드하세요.",
        "cad_no_checks": "선택된 비교 대상이 없습니다 (BOM 비교 또는 이전 리비전 비교 중 하나 이상 선택하세요).",
        "cad_result_title": "물성치 비교 결과",
        "cad_overall": "종합 신뢰도",
        "cad_no_prev": "이전 리비전 이력 없음 (최초 업로드)",
        "cad_first_upload_note": "이번 업로드 값은 다음 비교를 위해 이력에 저장됩니다.",
        "cad_saved_ok": "물성치 이력 저장 완료.",
        "raw_text_expander": "도면에서 추출한 원문 보기 (검증용)",
        "metric_total": "전체", "metric_match": "일치", "metric_review": "확인필요", "metric_bad": "불일치",
        "run_done": "완료: {n}건 비교",
    },
    "EN": {
        "console": "CONTROL CONSOLE",
        "title": "DWG-BOM Consistency Checker AI",
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
        "no_result": "No results yet. Upload documents in tab 1 and run comparison.",
        "export_btn": "Download result (Excel)",
        "pdf_missing": "pdfplumber is not installed, so PDF text extraction is unavailable.",
        "llm_check": "Re-check ambiguous materials with LLM",
        "vision_check": "Use Vision recognition for scanned (image) drawings (text-less drawings only, slow)",
        "vision_not_configured": "GEMINI_API_KEY is not set, so Vision recognition is disabled.",
        "vision_badge": "Vision-recognized",
        "history_empty": "No saved review history.",
        "history_sheets_error": "Review-history storage (Sheets) connection error — showing this session's history only.",
        "history_session_only": "Sheets not configured. This list only persists for the current session and will disappear on refresh.",
        "hist_time": "Time", "hist_bom_file": "BOM File", "hist_dwg_count": "Drawing Count",
        "hist_score": "Confidence(%)", "hist_grade": "Grade",
        "tab_cad": "5. 3D Property Compare",
        "cad_upload": "Upload CAD property Excel (xlsx/csv)",
        "cad_col_partno": "Part No. column", "cad_col_weight": "Weight column (kg)",
        "cad_col_x": "Width(X) column (mm)", "cad_col_y": "Depth(Y) column (mm)", "cad_col_z": "Height(Z) column (mm)",
        "cad_col_none": "Not used",
        "cad_compare_bom": "Compare vs BOM reference values",
        "cad_compare_prev": "Compare vs previous revision (Sheets history)",
        "cad_compare_dwg": "Compare vs drawing-labeled dimensions (requires tab-1 run; only explicit title-block values are recognized)",
        "cad_no_dwg_run": "Run the comparison in tab 1 first to use drawing dimensions.",
        "cad_bom_need_cols": "BOM reference column mapping (choose 'Not used' if BOM lacks that column)",
        "cad_tolerance": "Tolerance (%)",
        "cad_run_btn": "Run Property Comparison",
        "cad_no_file": "Please upload a CAD property file first.",
        "cad_no_bom": "Upload a BOM in tab 1 first to compare against BOM reference values.",
        "cad_no_checks": "No comparison target selected (pick at least one of BOM or previous-revision compare).",
        "cad_result_title": "Property Comparison Result",
        "cad_overall": "Overall Confidence",
        "cad_no_prev": "No previous revision history (first upload)",
        "cad_first_upload_note": "This upload's values will be saved to history for future comparisons.",
        "cad_saved_ok": "Property history saved.",
        "raw_text_expander": "View extracted drawing text (for verification)",
        "metric_total": "Total", "metric_match": "Match", "metric_review": "Review", "metric_bad": "Mismatch",
        "run_done": "Done: {n} compared",
    },
}

LOGIN_TEXT = {
    "KO": {
        "badge": "AUTO-DESIGN AI SYSTEM",
        "title": "AI 기반 DWG-BOM 정합성 자동 검토 시스템",
        "pwd_label": "비밀번호 입력",
        "auth_btn": "시스템 접속",
        "invalid": "비밀번호가 올바르지 않습니다.",
        "owner_only": "소유자 비번은 사용할 수 없습니다.",
        "enter_pwd_warn": "비번을 입력하세요.",
        "no_expiry": "무기한 사용 가능",
        "temp_mgr_title": "임시 비번 관리",
        "new_temp_pwd": "새 임시 비번",
        "expiry_period": "유효기간",
        "add_btn": "추가",
        "added_msg": "추가됨: ",
        "registered_pwds": "등록된 임시 비번",
        "no_registered": "등록된 임시 비번 없음",
        "left_label": "잔여",
        "expired_label": "만료됨",
        "unlimited_label": "무기한",
        "exp_1d": "1일", "exp_3d": "3일", "exp_7d": "7일", "exp_30d": "30일", "exp_none": "무기한",
    },
    "EN": {
        "badge": "AUTO-DESIGN AI SYSTEM",
        "title": "AI-based Automated DWG-BOM Consistency Review System",
        "pwd_label": "Enter Password",
        "auth_btn": "Authenticate",
        "invalid": "Invalid credentials.",
        "owner_only": "Cannot use owner password.",
        "enter_pwd_warn": "Enter a password.",
        "no_expiry": "No Expiry",
        "temp_mgr_title": "Temp Password Manager",
        "new_temp_pwd": "New Temp Password",
        "expiry_period": "Expires in",
        "add_btn": "Add",
        "added_msg": "Added: ",
        "registered_pwds": "Registered Passwords",
        "no_registered": "No temp passwords.",
        "left_label": "Left",
        "expired_label": "Expired",
        "unlimited_label": "No Expiry",
        "exp_1d": "1 Day", "exp_3d": "3 Days", "exp_7d": "7 Days", "exp_30d": "30 Days", "exp_none": "No Expiry",
    },
}

if "lang" not in st.session_state:
    st.session_state["lang"] = "KO"


def t(key: str) -> str:
    return LANG_DICT[st.session_state["lang"]].get(key, key)


# =========================================================
# 2. 콘솔 스타일 CSS (JOINT-AI-APP-6 원본과 동일)
# =========================================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Inter:wght@400;500;600;700&display=swap');

    .stApp {
        background-color: #0f0f0f !important;
        color: #ececec !important;
        font-family: 'Inter', sans-serif;
    }

    [data-testid="stAppViewContainer"] .main .block-container {
        max-width: 100% !important;
        width: 100% !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }

    [data-testid="stSidebar"] {
        background-color: #161616 !important;
        border-right: 1px solid #262626;
        min-width: 360px !important;
    }

    .scrollable-box {
        max-height: 400px;
        overflow-y: auto;
        padding: 15px;
        background-color: #161616;
        border: 1px solid #2e2e2e;
        border-radius: 6px;
        color: #ececec;
    }

    h1, h2, h3, h4 {
        font-family: 'Inter', sans-serif;
        font-weight: 600 !important;
        letter-spacing: -0.01em;
        color: #f2f2f2 !important;
    }

    .glass-card {
        background: #1a1a1a;
        border: 1px solid #2e2e2e;
        border-radius: 6px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }

    .glass-card-title {
        color: #ff9f1c;
        font-size: 0.9rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 12px;
        padding-bottom: 6px;
        border-bottom: 1px solid #262626;
    }

    .stButton>button, .stDownloadButton>button {
        height: 2.8rem !important;
        font-size: 0.9rem !important;
        border-radius: 4px !important;
        background: #10b981 !important;
        color: #ffffff !important;
        font-weight: 600;
        border: none !important;
        transition: all 0.2s ease;
        width: 100%;
    }

    label, .stTextInput label, .stSelectbox label, .stSlider label,
    .stNumberInput label, .stRadio label, .stFileUploader label,
    [data-testid="stWidgetLabel"] p {
        color: #b8b8b8 !important;
    }
    ::placeholder {
        color: #8a8a8a !important;
        opacity: 1 !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] span,
    [data-testid="stFileUploaderDropzoneInstructions"] small {
        color: #aaaaaa !important;
    }
    .stCaption, [data-testid="stCaptionContainer"] {
        color: #aaaaaa !important;
    }

    [data-testid="stProgress"] p,
    [data-testid="stProgress"] span,
    [data-testid="stProgress"] div,
    [data-testid="stProgress"] *,
    [data-testid="stSidebar"] [data-testid="stProgress"] * {
        color: #e8e8e8 !important;
    }
    [data-testid="stAlert"] p,
    [data-testid="stAlert"] span,
    [data-testid="stAlert"] div,
    [data-testid="stAlert"] *,
    [data-testid="stSidebar"] [data-testid="stAlert"] *,
    [data-testid="stAlertContentInfo"] p,
    [data-testid="stAlertContentWarning"] p,
    [data-testid="stAlertContentError"] p,
    [data-testid="stAlertContentSuccess"] p {
        color: #efefef !important;
    }
    [data-testid="stSidebar"] h3 {
        color: #ffb300 !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        background-color: #1a1a1a !important;
        border: 1px solid #2e2e2e !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        background-color: #262626 !important;
        color: #ececec !important;
        border: 1px solid #3d3d3d !important;
        height: 2.1rem !important;
        min-height: unset !important;
        padding: 0 14px !important;
        font-size: 0.8rem !important;
        width: auto !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] p,
    [data-testid="stFileUploaderDropzoneInstructions"] svg {
        color: #aaaaaa !important;
        fill: #aaaaaa !important;
    }
    [data-testid="stFileUploaderFile"] {
        background-color: #1a1a1a !important;
        border: 1px solid #2e2e2e !important;
        border-radius: 6px !important;
    }

    [data-testid="stFileUploaderFileName"],
    [data-testid="stFileUploaderFile"] span,
    [data-testid="stFileUploaderFile"] small,
    [data-testid="stFileUploaderFileErrorMessage"] {
        color: #d6d6d6 !important;
    }
    [data-testid="stExpander"] summary p,
    [data-testid="stExpander"] summary span,
    [data-testid="stExpander"] svg {
        color: #dcdcdc !important;
    }
    .stTabs [data-baseweb="tab"] p,
    .stTabs [data-baseweb="tab"] {
        color: #b8b8b8 !important;
    }
    .stTabs [aria-selected="true"] p {
        color: #ffffff !important;
    }
    [data-testid="stMarkdownContainer"] small,
    .stMarkdown small {
        color: #b8b8b8 !important;
    }

    .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #3a3a3a; gap: 8px; }
    .stTabs [data-baseweb="tab"], .stTabs button[data-baseweb="tab"], .stTabs [role="tab"] {
        background-color: #171717 !important; border: 1px solid #3a3a3a !important;
        border-bottom: none !important; border-radius: 8px 8px 0 0 !important;
        color: #ececec !important; font-weight: 700 !important; opacity: 1 !important;
        padding: 10px 22px !important;
    }
    .stTabs [data-baseweb="tab"] * { color: #ececec !important; opacity: 1 !important; }
    .stTabs [aria-selected="true"], .stTabs button[aria-selected="true"] {
        background-color: #3d2a0f44 !important; border-color: #ff9f1c !important; color: #ff9f1c !important;
    }
    .stTabs [aria-selected="true"] * { color: #ff9f1c !important; }
    [data-baseweb="tab"] { opacity: 1 !important; }
    [data-testid="stExpander"] { border: 1px solid #3a3a3a !important; border-radius: 8px !important; background: #1c1c1c !important; margin-bottom: 6px !important; overflow: hidden !important; }
    [data-testid="stExpander"] details { background: #1c1c1c !important; }
    [data-testid="stExpander"] details[open] { background: #1c1c1c !important; }
    .streamlit-expanderHeader, [data-testid="stExpander"] details summary { background: #1c1c1c !important; color: #d4d4d4 !important; font-weight: 600 !important; border: none !important; border-radius: 8px !important; padding: 12px 16px !important; }
    [data-testid="stExpander"] details[open] summary { background: #2a1f0f !important; color: #ff9f1c !important; border-bottom: 1px solid #3a3a3a !important; border-radius: 8px 8px 0 0 !important; }
    .streamlit-expanderHeader:focus, [data-testid="stExpander"] *:focus, [data-testid="stExpander"] summary:focus-visible { outline: none !important; box-shadow: none !important; }
    .streamlit-expanderContent, [data-testid="stExpander"] details > div { background: #131313 !important; border-top: 1px solid #3a3a3a !important; border-radius: 0 0 8px 8px !important; }
    </style>
""", unsafe_allow_html=True)


# =========================================================
# 3. 임시 비밀번호 Google Sheets 기반 영구 저장
#    (Streamlit Cloud는 Reboot 시 로컬 디스크가 초기화되므로 Sheets에 저장)
#    필요한 st.secrets: temp_pwd_sheet_id, [gcp_service_account]
# =========================================================
def _sanitize_secret_text(msg):
    """API 키 등 민감정보가 에러 메시지에 그대로 노출되는 것을 방지."""
    if not msg:
        return msg
    s = str(msg)
    s = re.sub(r'gsk_[A-Za-z0-9]{10,}', 'gsk_****REDACTED****', s)
    s = re.sub(r'\b(sk|key|token)[-_][A-Za-z0-9]{16,}\b', r'\1-****REDACTED****', s, flags=re.IGNORECASE)
    s = re.sub(r'Bearer\s+[A-Za-z0-9\._\-]{16,}', 'Bearer ****REDACTED****', s)
    s = re.sub(r'-----BEGIN PRIVATE KEY-----.*?-----END PRIVATE KEY-----', '****REDACTED PRIVATE KEY****', s, flags=re.DOTALL)
    return s


@st.cache_resource(show_spinner=False)
def _get_temp_pwd_worksheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["temp_pwd_sheet_id"])
    try:
        ws = sh.worksheet(_TEMP_PWD_WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=_TEMP_PWD_WORKSHEET, rows=200, cols=3)
        ws.update([["password", "expires", "created"]])
    return ws


def _load_temp_pwds():
    if not GSPREAD_OK:
        return {}
    try:
        ws = _get_temp_pwd_worksheet()
        records = ws.get_all_records()
        st.session_state['_sheets_last_error'] = None
        if not records:
            default = {
                _DEFAULT_TEMP_PWD: {
                    "expires": (datetime.now() + timedelta(days=7)).isoformat(),
                    "created": datetime.now().isoformat(),
                }
            }
            _save_temp_pwds(default)
            return {
                _DEFAULT_TEMP_PWD: {
                    "expires": datetime.now() + timedelta(days=7),
                    "created": datetime.now(),
                }
            }
        result = {}
        for row in records:
            pwd = str(row.get("password", "")).strip()
            if not pwd:
                continue
            exp = row.get("expires")
            cre = row.get("created")
            result[pwd] = {
                "expires": datetime.fromisoformat(exp) if exp else None,
                "created": datetime.fromisoformat(cre) if cre else datetime.now(),
            }
        return result
    except Exception as e:
        st.session_state['_sheets_last_error'] = _sanitize_secret_text(f"[로드 실패] {type(e).__name__}: {e}")
        return {}


def _save_temp_pwds(pwd_dict):
    if not GSPREAD_OK:
        return False
    try:
        ws = _get_temp_pwd_worksheet()
        rows = [["password", "expires", "created"]]
        for pwd, info in pwd_dict.items():
            exp = info.get("expires")
            cre = info.get("created")
            rows.append([
                pwd,
                exp.isoformat() if isinstance(exp, datetime) else (exp if isinstance(exp, str) else ""),
                cre.isoformat() if isinstance(cre, datetime) else (cre if isinstance(cre, str) else str(datetime.now())),
            ])
        ws.clear()
        ws.update(rows)
        st.session_state['_sheets_last_error'] = None
        return True
    except Exception as e:
        st.session_state['_sheets_last_error'] = _sanitize_secret_text(f"[저장 실패] {type(e).__name__}: {e}")
        return False


# =========================================================
# 3-1. 검토 이력 Google Sheets 영구 저장 (품번별 상세 결과, append-only 로그)
#    임시비번용과 별도 시트 사용. 필요한 st.secrets: review_history_sheet_id
#    (같은 gcp_service_account 재사용)
# =========================================================
_REVIEW_HISTORY_WORKSHEET = "review_history"
_REVIEW_HISTORY_HEADER = [
    "timestamp", "bom_file", "dwg_count", "part_no", "bom_material",
    "dwg_material", "match", "score", "grade", "process", "comment",
]


@st.cache_resource(show_spinner=False)
def _get_review_history_worksheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["review_history_sheet_id"])
    try:
        ws = sh.worksheet(_REVIEW_HISTORY_WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=_REVIEW_HISTORY_WORKSHEET, rows=1000, cols=len(_REVIEW_HISTORY_HEADER))
        ws.update([_REVIEW_HISTORY_HEADER])
    return ws


def _append_review_history(bom_filename, dwg_count, results):
    """비교 실행 1회 = 품번 개수만큼의 행을 append. clear() 없이 추가만 하므로 데이터 유실 위험 없음."""
    if not GSPREAD_OK:
        return False
    try:
        ws = _get_review_history_worksheet()
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = [
            [
                ts, bom_filename, dwg_count, r["part_no"], r["bom_material"],
                r["dwg_material"], "O" if r["match"] else "X", f"{r['score']:.1f}",
                r["grade"], r["process"], r["comment"],
            ]
            for r in results
        ]
        ws.append_rows(rows, value_input_option="RAW")
        st.session_state['_history_sheets_error'] = None
        return True
    except Exception as e:
        st.session_state['_history_sheets_error'] = _sanitize_secret_text(f"[이력저장 실패] {type(e).__name__}: {e}")
        return False


def _load_review_history():
    """저장된 검토 이력 전체를 최신순으로 반환. 실패 시 None (호출부에서 세션 내 요약으로 폴백)."""
    if not GSPREAD_OK:
        return None
    try:
        ws = _get_review_history_worksheet()
        records = ws.get_all_records()
        st.session_state['_history_sheets_error'] = None
        records.reverse()
        return records
    except Exception as e:
        st.session_state['_history_sheets_error'] = _sanitize_secret_text(f"[이력로드 실패] {type(e).__name__}: {e}")
        return None


# =========================================================
# 3-2. 3D 물성치(CAD export) 이력 저장 — 검토 이력과 같은 스프레드시트, 별도 워크시트 사용
#      "이전 리비전과 비교"를 위해 매 업로드 값을 append-only 로 쌓아둔다.
# =========================================================
_CAD_LOG_WORKSHEET = "cad_property_log"
_CAD_LOG_HEADER = ["timestamp", "part_no", "weight_kg", "dim_x_mm", "dim_y_mm", "dim_z_mm", "cad_file"]


@st.cache_resource(show_spinner=False)
def _get_cad_log_worksheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["review_history_sheet_id"])
    try:
        ws = sh.worksheet(_CAD_LOG_WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=_CAD_LOG_WORKSHEET, rows=2000, cols=len(_CAD_LOG_HEADER))
        ws.update([_CAD_LOG_HEADER])
    return ws


def _load_cad_log():
    if not GSPREAD_OK:
        return None
    try:
        ws = _get_cad_log_worksheet()
        records = ws.get_all_records()
        st.session_state['_cad_sheets_error'] = None
        return records
    except Exception as e:
        st.session_state['_cad_sheets_error'] = _sanitize_secret_text(f"[CAD 이력로드 실패] {type(e).__name__}: {e}")
        return None


def _append_cad_log(cad_filename, cad_rows):
    """cad_rows: [{part_no, weight_kg, dim_x_mm, dim_y_mm, dim_z_mm}, ...]"""
    if not GSPREAD_OK:
        return False
    try:
        ws = _get_cad_log_worksheet()
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = [
            [ts, r["part_no"], r.get("weight_kg", ""), r.get("dim_x_mm", ""),
             r.get("dim_y_mm", ""), r.get("dim_z_mm", ""), cad_filename]
            for r in cad_rows
        ]
        ws.append_rows(rows, value_input_option="RAW")
        st.session_state['_cad_sheets_error'] = None
        return True
    except Exception as e:
        st.session_state['_cad_sheets_error'] = _sanitize_secret_text(f"[CAD 이력저장 실패] {type(e).__name__}: {e}")
        return False


def find_previous_cad_entry(part_no, log_records, before_timestamp=None):
    """같은 품번의 로그 중 가장 최근 것 1건 반환 (before_timestamp 이전 것만, 없으면 전체 중 최신)"""
    if not log_records:
        return None
    candidates = [r for r in log_records if str(r.get("part_no", "")).strip() == part_no]
    if before_timestamp:
        candidates = [r for r in candidates if str(r.get("timestamp", "")) < before_timestamp]
    if not candidates:
        return None
    candidates.sort(key=lambda r: str(r.get("timestamp", "")))
    return candidates[-1]


def _to_float(v):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def compare_cad_property(current_val, ref_val, tolerance_pct):
    """
    수치 비교 -> (일치여부, 신뢰도점수 0~100, 차이%, 코멘트)
    허용오차 이내면 100점, 벗어난 정도에 비례해 감점 (JOINT 배지 4단계와 동일한 점수 스케일 사용)
    """
    if current_val is None or ref_val is None:
        if current_val is None and ref_val is None:
            reason = "CAD 실측값·기준값 모두 없음"
        elif current_val is None:
            reason = "CAD 파일에 실측값 없음 (컬럼 매핑 확인)"
        else:
            reason = "비교 대상 기준값 없음"
        return None, None, None, reason
    if ref_val == 0:
        diff_pct = 0.0 if current_val == 0 else 100.0
    else:
        diff_pct = abs(current_val - ref_val) / abs(ref_val) * 100
    if diff_pct <= tolerance_pct:
        score = 100.0 - (diff_pct / max(tolerance_pct, 0.01)) * 20.0  # 허용범위 내에서는 80~100점
        score = max(80.0, min(100.0, score))
        match = True
        comment = f"허용오차({tolerance_pct}%) 이내, 차이 {diff_pct:.1f}%"
    else:
        # 허용오차 초과분에 비례해 감점 (초과분 5%p당 -20점)
        over = diff_pct - tolerance_pct
        score = max(0.0, 80.0 - over * 4.0)
        match = False
        comment = f"허용오차({tolerance_pct}%) 초과, 차이 {diff_pct:.1f}%"
    return match, score, diff_pct, comment


def run_cad_comparison(cad_df, col_map, bom_df, bom_col_map, tolerance_pct, compare_bom, compare_prev, compare_dwg=False, dwg_dim_index=None):
    """
    col_map: {'part_no':col, 'weight':col_or_None, 'x':col_or_None, 'y':col_or_None, 'z':col_or_None}
    bom_col_map: {'part_no':col, 'weight':col_or_None, 'x':col_or_None, 'y':col_or_None, 'z':col_or_None}
    dwg_dim_index: {normalized_part_no: {"dims": {"x","y","z","weight"}, ...}} — run_comparison()에서 채워짐 (1번 탭 실행 결과)
    """
    prev_log = _load_cad_log() if compare_prev else None
    now_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    bom_lookup = {}
    if compare_bom and bom_df is not None and bom_col_map.get("part_no"):
        for _, brow in bom_df.iterrows():
            bpn = normalize(str(brow.get(bom_col_map["part_no"], "")))
            if not bpn:
                continue
            bom_lookup[bpn] = {
                "weight": _to_float(brow.get(bom_col_map.get("weight"))) if bom_col_map.get("weight") else None,
                "x": _to_float(brow.get(bom_col_map.get("x"))) if bom_col_map.get("x") else None,
                "y": _to_float(brow.get(bom_col_map.get("y"))) if bom_col_map.get("y") else None,
                "z": _to_float(brow.get(bom_col_map.get("z"))) if bom_col_map.get("z") else None,
            }

    results = []
    log_rows = []
    for _, row in cad_df.iterrows():
        part_no = str(row.get(col_map["part_no"], "")).strip()
        if not part_no:
            continue
        cur = {
            "weight": _to_float(row.get(col_map.get("weight"))) if col_map.get("weight") else None,
            "x": _to_float(row.get(col_map.get("x"))) if col_map.get("x") else None,
            "y": _to_float(row.get(col_map.get("y"))) if col_map.get("y") else None,
            "z": _to_float(row.get(col_map.get("z"))) if col_map.get("z") else None,
        }
        log_rows.append({
            "part_no": part_no, "weight_kg": cur["weight"] if cur["weight"] is not None else "",
            "dim_x_mm": cur["x"] if cur["x"] is not None else "",
            "dim_y_mm": cur["y"] if cur["y"] is not None else "",
            "dim_z_mm": cur["z"] if cur["z"] is not None else "",
        })

        checks = []
        # vs BOM 기준값
        if compare_bom:
            bref = bom_lookup.get(normalize(part_no))
            for dim_key, dim_label in [("weight", "무게"), ("x", "가로"), ("y", "세로"), ("z", "높이")]:
                ref_val = bref.get(dim_key) if bref else None
                match, score, diff_pct, comment = compare_cad_property(cur[dim_key], ref_val, tolerance_pct)
                checks.append({"source": "BOM", "dim": dim_label, "match": match, "score": score,
                                "diff_pct": diff_pct, "cur": cur[dim_key], "ref": ref_val, "comment": comment})
        # vs 이전 리비전
        if compare_prev:
            prev_entry = find_previous_cad_entry(part_no, prev_log, before_timestamp=now_ts) if prev_log is not None else None
            for dim_key, dim_label, log_key in [
                ("weight", "무게", "weight_kg"), ("x", "가로", "dim_x_mm"),
                ("y", "세로", "dim_y_mm"), ("z", "높이", "dim_z_mm"),
            ]:
                ref_val = _to_float(prev_entry.get(log_key)) if prev_entry else None
                match, score, diff_pct, comment = compare_cad_property(cur[dim_key], ref_val, tolerance_pct)
                checks.append({"source": "PREV", "dim": dim_label, "match": match, "score": score,
                                "diff_pct": diff_pct, "cur": cur[dim_key], "ref": ref_val, "comment": comment})
        # vs 도면 표기 치수 (타이틀블록/노트란에 명시적으로 라벨링된 요약값만 인식됨 — 1번 탭 실행 필요)
        if compare_dwg:
            dwg_entry = (dwg_dim_index or {}).get(normalize(part_no))
            dwg_dims = dwg_entry.get("dims") if dwg_entry else None
            for dim_key, dim_label in [("weight", "무게"), ("x", "가로"), ("y", "세로"), ("z", "높이")]:
                ref_val = dwg_dims.get(dim_key) if dwg_dims else None
                match, score, diff_pct, comment = compare_cad_property(cur[dim_key], ref_val, tolerance_pct)
                checks.append({"source": "DWG", "dim": dim_label, "match": match, "score": score,
                                "diff_pct": diff_pct, "cur": cur[dim_key], "ref": ref_val, "comment": comment})

        _scored = [c["score"] for c in checks if c["score"] is not None]
        overall_score = (sum(_scored) / len(_scored)) if _scored else None
        results.append({
            "part_no": part_no, "current": cur, "checks": checks, "overall_score": overall_score,
        })

    return results, log_rows


# =========================================================
# 4. 신뢰도 배지 (JOINT 원본과 동일한 4단계 색상 체계)
# =========================================================
def render_confidence_badge(score, label):
    _is_en = st.session_state.get('lang', 'KO') == 'EN'
    try:
        _s = float(score)
    except (TypeError, ValueError):
        _s = 0.0
    if _s >= 80:
        _color, _bg, _icon = "#10b981", "#0f2410", "🟢"
        _status = "High" if _is_en else "높음"
    elif _s >= 50:
        _color, _bg, _icon = "#facc15", "#2a2408", "🟡"
        _status = "Moderate" if _is_en else "보통"
    elif _s >= 20:
        _color, _bg, _icon = "#fb923c", "#2d1c08", "🟠"
        _status = "Low" if _is_en else "낮음"
    else:
        _color, _bg, _icon = "#f87171", "#2d0f0f", "🔴"
        _status = "Very Low" if _is_en else "매우 낮음"

    st.markdown(
        f"<div style='margin-top:6px;'>"
        f"<span style='color:#9c9c9c;font-size:0.85rem;'>{label}</span>"
        f"<div style='display:flex;align-items:baseline;gap:8px;margin-top:2px;'>"
        f"<span style='font-size:1.9rem;font-weight:800;color:{_color};'>{_s:.1f}%</span>"
        f"<span style='background:{_bg};color:{_color};font-size:0.72rem;font-weight:700;"
        f"padding:2px 8px;border-radius:10px;white-space:nowrap;'>{_icon} {_status}</span>"
        f"</div></div>",
        unsafe_allow_html=True,
    )


# =========================================================
# 5. 인증 시스템
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "temp_pwd_list" not in st.session_state:
    st.session_state.temp_pwd_list = _load_temp_pwds()
if "is_owner" not in st.session_state:
    st.session_state.is_owner = False
if "material_map" not in st.session_state:
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


def _check_temp_pwd(p):
    """임시 비번 유효성 검사 — Sheets에서 항상 최신 목록 확인"""
    _fresh = _load_temp_pwds()
    st.session_state.temp_pwd_list = _fresh
    info = _fresh.get(p)
    if info is None:
        return False
    if info['expires'] is None:
        return True
    return datetime.now() < info['expires']


if not st.session_state.authenticated:
    _, center, _ = st.columns([1, 1.8, 1])
    with center:
        st.markdown("<div style='height:40px;'></div>", unsafe_allow_html=True)

        st.markdown("""<style>
            div[data-testid="stSelectbox"] > div > div {
                background-color: #ececec !important;
                color: #262626 !important;
                border-color: #d4d4d4 !important;
            }
            div[data-testid="stSelectbox"] > div > div > div {
                color: #262626 !important;
            }
            div[data-baseweb="select"] > div {
                background-color: #ececec !important;
                border-color: #d4d4d4 !important;
            }
            div[data-baseweb="select"] span {
                color: #262626 !important;
            }
        </style>""", unsafe_allow_html=True)

        _, lang_select_col = st.columns([5, 1.1])
        with lang_select_col:
            lang_display_options = ["KO", "EN"]
            current_display = st.session_state["lang"]
            lang_choice_login = st.selectbox(
                "Language", lang_display_options,
                index=lang_display_options.index(current_display),
                label_visibility="collapsed",
                key="login_lang_select",
            )
            if lang_choice_login != st.session_state["lang"]:
                st.session_state["lang"] = lang_choice_login
                st.rerun()

        LT = LOGIN_TEXT[st.session_state["lang"]]

        st.markdown(
            f"""<div class='glass-card' style='text-align:center; padding:22px 36px; margin-top:12px;'>
                <div style='color:#ff9f1c; font-size:0.78rem; font-weight:700; letter-spacing:0.12em; text-transform:uppercase; margin-bottom:8px;'>{LT['badge']}</div>
                <h2 style='color:#f2f2f2; font-size:1.35rem; font-weight:600; line-height:1.4; margin:0 0 4px 0;'>{LT['title']}</h2>
                <div style='width:56px; height:2px; background:#10b981; margin:12px auto 0 auto;'></div>
            </div>""",
            unsafe_allow_html=True,
        )

        _pw_col, _btn_col = st.columns([4, 1])
        with _pw_col:
            pwd = st.text_input(LT['pwd_label'], type="password",
                                 label_visibility="collapsed",
                                 placeholder=LT['pwd_label'])
        with _btn_col:
            _login_btn = st.button(LT['auth_btn'], type="primary", use_container_width=True)
        if _login_btn:
            if pwd == OWNER_PWD:
                st.session_state.authenticated = True
                st.session_state.is_owner = True
                st.rerun()
            elif _check_temp_pwd(pwd):
                st.session_state.authenticated = True
                st.session_state.is_owner = False
                st.session_state.logged_temp_pwd = pwd
                st.rerun()
            else:
                st.error(LT['invalid'])
    st.stop()


# =========================================================
# 6. 재질 매칭 엔진 (DWG-BOM 고유 로직 — 변경 없음)
# =========================================================
def guess_col_index(columns, keywords, offset=0, default=0):
    """
    컬럼 매핑 셀렉트박스의 기본 선택값을 자동으로 맞춰주는 헬퍼.
    컬럼명에 keywords 중 하나가 포함되어 있으면 그 위치(+offset)를 반환, 없으면 default.
    offset은 "사용 안 함" 같은 항목이 리스트 맨 앞에 추가된 경우(+1) 보정용.
    "x"/"y"/"z" 같은 한 글자 영문 키워드는 오검출(예: "Taxonomy")을 막기 위해 완전 일치만 인정.
    """
    for i, c in enumerate(columns):
        cs = str(c).lower().strip()
        for kw in keywords:
            kw = kw.lower()
            if len(kw) == 1 and kw.isalpha():
                if cs == kw or cs.startswith(kw + "(") or cs.startswith(kw + " "):
                    return i + offset
            elif kw in cs:
                return i + offset
    return default


def normalize(s: str) -> str:
    if not isinstance(s, str):
        return ""
    return re.sub(r"[\s\-_/]", "", s).upper()


def build_alias_lookup(map_df: pd.DataFrame) -> dict:
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
    반환: (일치여부, 신뢰도점수 0~100, 등급라벨, 코멘트)
    점수 구간(JOINT 배지 기준과 동일): >=80 초록 / 50~79 노랑 / 20~49 주황 / <20 빨강
    """
    if not dwg_mat:
        return False, 0.0, "none", "도면에서 재질을 인식하지 못했습니다."

    n_bom, n_dwg = normalize(bom_mat), normalize(dwg_mat)

    if n_bom == n_dwg and n_bom != "":
        return True, 100.0, "exact", "표기 완전 일치"

    std_bom = alias_lookup.get(n_bom)
    std_dwg = alias_lookup.get(n_dwg)
    if std_bom and std_dwg and std_bom == std_dwg:
        return True, 95.0, "std", f"표준코드 매핑 일치 ({std_bom})"

    ratio = difflib.SequenceMatcher(None, n_bom, n_dwg).ratio()
    score = ratio * 100
    if ratio >= 0.5:
        return False, score, "low", f"유사도 {score:.0f}% - 수동 확인 필요"

    return False, score, "none", "재질 표기 불일치"


def llm_recheck(bom_mat: str, dwg_mat: str) -> str:
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
            model="openai/gpt-oss-20b",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=50,
            temperature=0,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        st.session_state['_llm_last_error'] = _sanitize_secret_text(f"[LLM 재확인 실패] {type(e).__name__}: {e}")
        return ""


# =========================================================
# 7. 도면 PDF 텍스트 추출 + 재질 인식
# =========================================================
MATERIAL_KEYWORDS = [
    r"재\s*질\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-\.]*)",
    r"MAT'?L\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-\.]*)",
    r"MATERIAL\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-\.]*)",
]
PARTNO_KEYWORDS = [
    r"품\s*번\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-]*)",
    r"도면\s*번호\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-]*)",
    r"PART\s*NO\.?\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-]*)",
    r"DWG\s*NO\.?\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-]*)",
]
# 도면 표기 치수/중량 — 벡터 좌표 기반 자동 치수 인식은 범위 밖(스코프 아웃).
# "치수(WxDxH): 120 x 90 x 40" 처럼 타이틀블록/노트란에 명시적으로 라벨링된 요약값만 인식.
DIM_SUMMARY_KEYWORDS = [
    r"치\s*수\s*\(?\s*W\s*x\s*D\s*x\s*H\s*\)?\s*[:：]?\s*([0-9.]+)\s*[xX×]\s*([0-9.]+)\s*[xX×]\s*([0-9.]+)",
    r"DIM(?:ENSION)?S?\s*\(?\s*W\s*x\s*D\s*x\s*H\s*\)?\s*[:：]?\s*([0-9.]+)\s*[xX×]\s*([0-9.]+)\s*[xX×]\s*([0-9.]+)",
    r"치\s*수\s*[:：]?\s*([0-9.]+)\s*[xX×]\s*([0-9.]+)\s*[xX×]\s*([0-9.]+)",
]
WEIGHT_KEYWORDS = [
    r"중\s*량\s*[:：]?\s*([0-9]+\.?[0-9]*)\s*(?:kg|KG)?",
    r"무\s*게\s*[:：]?\s*([0-9]+\.?[0-9]*)\s*(?:kg|KG)?",
    r"WEIGHT\s*[:：]?\s*([0-9]+\.?[0-9]*)\s*(?:kg|KG)?",
]


def extract_pdf_text(file) -> str:
    """페이지 전체 텍스트 (검증용 원문 표시에 사용)"""
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


def extract_titleblock_text(file) -> str:
    """
    타이틀블록(도면 하단-우측 영역)만 잘라서 텍스트 추출.
    페이지 전체 텍스트에서 품번/재질을 찾으면, 노트란 등 다른 영역의 텍스트가
    타이틀블록과 같은 높이(y좌표)에 있을 때 라벨-값이 뒤섞여 오검출될 수 있음
    (예: "재질:" 바로 다음에 엉뚱한 노트 문구가 값으로 잡히는 경우).
    KS/ISO 도면 표준상 타이틀블록은 대부분 우하단에 있으므로, 그 영역만 먼저 검색해서
    이 문제를 줄이고, 못 찾으면 전체 텍스트로 폴백한다.
    """
    if not PDF_OK:
        return ""
    try:
        with pdfplumber.open(file) as pdf:
            page = pdf.pages[-1]
            w, h = page.width, page.height
            # 우하단 영역 (타이틀블록이 보통 위치하는 곳). x축 시작점을 60%로 잡아
            # 좌측 노트란 문장이 경계선에 걸려 한두 글자만 잘려 들어오는 것을 방지.
            crop_box = (w * 0.6, h * 0.55, w, h)
            cropped = page.within_bbox(crop_box)
            return cropped.extract_text() or ""
    except Exception:
        return ""


def guess_field(text: str, patterns) -> str:
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def guess_field_prefer_titleblock(titleblock_text: str, full_text: str, patterns) -> str:
    """타이틀블록 영역에서 먼저 찾고, 없으면 전체 텍스트로 폴백."""
    val = guess_field(titleblock_text, patterns)
    if val:
        return val
    return guess_field(full_text, patterns)


def extract_drawing_dims(titleblock_text: str, full_text: str):
    """도면 텍스트에서 '치수(WxDxH)' / '중량' 요약 표기를 찾아 dict로 반환.
    못 찾으면 해당 키는 None (자동 벡터 치수 인식은 범위 밖)."""
    result = {"x": None, "y": None, "z": None, "weight": None}
    for source in (titleblock_text, full_text):
        if result["x"] is not None:
            break
        for p in DIM_SUMMARY_KEYWORDS:
            m = re.search(p, source, re.IGNORECASE)
            if m:
                try:
                    result["x"] = float(m.group(1))
                    result["y"] = float(m.group(2))
                    result["z"] = float(m.group(3))
                except ValueError:
                    pass
                break
    for source in (titleblock_text, full_text):
        if result["weight"] is not None:
            break
        for p in WEIGHT_KEYWORDS:
            m = re.search(p, source, re.IGNORECASE)
            if m:
                try:
                    result["weight"] = float(m.group(1))
                except ValueError:
                    pass
                break
    return result


def is_probably_scanned(text: str) -> bool:
    """텍스트 레이어가 거의/전혀 없으면 스캔본(이미지 PDF)일 가능성이 높다고 판단."""
    return len((text or "").strip()) < 20


def rasterize_pdf_first_page(file, resolution=150):
    """스캔 도면을 Vision 모델에 보내기 위해 첫 페이지를 이미지(PNG bytes)로 렌더링."""
    if not PDF_OK:
        return None
    try:
        with pdfplumber.open(file) as pdf:
            page = pdf.pages[0]
            img = page.to_image(resolution=resolution)
            buf = io.BytesIO()
            img.original.save(buf, format="PNG")
            return buf.getvalue()
    except Exception as e:
        st.session_state['_vision_last_error'] = _sanitize_secret_text(f"[이미지 렌더링 실패] {type(e).__name__}: {e}")
        return None


_VISION_PROMPT = (
    "You are reading a mechanical engineering drawing (도면) title block. "
    "Extract these fields if visible: part number (품번/도면번호/PART NO/DWG NO), "
    "material (재질/MATERIAL), overall dimensions width x depth x height in mm if shown as a single "
    "summary value (not individual dimension lines), and weight in kg if shown. "
    "Respond ONLY with compact JSON, no markdown fences, no explanation, in this exact shape: "
    '{"part_no": "<string or null>", "material": "<string or null>", '
    '"dim_x_mm": <number or null>, "dim_y_mm": <number or null>, "dim_z_mm": <number or null>, '
    '"weight_kg": <number or null>}'
)


def vision_extract_fields(image_bytes):
    """
    스캔(이미지) 도면에서 Gemini Vision(google-genai SDK)으로 품번/재질/치수/중량을 추출.
    실패 시 빈 dict 반환 (앱은 계속 동작, 해당 파일은 미검출로 처리됨).
    """
    if not image_bytes:
        st.session_state['_vision_last_error'] = "[Vision 인식 실패] 렌더링된 이미지가 없습니다 (위 렌더링 실패 메시지 확인)."
        return {}
    if not (GEMINI_OK and GEMINI_API_KEY):
        st.session_state['_vision_last_error'] = "[Vision 인식 실패] GEMINI_OK/GEMINI_API_KEY 조건 불충족 (라이브러리 또는 키 확인 필요)."
        return {}
    try:
        client = gemini_genai.Client(api_key=GEMINI_API_KEY)
        resp = client.models.generate_content(
            model="gemini-flash-latest",
            contents=[
                _VISION_PROMPT,
                gemini_types.Part.from_bytes(data=image_bytes, mime_type="image/png"),
            ],
            config=gemini_types.GenerateContentConfig(
                response_mime_type="application/json",
            ),
        )
        raw = (resp.text or "").strip()
        raw = re.sub(r"^```(?:json)?|```$", "", raw, flags=re.MULTILINE).strip()
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        st.session_state['_vision_last_error'] = _sanitize_secret_text(f"[Vision 인식 실패] {type(e).__name__}: {e}")
        return {}


def run_comparison(bom_df, partno_col, material_col, process_col, dwg_files, use_llm, use_vision=False):
    alias_lookup = build_alias_lookup(st.session_state["material_map"])
    if use_vision:
        st.session_state['_vision_last_error'] = None  # 이전 실행의 잔여 에러 메시지 제거
    if use_llm:
        st.session_state['_llm_last_error'] = None

    dwg_index = {}
    for f in dwg_files:
        text = extract_pdf_text(f)
        tb_text = extract_titleblock_text(f)
        pn = guess_field_prefer_titleblock(tb_text, text, PARTNO_KEYWORDS)
        mat = guess_field_prefer_titleblock(tb_text, text, MATERIAL_KEYWORDS)
        dims = extract_drawing_dims(tb_text, text)
        via_vision = False

        # 텍스트 레이어가 없는(=스캔본으로 추정되는) 도면은 Vision으로 재시도
        if use_vision and not pn and not mat and is_probably_scanned(text):
            f.seek(0) if hasattr(f, "seek") else None
            img_bytes = rasterize_pdf_first_page(f)
            f.seek(0) if hasattr(f, "seek") else None
            v = vision_extract_fields(img_bytes)
            if v:
                pn = pn or (v.get("part_no") or "")
                mat = mat or (v.get("material") or "")
                if dims.get("x") is None and v.get("dim_x_mm") is not None:
                    dims = {
                        "x": v.get("dim_x_mm"), "y": v.get("dim_y_mm"),
                        "z": v.get("dim_z_mm"), "weight": v.get("weight_kg"),
                    }
                elif dims.get("weight") is None and v.get("weight_kg") is not None:
                    dims["weight"] = v.get("weight_kg")
                via_vision = True

        key = normalize(pn) if pn else normalize(f.name.rsplit(".", 1)[0])
        dwg_index[key] = {"material": mat, "raw_text": text, "filename": f.name, "dims": dims, "via_vision": via_vision}

    # 5단계(3D 물성치 비교) 탭에서 "도면 표기 치수와 비교"에 재사용할 수 있도록 세션에 저장
    st.session_state["dwg_dim_index"] = dwg_index

    results = []
    for _, row in bom_df.iterrows():
        part_no = str(row.get(partno_col, "")).strip()
        bom_mat = str(row.get(material_col, "")).strip()
        process = str(row.get(process_col, "")).strip()

        dwg_entry = dwg_index.get(normalize(part_no), {})
        dwg_mat = dwg_entry.get("material", "")

        match, score, grade, comment = match_material(bom_mat, dwg_mat, alias_lookup)

        if grade == "low" and use_llm:
            llm_out = llm_recheck(bom_mat, dwg_mat)
            if llm_out.startswith("동일"):
                match, score, grade = True, 92.0, "std"
                comment = "LLM 재확인: " + llm_out.split("|", 1)[-1].strip()
            elif llm_out.startswith("다름"):
                score, grade = min(score, 15.0), "none"
                comment = "LLM 재확인: " + llm_out.split("|", 1)[-1].strip()

        results.append({
            "part_no": part_no,
            "bom_material": bom_mat,
            "dwg_material": dwg_mat or "-",
            "match": match,
            "score": score,
            "grade": grade,
            "process": process,
            "comment": comment,
            "raw_text": dwg_entry.get("raw_text", ""),
            "dwg_filename": dwg_entry.get("filename", ""),
            "via_vision": dwg_entry.get("via_vision", False),
        })
    return results


# =========================================================
# 8. 사이드바
# =========================================================
with st.sidebar:
    st.markdown(f"<h2 style='color: #ffffff; font-size:1.15rem; margin-bottom: 20px;'>{t('console')}</h2>", unsafe_allow_html=True)

    if not st.session_state.get('is_owner', False):
        _LT_exp = LOGIN_TEXT[st.session_state["lang"]]
        _logged_pwd = st.session_state.get('logged_temp_pwd', '')
        _tp_info = st.session_state.temp_pwd_list.get(_logged_pwd, {})
        _exp_dt = _tp_info.get('expires')
        if _exp_dt is None:
            st.sidebar.markdown(f"🟢 {_LT_exp['no_expiry']}")
        else:
            _remain = _exp_dt - datetime.now()
            if _remain.total_seconds() > 0:
                _total_h = int(_remain.total_seconds() // 3600)
                _days_r, _hrs_r = _total_h // 24, _total_h % 24
                _time_str = f"{_days_r}일 {_hrs_r}시간" if st.session_state["lang"] == "KO" else f"{_days_r}d {_hrs_r}h"
                st.sidebar.markdown(f"🟡 {_time_str}")
            else:
                st.sidebar.markdown(f"🔴 {_LT_exp['access_expired'] if 'access_expired' in _LT_exp else _LT_exp['expired_label']}")

    st.sidebar.markdown("---")
    lang_choice_sb = st.sidebar.selectbox(
        "🌐 Language", ["KO", "EN"],
        index=0 if st.session_state["lang"] == "KO" else 1,
    )
    if lang_choice_sb != st.session_state["lang"]:
        st.session_state["lang"] = lang_choice_sb
        st.rerun()

    if st.sidebar.button("🚪 " + ("로그아웃" if st.session_state["lang"] == "KO" else "Log out")):
        st.session_state.authenticated = False
        st.rerun()

    # ── 소유자 전용: 임시 비번 관리 패널 ──────────────────────────
    if st.session_state.get('is_owner', False):
        st.sidebar.markdown("---")
        _LT_sb = LOGIN_TEXT[st.session_state["lang"]]
        st.sidebar.markdown(
            "<div style='color:#ffb300;font-weight:700;font-size:0.9rem;margin-bottom:8px;'>🔐 " +
            _LT_sb['temp_mgr_title'] + "</div>",
            unsafe_allow_html=True,
        )
        _sheets_err = st.session_state.get('_sheets_last_error')
        if _sheets_err:
            st.sidebar.error(f"⚠️ Google Sheets 오류\n\n{_sheets_err}")
        elif not GSPREAD_OK:
            st.sidebar.caption("⚪ Google Sheets 라이브러리 미설치 (세션 전용)")
        else:
            st.sidebar.caption("🟢 Google Sheets 연결 정상")

        if st.sidebar.button("🔄 Sheets 연결 테스트", key="sb_test_sheets"):
            try:
                _test_ws = _get_temp_pwd_worksheet()
                _test_ws.get_all_records()
                st.session_state['_sheets_last_error'] = None
                st.sidebar.success("✅ Sheets 연결 성공")
            except Exception as _e_test:
                st.session_state['_sheets_last_error'] = _sanitize_secret_text(f"[테스트 실패] {type(_e_test).__name__}: {_e_test}")
                st.sidebar.error(f"⚠️ {st.session_state['_sheets_last_error']}")

        _new_tp = st.sidebar.text_input(_LT_sb['new_temp_pwd'], key="sb_new_tp")
        _exp_opt = st.sidebar.selectbox(
            _LT_sb['expiry_period'],
            [_LT_sb['exp_1d'], _LT_sb['exp_3d'], _LT_sb['exp_7d'], _LT_sb['exp_30d'], _LT_sb['exp_none']],
            key="sb_exp_sel",
        )
        _day_map_sb = {
            _LT_sb['exp_1d']: 1, _LT_sb['exp_3d']: 3, _LT_sb['exp_7d']: 7,
            _LT_sb['exp_30d']: 30, _LT_sb['exp_none']: None,
        }
        if st.sidebar.button("➕ " + _LT_sb['add_btn'], key="sb_add_tp"):
            if _new_tp and _new_tp != OWNER_PWD:
                _days_sb = _day_map_sb.get(_exp_opt)
                _exp_dt_sb = (datetime.now() + timedelta(days=_days_sb)) if _days_sb else None
                st.session_state.temp_pwd_list[_new_tp] = {
                    'expires': _exp_dt_sb,
                    'created': datetime.now(),
                }
                _saved_ok = _save_temp_pwds(st.session_state.temp_pwd_list)
                if _saved_ok:
                    st.sidebar.success(_LT_sb['added_msg'] + _new_tp)
                else:
                    st.sidebar.error("⚠️ Sheets 저장 실패 — 아래 오류 메시지를 확인하세요 (재시작 시 사라질 수 있습니다)")
                st.rerun()
            elif _new_tp == OWNER_PWD:
                st.sidebar.error(_LT_sb['owner_only'])
            else:
                st.sidebar.warning(_LT_sb['enter_pwd_warn'])

        if st.session_state.temp_pwd_list:
            st.sidebar.markdown(
                "<div style='font-size:0.78rem;color:#9c9c9c;margin:8px 0 4px;'>" +
                _LT_sb['registered_pwds'] + "</div>",
                unsafe_allow_html=True,
            )
            for _tp_k, _tp_v in list(st.session_state.temp_pwd_list.items()):
                _exp_v = _tp_v['expires']
                if _exp_v is None:
                    _st_icon, _st_txt = "🟢", _LT_sb['unlimited_label']
                elif datetime.now() < _exp_v:
                    _hrs_v = int((_exp_v - datetime.now()).total_seconds() // 3600)
                    _st_icon, _st_txt = "🟡", f"{_LT_sb['left_label']}: {_hrs_v}h"
                else:
                    _st_icon, _st_txt = "🔴", _LT_sb['expired_label']
                _rc1, _rc2 = st.sidebar.columns([3, 1])
                _rc1.markdown(
                    f"<span style='font-size:0.8rem;'>{_st_icon} <code>{_tp_k}</code><br>"
                    f"<span style='color:#8a8a8a;font-size:0.72rem;'>{_st_txt}</span></span>",
                    unsafe_allow_html=True,
                )
                if _rc2.button("🗑️", key=f"sb_del_{_tp_k}"):
                    del st.session_state.temp_pwd_list[_tp_k]
                    _save_temp_pwds(st.session_state.temp_pwd_list)
                    st.rerun()
        else:
            st.sidebar.caption(_LT_sb['no_registered'])


# =========================================================
# 9. 메인 뷰포트
# =========================================================
st.markdown(f"<h1 style='margin-bottom:20px; font-size:1.8rem;'>{t('title')}</h1>", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [t("tab_upload"), t("tab_map"), t("tab_result"), t("tab_history"), t("tab_cad")]
)

# ---------------- Tab 1: 업로드 ----------------
with tab1:
    if not PDF_OK:
        st.warning(t("pdf_missing"))

    st.markdown(f"<div class='glass-card'><div class='glass-card-title'>{t('bom_upload')}</div>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        bom_file = st.file_uploader(t("bom_upload"), type=["xlsx", "csv"], label_visibility="collapsed")
    with col_b:
        dwg_files = st.file_uploader(t("dwg_upload"), type=["pdf"], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

    bom_df = None
    if bom_file is not None:
        try:
            if bom_file.name.lower().endswith(".csv"):
                bom_df = pd.read_csv(bom_file)
            else:
                bom_df = pd.read_excel(bom_file)
            st.dataframe(bom_df.head(20), use_container_width=True)
            st.session_state["bom_df"] = bom_df
            st.session_state["bom_file_name"] = bom_file.name
        except Exception as e:
            st.error(_sanitize_secret_text(f"BOM 파일을 읽는 중 오류: {e}"))

    if bom_df is not None:
        st.markdown(f"<div class='glass-card'><div class='glass-card-title'>{t('col_map_header')}</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        cols = list(bom_df.columns)
        with c1:
            partno_col = st.selectbox(t("col_partno"), cols, key="partno_col",
                                       index=guess_col_index(cols, ["품번", "part no", "partno", "도면번호"]))
        with c2:
            material_col = st.selectbox(t("col_material"), cols, key="material_col",
                                         index=guess_col_index(cols, ["재질", "material", "mat'l"]))
        with c3:
            process_col = st.selectbox(t("col_process"), cols, key="process_col",
                                        index=guess_col_index(cols, ["가공", "공법", "process"]))
        st.markdown("</div>", unsafe_allow_html=True)

        use_llm = st.checkbox(t("llm_check"), value=bool(GROQ_API_KEY))
        use_vision = st.checkbox(t("vision_check"), value=False)
        if use_vision and not (GEMINI_OK and GEMINI_API_KEY):
            st.caption(f"⚠️ {t('vision_not_configured')}")

        if st.button(t("run_btn"), type="primary"):
            if not dwg_files:
                st.warning(t("no_dwg"))
            else:
                with st.spinner("..."):
                    results = run_comparison(bom_df, partno_col, material_col, process_col, dwg_files, use_llm, use_vision=use_vision)
                st.session_state["last_result"] = results
                # 세션 내 요약 (Sheets 미설정 시 폴백용)
                st.session_state["review_history"].append({
                    "시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "BOM파일": bom_file.name,
                    "도면수": len(dwg_files),
                    "검토건수": len(results),
                })
                # Google Sheets에 품번별 상세 이력 영구 저장 (실패해도 앱은 계속 동작)
                _append_review_history(bom_file.name, len(dwg_files), results)
                st.success(t("run_done").format(n=len(results)))

        _vision_err = st.session_state.get('_vision_last_error')
        if _vision_err:
            st.warning(f"⚠️ {_vision_err}")
        _llm_err = st.session_state.get('_llm_last_error')
        if _llm_err:
            st.warning(f"⚠️ {_llm_err}")
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
        n_total = len(results)
        n_ok = sum(1 for r in results if r["match"])
        n_low = sum(1 for r in results if r["grade"] == "low")
        n_bad = sum(1 for r in results if r["grade"] == "none")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric(t("metric_total"), n_total)
        c2.metric(t("metric_match"), n_ok)
        c3.metric(t("metric_review"), n_low)
        c4.metric(t("metric_bad"), n_bad)

        st.markdown("---")

        for r in results:
            vision_tag = (
                f"&nbsp;<span style='background:#1a2e4a;color:#7db8ff;font-size:0.68rem;"
                f"font-weight:700;padding:2px 7px;border-radius:8px;'>👁 {t('vision_badge')}</span>"
                if r.get("via_vision") else ""
            )
            st.markdown(
                f"""<div class='glass-card'>
  <div class='glass-card-title'>{t('result_partno')}: {r['part_no']}{vision_tag}</div>
  {t('result_bom_mat')}: <code>{r['bom_material']}</code>
  &nbsp;&nbsp;→&nbsp;&nbsp;
  {t('result_dwg_mat')}: <code>{r['dwg_material']}</code>
  <br>
  {t('result_process')}: {r['process']}
  <br>
  {t('result_comment')}: {r['comment']}
</div>""",
                unsafe_allow_html=True,
            )
            render_confidence_badge(r["score"], t("result_conf"))
            if r["raw_text"]:
                with st.expander(f"{t('raw_text_expander')} — {r['dwg_filename']}"):
                    st.text(r["raw_text"][:2000])
            st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        st.markdown("---")
        export_df = pd.DataFrame([
            {
                t("result_partno"): r["part_no"],
                t("result_bom_mat"): r["bom_material"],
                t("result_dwg_mat"): r["dwg_material"],
                t("result_match"): "O" if r["match"] else "X",
                t("result_process"): r["process"],
                t("result_conf"): f"{r['score']:.1f}%",
                t("result_comment"): r["comment"],
            }
            for r in results
        ])
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="result")
        st.download_button(
            t("export_btn"),
            data=buf.getvalue(),
            file_name=f"dwg_bom_check_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------------- Tab 4: 검토 이력 ----------------
with tab4:
    sheet_hist = _load_review_history()
    _hist_err = st.session_state.get('_history_sheets_error')
    if _hist_err:
        st.warning(f"⚠️ {t('history_sheets_error')}\n\n{_hist_err}")
    if sheet_hist is not None:
        if sheet_hist:
            hist_df = pd.DataFrame(sheet_hist)
            rename_map = {
                "timestamp": t("hist_time"), "bom_file": t("hist_bom_file"),
                "dwg_count": t("hist_dwg_count"), "part_no": t("result_partno"),
                "bom_material": t("result_bom_mat"), "dwg_material": t("result_dwg_mat"),
                "match": t("result_match"), "score": t("hist_score"),
                "grade": t("hist_grade"), "process": t("result_process"),
                "comment": t("result_comment"),
            }
            hist_df = hist_df.rename(columns=rename_map)
            st.dataframe(hist_df, use_container_width=True)
        else:
            st.info(t("history_empty"))
    else:
        # Sheets 미설정/실패 시 세션 내 요약으로 폴백
        hist = st.session_state["review_history"]
        if not hist:
            st.info(t("history_empty"))
        else:
            st.caption(t("history_session_only"))
            st.dataframe(pd.DataFrame(hist), use_container_width=True)

# ---------------- Tab 5: 3D 물성치 비교 ----------------
with tab5:
    st.markdown(f"<div class='glass-card'><div class='glass-card-title'>{t('cad_upload')}</div>", unsafe_allow_html=True)
    cad_file = st.file_uploader(t("cad_upload"), type=["xlsx", "csv"], label_visibility="collapsed", key="cad_file_uploader")
    st.markdown("</div>", unsafe_allow_html=True)

    cad_df = None
    if cad_file is not None:
        try:
            if cad_file.name.lower().endswith(".csv"):
                cad_df = pd.read_csv(cad_file)
            else:
                cad_df = pd.read_excel(cad_file)
            st.dataframe(cad_df.head(20), use_container_width=True)
        except Exception as e:
            st.error(_sanitize_secret_text(f"CAD 파일을 읽는 중 오류: {e}"))

    if cad_df is None:
        st.info(t("cad_no_file"))
    else:
        cad_cols = list(cad_df.columns)
        cad_cols_opt = [t("cad_col_none")] + cad_cols

        st.markdown(f"<div class='glass-card'><div class='glass-card-title'>{t('col_map_header')}</div>", unsafe_allow_html=True)
        cc1, cc2, cc3, cc4, cc5 = st.columns(5)
        with cc1:
            cad_partno_col = st.selectbox(t("cad_col_partno"), cad_cols, key="cad_partno_col",
                                           index=guess_col_index(cad_cols, ["품번", "part no", "partno"]))
        with cc2:
            cad_weight_col = st.selectbox(t("cad_col_weight"), cad_cols_opt, key="cad_weight_col",
                                           index=guess_col_index(cad_cols, ["무게", "weight"], offset=1))
        with cc3:
            cad_x_col = st.selectbox(t("cad_col_x"), cad_cols_opt, key="cad_x_col",
                                      index=guess_col_index(cad_cols, ["가로", "width", "x"], offset=1))
        with cc4:
            cad_y_col = st.selectbox(t("cad_col_y"), cad_cols_opt, key="cad_y_col",
                                      index=guess_col_index(cad_cols, ["세로", "depth", "y"], offset=1))
        with cc5:
            cad_z_col = st.selectbox(t("cad_col_z"), cad_cols_opt, key="cad_z_col",
                                      index=guess_col_index(cad_cols, ["높이", "height", "z"], offset=1))
        st.markdown("</div>", unsafe_allow_html=True)

        col_map = {
            "part_no": cad_partno_col,
            "weight": None if cad_weight_col == t("cad_col_none") else cad_weight_col,
            "x": None if cad_x_col == t("cad_col_none") else cad_x_col,
            "y": None if cad_y_col == t("cad_col_none") else cad_y_col,
            "z": None if cad_z_col == t("cad_col_none") else cad_z_col,
        }

        compare_bom = st.checkbox(t("cad_compare_bom"), value=True)
        compare_prev = st.checkbox(t("cad_compare_prev"), value=True)
        compare_dwg = st.checkbox(t("cad_compare_dwg"), value=False)
        if compare_dwg and not st.session_state.get("dwg_dim_index"):
            st.warning(t("cad_no_dwg_run"))

        bom_col_map = {"part_no": None, "weight": None, "x": None, "y": None, "z": None}
        session_bom_df = st.session_state.get("bom_df")
        if compare_bom:
            if session_bom_df is None:
                st.warning(t("cad_no_bom"))
                compare_bom = False
            else:
                st.caption(t("cad_bom_need_cols"))
                bc1, bc2, bc3, bc4, bc5 = st.columns(5)
                bom_cols = list(session_bom_df.columns)
                bom_cols_opt = [t("cad_col_none")] + bom_cols
                with bc1:
                    b_pn = st.selectbox(t("cad_col_partno"), bom_cols, key="bom_ref_partno_col",
                                         index=guess_col_index(bom_cols, ["품번", "part no", "partno"]))
                with bc2:
                    b_w = st.selectbox(t("cad_col_weight"), bom_cols_opt, key="bom_ref_weight_col",
                                        index=guess_col_index(bom_cols, ["기준무게", "무게", "weight"], offset=1))
                with bc3:
                    b_x = st.selectbox(t("cad_col_x"), bom_cols_opt, key="bom_ref_x_col",
                                        index=guess_col_index(bom_cols, ["기준가로", "가로", "width", "x"], offset=1))
                with bc4:
                    b_y = st.selectbox(t("cad_col_y"), bom_cols_opt, key="bom_ref_y_col",
                                        index=guess_col_index(bom_cols, ["기준세로", "세로", "depth", "y"], offset=1))
                with bc5:
                    b_z = st.selectbox(t("cad_col_z"), bom_cols_opt, key="bom_ref_z_col",
                                        index=guess_col_index(bom_cols, ["기준높이", "높이", "height", "z"], offset=1))
                bom_col_map = {
                    "part_no": b_pn,
                    "weight": None if b_w == t("cad_col_none") else b_w,
                    "x": None if b_x == t("cad_col_none") else b_x,
                    "y": None if b_y == t("cad_col_none") else b_y,
                    "z": None if b_z == t("cad_col_none") else b_z,
                }

        tolerance_pct = st.number_input(t("cad_tolerance"), min_value=0.1, max_value=50.0, value=5.0, step=0.5)

        if st.button(t("cad_run_btn"), type="primary"):
            if not compare_bom and not compare_prev and not compare_dwg:
                st.warning(t("cad_no_checks"))
            else:
                with st.spinner("..."):
                    cad_results, log_rows = run_cad_comparison(
                        cad_df, col_map, session_bom_df, bom_col_map,
                        tolerance_pct, compare_bom, compare_prev,
                        compare_dwg=compare_dwg, dwg_dim_index=st.session_state.get("dwg_dim_index"),
                    )
                    _cad_saved = _append_cad_log(cad_file.name, log_rows)
                st.session_state["last_cad_result"] = cad_results
                if _cad_saved:
                    st.success(t("cad_saved_ok"))

        cad_sheets_err = st.session_state.get('_cad_sheets_error')
        if cad_sheets_err:
            st.caption(f"⚠️ {cad_sheets_err}")

        cad_results = st.session_state.get("last_cad_result")
        if cad_results:
            st.markdown("---")
            st.markdown(f"### {t('cad_result_title')}")
            for r in cad_results:
                st.markdown(
                    f"<div class='glass-card'><div class='glass-card-title'>{t('result_partno')}: {r['part_no']}</div>",
                    unsafe_allow_html=True,
                )
                if not r["checks"]:
                    st.caption(t("cad_no_checks"))
                else:
                    for c in r["checks"]:
                        _src_map_ko = {"BOM": "BOM", "PREV": "이전 리비전", "DWG": "도면 표기"}
                        _src_map_en = {"BOM": "BOM", "PREV": "Prev Rev", "DWG": "Drawing"}
                        src_label = (_src_map_ko if st.session_state["lang"] == "KO" else _src_map_en)[c["source"]]
                        cur_str = f"{c['cur']:.2f}" if c['cur'] is not None else "-"
                        ref_str = f"{c['ref']:.2f}" if c['ref'] is not None else "-"
                        st.markdown(
                            f"[{src_label}] {c['dim']}: <code>{cur_str}</code> ↔ <code>{ref_str}</code> — {c['comment']}",
                            unsafe_allow_html=True,
                        )
                st.markdown("</div>", unsafe_allow_html=True)
                if r["overall_score"] is not None:
                    render_confidence_badge(r["overall_score"], t("cad_overall"))
                st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
