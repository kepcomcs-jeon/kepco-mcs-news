"""한전MCS AI 위험성평가 및 작업승인(PTW) 폐루프 통합 앱.

고용노동부 고시 제2024-76호 법정 양식 + BM 특허 핵심 구조 통합:
1. [GPS 자동 수집 & 1차 점검] 스마트폰 GPS 주소 자동 수집 + Gemini 멀티모달 위험요인 탐지 + 번호 마커 렌더링
2. [피드백 보정] 안전관리자의 오탐·점수 수정 이력을 SQLite DB에 누적해 구역·공종별 위험도 가중치 자동 보정
3. [전/후 AI 검증 루프] 개선 전(마킹) 사진과 개선 후 사진을 AI가 대조해 위험 해소 여부 판정 및 미해소 시 승인 차단
4. [전자서명 연동 PTW 승인] AI 검증 통과 + 참여 근로자 전자서명 완료 시 SHA-256 해시 발급 및 법정 양식(.docx) 자동 생성
"""

from __future__ import annotations

import base64
import datetime
import hashlib
import io
import json
import os
import sqlite3
from typing import Any

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from geopy.geocoders import Nominatim
from google import genai
from google.genai import types
from PIL import Image, ImageDraw, ImageFont, ImageOps
import pytz
import streamlit as st
import streamlit.components.v1 as components
from streamlit_drawable_canvas import st_canvas
from streamlit_geolocation import streamlit_geolocation

# ---------------------------------------------------------------------------
# 1. 페이지 설정 및 모바일 아이콘 / 커스텀 CSS
# ---------------------------------------------------------------------------
try:
    icon_img = Image.open("safety_check_app/app_icon.png")
    st.set_page_config(page_title="한전MCS AI 위험성평가", page_icon=icon_img, layout="centered")
except Exception:
    st.set_page_config(page_title="한전MCS AI 위험성평가", page_icon="⚡", layout="centered")

app_icon_url = "https://raw.githubusercontent.com/kepcomcs-jeon/kepco-mcs-news/main/safety_check_app/app_icon.png"

components.html(f"""
<script>
    const parentDoc = window.parent.document;
    const oldApple = parentDoc.querySelector('link[rel="apple-touch-icon"]');
    if(oldApple) oldApple.remove();
    const oldAndroid = parentDoc.querySelector('link[rel="icon"][sizes="192x192"]');
    if(oldAndroid) oldAndroid.remove();

    const appleIcon = parentDoc.createElement('link');
    appleIcon.rel = 'apple-touch-icon';
    appleIcon.href = '{app_icon_url}';
    parentDoc.head.appendChild(appleIcon);

    const androidIcon = parentDoc.createElement('link');
    androidIcon.rel = 'icon';
    androidIcon.sizes = '192x192';
    androidIcon.href = '{app_icon_url}';
    parentDoc.head.appendChild(androidIcon);
</script>
""", height=0, width=0)

st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; }
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    div.stButton > button:first-child {
        border-radius: 8px; font-weight: bold;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px); box-shadow: 0 6px 10px rgba(0,0,0,0.1);
    }
    hr { margin-top: 1.5em; margin-bottom: 1.5em; border-color: #E0E6ED; }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# 2. SQLite 이력 DB 초기화 (구역·공종별 피드백 가중치 보정 & PTW 기록)
# ---------------------------------------------------------------------------
DB_PATH = "safety_history.db"
DEFAULT_MODEL = "gemini-3.5-flash"


def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS hazard_feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                zone_code TEXT NOT NULL,
                work_type TEXT NOT NULL,
                hazard_title TEXT NOT NULL,
                ai_score INTEGER NOT NULL,
                adjusted_score INTEGER NOT NULL,
                is_false_positive INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS ptw_records (
                ptw_id TEXT PRIMARY KEY,
                zone_code TEXT NOT NULL,
                work_type TEXT NOT NULL,
                gps_coords TEXT,
                address TEXT,
                verification_status TEXT NOT NULL,
                signer_count INTEGER NOT NULL,
                integrity_hash TEXT NOT NULL,
                approved_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def get_zone_weight_multiplier(zone_code: str, work_type: str) -> float:
    if not zone_code.strip():
        return 1.0
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT AVG(CAST(adjusted_score AS REAL) / NULLIF(ai_score, 0))
            FROM hazard_feedback
            WHERE zone_code = ? AND work_type = ? AND is_false_positive = 0
            """,
            (zone_code.strip(), work_type),
        ).fetchone()
    if row and row[0]:
        return round(max(0.7, min(1.3, float(row[0]))), 2)
    return 1.0


def save_feedback_records(
    zone_code: str,
    work_type: str,
    hazards: list[dict[str, Any]],
) -> None:
    now_iso = datetime.datetime.now(pytz.timezone("Asia/Seoul")).isoformat(timespec="seconds")
    with sqlite3.connect(DB_PATH) as conn:
        for item in hazards:
            conn.execute(
                """
                INSERT INTO hazard_feedback
                (zone_code, work_type, hazard_title, ai_score, adjusted_score, is_false_positive, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    zone_code.strip() or "DEFAULT",
                    work_type,
                    item["hazard_name"],
                    int(item["ai_raw_score"]),
                    int(item["final_score"]),
                    1 if item.get("excluded", False) else 0,
                    now_iso,
                ),
            )
        conn.commit()


def save_ptw_record(
    ptw_id: str,
    zone_code: str,
    work_type: str,
    gps_coords: str,
    address: str,
    verification_status: str,
    signer_count: int,
    integrity_hash: str,
) -> None:
    now_iso = datetime.datetime.now(pytz.timezone("Asia/Seoul")).isoformat(timespec="seconds")
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO ptw_records
            (ptw_id, zone_code, work_type, gps_coords, address, verification_status, signer_count, integrity_hash, approved_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                ptw_id,
                zone_code,
                work_type,
                gps_coords,
                address,
                verification_status,
                signer_count,
                integrity_hash,
                now_iso,
            ),
        )
        conn.commit()


# ---------------------------------------------------------------------------
# 3. 이미지 최적화, 번호 마커 렌더링 및 워드 서식 유틸리티
# ---------------------------------------------------------------------------
def compress_image(uploaded_file: Any, max_size: tuple[int, int] = (1024, 1024), quality: int = 80) -> Image.Image:
    img = Image.open(uploaded_file)
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail(max_size)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    buf.seek(0)
    return Image.open(buf).convert("RGB")


def image_to_jpeg_bytes(img: Image.Image, quality: int = 80) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality)
    return buf.getvalue()


def canvas_rgba_to_jpeg_bytes(canvas_image_data: Any) -> bytes:
    """RGBA 서명 배열을 흰색 배경과 합성하여 워드 호환 JPEG 바이트로 변환합니다."""
    rgba_img = Image.fromarray(canvas_image_data.astype("uint8"), mode="RGBA")
    white_bg = Image.new("RGBA", rgba_img.size, (255, 255, 255, 255))
    composited = Image.alpha_composite(white_bg, rgba_img).convert("RGB")
    buf = io.BytesIO()
    composited.save(buf, format="JPEG", quality=90)
    return buf.getvalue()


def draw_hazard_markers(img: Image.Image, hazards: list[dict[str, Any]]) -> Image.Image:
    marked = img.copy()
    draw = ImageDraw.Draw(marked)
    w, h = marked.size
    radius = max(16, int(min(w, h) * 0.032))

    try:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf", size=int(radius * 1.15))
    except OSError:
        font = ImageFont.load_default()

    grade_colors = {
        "상": (220, 38, 38),
        "중": (234, 88, 12),
        "하": (37, 99, 235),
    }

    for item in hazards:
        if item.get("excluded", False):
            continue
        cx = int(max(0, min(1000, item.get("x", 500))) / 1000.0 * w)
        cy = int(max(0, min(1000, item.get("y", 500))) / 1000.0 * h)
        color = grade_colors.get(item.get("grade", "중"), (220, 38, 38))

        draw.ellipse(
            (cx - radius - 2, cy - radius - 2, cx + radius + 2, cy + radius + 2),
            fill=(255, 255, 255),
        )
        draw.ellipse(
            (cx - radius, cy - radius, cx + radius, cy + radius),
            fill=color,
        )
        label = str(item["id"])
        bbox = draw.textbbox((0, 0), label, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((cx - tw / 2, cy - th / 2 - 2), label, fill=(255, 255, 255), font=font)

    return marked


def score_to_grade(score: int) -> str:
    if score >= 12:
        return "상"
    if score >= 6:
        return "중"
    return "하"


def apply_cell_shading(cell: Any, color: str = "D9D9D9") -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)


def format_table(table: Any, is_header_top: bool = True) -> None:
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if is_header_top and len(table.rows) > 0:
        for cell in table.rows[0].cells:
            apply_cell_shading(cell, "D9D9D9")


def insert_signatures_to_cell(
    cell: Any,
    signatures: list[dict[str, Any]],
    date_str: str = "",
    is_horizontal: bool = False,
) -> None:
    """등록된 참여 근로자 성명, 날짜, 전자서명 이미지를 셀에 자동 삽입합니다."""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not signatures:
        p.add_run("\n\n          (서명)")
        return

    if is_horizontal:
        p.add_run("위 판독 기준으로 정함.     ")
        for idx, sig in enumerate(signatures):
            if idx > 0:
                p.add_run("   /   ")
            p.add_run(f"근로자: {sig['name']}  ")
            img_buf = io.BytesIO(sig["jpeg_bytes"])
            p.add_run().add_picture(img_buf, width=Mm(16))
    else:
        for idx, sig in enumerate(signatures):
            if idx > 0:
                p.add_run("\n")
            if date_str:
                p.add_run(f"{sig['name']}\n{date_str}\n")
            else:
                p.add_run(f"{sig['name']}\n")
            img_buf = io.BytesIO(sig["jpeg_bytes"])
            p.add_run().add_picture(img_buf, width=Mm(18))


# ---------------------------------------------------------------------------
# 4. 보안 로그인 함수
# ---------------------------------------------------------------------------
def check_password() -> bool:
    def password_entered() -> None:
        if st.session_state["password"] == "0325":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### 🔒 한전MCS 안전점검 시스템")
        st.text_input("관리자 비밀번호 4자리를 입력하세요", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔒 한전MCS 안전점검 시스템")
        st.text_input("관리자 비밀번호 4자리를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("비밀번호가 일치하지 않습니다. 다시 시도해 주세요.")
        return False
    else:
        return True


# ---------------------------------------------------------------------------
# 5. Gemini API 연동 (1차 위험요인·법정기준 추출 & 2차 개선 전/후 비교 검증)
# ---------------------------------------------------------------------------
def get_gemini_client() -> genai.Client:
    api_key = ""
    try:
        api_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        api_key = ""
    api_key = api_key or os.environ.get("GEMINI_API_KEY", "")
    if not api_key:
        raise ValueError("GEMINI_API_KEY가 설정되지 않았습니다. Secrets 또는 환경변수를 확인해 주세요.")
    return genai.Client(api_key=api_key)


def analyze_initial_hazards(
    client: genai.Client,
    before_img: Image.Image,
    work_type: str,
    zone_code: str,
    weight_multiplier: float,
    model_name: str = DEFAULT_MODEL,
) -> dict[str, Any]:
    """1차 현장 사진에서 법정 판단기준, 유해·위험요인, 마커 좌표(0~1000), TBM 사항을 JSON으로 추출합니다."""
    prompt = f"""
당신은 한전MCS의 산업안전보건 및 현장점검 전문가입니다.
첨부된 작업현장 사진(공종: {work_type}, 구역: {zone_code})을 분석하여
고용노동부 고시 제2024-76호 '위험성평가표' 작성을 위한 데이터를 JSON 형식으로 반환하세요.

반드시 아래 JSON 객체 형식만 반환하세요(마크다운 코드블록 제외):
{{
  "criteria": {{
    "high": {{"condition": "상 수준 판단 기준", "action": "상 수준 조치 사항"}},
    "medium": {{"condition": "중 수준 판단 기준", "action": "중 수준 조치 사항"}},
    "low": {{"condition": "하 수준 판단 기준", "action": "하 수준 조치 사항"}},
    "legal_basis": "판단 근거 (관련 산업안전보건법 및 안전보건규칙 조항)"
  }},
  "hazards": [
    {{
      "id": 1,
      "hazard_name": "유해·위험요인 요약 (사진에서 보이는 것)",
      "accident_type": "재해형태 (추락/낙하/감전/협착/화재 등)",
      "current_measure": "현재 조치 상태",
      "frequency": 4,
      "severity": 4,
      "countermeasure": "구체적인 감소대책 (제12조)",
      "x": 520,
      "y": 380
    }}
  ],
  "tbm": ["작업 전 공유(TBM) 사항 1", "작업 전 공유(TBM) 사항 2", "작업 전 공유(TBM) 사항 3"]
}}

제약사항:
- hazards는 사진에서 확인되는 유해·위험요인을 2~5개 도출
- frequency(빈도): 1~5 정수, severity(강도): 1~4 정수
- x, y: 사진 내 해당 위험 지점의 정규화 중심 좌표 (0~1000 정수)
"""
    response = client.models.generate_content(
        model=model_name,
        contents=[
            types.Part.from_bytes(data=image_to_jpeg_bytes(before_img), mime_type="image/jpeg"),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.2,
        ),
    )
    raw_data = json.loads(response.text)
    raw_items = raw_data.get("hazards", [])
    enriched: list[dict[str, Any]] = []

    for idx, item in enumerate(raw_items, start=1):
        freq = int(item.get("frequency", 3))
        sev = int(item.get("severity", 3))
        raw_score = freq * sev
        adjusted_score = max(1, min(20, int(round(raw_score * weight_multiplier))))
        enriched.append({
            "id": idx,
            "hazard_name": item.get("hazard_name", "미확인 위험요인"),
            "accident_type": item.get("accident_type", "기타"),
            "current_measure": item.get("current_measure", "미조치"),
            "frequency": freq,
            "severity": sev,
            "ai_raw_score": raw_score,
            "final_score": adjusted_score,
            "grade": score_to_grade(adjusted_score),
            "countermeasure": item.get("countermeasure", ""),
            "x": int(item.get("x", 500)),
            "y": int(item.get("y", 500)),
            "excluded": False,
        })

    return {
        "criteria": raw_data.get("criteria", {}),
        "hazards": enriched,
        "tbm": raw_data.get("tbm", []),
    }


def verify_before_after_mitigation(
    client: genai.Client,
    before_marked_img: Image.Image,
    after_img: Image.Image,
    active_hazards: list[dict[str, Any]],
    model_name: str = DEFAULT_MODEL,
) -> dict[str, Any]:
    """[핵심 BM 로직] 개선 전(번호 마킹) 사진과 개선 후 사진을 AI가 대조해 항목별 위험 해소 여부를 판정합니다."""
    hazard_summary = json.dumps(
        [
            {
                "id": h["id"],
                "hazard_name": h["hazard_name"],
                "grade": h["grade"],
                "required_countermeasure": h["countermeasure"],
            }
            for h in active_hazards
        ],
        ensure_ascii=False,
    )

    prompt = f"""
당신은 산업현장 위험성평가 개선조치 검증 심사관입니다.
첫 번째 이미지는 위험 지점에 번호 마커가 표시된 [개선 전 사진]이고,
두 번째 이미지는 안전조치 완료 후 촬영한 [개선 후 사진]입니다.

아래 [식별된 위험요인 목록]의 각 번호별로 개선 후 사진에서 위험이 실제로 해소되었는지 엄격히 판정하세요.
식별된 위험요인 목록:
{hazard_summary}

반드시 아래 JSON 객체 형식만 반환하세요:
{{
  "same_location_verified": true,
  "items": [
    {{
      "id": 1,
      "status": "RESOLVED",
      "confidence": 0.92,
      "reason": "판정 근거를 1~2문장으로 기재",
      "rework_instruction": "미해소/부분해소 시 재조치 지시 (해소 시 빈 문자열)"
    }}
  ],
  "overall_summary": "전체 개선조치 검증 요약"
}}

판정 기준:
- status는 반드시 "RESOLVED"(완전 해소), "PARTIAL"(부분 해소), "UNRESOLVED"(미해소/확인불가) 중 하나
- 사진 각도나 위치가 완전히 다른 현장이면 same_location_verified를 false로 설정
"""
    response = client.models.generate_content(
        model=model_name,
        contents=[
            types.Part.from_bytes(data=image_to_jpeg_bytes(before_marked_img), mime_type="image/jpeg"),
            types.Part.from_bytes(data=image_to_jpeg_bytes(after_img), mime_type="image/jpeg"),
            prompt,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.1,
        ),
    )
    result = json.loads(response.text)

    hazard_grade_map = {h["id"]: h["grade"] for h in active_hazards}
    items = result.get("items", [])
    high_risk_unresolved = any(
        hazard_grade_map.get(it.get("id")) == "상" and it.get("status") != "RESOLVED"
        for it in items
    )
    resolved_count = sum(1 for it in items if it.get("status") == "RESOLVED")
    resolution_rate = (resolved_count / len(items)) if items else 0.0
    same_loc = bool(result.get("same_location_verified", False))

    gate_passed = same_loc and (not high_risk_unresolved) and (resolution_rate >= 0.8)
    result["resolution_rate"] = round(resolution_rate * 100, 1)
    result["high_risk_unresolved"] = high_risk_unresolved
    result["gate_passed"] = gate_passed
    return result


# ---------------------------------------------------------------------------
# 6. 무결성 해시 및 법정 양식(고용노동부 고시 제2024-76호) 워드 보고서 생성
# ---------------------------------------------------------------------------
def compute_integrity_hash(
    ptw_id: str,
    zone_code: str,
    gps_coords: str,
    hazards: list[dict[str, Any]],
    verification: dict[str, Any],
    signer_names: list[str],
) -> str:
    payload = json.dumps(
        {
            "ptw_id": ptw_id,
            "zone_code": zone_code,
            "gps_coords": gps_coords,
            "hazards": [
                {"id": h["id"], "name": h["hazard_name"], "score": h["final_score"]}
                for h in hazards
                if not h.get("excluded")
            ],
            "verification_gate": verification.get("gate_passed"),
            "signers": signer_names,
        },
        ensure_ascii=False,
        sort_keys=True,
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def build_legal_ptw_docx(
    ptw_id: str,
    zone_code: str,
    work_type: str,
    current_address: str,
    gps_coords: str,
    criteria: dict[str, Any],
    hazards: list[dict[str, Any]],
    tbm_list: list[str],
    verification: dict[str, Any],
    before_marked_img: Image.Image,
    after_img: Image.Image,
    signatures: list[dict[str, Any]],
    integrity_hash: str,
) -> bytes:
    """고용노동부 고시 제2024-76호 법정 양식 + 전/후 AI 검증 + PTW 무결성 해시 통합 워드 생성"""
    korea_tz = pytz.timezone("Asia/Seoul")
    now = datetime.datetime.now(korea_tz)
    current_time_str = now.strftime("%Y. %m. %d. %H:%M")
    date_only_str = now.strftime("%Y. %m. %d.")

    active_hazards = [h for h in hazards if not h.get("excluded")]
    total_count = len(active_hazards)
    count_high = sum(1 for h in active_hazards if h.get("grade") == "상")
    count_mid = sum(1 for h in active_hazards if h.get("grade") == "중")
    count_low = sum(1 for h in active_hazards if h.get("grade") == "하")

    doc = Document()
    for section in doc.sections:
        section.top_margin = Mm(15)
        section.bottom_margin = Mm(15)
        section.left_margin = Mm(18)
        section.right_margin = Mm(18)

    style = doc.styles["Normal"]
    style.font.name = "맑은 고딕"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    style.font.size = Pt(10.5)

    title = doc.add_heading("현장 위험성 평가표 및 작업개시 승인(PTW)서", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.font.name = "맑은 고딕"
    title_run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    title_run.font.bold = True
    title_run.font.color.rgb = None

    subtitle = doc.add_paragraph(
        f"사업장 위험성평가에 관한 지침(고용노동부고시 제2024-76호)의 절차를 따름\n"
        f"[승인번호: {ptw_id} | 무결성 해시(SHA-256): {integrity_hash[:24]}...]"
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 표 1: 기본 정보 및 서명란
    table1 = doc.add_table(rows=5, cols=4)
    table1.style = "Table Grid"
    table1.cell(0, 0).text, table1.cell(0, 1).text = "사업장 / 구역", f"[          지점]\n{zone_code}"
    table1.cell(0, 2).text, table1.cell(0, 3).text = "평가 일시", current_time_str
    table1.cell(1, 0).text, table1.cell(1, 1).text = "평가 대상(GPS)", f"{current_address}\n({gps_coords})" if gps_coords else current_address
    table1.cell(1, 2).text, table1.cell(1, 3).text = "공종 / 평가 종류", f"{work_type}\n□ 최초 □ 수시 □ 정기 □ 상시"
    table1.cell(2, 0).text, table1.cell(2, 1).text = "평가 방법", "현장점검·전/후 AI 대조검증"
    table1.cell(2, 2).text, table1.cell(2, 3).text = "평가 결과", f"{total_count}건 (상 {count_high}·중 {count_mid}·하 {count_low})"
    table1.cell(3, 0).text, table1.cell(3, 1).text = "AI 개선 검증", f"통과 (해소율 {verification.get('resolution_rate', 0)}%)"
    table1.cell(3, 2).text, table1.cell(3, 3).text = "PTW 승인 상태", "작업개시 승인 완료"

    table1.cell(4, 0).text = "평가 책임자"
    table1.cell(4, 1).text = "\n\n          (서명)"
    table1.cell(4, 2).text = "참여 근로자"
    insert_signatures_to_cell(table1.cell(4, 3), signatures, date_only_str)

    format_table(table1, is_header_top=False)
    for r in range(5):
        apply_cell_shading(table1.cell(r, 0), "D9D9D9")
        apply_cell_shading(table1.cell(r, 2), "D9D9D9")
    doc.add_paragraph()

    # 1. 사업장 위험성 평가 절차
    doc.add_heading("1. 사업장 위험성 평가 절차", level=2)
    table2 = doc.add_table(rows=7, cols=3)
    table2.style = "Table Grid"
    hdr2 = table2.rows[0].cells
    hdr2[0].text, hdr2[1].text, hdr2[2].text = "절차", "AI 판독 및 검증", "관리자 확인"
    table2.cell(1, 0).text, table2.cell(1, 1).text, table2.cell(1, 2).text = "가. 사전준비(제9조)", "위험성 수준 기준 기본안", "위험성의 수준과 그 수준을 판단하는 기준, 허용 가능한 위험성의 수준을 확정(제2항)"
    table2.cell(2, 0).text, table2.cell(2, 1).text, table2.cell(2, 2).text = "나. 유해·위험요인 파악(제10조)", f"현장점검(제10조 제1호)간 첨부된 사진에서 {total_count}건", "사진에 포함되지 않은 현장 위험물(화학물질·소음·중량물 등)과 근로자가 제안한 위험요인도 추가하여 파악"
    table2.cell(3, 0).text, table2.cell(3, 1).text, table2.cell(3, 2).text = "다. 위험성 결정(제11조)", "상·중·하 수준 및 구역 가중치 반영", "허용 가능한지 결정, 현장 근로자 참여(제6조)"
    table2.cell(4, 0).text, table2.cell(4, 1).text, table2.cell(4, 2).text = "라. 위험성 감소대책 수립·실행(제12조)", "감소대책 수립 및 개선 전·후 사진 AI 대조 검증", "감소대책에 대한 담당자 및 완료일 작성, 실행 후 위험성 수준 재확인"
    table2.cell(5, 0).text, table2.cell(5, 1).text, table2.cell(5, 2).text = "마. 공유(제13조)", "작업 전 위험요인과 감소대책을 공유", "작업 전 안전점검회의(TBM)에서 공유하고 서명"
    table2.cell(6, 0).text, table2.cell(6, 1).text, table2.cell(6, 2).text = "바. 기록·보존(제14조)", "본 문서 (SHA-256 무결성 해시 포함)", "서명한 본 문서를 보관"
    format_table(table2)
    doc.add_paragraph("※ 매월 1회 이상 위험개소 순회점검 통해 유해·위험요인을 발굴하여 제11조의 위험성 결정 및 제12조의 위험성 감소대책을 수립·실행 할 것.(사업장 위험성 평가에 관한 지침 제15조 제4항)")

    # 2. 법적 의무
    doc.add_heading("2. 법적 의무", level=2)
    table3 = doc.add_table(rows=4, cols=4)
    table3.style = "Table Grid"
    hdr3 = table3.rows[0].cells
    hdr3[0].text, hdr3[1].text, hdr3[2].text, hdr3[3].text = "의무", "조문", "과태료", "비고"
    table3.cell(1, 0).text, table3.cell(1, 1).text, table3.cell(1, 2).text, table3.cell(1, 3).text = "위험성평가 실시", "제36조 제1항", "1천만원 이하", "위험성평가표 전체"
    table3.cell(2, 0).text, table3.cell(2, 1).text, table3.cell(2, 2).text, table3.cell(2, 3).text = "근로자의 참여\n근로자대표 참여\n결과 알림", "제36조 제2항ㆍ제3항ㆍ제4항 전단", "500만원 이하", "참여 근로자 서명\n작업 전 공유"
    table3.cell(3, 0).text, table3.cell(3, 1).text, table3.cell(3, 2).text, table3.cell(3, 3).text = "결과 기록·보존", "제36조 제5항", "300만원 이하", "서명된 본 평가표 보관"
    format_table(table3)
    doc.add_paragraph("※ 과태료 시행 : 상시 50명 이상(건설업은 공사 금액 50억원 이상) 2027년 1월 1일 시행(산업안전보건법 부칙 제1조). 중대재해로 이어질 수 있는 위험은 작업 전 안전점검회의(TBM) 등을 통해 상시 알리도록 노력해야 합니다(제36조 제4항 후단).")

    # 3. 위험성 수준 기준(AI 기준)
    doc.add_heading("3. 위험성 수준 기준(AI 기준)", level=2)
    table4 = doc.add_table(rows=6, cols=4)
    table4.style = "Table Grid"
    hdr4 = table4.rows[0].cells
    hdr4[0].text, hdr4[1].text, hdr4[2].text, hdr4[3].text = "수준", "기준 제시", "판단", "조치"
    table4.cell(1, 0).text, table4.cell(1, 1).text, table4.cell(1, 2).text, table4.cell(1, 3).text = "상", criteria.get("high", {}).get("condition", ""), "허용 불가", criteria.get("high", {}).get("action", "")
    table4.cell(2, 0).text, table4.cell(2, 1).text, table4.cell(2, 2).text, table4.cell(2, 3).text = "중", criteria.get("medium", {}).get("condition", ""), "허용 불가", criteria.get("medium", {}).get("action", "")
    table4.cell(3, 0).text, table4.cell(3, 1).text, table4.cell(3, 2).text, table4.cell(3, 3).text = "하", criteria.get("low", {}).get("condition", ""), "허용 가능", criteria.get("low", {}).get("action", "")

    table4.cell(4, 0).text = "판단 근거"
    table4.cell(4, 1).text = criteria.get("legal_basis", "")
    table4.cell(4, 1).merge(table4.cell(4, 3))

    table4.cell(5, 0).merge(table4.cell(5, 3))
    insert_signatures_to_cell(table4.cell(5, 0), signatures, date_only_str, is_horizontal=True)

    format_table(table4)
    apply_cell_shading(table4.cell(5, 0), "D9D9D9")
    doc.add_page_break()

    # 4. AI 판독 및 개선 전·후 비교 사진
    doc.add_heading("4. AI 판독 및 개선 전·후 비교", level=2)
    doc.add_paragraph("※ 개선 전 사진에 표시된 원형 번호 마커는 아래 '4-1. 위험성 평가표'의 번호와 일치합니다.")

    pic_table = doc.add_table(rows=2, cols=2)
    pic_table.style = "Table Grid"
    pic_table.cell(0, 0).text = "[개선 전] 위험요인 번호 마킹"
    pic_table.cell(0, 1).text = "[개선 후] 안전조치 완료 사진"

    b_buf = io.BytesIO(image_to_jpeg_bytes(before_marked_img, quality=85))
    a_buf = io.BytesIO(image_to_jpeg_bytes(after_img, quality=85))
    pic_table.cell(1, 0).paragraphs[0].add_run().add_picture(b_buf, width=Mm(78))
    pic_table.cell(1, 1).paragraphs[0].add_run().add_picture(a_buf, width=Mm(78))
    format_table(pic_table)
    doc.add_paragraph()

    # 4-1. 위험성 평가표(AI 분석 및 개선 검증)
    doc.add_heading("4-1. 위험성 평가표(AI 분석 및 개선 검증)", level=2)
    verify_map = {it.get("id"): it for it in verification.get("items", [])}
    status_label_map = {"RESOLVED": "해소", "PARTIAL": "부분해소", "UNRESOLVED": "미해소"}

    table5 = doc.add_table(rows=1, cols=8)
    table5.style = "Table Grid"
    hdr5 = table5.rows[0].cells
    hdr5[0].text = "번호"
    hdr5[1].text = "유해·위험요인\n(사진에서 보이는 것)"
    hdr5[2].text = "재해형태"
    hdr5[3].text = "현재 조치"
    hdr5[4].text = "위험 수준\n(점수)"
    hdr5[5].text = "감소 대책\n(제12조)"
    hdr5[6].text = "AI 개선\n검증 결과"
    hdr5[7].text = "담당\n완료일"

    for item in active_hazards:
        v_item = verify_map.get(item["id"], {})
        v_status = status_label_map.get(v_item.get("status", ""), v_item.get("status", "-"))
        v_reason = v_item.get("reason", "")

        row = table5.add_row().cells
        row[0].text = str(item.get("id", ""))
        row[1].text = item.get("hazard_name", "")
        row[2].text = item.get("accident_type", "")
        row[3].text = item.get("current_measure", "")
        row[4].text = f"{item.get('grade', '')}\n({item.get('final_score', '')}점)"
        row[5].text = item.get("countermeasure", "")
        row[6].text = f"[{v_status}]\n{v_reason}"
        insert_signatures_to_cell(row[7], signatures, date_only_str)

    format_table(table5)
    doc.add_paragraph()

    # 5. 안전점검회의(TBM) 공유
    doc.add_heading("5. 안전점검회의(TBM) 공유", level=2)
    for i, tbm_text in enumerate(tbm_list):
        doc.add_paragraph(f"{i + 1}. {tbm_text}")

    table6 = doc.add_table(rows=1, cols=3)
    table6.style = "Table Grid"
    hdr6 = table6.rows[0].cells
    hdr6[0].text, hdr6[1].text, hdr6[2].text = "공유 일시", "공유한 사람 (관리자)", "참석자 서명 (근로자)"

    row6 = table6.add_row().cells
    row6[0].text = current_time_str
    row6[1].text = "\n\n          (서명)"
    insert_signatures_to_cell(row6[2], signatures, "")
    format_table(table6)

    doc.add_paragraph("\n※ 이 문서는 초안입니다. 유해·위험요인은 본 AI 모델이 사진에서 분석한 결과이고, 위험성 수준과 감소대책 후보는 코드가 정해진 규칙으로 고른 것입니다. 허용 가능 여부의 결정, 대책 선택과 실행, 근로자 참여는 관리자가 시행합니다.(사업장 위험성 평가에 관한 지침 제6조·제11조·제12조)")
    doc.add_paragraph("※ 사진에 보이지 않는 위험(화학물질·소음·분진·중량물 취급 작업 자세 등)은 이 표에 들어 있지 않습니다. 파악 단계에서 추가 작성해 주세요.")
    doc.add_paragraph(f"※ 본 문서의 위변조 검증 해시(SHA-256): {integrity_hash}")

    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue()


# ---------------------------------------------------------------------------
# 7. 메인 앱 실행
# ---------------------------------------------------------------------------
def main() -> None:
    if not check_password():
        return

    init_db()

    # 세션 상태 초기화
    for key, default in [
        ("gps_coords", ""),
        ("current_address", ""),
        ("before_img", None),
        ("before_marked_img", None),
        ("after_img", None),
        ("criteria", {}),
        ("hazards", []),
        ("tbm_list", []),
        ("verification", None),
        ("signatures", []),
        ("ptw_approved", False),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    # 사이드바
    with st.sidebar:
        try:
            st.image("safety_check_app/logo.png", use_container_width=True)
        except Exception:
            st.title("한전MCS 안전처")
        st.divider()
        st.header("📌 사용 방법")
        st.markdown(
            "**1단계:** 현장 GPS 주소 자동 수집 및 구역·공종 입력  \n"
            "**2단계:** 개선 전 사진 업로드 → AI 분석 및 가중치 보정  \n"
            "**3단계:** 안전조치 후 사진 업로드 → AI 전·후 대조 검증  \n"
            "**4단계:** 참여 근로자 전자서명 → PTW 승인 및 법정 양식 발급"
        )
        st.divider()
        st.header("⚙️ 설정")
        model_name = st.text_input("사용 AI 모델", value=DEFAULT_MODEL)
        required_signers = st.number_input("작업개시 필수 서명 인원", min_value=1, max_value=20, value=1)
        if st.button("🔄 점검 세션 초기화", use_container_width=True):
            for k in [
                "before_img", "before_marked_img", "after_img",
                "criteria", "hazards", "tbm_list", "verification", "signatures", "ptw_approved",
            ]:
                if k in ("hazards", "tbm_list", "signatures"):
                    st.session_state[k] = []
                elif k == "criteria":
                    st.session_state[k] = {}
                elif k == "ptw_approved":
                    st.session_state[k] = False
                else:
                    st.session_state[k] = None
            st.rerun()
        st.divider()
        st.caption("© 2027 한전MCS. All rights reserved.")

    # 상단 로고 헤더
    try:
        with open("safety_check_app/logo.png", "rb") as image_file:
            logo_base64 = base64.b64encode(image_file.read()).decode()
        st.markdown(f"""
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 15px;">
                <img src="data:image/png;base64,{logo_base64}" width="60">
                <h1 style="margin: 0; font-size: 24px; letter-spacing: -0.5px;">AI 위험성평가 & 작업승인(PTW)</h1>
            </div>
        """, unsafe_allow_html=True)
    except Exception:
        st.title("AI 위험성평가 & 작업승인(PTW)")

    st.markdown("■ 고용노동부 고시 제2024-76호 지침에 따른 **위험성평가 및 개선 전·후 AI 검증**을 지원합니다.")
    st.divider()

    # -----------------------------------------------------------------------
    # STEP 1: 현장 위치 자동 수집 (GPS) & 메타데이터
    # -----------------------------------------------------------------------
    st.subheader("📍 1단계: 현장 위치 자동 수집 및 작업 정보")
    st.info("현재 위치를 평가표에 자동 입력하려면 아래 버튼을 눌러주세요. (스마트폰 GPS 켜기 필수)")
    location = streamlit_geolocation()

    if location is not None and location.get("latitude") is not None:
        lat, lon = location["latitude"], location["longitude"]
        coords_str = f"{lat:.5f}, {lon:.5f}"
        if st.session_state["gps_coords"] != coords_str:
            st.session_state["gps_coords"] = coords_str
            try:
                geolocator = Nominatim(user_agent="kepco_safety_app")
                addr = geolocator.reverse((lat, lon))
                st.session_state["current_address"] = addr.address if addr else f"위도: {lat}, 경도: {lon}"
            except Exception:
                st.session_state["current_address"] = f"위도: {lat}, 경도: {lon}"

    if st.session_state["current_address"]:
        st.success(f"✅ 현재 현장 주소: {st.session_state['current_address']} ({st.session_state['gps_coords']})")

    col1, col2 = st.columns(2)
    zone_code = col1.text_input("작업 구역 코드", value="배전선로-A구역")
    work_type = col2.selectbox("공종 선택", ["전기작업", "고소작업", "화기작업", "밀폐공간", "중량물 취급", "일반작업"])

    weight_mul = get_zone_weight_multiplier(zone_code, work_type)
    st.caption(f"📊 동일 구역·공종 과거 피드백 보정 가중치: **×{weight_mul}** (보정 이력 누적 시 자동 갱신)")

    st.divider()

    # -----------------------------------------------------------------------
    # STEP 2: 1차 위험요인 탐지 (번호 마커 + 안전관리자 피드백 보정)
    # -----------------------------------------------------------------------
    st.subheader("📸 2단계: [개선 전] 현장 사진 분석 및 위험요인 보정")
    before_file = st.file_uploader(
        "Tip: 사진에 작업자의 행동이나 주요 설비(전신주, 개구부 등)가 잘 보일수록 AI가 더 정확히 판독합니다.",
        type=["jpg", "jpeg", "png"],
        key="before_up",
    )

    if before_file and st.button("🔍 AI 위험요인 분석 및 번호 마킹 실행", type="primary", use_container_width=True):
        try:
            client = get_gemini_client()
            with st.spinner("이미지 최적화 및 Gemini 멀티모달 위험성 분석 중..."):
                before_img = compress_image(before_file)
                result = analyze_initial_hazards(
                    client, before_img, work_type, zone_code, weight_mul, model_name=model_name
                )
                marked_img = draw_hazard_markers(before_img, result["hazards"])
                st.session_state["before_img"] = before_img
                st.session_state["before_marked_img"] = marked_img
                st.session_state["criteria"] = result["criteria"]
                st.session_state["hazards"] = result["hazards"]
                st.session_state["tbm_list"] = result["tbm"]
                st.session_state["verification"] = None
                st.session_state["ptw_approved"] = False
            st.success(f"총 {len(result['hazards'])}건의 유해·위험요인을 탐지하고 번호 마커를 표시했습니다.")
        except Exception as err:
            st.error(f"분석 중 오류가 발생했습니다: {err}")

    if st.session_state["before_marked_img"] is not None and st.session_state["hazards"]:
        st.image(st.session_state["before_marked_img"], caption="1차 위험요인 번호 마킹 결과", use_container_width=True)
        st.markdown("**안전관리자 검토 및 오탐·점수 보정 (구역 이력 DB 자동 학습)**")
        for h in st.session_state["hazards"]:
            with st.expander(f"#{h['id']} [{h['grade']} / {h['final_score']}점] {h['hazard_name']}", expanded=False):
                c_a, c_b = st.columns([1, 2])
                h["excluded"] = c_a.checkbox("오탐 제외", value=h["excluded"], key=f"exc_{h['id']}")
                h["final_score"] = c_b.slider(
                    "위험성 점수 보정 (빈도×강도)",
                    min_value=1,
                    max_value=20,
                    value=int(h["final_score"]),
                    key=f"scr_{h['id']}",
                )
                h["grade"] = score_to_grade(h["final_score"])
                st.write(f"**재해형태:** {h['accident_type']} | **현재 조치:** {h['current_measure']}")
                st.write(f"**감소 대책:** {h['countermeasure']}")

        if st.button("💾 보정 결과 확정 및 구역 이력 DB 반영"):
            save_feedback_records(zone_code, work_type, st.session_state["hazards"])
            st.session_state["before_marked_img"] = draw_hazard_markers(
                st.session_state["before_img"], st.session_state["hazards"]
            )
            st.success("보정 이력이 저장되었으며 번호 마커가 업데이트되었습니다.")

    st.divider()

    # -----------------------------------------------------------------------
    # STEP 3: 개선 전/후 AI 자동 비교 검증 루프 (승인 게이트 제어)
    # -----------------------------------------------------------------------
    if st.session_state["before_marked_img"] is not None:
        st.subheader("✅ 3단계: [개선 후] 사진 대조 및 AI 자동 검증")
        st.info("위험요인에 대한 안전조치를 완료한 뒤 현장 사진을 업로드하세요. 고위험(상) 항목이 해소되어야 서명 단계가 열립니다.")
        after_file = st.file_uploader("안전조치 완료 후 현장 사진 업로드", type=["jpg", "jpeg", "png"], key="after_up")

        if after_file and st.button("🔄 개선 전/후 AI 자동 검증 실행", type="primary", use_container_width=True):
            active_hazards = [h for h in st.session_state["hazards"] if not h.get("excluded")]
            if not active_hazards:
                st.warning("검증할 유효 위험요인이 없습니다.")
            else:
                try:
                    client = get_gemini_client()
                    with st.spinner("개선 전 마킹 사진과 개선 후 사진을 AI가 대조 검증 중..."):
                        after_img = compress_image(after_file)
                        verification = verify_before_after_mitigation(
                            client,
                            st.session_state["before_marked_img"],
                            after_img,
                            active_hazards,
                            model_name=model_name,
                        )
                        st.session_state["after_img"] = after_img
                        st.session_state["verification"] = verification
                        st.session_state["ptw_approved"] = False
                except Exception as err:
                    st.error(f"전/후 검증 중 오류가 발생했습니다: {err}")

    if st.session_state["verification"] is not None:
        v = st.session_state["verification"]
        c_before, c_after = st.columns(2)
        c_before.image(st.session_state["before_marked_img"], caption="개선 전 (마킹)", use_container_width=True)
        c_after.image(st.session_state["after_img"], caption="개선 후", use_container_width=True)

        if v["gate_passed"]:
            st.success(f"🟢 AI 검증 통과 (해소율 {v['resolution_rate']}%) — 4단계 참여 근로자 전자서명이 잠금 해제되었습니다.")
        else:
            st.error(
                f"🔴 작업 개시 승인 차단 (해소율 {v['resolution_rate']}%, "
                f"동일현장={v.get('same_location_verified')}, 고위험 미해소={v.get('high_risk_unresolved')}) "
                "— 미해소 항목 재조치 후 사진을 다시 업로드하세요."
            )

        for item in v.get("items", []):
            status_badge = {"RESOLVED": "✅ 해소", "PARTIAL": "⚠️ 부분해소", "UNRESOLVED": "❌ 미해소"}.get(
                item.get("status", ""), item.get("status", "")
            )
            st.write(f"- **#{item['id']} [{status_badge}]** (신뢰도 {int(item.get('confidence', 0) * 100)}%): {item.get('reason', '')}")
            if item.get("rework_instruction"):
                st.warning(f"재조치 지시: {item['rework_instruction']}")

    # -----------------------------------------------------------------------
    # STEP 4: 참여 근로자 전자서명 & 최종 PTW 승인 / 법정 양식(.docx) 발급
    # -----------------------------------------------------------------------
    if st.session_state["verification"] and st.session_state["verification"]["gate_passed"]:
        st.divider()
        st.subheader("✍️ 4단계: 참여 근로자 전자서명 및 법정 평가표 발급")
        st.info("아래 패드에 참여 근로자의 서명을 입력해 주세요. (평가 책임자는 출력 후 수기 서명)")

        current_count = len(st.session_state["signatures"])
        st.write(f"현재 서명 완료 인원: **{current_count} / {required_signers}명**")

        worker_name = st.text_input("참여 근로자 성명", placeholder="예: 홍길동")
        st.markdown("**참여 근로자 서명**")
        canvas_res = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#f0f2f6",
            height=150,
            width=260,
            drawing_mode="freedraw",
            return_image_data=True,
            key=f"worker_sign_{current_count}",
        )

        if st.button("✍️ 위험 인지 확인 및 서명 등록"):
            if not worker_name.strip():
                st.warning("참여 근로자 성명을 입력해 주세요.")
            elif canvas_res.image_data is None:
                st.warning("서명 패드에 서명을 입력해 주세요.")
            else:
                jpeg_bytes = canvas_rgba_to_jpeg_bytes(canvas_res.image_data)
                korea_tz = pytz.timezone("Asia/Seoul")
                st.session_state["signatures"].append({
                    "name": worker_name.strip(),
                    "signed_at": datetime.datetime.now(korea_tz).strftime("%Y-%m-%d %H:%M:%S"),
                    "jpeg_bytes": jpeg_bytes,
                })
                st.rerun()

        if st.session_state["signatures"]:
            st.write("**서명 완료 명단:** " + ", ".join(s["name"] for s in st.session_state["signatures"]))

        if len(st.session_state["signatures"]) >= required_signers:
            if st.button("🚀 최종 작업개시 승인(PTW) 및 법정 위험성평가표 생성", type="primary", use_container_width=True):
                korea_tz = pytz.timezone("Asia/Seoul")
                now_str = datetime.datetime.now(korea_tz).strftime("%Y%m%d-%H%M%S")
                ptw_id = f"PTW-{now_str}"
                signer_names = [s["name"] for s in st.session_state["signatures"]]

                integrity_hash = compute_integrity_hash(
                    ptw_id,
                    zone_code,
                    st.session_state["gps_coords"],
                    st.session_state["hazards"],
                    st.session_state["verification"],
                    signer_names,
                )
                save_ptw_record(
                    ptw_id,
                    zone_code,
                    work_type,
                    st.session_state["gps_coords"],
                    st.session_state["current_address"],
                    "PASSED",
                    len(signer_names),
                    integrity_hash,
                )
                docx_bytes = build_legal_ptw_docx(
                    ptw_id=ptw_id,
                    zone_code=zone_code,
                    work_type=work_type,
                    current_address=st.session_state["current_address"] or "현장 주소 미수집",
                    gps_coords=st.session_state["gps_coords"],
                    criteria=st.session_state["criteria"],
                    hazards=st.session_state["hazards"],
                    tbm_list=st.session_state["tbm_list"],
                    verification=st.session_state["verification"],
                    before_marked_img=st.session_state["before_marked_img"],
                    after_img=st.session_state["after_img"],
                    signatures=st.session_state["signatures"],
                    integrity_hash=integrity_hash,
                )
                st.session_state["ptw_approved"] = True
                st.success(f"🎉 작업개시 승인 및 법정 평가표 작성 완료! (승인번호: **{ptw_id}**)")
                st.download_button(
                    label="📄 현장 위험성 평가표 및 PTW 증빙서(.docx) 다운로드",
                    data=docx_bytes,
                    file_name=f"현장_위험성_평가표_{ptw_id}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )


if __name__ == "__main__":
    main()