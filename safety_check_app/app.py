"""한전MCS 전국 지점 통합 AI 위험성평가 & 작업승인(PTW) 플랫폼."""

from __future__ import annotations

import datetime
import hashlib
import io
import json
import os
from typing import Any

from docx import Document
from docx.shared import Inches
from google import genai
from google.genai import types
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

import db_supabase as db

try:
    from geopy.geocoders import Nominatim
except ImportError:
    Nominatim = None

try:
    from streamlit_geolocation import streamlit_geolocation
except ImportError:
    streamlit_geolocation = None

try:
    from streamlit_drawable_canvas import st_canvas
except ImportError:
    st_canvas = None

# ==========================================
# 페이지 기본 설정 및 상수
# ==========================================
st.set_page_config(
    page_title="한전MCS AI 위험성평가 & 작업승인(PTW)",
    page_icon="⚡",
    layout="wide",
)

WORK_TYPES = [
    "전력량계 검침",
    "청구서 송달/현장방문",
    "전기작업(단전,전류제한기)",
    "배전공사 고소작업",
    "도서발전소 설비점검",
    "일반작업",
]

STATUS_LABELS = {
    "PENDING_REVIEW": "🟠 관리자 검토 대기",
    "READY_FOR_ACTION": "🔵 조치 및 2차 검증 가능",
    "APPROVED": "🟢 PTW 최종 승인 완료",
}


# ==========================================
# AI / 이미지 / 법정 양식 유틸리티 함수
# ==========================================
def get_gemini_client() -> genai.Client:
    api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
    if not api_key:
        raise ValueError("GEMINI_API_KEY가 Secrets에 설정되지 않았습니다.")
    return genai.Client(api_key=api_key)


def calc_grade(score: int) -> str:
    if score >= 12:
        return "상"
    if score >= 6:
        return "중"
    return "하"


def draw_hazard_markers(img: Image.Image, hazards: list[dict[str, Any]]) -> Image.Image:
    """개선 전 사진 위에 유효한 위험요인의 등급별 색상 번호 마커를 렌더링합니다."""
    marked = img.copy().convert("RGB")
    draw = ImageDraw.Draw(marked)
    w, h = marked.size
    radius = max(18, int(min(w, h) * 0.03))

    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", size=int(radius * 1.1))
    except Exception:
        try:
            font = ImageFont.load_default(size=int(radius * 1.1))
        except Exception:
            font = ImageFont.load_default()

    color_map = {
        "상": (220, 38, 38),   # Red
        "중": (234, 88, 12),   # Orange
        "하": (37, 99, 235),   # Blue
    }

    for item in hazards:
        if item.get("excluded", False):
            continue
        x_pct = max(5, min(95, float(item.get("x_pct", 50))))
        y_pct = max(5, min(95, float(item.get("y_pct", 50))))
        cx, cy = int(w * x_pct / 100.0), int(h * y_pct / 100.0)
        fill_color = color_map.get(item.get("grade", "중"), (234, 88, 12))

        draw.ellipse(
            [cx - radius, cy - radius, cx + radius, cy + radius],
            fill=fill_color,
            outline=(255, 255, 255),
            width=max(2, radius // 6),
        )
        label = str(item.get("id", ""))
        bbox = draw.textbbox((0, 0), label, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((cx - tw / 2, cy - th / 2 - 2), label, fill=(255, 255, 255), font=font)

    return marked


def analyze_initial_hazards(
    img: Image.Image,
    branch_name: str,
    zone_code: str,
    work_type: str,
    weight_multiplier: float,
) -> tuple[list[dict[str, Any]], list[str]]:
    """1차 현장 사진 AI 위험성평가 분석 + 구역 과거 이력 가중치 자동 반영."""
    client = get_gemini_client()
    prompt = f"""
당신은 대한민국 산업안전보건법(고용노동부 고시 제2024-76호) 기준의 베테랑 전기·현장 안전관리자입니다.
첨부된 작업 전 현장 사진을 정밀 분석하여 유해·위험요인(최대 5개)과 핵심 TBM(작업 전 안전점검회의) 전달사항 3가지를 도출하세요.

[현장 정보]
- 지점명: {branch_name}
- 현장명: {zone_code}
- 공종(작업): {work_type}

[출력 JSON 스키마 - 반드시 유효한 JSON만 출력]
{{
  "hazards": [
    {{
      "id": 1,
      "hazard_name": "구체적인 유해·위험요인 명칭 (예: 개폐기 인근 노출 충전부 접촉 위험)",
      "accident_type": "재해유형 (감전, 추락, 전도, 협착, 충돌, 낙하 등)",
      "frequency": 3,
      "severity": 4,
      "countermeasure": "작업자가 즉시 실행하고 사진으로 증명할 수 있는 구체적 감소대책",
      "x_pct": 50,
      "y_pct": 50
    }}
  ],
  "tbm_points": [
    "작업 전 근로자 공유 핵심 안전수칙 1",
    "작업 전 근로자 공유 핵심 안전수칙 2",
    "작업 전 근로자 공유 핵심 안전수칙 3"
  ]
}}
- frequency(가능성): 1~5 정수
- severity(중대성): 1~4 정수
- x_pct, y_pct: 사진 내 위험요인 중심 좌표 (0~100 백분율)
"""
    response = client.models.generate_content(
        model="gemini-3.5-flash-lite",
        contents=[prompt, img],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.2,
        ),
    )
    data = json.loads(response.text)
    raw_hazards = data.get("hazards", [])
    tbm_points = data.get("tbm_points", [])

    processed: list[dict[str, Any]] = []
    for idx, h in enumerate(raw_hazards, start=1):
        freq = max(1, min(5, int(h.get("frequency", 2))))
        sev = max(1, min(4, int(h.get("severity", 2))))
        ai_raw = freq * sev
        weighted = max(1, min(20, int(round(ai_raw * weight_multiplier))))
        processed.append({
            "id": idx,
            "hazard_name": h.get("hazard_name", f"위험요인 {idx}"),
            "accident_type": h.get("accident_type", "기타"),
            "frequency": freq,
            "severity": sev,
            "ai_raw_score": ai_raw,
            "weight_applied": weight_multiplier,
            "final_score": weighted,
            "grade": calc_grade(weighted),
            "countermeasure": h.get("countermeasure", "안전조치 이행 필요"),
            "x_pct": float(h.get("x_pct", 50)),
            "y_pct": float(h.get("y_pct", 50)),
            "excluded": False,
        })
    return processed, tbm_points


def verify_before_after_images(
    before_img: Image.Image,
    after_img: Image.Image,
    active_hazards: list[dict[str, Any]],
) -> dict[str, Any]:
    """2차 개선 전·후 사진 AI 대조 검증 및 폐루프(Closed-loop) 승인 게이트 판정."""
    client = get_gemini_client()
    hazard_summary = [
        {
            "id": h["id"],
            "hazard_name": h["hazard_name"],
            "grade": h["grade"],
            "countermeasure": h["countermeasure"],
        }
        for h in active_hazards
    ]
    prompt = f"""
당신은 작업승인(PTW) 최종 검증 AI 감사관입니다.
첫 번째 사진은 [개선 전], 두 번째 사진은 [개선 후] 현장 사진입니다.
아래 확정된 위험요인 및 필수 감소대책 목록을 기준으로, 두 번째 사진에서 실제로 안전조치가 완료되었는지 엄격하게 대조 판정하세요.

[검증 대상 목록]
{json.dumps(hazard_summary, ensure_ascii=False)}

[출력 JSON 스키마 - 반드시 유효한 JSON만 출력]
{{
  "results": [
    {{
      "id": 1,
      "resolved": true,
      "reason": "개선 후 사진에서 확인된 구체적 조치 근거 또는 미조치 사유"
    }}
  ],
  "overall_comment": "전체 안전조치 이행 상태에 대한 종합 의견"
}}
"""
    response = client.models.generate_content(
        model="gemini-3.5-flash-lite",
        contents=[prompt, before_img, after_img],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.1,
        ),
    )
    data = json.loads(response.text)
    res_map = {int(r["id"]): r for r in data.get("results", [])}

    verified_items: list[dict[str, Any]] = []
    high_total = 0
    high_resolved = 0
    resolved_count = 0

    for h in active_hazards:
        r = res_map.get(int(h["id"]), {"resolved": False, "reason": "판독 불가"})
        is_res = bool(r.get("resolved", False))
        if is_res:
            resolved_count += 1
        if h["grade"] == "상":
            high_total += 1
            if is_res:
                high_resolved += 1
        verified_items.append({
            "id": h["id"],
            "hazard_name": h["hazard_name"],
            "grade": h["grade"],
            "countermeasure": h["countermeasure"],
            "resolved": is_res,
            "reason": r.get("reason", ""),
        })

    total = len(active_hazards)
    resolution_rate = round((resolved_count / total) * 100, 1) if total > 0 else 100.0
    high_passed = (high_resolved == high_total)
    rate_passed = (resolution_rate >= 80.0)
    passed = high_passed and rate_passed

    return {
        "passed": passed,
        "resolution_rate": resolution_rate,
        "high_total": high_total,
        "high_resolved": high_resolved,
        "items": verified_items,
        "overall_comment": data.get("overall_comment", ""),
    }


def render_hazard_cards(hazards: list[dict[str, Any]]) -> None:
    """현장 작업자가 즉시 보고 조치할 수 있도록 유해·위험요인 및 필수 감소대책을 카드 UI로 표시."""
    st.markdown("#### 🚨 AI 판독 결과 및 필수 감소대책 (현장 조치 가이드)")
    active = [h for h in hazards if not h.get("excluded", False)]
    if not active:
        st.info("유효한 위험요인이 없습니다 (모두 오탐 제외 처리됨).")
        return

    for h in active:
        grade = h.get("grade", "중")
        badge = "🔴 상" if grade == "상" else ("🟠 중" if grade == "중" else "🔵 하")
        border_color = "#DC2626" if grade == "상" else ("#EA580C" if grade == "중" else "#2563EB")
        st.markdown(
            f"""
            <div style="padding: 12px 15px; border-radius: 8px; border-left: 5px solid {border_color};
                        background-color: #F8FAFC; margin-bottom: 10px; color: #0F172A;">
                <div style="font-weight: bold; font-size: 15px; margin-bottom: 4px;">
                    #{h['id']} [{badge} / {h['final_score']}점] {h['hazard_name']} ({h.get('accident_type', '')})
                </div>
                <div style="font-size: 14px; color: #334155;">
                    👉 <b>필수 감소대책:</b> {h['countermeasure']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def capture_image_widget(key_prefix: str, label: str) -> Image.Image | None:
    """모바일 카메라 메모리 초기화 방지를 위해 [직접 촬영]과 [앨범 선택] 탭을 제공합니다."""
    tab_cam, tab_file = st.tabs([f"📷 {label} (브라우저 직접 촬영)", f"🖼️ {label} (앨범/갤러리 선택)"])
    img_data = None
    with tab_cam:
        cam_file = st.camera_input("카메라로 현장 촬영", key=f"{key_prefix}_cam")
        if cam_file is not None:
            img_data = cam_file
    with tab_file:
        up_file = st.file_uploader("미리 찍어둔 현장 사진 업로드 (권장)", type=["jpg", "jpeg", "png"], key=f"{key_prefix}_file")
        if up_file is not None:
            img_data = up_file

    if img_data is not None:
        return Image.open(img_data).convert("RGB")
    return None


def build_legal_ptw_docx(
    ptw_id: str,
    branch_name: str,
    zone_code: str,
    work_type: str,
    gps_coords: str,
    address: str,
    hazards: list[dict[str, Any]],
    tbm_list: list[str],
    verification: dict[str, Any],
    signers: list[str],
    integrity_hash: str,
    before_marked_img: Image.Image | None = None,
    after_img: Image.Image | None = None,
) -> bytes:
    """산업안전보건법 고시 제2024-76호 기반 위험성평가표 및 작업승인서(PTW) 워드 문서 생성."""
    doc = Document()
    doc.add_heading("AI 기반 수시 위험성평가 및 작업승인서 (PTW)", level=1)
    doc.add_paragraph(f"승인번호: {ptw_id} | 발급일시: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 표 1: 현장 기본정보
    t1 = doc.add_table(rows=4, cols=2)
    t1.style = "Table Grid"
    rows_info = [
        ("소속 지점 / 현장(전주번호)", f"{branch_name} / {zone_code}"),
        ("공종 (작업명)", work_type),
        ("현장 위치 (GPS)", f"{address} ({gps_coords})"),
        ("검증 결과 / 참여 근로자", f"승인 통과 (해소율 {verification.get('resolution_rate', 0)}%) / {', '.join(signers)}"),
    ]
    for idx, (k, v) in enumerate(rows_info):
        t1.cell(idx, 0).text = k
        t1.cell(idx, 1).text = str(v)

    # 표 2: 유해·위험요인 및 개선 검증 결과
    doc.add_heading("1. 유해·위험요인 파악 및 개선 전·후 AI 검증 내역", level=2)
    active = [h for h in hazards if not h.get("excluded", False)]
    ver_map = {v["id"]: v for v in verification.get("items", [])}
    t2 = doc.add_table(rows=1 + len(active), cols=5)
    t2.style = "Table Grid"
    headers = ["번호", "유해·위험요인", "등급/점수", "필수 감소대책", "개선 후 AI 검증"]
    for c, text in enumerate(headers):
        t2.cell(0, c).text = text

    for r_idx, h in enumerate(active, start=1):
        v = ver_map.get(h["id"], {})
        status_str = "✅ 해소" if v.get("resolved") else "❌ 미조치"
        t2.cell(r_idx, 0).text = str(h["id"])
        t2.cell(r_idx, 1).text = h["hazard_name"]
        t2.cell(r_idx, 2).text = f"{h['grade']} ({h['final_score']}점)"
        t2.cell(r_idx, 3).text = h["countermeasure"]
        t2.cell(r_idx, 4).text = f"{status_str} - {v.get('reason', '')}"

    # TBM 전달사항
    if tbm_list:
        doc.add_heading("2. 작업 전 안전점검회의 (TBM) 핵심 전달사항", level=2)
        for pt in tbm_list:
            doc.add_paragraph(f"• {pt}")

    # 사진 증빙 첨부
    doc.add_heading("3. 개선 전·후 현장 사진 증빙", level=2)
    if before_marked_img is not None:
        doc.add_paragraph("[개선 전 - 위험요인 번호 마킹]")
        b_buf = io.BytesIO()
        before_marked_img.save(b_buf, format="JPEG", quality=80)
        b_buf.seek(0)
        doc.add_picture(b_buf, width=Inches(4.5))
    if after_img is not None:
        doc.add_paragraph("[개선 후 - 안전조치 완료 검증]")
        a_buf = io.BytesIO()
        after_img.save(a_buf, format="JPEG", quality=80)
        a_buf.seek(0)
        doc.add_picture(a_buf, width=Inches(4.5))

    # 무결성 해시 기록
    doc.add_heading("4. 법적 증빙 및 위변조 방지 해시 (SHA-256)", level=2)
    doc.add_paragraph(f"SHA-256: {integrity_hash}")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ==========================================
# 사이드바: QR 지점 자동인식 및 역할별 인증
# ==========================================
def render_sidebar_auth(branches: list[dict[str, Any]]) -> dict[str, Any]:
    st.sidebar.title("⚡ 한전MCS 안전 플랫폼")
    mode = st.sidebar.radio(
        "접속 모드 선택",
        ["👷 현장 작업자 모드", "🛡️ 지점 안전관리자 모드", "🏢 본사 통합 관제"],
    )

    branch_names = [b["branch_name"] for b in branches]
    qr_branch = st.query_params.get("branch", "")
    default_idx = branch_names.index(qr_branch) if qr_branch in branch_names else 0

    if mode == "🏢 본사 통합 관제":
        hq_pin = st.sidebar.text_input("본사 마스터 PIN 입력", type="password")
        expected_hq = st.secrets.get("HQ_MASTER_PIN", os.environ.get("HQ_MASTER_PIN", "0000"))
        if hq_pin == str(expected_hq):
            st.sidebar.success("본사 관제 인증 완료")
            return {"mode": "HQ", "authenticated": True, "branch_name": None}
        st.sidebar.info("본사 마스터 PIN을 입력하세요.")
        return {"mode": "HQ", "authenticated": False, "branch_name": None}

    if not branch_names:
        st.sidebar.error("Supabase branches 테이블에 등록된 지점이 없습니다.")
        return {"mode": "NONE", "authenticated": False, "branch_name": None}

    if qr_branch in branch_names:
        st.sidebar.success(f"QR 지점 자동 인식: **{qr_branch}**")
    selected_branch = st.sidebar.selectbox("소속 지점 선택", branch_names, index=default_idx)
    branch_info = next(b for b in branches if b["branch_name"] == selected_branch)

    if mode == "👷 현장 작업자 모드":
        pin = st.sidebar.text_input(f"[{selected_branch}] 근로자 공용 PIN", type="password")
        if pin == str(branch_info["worker_pin"]):
            st.sidebar.success(f"{selected_branch} 근로자 인증 완료")
            return {"mode": "WORKER", "authenticated": True, "branch_name": selected_branch, "branch_info": branch_info}
        st.sidebar.caption("※ 지점 공용 PIN을 입력하면 점검을 시작할 수 있습니다.")
        return {"mode": "WORKER", "authenticated": False, "branch_name": selected_branch}

    # 지점 안전관리자 모드
    admin_pin = st.sidebar.text_input(f"[{selected_branch}] 안전관리자 전용 PIN", type="password")
    if admin_pin == str(branch_info["admin_pin"]):
        st.sidebar.success(f"{selected_branch} 안전관리자 인증 완료")
        return {"mode": "ADMIN", "authenticated": True, "branch_name": selected_branch, "branch_info": branch_info}
    st.sidebar.caption("※ 해당 지점 안전관리자 전용 PIN이 필요합니다.")
    return {"mode": "ADMIN", "authenticated": False, "branch_name": selected_branch}


# ==========================================
# 화면 1: 현장 작업자 모드
# ==========================================
def render_worker_view(branch_name: str) -> None:
    st.title(f"👷 현장 작업자 모드 — [{branch_name}]")
    tab_new, tab_action = st.tabs(["📸 1~2단계: 신규 점검 등록", "✅ 3~4단계: 조치 후 AI 검증 & PTW 발급"])

    # --- 탭 1: 신규 점검 등록 ---
    with tab_new:
        st.subheader("📍 1단계: 현장 위치 수집 및 작업 정보")
        gps_coords = st.session_state.get("gps_coords", "")
        address = st.session_state.get("gps_address", "")

        col_gps1, col_gps2 = st.columns([1, 3])
        with col_gps1:
            if streamlit_geolocation is not None:
                loc = streamlit_geolocation()
                if loc and loc.get("latitude") and loc.get("longitude"):
                    lat, lon = loc["latitude"], loc["longitude"]
                    gps_coords = f"{lat:.5f}, {lon:.5f}"
                    st.session_state["gps_coords"] = gps_coords
                    if Nominatim is not None and not st.session_state.get("gps_address"):
                        try:
                            geolocator = Nominatim(user_agent="kepcomcs_safety_ptw")
                            place = geolocator.reverse(f"{lat}, {lon}", language="ko", timeout=5)
                            if place:
                                address = place.address
                                st.session_state["gps_address"] = address
                        except Exception:
                            pass
        with col_gps2:
            address = st.text_input("현장 주소 (GPS 자동 수집 또는 직접 입력)", value=address, placeholder="예: 전남 영광군 영광읍 녹사리")
            if gps_coords:
                st.caption(f"수집된 GPS 좌표: {gps_coords}")

        c1, c2, c3 = st.columns(3)
        worker_name = c1.text_input("점검 등록 작업자 성명", placeholder="예: 홍길동")
        zone_code = c2.text_input("지점명 / 현장명", placeholder=f"예: {branch_name}-나주 25분구")
        work_type = c3.selectbox("공종(작업) 선택", WORK_TYPES)

        weight, basis = db.get_zone_weight_multiplier(branch_name, zone_code, work_type)
        st.info(f"📊 **AI 과거 피드백 보정 가중치: ×{weight}** ({basis})")

        st.markdown("---")
        st.subheader("📷 2단계: 개선 전 현장 사진 AI 1차 분석 및 등록")
        before_img = capture_image_widget("worker_before", "개선 전 현장 사진")

        if before_img is not None:
            st.image(before_img, caption="업로드된 개선 전 사진", width=450)
            if st.button("🤖 AI 1차 위험성평가 실행 및 중앙 DB 등록", type="primary", use_container_width=True):
                if not zone_code.strip():
                    st.warning("지점명 / 현장명을 먼저 입력해 주세요.")
                else:
                    with st.spinner("Gemini AI가 유해·위험요인과 감소대책을 분석 중입니다..."):
                        hazards, tbm_list = analyze_initial_hazards(
                            img=before_img,
                            branch_name=branch_name,
                            zone_code=zone_code,
                            work_type=work_type,
                            weight_multiplier=weight,
                        )
                        marked_img = draw_hazard_markers(before_img, hazards)
                        record, alert_msg = db.create_inspection_record(
                            branch_name=branch_name,
                            worker_name=worker_name,
                            zone_code=zone_code,
                            work_type=work_type,
                            gps_coords=gps_coords,
                            address=address,
                            before_img=before_img,
                            before_marked_img=marked_img,
                            hazards=hazards,
                            tbm_list=tbm_list,
                        )
                        st.session_state["last_created"] = {
                            "record": record,
                            "marked_img": marked_img,
                            "hazards": hazards,
                            "tbm_list": tbm_list,
                            "alert_msg": alert_msg,
                        }

        if "last_created" in st.session_state:
            lc = st.session_state["last_created"]
            st.success(f"✅ 점검 등록 완료! {lc['alert_msg']}")
            st.image(lc["marked_img"], caption="1차 위험요인 번호 마킹 결과", use_container_width=True)
            render_hazard_cards(lc["hazards"])
            if lc["tbm_list"]:
                st.markdown("#### 🗣️ 작업 전 TBM 핵심 전달사항")
                for pt in lc["tbm_list"]:
                    st.markdown(f"- {pt}")

    # --- 탭 2: 조치 완료 및 PTW 발급 ---
    with tab_action:
        col_title, col_ref = st.columns([4, 1])
        col_title.subheader(f"📋 [{branch_name}] 진행 중인 점검 목록")
        if col_ref.button("🔄 상태 새로고침", use_container_width=True):
            st.rerun()

        active_list = db.fetch_inspections_by_branch(
            branch_name=branch_name,
            statuses=["PENDING_REVIEW", "READY_FOR_ACTION"],
        )
        if not active_list:
            st.info("현재 진행 중인 점검 건이 없습니다. [1~2단계] 탭에서 신규 점검을 등록해 주세요.")
            return

        options = {
            f"{STATUS_LABELS[r['status']]} | {r['zone_code']} ({r['work_type']}) - {r['created_at'][:16].replace('T', ' ')}": r
            for r in active_list
        }
        chosen_label = st.selectbox("작업할 점검 건을 선택하세요", list(options.keys()))
        rec = options[chosen_label]

        if rec["status"] == "PENDING_REVIEW":
            st.warning(
                f"⏳ **[{branch_name}] 안전관리자의 원격 검토 및 점수 보정을 기다리고 있습니다.**\n\n"
                "관리자가 확정 버튼을 누르면 우측 상단 [상태 새로고침]을 눌러 3단계(개선 후 AI 검증)를 진행할 수 있습니다."
            )
            if rec.get("before_marked_url"):
                st.image(rec["before_marked_url"], caption="관리자 검토 대기 중인 마킹 사진", width=480)
            render_hazard_cards(rec.get("hazards_json") or [])
            return

        # READY_FOR_ACTION 상태 -> 3~4단계 진행 가능
        st.success("✅ 안전관리자 검토(또는 중·하 자동 통과) 완료! 아래 감소대책을 이행한 뒤 개선 후 사진을 올려주세요.")
        if rec.get("before_marked_url"):
            st.image(rec["before_marked_url"], caption="최종 확정된 1차 마킹 사진", width=480)
        hazards = rec.get("hazards_json") or []
        active_hazards = [h for h in hazards if not h.get("excluded", False)]
        render_hazard_cards(hazards)

        st.markdown("---")
        st.subheader("🔍 3단계: 안전조치 완료 후 사진 AI 대조 검증")
        after_img = capture_image_widget(f"after_{rec['id']}", "개선 후 현장 사진")

        ver_key = f"verification_{rec['id']}"
        after_key = f"after_img_{rec['id']}"

        if after_img is not None:
            st.image(after_img, caption="업로드된 개선 후 사진", width=450)
            if st.button("🔎 개선 전·후 AI 대조 검증 실행", type="primary", use_container_width=True):
                with st.spinner("AI가 개선 전·후 사진을 대조하여 안전조치 이행 여부를 검증 중입니다..."):
                    before_orig = db.download_image_from_url(rec.get("before_img_url", ""))
                    if before_orig is None:
                        st.error("개선 전 원본 사진을 불러오는 데 실패했습니다.")
                    else:
                        verification = verify_before_after_images(before_orig, after_img, active_hazards)
                        st.session_state[ver_key] = verification
                        st.session_state[after_key] = after_img

        if ver_key in st.session_state:
            ver = st.session_state[ver_key]
            saved_after = st.session_state[after_key]
            st.markdown("#### 📊 AI 대조 검증 판정 결과")
            m1, m2, m3 = st.columns(3)
            m1.metric("전체 해소율 (기준 80%)", f"{ver['resolution_rate']}%")
            m2.metric("고위험(상) 해소 (기준 100%)", f"{ver['high_resolved']} / {ver['high_total']}")
            m3.metric("PTW 승인 게이트", "🟢 통과" if ver["passed"] else "🔴 보완 필요")

            for item in ver["items"]:
                icon = "✅" if item["resolved"] else "❌"
                st.write(f"{icon} **#{item['id']} [{item['grade']}] {item['hazard_name']}** — {item['reason']}")
            if ver.get("overall_comment"):
                st.info(f"💬 AI 종합 소견: {ver['overall_comment']}")

            if not ver["passed"]:
                st.error("🚫 고위험(상) 100% 해소 및 전체 80% 이상 해소 조건을 충족하지 못했습니다. 미조치 항목을 보완한 뒤 다시 촬영해 주세요.")
                return

            # --- 4단계: 전자서명 및 PTW 최종 발급 ---
            st.markdown("---")
            st.subheader("✍️ 4단계: 참여 근로자 전자서명 및 작업승인(PTW) 발급")
            signers_raw = st.text_input(
                "참여 근로자 성명 (쉼표로 구분)",
                value=rec.get("worker_name", ""),
                placeholder="예: 홍길동, 김철수, 이영희",
            )
            sig_confirmed = False
            if st_canvas is not None:
                st.caption("대표 작업자 자필 서명")
                canvas_result = st_canvas(
                    fill_color="rgba(255, 255, 255, 0)",
                    stroke_width=3,
                    stroke_color="#000000",
                    background_color="#FFFFFF",
                    height=140,
                    width=380,
                    drawing_mode="freedraw",
                    key=f"sig_{rec['id']}",
                )
                if canvas_result.image_data is not None and canvas_result.image_data[:, :, 3].sum() > 0:
                    sig_confirmed = True
            sig_check = st.checkbox("위 유해·위험요인 및 TBM 내용을 전원 숙지하고 안전조치를 완료했음을 확인합니다.")

            if st.button("🛡️ 최종 작업승인(PTW) 발급 및 법정 양식 생성", type="primary", use_container_width=True):
                signers = [s.strip() for s in signers_raw.split(",") if s.strip()]
                if not signers:
                    st.warning("참여 근로자 성명을 1명 이상 입력해 주세요.")
                elif not sig_check:
                    st.warning("안전조치 완료 확인 체크박스에 동의해 주세요.")
                else:
                    ptw_id = f"PTW-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}"
                    hash_payload = json.dumps(
                        {
                            "ptw_id": ptw_id,
                            "branch": branch_name,
                            "zone": rec["zone_code"],
                            "work_type": rec["work_type"],
                            "gps": rec.get("gps_coords", ""),
                            "signers": signers,
                            "verification": ver,
                        },
                        ensure_ascii=False,
                        sort_keys=True,
                    )
                    integrity_hash = hashlib.sha256(hash_payload.encode("utf-8")).hexdigest()

                    with st.spinner("최종 PTW 기록을 중앙 DB에 보존하고 워드 보고서를 생성 중입니다..."):
                        db.complete_ptw_approval(
                            ptw_id=ptw_id,
                            inspection_id=rec["id"],
                            branch_name=branch_name,
                            zone_code=rec["zone_code"],
                            work_type=rec["work_type"],
                            gps_coords=rec.get("gps_coords", ""),
                            address=rec.get("address", ""),
                            after_img=saved_after,
                            verification=ver,
                            signers=signers,
                            integrity_hash=integrity_hash,
                        )
                        before_marked = db.download_image_from_url(rec.get("before_marked_url", ""))
                        docx_bytes = build_legal_ptw_docx(
                            ptw_id=ptw_id,
                            branch_name=branch_name,
                            zone_code=rec["zone_code"],
                            work_type=rec["work_type"],
                            gps_coords=rec.get("gps_coords", ""),
                            address=rec.get("address", ""),
                            hazards=hazards,
                            tbm_list=rec.get("tbm_json") or [],
                            verification=ver,
                            signers=signers,
                            integrity_hash=integrity_hash,
                            before_marked_img=before_marked,
                            after_img=saved_after,
                        )
                    st.success(f"🎉 작업승인(PTW) 발급 완료! 승인번호: **{ptw_id}**")
                    st.code(f"위변조 방지 해시(SHA-256): {integrity_hash}")
                    st.download_button(
                        label="📥 법정 위험성평가표 & PTW 워드문서(.docx) 다운로드",
                        data=docx_bytes,
                        file_name=f"{ptw_id}_{branch_name}_{rec['zone_code']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )


# ==========================================
# 화면 2: 지점 안전관리자 모드 (1:1 지점 격리)
# ==========================================
def render_admin_view(branch_name: str, branch_info: dict[str, Any]) -> None:
    st.title(f"🛡️ 지점 안전관리자 모드 — [{branch_name}]")
    tab_review, tab_history, tab_settings = st.tabs([
        "🚨 검토 대기함 (오탐·점수 보정)",
        "📁 우리 지점 PTW 승인 이력",
        "⚙️ 담당자 인수인계 설정",
    ])

    # --- 탭 1: 검토 대기함 ---
    with tab_review:
        col_h, col_b = st.columns([4, 1])
        col_h.subheader(f"[{branch_name}] 고위험(상) 검토 대기 건")
        if col_b.button("🔄 대기함 새로고침", use_container_width=True):
            st.rerun()

        pending = db.fetch_inspections_by_branch(branch_name=branch_name, statuses=["PENDING_REVIEW"])
        if not pending:
            st.success(f"현재 [{branch_name}]에 검토 대기 중인 고위험 점검 건이 없습니다.")
        else:
            opt_map = {
                f"🟠 {r['zone_code']} | {r['work_type']} | 작업자: {r.get('worker_name') or '미입력'} ({r['created_at'][:16].replace('T', ' ')})": r
                for r in pending
            }
            sel_key = st.selectbox("검토할 현장 점검 건을 선택하세요", list(opt_map.keys()))
            rec = opt_map[sel_key]

            c_img, c_info = st.columns([1.2, 1])
            with c_img:
                if rec.get("before_marked_url"):
                    st.image(rec["before_marked_url"], caption="현장 마킹 사진 (1차 AI 판독)", use_container_width=True)
            with c_info:
                st.markdown(f"**지점/현장:** {rec['branch_name']} / {rec['zone_code']}")
                st.markdown(f"**공종:** {rec['work_type']}")
                st.markdown(f"**위치:** {rec.get('address', '')} ({rec.get('gps_coords', '')})")
                st.markdown(f"**작업자:** {rec.get('worker_name') or '미입력'}")

            st.markdown("---")
            st.markdown("### 🔧 항목별 오탐 제외 및 위험성 점수 보정 (구역 이력 DB 자동 학습)")
            hazards = rec.get("hazards_json") or []
            updated_hazards: list[dict[str, Any]] = []

            for h in hazards:
                with st.expander(
                    f"#{h['id']} [{h['grade']} / {h['final_score']}점] {h['hazard_name']} (AI 원점수: {h['ai_raw_score']}점)",
                    expanded=True,
                ):
                    st.write(f"**필수 감소대책:** {h['countermeasure']}")
                    col_ex, col_sc = st.columns([1, 2])
                    excluded = col_ex.checkbox(
                        "오탐 제외 (보고서 및 검증에서 삭제)",
                        value=bool(h.get("excluded", False)),
                        key=f"ex_{rec['id']}_{h['id']}",
                    )
                    new_score = col_sc.slider(
                        "확정 위험성 점수 (12점 이상=상, 6~11점=중, 5점 이하=하)",
                        min_value=1,
                        max_value=20,
                        value=int(h["final_score"]),
                        disabled=excluded,
                        key=f"sc_{rec['id']}_{h['id']}",
                    )
                    new_cm = st.text_input(
                        "감소대책 보완 수정 (필요 시)",
                        value=h["countermeasure"],
                        disabled=excluded,
                        key=f"cm_{rec['id']}_{h['id']}",
                    )
                    item_copy = dict(h)
                    item_copy["excluded"] = excluded
                    item_copy["final_score"] = new_score
                    item_copy["grade"] = calc_grade(new_score)
                    item_copy["countermeasure"] = new_cm
                    updated_hazards.append(item_copy)

            if st.button("✅ 보정 결과 확정 및 현장 조치 승인 (DB 학습 반영)", type="primary", use_container_width=True):
                with st.spinner("마커를 재렌더링하고 학습 DB에 반영 중입니다..."):
                    orig_img = db.download_image_from_url(rec.get("before_img_url", ""))
                    if orig_img is None:
                        st.error("원본 이미지를 불러올 수 없습니다.")
                    else:
                        new_marked = draw_hazard_markers(orig_img, updated_hazards)
                        db.admin_confirm_feedback(
                            inspection_id=rec["id"],
                            branch_name=branch_name,
                            zone_code=rec["zone_code"],
                            work_type=rec["work_type"],
                            updated_hazards=updated_hazards,
                            updated_marked_img=new_marked,
                        )
                        st.success("확정 완료! 피드백이 가중치 DB에 누적되었으며, 현장 작업자 화면이 '조치 가능' 상태로 전환되었습니다.")
                        st.rerun()

    # --- 탭 2: 우리 지점 승인 이력 ---
    with tab_history:
        st.subheader(f"📁 [{branch_name}] 최종 승인(PTW) 보존 이력")
        records = db.fetch_ptw_records(branch_name=branch_name)
        if not records:
            st.info("아직 승인 완료된 PTW 이력이 없습니다.")
        else:
            table_rows = [
                {
                    "승인번호": r["ptw_id"],
                    "현장명": r["zone_code"],
                    "공종": r["work_type"],
                    "참여인원": f"{r['signer_count']}명",
                    "승인일시": r["approved_at"][:19].replace("T", " "),
                    "무결성 해시(앞 12자리)": r["integrity_hash"][:12] + "...",
                }
                for r in records
            ]
            st.dataframe(table_rows, use_container_width=True)

    # --- 탭 3: 담당자 인수인계 설정 ---
    with tab_settings:
        st.subheader(f"⚙️ [{branch_name}] 안전관리자 인수인계 및 알림 수신처 설정")
        st.caption("인사이동 시 본사에 요청하지 않고 전임자가 후임자 정보와 PIN을 직접 변경할 수 있습니다.")
        with st.form("handover_form"):
            m_name = st.text_input("안전관리자 성명", value=branch_info.get("manager_name") or "")
            m_email = st.text_input("알림 수신 이메일", value=branch_info.get("manager_email") or "", placeholder="예: safety@kepcomcs.co.kr")
            m_webhook = st.text_input("지점 메신저 웹훅 URL (잔디/팀즈/슬랙)", value=branch_info.get("webhook_url") or "")
            m_pin = st.text_input("새 관리자 전용 PIN (변경 시에만 입력)", type="password")
            submitted = st.form_submit_button("💾 설정 저장", type="primary")
            if submitted:
                db.update_branch_manager_settings(
                    branch_name=branch_name,
                    manager_name=m_name,
                    manager_email=m_email,
                    webhook_url=m_webhook,
                    new_admin_pin=m_pin,
                )
                st.success("지점 안전관리자 정보가 성공적으로 업데이트되었습니다.")


# ==========================================
# 화면 3: 본사 통합 관제 모드 (전국 180개 지점)
# ==========================================
def render_hq_view(branches: list[dict[str, Any]]) -> None:
    st.title("🏢 본사 안전재난부 전국 지점 통합 관제")
    all_insp = db.fetch_inspections_by_branch(branch_name=None, limit=200)
    all_ptw = db.fetch_ptw_records(branch_name=None, limit=200)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("등록 지점 수", f"{len(branches)}개")
    c2.metric("검토 대기(고위험)", sum(1 for r in all_insp if r["status"] == "PENDING_REVIEW"))
    c3.metric("조치 진행 중", sum(1 for r in all_insp if r["status"] == "READY_FOR_ACTION"))
    c4.metric("PTW 최종 승인", len(all_ptw))

    st.markdown("---")
    st.subheader("전국 실시간 점검 현황")
    hq_filter = st.selectbox("지점 필터", ["전체"] + [b["branch_name"] for b in branches])
    filtered = [r for r in all_insp if hq_filter == "전체" or r["branch_name"] == hq_filter]

    if filtered:
        st.dataframe(
            [
                {
                    "지점명": r["branch_name"],
                    "상태": STATUS_LABELS.get(r["status"], r["status"]),
                    "현장명": r["zone_code"],
                    "공종": r["work_type"],
                    "작업자": r.get("worker_name") or "-",
                    "등록시각": r["created_at"][:19].replace("T", " "),
                }
                for r in filtered
            ],
            use_container_width=True,
        )
    else:
        st.info("조회된 전국 점검 기록이 없습니다.")


# ==========================================
# 메인 실행 진입점
# ==========================================
def main() -> None:
    try:
        branches = db.fetch_all_branches()
    except Exception as exc:
        st.error(f"Supabase 연결 오류가 발생했습니다. Secrets 설정과 SQL 테이블 생성을 확인해 주세요.\n\n상세 내용: {exc}")
        return

    auth = render_sidebar_auth(branches)
    if not auth["authenticated"]:
        st.title("⚡ 한전MCS AI 위험성평가 & 작업승인(PTW)")
        st.info("좌측 사이드바에서 접속 모드와 소속 지점을 선택하고 PIN 번호를 입력해 로그인하세요.")
        st.markdown(
            """
            **이용 안내:**
            * **현장 작업자:** 소속 지점 QR코드를 스캔하거나 지점을 선택한 뒤 지점 공용 PIN을 입력하세요.
            * **지점 안전관리자:** 지점 안전관리자 모드를 선택하고 전용 승인 PIN으로 로그인하세요.
            * **PIN 분실 문의:** 각 지점 안전관리자 또는 본사 안전재난부
            """
        )
        return

    if auth["mode"] == "WORKER":
        render_worker_view(auth["branch_name"])
    elif auth["mode"] == "ADMIN":
        render_admin_view(auth["branch_name"], auth["branch_info"])
    elif auth["mode"] == "HQ":
        render_hq_view(branches)


if __name__ == "__main__":
    main()