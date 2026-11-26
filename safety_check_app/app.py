"""한전MCS 전국 지점 통합 AI 위험성평가 & 작업승인(PTW) 플랫폼."""

from __future__ import annotations

import datetime
import hashlib
import io
import json
import os
from typing import Any

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from google import genai
from google.genai import types
from PIL import Image, ImageDraw, ImageFont, ImageOps
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

KST = datetime.timezone(datetime.timedelta(hours=9))

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

ALTERNATIVE_MEASURE_PRESETS = [
    "보호구 착용 강화(턱끈 체결 안전모·절연장갑) 및 안전거리 확보 후 작업",
    "고객 적치물·장애물 우회 동선 확보 및 주의 작업 실시",
    "타인 소유 노후 설비·구조물 접촉 금지 및 원거리 육안·장비 점검",
    "고객 안내 후 안전 위치 확보 및 작업시간 최소화",
    "직접 입력",
]


def now_kst() -> datetime.datetime:
    return datetime.datetime.now(KST)


# ==========================================
# AI / 이미지 / 서명 추출 유틸리티 함수
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


def extract_signature_png_bytes(
    canvas_result: Any, width: int = 380, height: int = 140
) -> bytes | None:
    """st_canvas 결과에서 자필 서명 PNG 바이트를 추출합니다 (image_data 오류 시 json_data로 자동 렌더링)."""
    if canvas_result is None:
        return None

    json_data = getattr(canvas_result, "json_data", None)
    has_json_objects = isinstance(json_data, dict) and len(json_data.get("objects", [])) > 0

    # 1순위: image_data가 정상 반환되는 환경
    try:
        arr = canvas_result.image_data
        if arr is not None:
            ink_mask = (arr[:, :, :3] < 235).any(axis=2) & (arr[:, :, 3] > 0)
            if ink_mask.any():
                ys, xs = ink_mask.nonzero()
                pad = 14
                top = max(0, int(ys.min()) - pad)
                bottom = min(arr.shape[0], int(ys.max()) + pad)
                left = max(0, int(xs.min()) - pad)
                right = min(arr.shape[1], int(xs.max()) + pad)

                img = Image.fromarray(arr.astype("uint8"), mode="RGBA")
                bg = Image.new("RGB", img.size, (245, 246, 248))
                bg.paste(img, mask=img.split()[3])
                if right > left + 5 and bottom > top + 5:
                    bg = bg.crop((left, top, right, bottom))
                buf = io.BytesIO()
                bg.save(buf, format="PNG")
                return buf.getvalue()
    except Exception:
        pass

    # 2순위: Streamlit Cloud 등에서 RuntimeError 발생 시 json_data(Fabric.js path)로 직접 렌더링
    if not has_json_objects:
        return None

    img = Image.new("RGB", (width, height), (245, 246, 248))
    draw = ImageDraw.Draw(img)
    all_pts: list[tuple[float, float]] = []

    for obj in json_data.get("objects", []):
        if obj.get("type") != "path":
            continue
        stroke_w = max(3, int(obj.get("strokeWidth", 3)))
        path_cmds = obj.get("path", [])
        pts: list[tuple[float, float]] = []
        cur_x, cur_y = 0.0, 0.0

        for cmd in path_cmds:
            if not cmd:
                continue
            op = cmd[0]
            if op in ("M", "L") and len(cmd) >= 3:
                cur_x, cur_y = float(cmd[1]), float(cmd[2])
                pts.append((cur_x, cur_y))
            elif op == "Q" and len(cmd) >= 5:
                cx, cy, nx, ny = float(cmd[1]), float(cmd[2]), float(cmd[3]), float(cmd[4])
                for step in range(1, 6):
                    t = step / 5.0
                    px = ((1 - t) ** 2) * cur_x + 2 * (1 - t) * t * cx + (t ** 2) * nx
                    py = ((1 - t) ** 2) * cur_y + 2 * (1 - t) * t * cy + (t ** 2) * ny
                    pts.append((px, py))
                cur_x, cur_y = nx, ny
            elif op == "C" and len(cmd) >= 7:
                c1x, c1y, c2x, c2y, nx, ny = (
                    float(cmd[1]), float(cmd[2]), float(cmd[3]),
                    float(cmd[4]), float(cmd[5]), float(cmd[6]),
                )
                for step in range(1, 7):
                    t = step / 6.0
                    px = ((1 - t) ** 3) * cur_x + 3 * ((1 - t) ** 2) * t * c1x + 3 * (1 - t) * (t ** 2) * c2x + (t ** 3) * nx
                    py = ((1 - t) ** 3) * cur_y + 3 * ((1 - t) ** 2) * t * c1y + 3 * (1 - t) * (t ** 2) * c2y + (t ** 3) * ny
                    pts.append((px, py))
                cur_x, cur_y = nx, ny

        if len(pts) >= 2:
            draw.line(pts, fill=(15, 15, 15), width=stroke_w)
            all_pts.extend(pts)

    if not all_pts:
        return None

    xs = [p[0] for p in all_pts]
    ys = [p[1] for p in all_pts]
    pad = 16
    left = max(0, int(min(xs)) - pad)
    top = max(0, int(min(ys)) - pad)
    right = min(width, int(max(xs)) + pad)
    bottom = min(height, int(max(ys)) + pad)
    if right > left + 8 and bottom > top + 8:
        img = img.crop((left, top, right, bottom))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def analyze_initial_hazards(
    img: Image.Image,
    branch_name: str,
    zone_code: str,
    work_type: str,
    weight_multiplier: float,
) -> tuple[list[dict[str, Any]], list[str]]:
    """1차 현장 사진 AI 위험성평가 분석 + 고객 소유물 불간섭 원칙 및 구역 이력 가중치 반영."""
    client = get_gemini_client()
    prompt = f"""
당신은 대한민국 산업안전보건법(고용노동부 고시 제2024-76호) 기준의 베테랑 전기·현장 안전관리자입니다.
첨부된 작업 전 현장 사진을 정밀 분석하여 유해·위험요인(최대 5개)과 핵심 TBM(작업 전 안전점검회의) 전달사항 3~4가지를 도출하세요.

[현장 정보]
- 지점명: {branch_name}
- 현장명: {zone_code}
- 공종(작업): {work_type}

[현장 제약 조건 - 타인 재산 불간섭 원칙 (필수 준수)]
- 작업자는 고객 사유지·외부 설비를 방문 점검하는 한전MCS 직원입니다.
- 고객 소유의 적치물(장작, 잡자재, 화분, 차량 등)이나 타인 소유의 노후 건축물·처마·담장 등을 작업자가 임의로 치우거나 철거·보수하라는 대책은 제시하지 마세요.
- 타인 재산·구조물과 관련된 위험요인은 반드시 작업자가 현장에서 즉시 통제 가능한 관리적 대책과 보호구 착용 대책(예: 턱끈 체결 안전모 필착, 절연장갑 착용, 안전거리 확보, 적치물 우회 동선 확보, 노출 설비 접촉 금지 및 한전 지사 보고 등) 중심으로 감소대책을 작성하세요.

[출력 JSON 스키마 - 반드시 유효한 JSON만 출력]
{{
  "hazards": [
    {{
      "id": 1,
      "hazard_name": "사진에서 보이는 구체적인 유해·위험요인 (예: 노후 전력량계함 배선 노출부 접촉에 의한 감전 위험)",
      "accident_type": "재해형태 (예: 감전, 추락, 넘어짐(전도), 낙하·비래/부딪힘, 협착 등)",
      "current_measure": "사진상 확인되는 현재 조치 상태 (예: 일반 작업복 및 작업 장갑 착용 상태, 육안 주의에 의존 등)",
      "frequency": 3,
      "severity": 4,
      "countermeasure": "타인 재산을 훼손·이동하지 않고 작업자가 즉시 실행 가능한 구체적 감소대책",
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
            "current_measure": h.get("current_measure", "현장 육안 확인 필요"),
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


def build_non_photo_verification(
    active_hazards: list[dict[str, Any]],
    mode: str,
    alt_measures: dict[int, str] | None = None,
) -> dict[str, Any]:
    """하 등급 전용 패스트트랙 또는 타인 재산 제거 불가 대체조치 완료 시 검증 결과 객체를 생성합니다."""
    alt_measures = alt_measures or {}
    items: list[dict[str, Any]] = []
    for h in active_hazards:
        h_id = int(h["id"])
        if mode == "LOW_ONLY_FAST_TRACK":
            reason = "하 등급(허용 가능 수준)으로 개선 후 사진 검증 생략 — 기본 안전수칙 준수 및 TBM 확인"
            is_alt = False
            alt_text = ""
        else:
            alt_text = alt_measures.get(h_id, ALTERNATIVE_MEASURE_PRESETS[0])
            reason = f"타인 재산·시설물로 물리적 제거 불가 → 대체조치 적용: {alt_text}"
            is_alt = True

        items.append({
            "id": h_id,
            "hazard_name": h["hazard_name"],
            "grade": h["grade"],
            "countermeasure": h["countermeasure"],
            "resolved": True,
            "is_alternative": is_alt,
            "alternative_measure": alt_text,
            "reason": reason,
        })

    if mode == "LOW_ONLY_FAST_TRACK":
        mode_label = "하 등급 전용 패스트트랙 (개선 후 사진 생략)"
        overall_comment = "허용 가능 수준(하 등급)만 존재하여 기본 안전수칙 준수 및 TBM 확인으로 즉시 작업승인 가능"
    else:
        mode_label = "타인 재산·시설물 제거 불가에 따른 대체 안전조치 확정"
        overall_comment = "고객 소유물·외부 구조물로 물리적 제거가 불가하여 작업자 보호구·우회 접근 등 대체 안전조치로 승인 통과"

    return {
        "passed": True,
        "mode": mode,
        "mode_label": mode_label,
        "resolution_rate": 100.0,
        "high_total": 0,
        "high_resolved": 0,
        "items": items,
        "overall_comment": overall_comment,
    }


def verify_before_after_images(
    before_img: Image.Image,
    after_img: Image.Image,
    active_hazards: list[dict[str, Any]],
    alt_measures: dict[int, str] | None = None,
) -> dict[str, Any]:
    """2차 개선 전·후 사진 AI 대조 검증 (타인 재산 불간섭 원칙 + 중·하 대체조치 통합 판정)."""
    alt_measures = alt_measures or {}
    photo_targets = [
        h for h in active_hazards
        if not (h["grade"] in ("중", "하") and int(h["id"]) in alt_measures)
    ]

    res_map: dict[int, dict[str, Any]] = {}
    overall_comment = ""

    if photo_targets:
        client = get_gemini_client()
        hazard_summary = [
            {
                "id": h["id"],
                "hazard_name": h["hazard_name"],
                "grade": h["grade"],
                "countermeasure": h["countermeasure"],
            }
            for h in photo_targets
        ]
        prompt = f"""
당신은 작업승인(PTW) 최종 검증 AI 감사관입니다.
첫 번째 사진은 [개선 전], 두 번째 사진은 [개선 후] 현장 사진입니다.
아래 확정된 위험요인 및 필수 감소대책 목록을 기준으로, 두 번째 사진에서 실제로 안전조치가 완료되었는지 대조 판정하세요.

[현장 검증 원칙 - 타인 재산 불간섭 기준 (필수 준수)]
- 작업자는 고객 사유지·외부 현장을 방문 중이므로, 타인 소유의 적치물·노후 건축물·처마·외부 설비가 배경에 그대로 남아 있더라도 물리적 철거·이동을 요구하지 마세요.
- 개선 후 사진에서 작업자가 보호구(턱끈 체결 안전모, 절연장갑 등)를 착용했거나, 위험 구역에서 안전거리를 확보·우회 접근 위치로 이동하는 등 감소대책에 명시된 작업자 통제 가능 조치가 확인되면 해소(resolved: true)로 판정하세요.

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
        overall_comment = data.get("overall_comment", "")

    verified_items: list[dict[str, Any]] = []
    high_total = 0
    high_resolved = 0
    resolved_count = 0
    has_alt = False

    for h in active_hazards:
        h_id = int(h["id"])
        if h["grade"] in ("중", "하") and h_id in alt_measures:
            has_alt = True
            alt_text = alt_measures[h_id]
            resolved_count += 1
            verified_items.append({
                "id": h_id,
                "hazard_name": h["hazard_name"],
                "grade": h["grade"],
                "countermeasure": h["countermeasure"],
                "resolved": True,
                "is_alternative": True,
                "alternative_measure": alt_text,
                "reason": f"타인 재산·시설물로 물리적 제거 불가 → 대체조치 적용: {alt_text}",
            })
        else:
            r = res_map.get(h_id, {"resolved": False, "reason": "판독 불가"})
            is_res = bool(r.get("resolved", False))
            if is_res:
                resolved_count += 1
            if h["grade"] == "상":
                high_total += 1
                if is_res:
                    high_resolved += 1
            verified_items.append({
                "id": h_id,
                "hazard_name": h["hazard_name"],
                "grade": h["grade"],
                "countermeasure": h["countermeasure"],
                "resolved": is_res,
                "is_alternative": False,
                "alternative_measure": "",
                "reason": r.get("reason", ""),
            })

    total = len(active_hazards)
    resolution_rate = round((resolved_count / total) * 100, 1) if total > 0 else 100.0
    high_passed = (high_resolved == high_total)
    rate_passed = (resolution_rate >= 80.0)
    passed = high_passed and rate_passed

    mode_label = (
        "2차 AI 개선 전·후 대조 + 타인 재산 대체조치 병행"
        if has_alt
        else "2차 AI 개선 전·후 사진 대조 검증"
    )

    return {
        "passed": passed,
        "mode": "PHOTO_AND_ALT" if has_alt else "PHOTO_ONLY",
        "mode_label": mode_label,
        "resolution_rate": resolution_rate,
        "high_total": high_total,
        "high_resolved": high_resolved,
        "items": verified_items,
        "overall_comment": overall_comment or "개선 후 사진 및 대체 안전조치 기준 검증 완료",
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
        cur_m = h.get("current_measure", "")
        cur_html = (
            f'<div style="font-size: 13px; color: #475569; margin-bottom: 3px;">🔎 <b>현재 조치:</b> {cur_m}</div>'
            if cur_m
            else ""
        )
        st.markdown(
            f"""
            <div style="padding: 12px 15px; border-radius: 8px; border-left: 5px solid {border_color};
                        background-color: #F8FAFC; margin-bottom: 10px; color: #0F172A;">
                <div style="font-weight: bold; font-size: 15px; margin-bottom: 4px;">
                    #{h['id']} [{badge} / {h['final_score']}점] {h['hazard_name']} ({h.get('accident_type', '')})
                </div>
                {cur_html}
                <div style="font-size: 14px; color: #334155;">
                    👉 <b>필수 감소대책:</b> {h['countermeasure']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def capture_image_widget(key_prefix: str, label: str) -> Image.Image | None:
    """모바일 카메라 촬영 및 앨범/갤러리 선택을 지원하며 EXIF 방향을 자동 보정합니다."""
    up_file = st.file_uploader(
        f"🖼️ {label} (사진 촬영 또는 앨범 선택)",
        type=["jpg", "jpeg", "png"],
        key=f"{key_prefix}_file",
    )
    if up_file is not None:
        img = Image.open(up_file)
        img = ImageOps.exif_transpose(img)
        return img.convert("RGB")
    return None


def _get_fallback_after_img(rec: dict[str, Any], before_marked: Image.Image | None) -> Image.Image:
    """db_supabase.complete_ptw_approval 호환을 위해 개선 후 사진 생략 시 대체 이미지 반환."""
    if before_marked is not None:
        return before_marked
    orig = db.download_image_from_url(rec.get("before_img_url", ""))
    if orig is not None:
        return orig
    return Image.new("RGB", (400, 120), (245, 246, 248))


# ==========================================
# 법정 '현장 위험성 평가표' DOCX 생성 유틸리티
# ==========================================
def _set_font(run, size: float = 9.0, bold: bool = False, color: tuple[int, int, int] | None = None) -> None:
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def _set_cell_shading(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def _set_cell_margins(cell, top: int = 55, bottom: int = 55, left: int = 85, right: int = 85) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def _set_table_borders(table, color: str = "B8C2CC", sz: str = "4") -> None:
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tbl_pr.append(borders)


def _write_cell(
    cell,
    text: str,
    size: float = 8.4,
    bold: bool = False,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT,
    bg: str | None = None,
    color: tuple[int, int, int] | None = None,
) -> None:
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_cell_margins(cell)
    if bg:
        _set_cell_shading(cell, bg)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    _set_font(run, size=size, bold=bold, color=color)


def _append_sig_image(paragraph, sig_png_bytes: bytes | None, width_cm: float = 1.45) -> None:
    if not sig_png_bytes:
        return
    run = paragraph.add_run()
    run.add_picture(io.BytesIO(sig_png_bytes), width=Cm(width_cm))


def _set_col_widths(table, widths_cm: list[float]) -> None:
    table.autofit = False
    for row in table.rows:
        for idx, w in enumerate(widths_cm):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(w)


def _add_doc_heading(doc: Document, title: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    _set_font(p.add_run(title), size=11, bold=True, color=(15, 44, 89))


def _add_doc_note(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(4)
    _set_font(p.add_run(text), size=7.8, color=(80, 88, 98))


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
    sig_png_bytes: bytes | None = None,
    eval_dt: datetime.datetime | None = None,
) -> bytes:
    """고용노동부고시 제2024-76호 기준 '현장 위험성 평가표' + PTW 이행 검증 워드 문서 생성."""
    doc = Document()
    for sec in doc.sections:
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(1.4)
        sec.bottom_margin = Cm(1.4)
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)

    dt = eval_dt or now_kst()
    dt_full = dt.strftime("%Y. %m. %d. %H:%M")
    dt_date = dt.strftime("%Y. %m. %d.")

    active = [h for h in (hazards or []) if not h.get("excluded", False)]
    ver_items = verification.get("items", [])
    ver_map = {int(v["id"]): v for v in ver_items}

    cnt_all = len(active)
    cnt_high = sum(1 for h in active if h.get("grade") == "상")
    cnt_mid = sum(1 for h in active if h.get("grade") == "중")
    cnt_low = sum(1 for h in active if h.get("grade") == "하")

    signer_display = ", ".join(signers) if signers else ""

    target_lines: list[str] = []
    if gps_coords and "," in gps_coords:
        lat_s, lon_s = [p.strip() for p in gps_coords.split(",", 1)]
        target_lines.append(f"위도: {lat_s}, 경도: {lon_s}")
    elif gps_coords:
        target_lines.append(gps_coords)
    if address:
        target_lines.append(address)
    if zone_code:
        target_lines.append(f"{zone_code} ({work_type})")
    target_text = "\n".join(target_lines) if target_lines else work_type

    # ---- 문서 제목 ----
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    _set_font(p_title.add_run("현장 위험성 평가표"), size=17, bold=True, color=(15, 44, 89))

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(8)
    _set_font(
        p_sub.add_run("사업장 위험성평가에 관한 지침(고용노동부고시 제2024-76호)의 절차를 따름"),
        size=8.5,
        color=(90, 98, 110),
    )

    # ---- 상단 기본정보표 (4x4) ----
    meta_tbl = doc.add_table(rows=4, cols=4)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(meta_tbl)

    branch_display = f"[{branch_name}]" if branch_name else "[          지점]"
    photo_method = "현장점검·사진 2장(전·후)" if after_img is not None else "현장점검·사진 1장"

    meta_rows = [
        ("사업장", branch_display, "평가 일시", dt_full),
        ("평가 대상", target_text, "평가 종류", "□ 최초 □ 수시\n□ 정기 □ 상시"),
        ("평가 방법", photo_method, "평가 결과", f"{cnt_all}건, 상 {cnt_high}·중 {cnt_mid}·하 {cnt_low}"),
        ("평가 책임자", "          (서명)", "참여 근로자", ""),
    ]
    for r_idx, (k1, v1, k2, v2) in enumerate(meta_rows):
        _write_cell(meta_tbl.cell(r_idx, 0), k1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")
        _write_cell(meta_tbl.cell(r_idx, 1), v1)
        _write_cell(meta_tbl.cell(r_idx, 2), k2, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")
        if r_idx < 3:
            _write_cell(meta_tbl.cell(r_idx, 3), v2)
        else:
            c_worker = meta_tbl.cell(3, 3)
            _write_cell(c_worker, f"{signer_display}\n{dt_date}" if signer_display else dt_date, size=8.4)
            if sig_png_bytes:
                p_sig = c_worker.add_paragraph()
                p_sig.paragraph_format.space_before = Pt(2)
                p_sig.paragraph_format.space_after = Pt(1)
                _append_sig_image(p_sig, sig_png_bytes, width_cm=1.5)
    _set_col_widths(meta_tbl, [2.8, 6.2, 2.8, 6.2])

    # ---- 1. 사업장 위험성 평가 절차 ----
    _add_doc_heading(doc, "1. 사업장 위험성 평가 절차")
    proc_tbl = doc.add_table(rows=7, cols=3)
    proc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(proc_tbl)

    for c_idx, head in enumerate(["절차", "AI 판독", "평가 책임자 확인"]):
        _write_cell(proc_tbl.cell(0, c_idx), head, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")

    proc_rows = [
        ("가. 사전준비(제9조)", "위험성 수준 기준 기본안", "위험성의 수준과 그 수준을 판단하는 기준, 허용 가능한 위험성의 수준을 확정(제2항)"),
        ("나. 유해·위험요인 파악(제10조)", f"현장점검(제10조 제1호)간 첨부된 사진에서 {cnt_all}건", "사진에 포함되지 않은 현장 위험물(화학물질·소음·중량물 등)과 근로자가 제안한 위험요인도 추가하여 파악"),
        ("다. 위험성 결정(제11조)", "상·중·하 수준으로 분류", "허용 가능한지 결정, 현장 근로자 참여(제6조)"),
        ("라. 위험성 감소대책 수립·실행(제12조)", "감소대책 수립(제거·대체 → 공학적 대책 → 관리적 대책 → 보호구 착용)", "감소대책에 대한 담당자 및 완료일 작성, 실행 후 위험성 수준 재확인"),
        ("마. 공유(제13조)", "작업 전 위험요인과 감소대책을 공유", "작업 전 안전점검회의(TBM)에서 공유하고 서명"),
        ("바. 기록·보존(제14조)", "본 문서", "서명한 본 문서를 보관"),
    ]
    for idx, (col1, col2, col3) in enumerate(proc_rows, start=1):
        _write_cell(proc_tbl.cell(idx, 0), col1, size=8.1, bold=True)
        _write_cell(proc_tbl.cell(idx, 1), col2, size=8.1)
        _write_cell(proc_tbl.cell(idx, 2), col3, size=8.1)
    _set_col_widths(proc_tbl, [4.3, 5.7, 8.0])

    _add_doc_note(
        doc,
        "※ 평가 책임자는 제10조의 사업장 순회점검을 통해 유해·위험요인을 발굴, 제11조의 위험성 결정 및 "
        "제12조의 위험성 감소대책을 수립·실행 할 것.(사업장 위험성 평가에 관한 지침 제15조 제4항)",
    )

    # ---- 2. 법적 의무 ----
    _add_doc_heading(doc, "2. 법적 의무")
    law_tbl = doc.add_table(rows=4, cols=4)
    law_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(law_tbl)

    for c_idx, head in enumerate(["의무", "조문", "과태료", "비고"]):
        _write_cell(law_tbl.cell(0, c_idx), head, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")

    law_rows = [
        ("위험성평가 실시", "제36조 제1항", "1천만원 이하", "위험성평가표 전체"),
        ("근로자의 참여\n근로자대표 참여\n결과 알림", "제36조 제2항ㆍ제3항ㆍ제4항 전단", "500만원 이하", "참여 근로자 서명\n작업 전 공유"),
        ("결과 기록·보존", "제36조 제5항", "300만원 이하", "서명된 본 평가표 보관"),
    ]
    for idx, row_vals in enumerate(law_rows, start=1):
        for c_idx, val in enumerate(row_vals):
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (1, 2) else WD_ALIGN_PARAGRAPH.LEFT
            _write_cell(law_tbl.cell(idx, c_idx), val, size=8.1, align=align)
    _set_col_widths(law_tbl, [4.2, 4.8, 3.0, 6.0])

    _add_doc_note(
        doc,
        "※ 과태료 시행 : 상시 50명 이상(건설업은 공사 금액 50억원 이상) 2027년 1월 1일 시행(산업안전보건법 부칙 제1조). "
        "중대재해로 이어질 수 있는 위험은 작업 전 안전점검회의(TBM) 등을 통해 상시 알리도록 노력해야 합니다(제36조 제4항 후단).",
    )

    # ---- 3. 위험성 수준 기준(AI 기준) ----
    _add_doc_heading(doc, "3. 위험성 수준 기준(AI 기준)")
    std_tbl = doc.add_table(rows=6, cols=4)
    std_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(std_tbl)

    for c_idx, head in enumerate(["수준", "기준 제시", "판단", "조치"]):
        _write_cell(std_tbl.cell(0, c_idx), head, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")

    std_rows = [
        ("상", "사망 또는 영구 장애를 초래할 수 있는 고위험 요인이 존재하거나, 다수의 위험에 동시 노출되어 즉시 작업 중단이 필요한 상태", "허용 불가", "즉시 작업을 중지하고 위험 요인 제거 및 현장 안전조치(보호구 완비, 통로 확보 등) 후 재작업", "FEE4E2", (180, 35, 24)),
        ("중", "휴업을 수반하는 부상(골절, 감전, 열상 등)이 발생할 가능성이 높은 유해·위험요인이 존재하는 상태", "허용 불가", "작업 전 안전보호구 착용 및 환경 개선(장애물 정리, 안전거리 유지 등) 후 작업 실시", "FEF0C7", (181, 71, 8)),
        ("하", "경미한 부상(찰과상, 단순 타박상 등) 발생 가능성이 있으나 통상적인 주의로 예방 가능한 상태", "허용 가능", "기본 안전수칙 준수 및 상시 주의력 유지", "D1FADF", (2, 122, 72)),
    ]
    for idx, (lvl, desc, judge, act, bg_hex, txt_col) in enumerate(std_rows, start=1):
        _write_cell(std_tbl.cell(idx, 0), lvl, size=8.4, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg=bg_hex, color=txt_col)
        _write_cell(std_tbl.cell(idx, 1), desc, size=8.1)
        _write_cell(std_tbl.cell(idx, 2), judge, size=8.1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _write_cell(std_tbl.cell(idx, 3), act, size=8.1)

    _write_cell(std_tbl.cell(4, 0), "판단 근거", size=8.1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")
    merged_basis = std_tbl.cell(4, 1).merge(std_tbl.cell(4, 3))
    _write_cell(
        merged_basis,
        "산업안전보건법 제36조(위험성평가의 실시), 산업안전보건기준에 관한 규칙 제32조(보호구의 지급 등), 제301조(전기 기계·기구 등의 유지관리)",
        size=8.1,
    )

    merged_confirm = std_tbl.cell(5, 0).merge(std_tbl.cell(5, 3))
    _write_cell(
        merged_confirm,
        f"위 판독 기준으로 정함.     참여 근로자: {signer_display}   ",
        size=8.4,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
        bg="F8FAFC",
    )
    _append_sig_image(merged_confirm.paragraphs[0], sig_png_bytes, width_cm=1.4)
    _set_col_widths(std_tbl, [1.8, 7.2, 2.2, 6.8])

    # ---- 4. AI 판독 ----
    _add_doc_heading(doc, "4. AI 판독")
    _add_doc_note(doc, "※ 사진에 표시된 원과 번호는 아래 '위험성 평가표'의 번호와 일치합니다.")
    if before_marked_img is not None:
        b_buf = io.BytesIO()
        before_marked_img.convert("RGB").save(b_buf, format="JPEG", quality=85)
        b_buf.seek(0)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(b_buf, width=Cm(15.5))

    # ---- 4-1. 위험성 평가표(AI 분석) ----
    _add_doc_heading(doc, "4-1. 위험성 평가표(AI 분석)")
    hz_tbl = doc.add_table(rows=1 + len(active), cols=7)
    hz_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(hz_tbl)

    hz_headers = [
        "번호",
        "유해·위험요인\n(사진에서 보이는 것)",
        "재해형태",
        "현재 조치",
        "위험\n수준",
        "감소 대책(제12조)",
        "담당\n완료일",
    ]
    for c_idx, head in enumerate(hz_headers):
        _write_cell(hz_tbl.cell(0, c_idx), head, size=8.1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")

    for r_idx, h in enumerate(active, start=1):
        grade = str(h.get("grade", "중"))
        grade_bg = "FEE4E2" if grade == "상" else ("FEF0C7" if grade == "중" else "D1FADF")
        grade_col = (180, 35, 24) if grade == "상" else ((181, 71, 8) if grade == "중" else (2, 122, 72))

        measure_text = str(h.get("countermeasure", ""))
        v_item = ver_map.get(int(h.get("id", r_idx)), {})
        if v_item.get("is_alternative"):
            alt_m = v_item.get("alternative_measure", "")
            measure_text += f"\n※ [대체조치] 고객 소유물·외부 시설물로 직접 제거 불가하여 대체 적용 ({alt_m})"

        _write_cell(hz_tbl.cell(r_idx, 0), str(h.get("id", r_idx)), size=8.1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _write_cell(hz_tbl.cell(r_idx, 1), str(h.get("hazard_name", "")), size=7.9)
        _write_cell(hz_tbl.cell(r_idx, 2), str(h.get("accident_type", "-")), size=7.9, align=WD_ALIGN_PARAGRAPH.CENTER)
        _write_cell(hz_tbl.cell(r_idx, 3), str(h.get("current_measure") or "현장 확인 필요"), size=7.9)
        _write_cell(hz_tbl.cell(r_idx, 4), grade, size=8.1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg=grade_bg, color=grade_col)
        _write_cell(hz_tbl.cell(r_idx, 5), measure_text, size=7.9)

        c_done = hz_tbl.cell(r_idx, 6)
        _write_cell(c_done, f"{signer_display}\n{dt_date}" if signer_display else dt_date, size=7.7, align=WD_ALIGN_PARAGRAPH.CENTER)
        if sig_png_bytes:
            p_row_sig = c_done.add_paragraph()
            p_row_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_row_sig.paragraph_format.space_before = Pt(2)
            _append_sig_image(p_row_sig, sig_png_bytes, width_cm=1.35)
    _set_col_widths(hz_tbl, [1.0, 4.3, 1.9, 2.5, 1.2, 4.9, 2.2])

    # ---- 5. 안전점검회의(TBM) 공유 ----
    _add_doc_heading(doc, "5. 안전점검회의(TBM) 공유")
    for idx, pt in enumerate(tbm_list or [], start=1):
        clean_pt = str(pt).strip()
        if not clean_pt:
            continue
        p_tbm = doc.add_paragraph()
        p_tbm.paragraph_format.space_before = Pt(1)
        p_tbm.paragraph_format.space_after = Pt(2)
        prefix = "" if clean_pt[:2].rstrip(".").isdigit() else f"{idx}. "
        _set_font(p_tbm.add_run(f"{prefix}{clean_pt}"), size=8.4)

    tbm_tbl = doc.add_table(rows=2, cols=3)
    tbm_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(tbm_tbl)
    for c_idx, head in enumerate(["공유 일시", "공유한 사람 (평가 책임자)", "참석자 서명 (근로자)"]):
        _write_cell(tbm_tbl.cell(0, c_idx), head, size=8.2, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")

    _write_cell(tbm_tbl.cell(1, 0), dt_full, size=8.2, align=WD_ALIGN_PARAGRAPH.CENTER)
    _write_cell(tbm_tbl.cell(1, 1), "          (서명)", size=8.2, align=WD_ALIGN_PARAGRAPH.CENTER)
    c_tbm_sig = tbm_tbl.cell(1, 2)
    _write_cell(c_tbm_sig, f"{signer_display}  ", size=8.2, align=WD_ALIGN_PARAGRAPH.CENTER)
    _append_sig_image(c_tbm_sig.paragraphs[0], sig_png_bytes, width_cm=1.5)
    _set_col_widths(tbm_tbl, [5.0, 6.5, 6.5])

    _add_doc_note(
        doc,
        "※ 이 문서는 초안입니다. 유해·위험요인은 본 AI 모델이 사진에서 분석한 결과이고, 위험성 수준과 감소대책 후보는 "
        "코드가 정해진 규칙으로 고른 것입니다. 허용 가능 여부의 결정, 대책 선택과 실행, 근로자 참여는 관리자가 시행합니다."
        "(사업장 위험성 평가에 관한 지침 제6조·제11조·제12조)",
    )
    _add_doc_note(
        doc,
        "※ 사진에 보이지 않는 위험(화학물질 소음 분진·중량물 취급 작업 자세 등)은 이 표에 들어 있지 않습니다. 파악 단계에서 추가 작성해 주세요.",
    )

    # ---- 6. 안전조치 이행 검증 및 작업승인(PTW) 기록 ----
    _add_doc_heading(doc, "6. 안전조치 이행 검증 및 작업승인(PTW) 기록")
    ptw_tbl = doc.add_table(rows=3, cols=2)
    ptw_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(ptw_tbl)
    mode_label = verification.get("mode_label", "2차 AI 개선 전·후 사진 대조 검증")
    ptw_info_rows = [
        ("PTW 승인번호 / 일시", f"{ptw_id} ({dt_full})"),
        (
            "이행 검증 방식 / 결과",
            f"{mode_label} — 승인 통과 (전체 해소율 {verification.get('resolution_rate', 0)}%, "
            f"고위험 {verification.get('high_resolved', 0)}/{verification.get('high_total', 0)} 해소)",
        ),
        ("위·변조 방지 해시(SHA-256)", integrity_hash),
    ]
    for r_idx, (k, v) in enumerate(ptw_info_rows):
        _write_cell(ptw_tbl.cell(r_idx, 0), k, size=8.1, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="EEF2F7")
        _write_cell(ptw_tbl.cell(r_idx, 1), str(v), size=8.0)
    _set_col_widths(ptw_tbl, [4.6, 13.4])

    if ver_items:
        p_ver = doc.add_paragraph()
        p_ver.paragraph_format.space_before = Pt(5)
        p_ver.paragraph_format.space_after = Pt(2)
        _set_font(p_ver.add_run("[항목별 안전조치 이행 검증 내역]"), size=8.4, bold=True)
        for item in ver_items:
            mark = "대체조치" if item.get("is_alternative") else ("해소" if item.get("resolved") else "미조치")
            p_item = doc.add_paragraph()
            p_item.paragraph_format.space_before = Pt(1)
            p_item.paragraph_format.space_after = Pt(1)
            _set_font(
                p_item.add_run(f"• #{item.get('id')} [{item.get('grade')}] {mark} - {item.get('reason', '')}"),
                size=8.0,
            )

    if after_img is not None:
        p_cap = doc.add_paragraph()
        p_cap.paragraph_format.space_before = Pt(6)
        p_cap.paragraph_format.space_after = Pt(3)
        _set_font(p_cap.add_run("[개선 후 현장 사진 - 안전조치 완료 증빙]"), size=8.4, bold=True)
        a_buf = io.BytesIO()
        after_img.convert("RGB").save(a_buf, format="JPEG", quality=85)
        a_buf.seek(0)
        p_after = doc.add_paragraph()
        p_after.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_after.add_run().add_picture(a_buf, width=Cm(14.5))

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
        if "last_ptw_download" in st.session_state:
            dl = st.session_state["last_ptw_download"]
            st.success(f"🎉 작업승인(PTW) 발급 완료! 승인번호: **{dl['ptw_id']}**")
            st.code(f"위변조 방지 해시(SHA-256): {dl['integrity_hash']}")
            st.download_button(
                label="📥 현장 위험성 평가표 & PTW 워드문서(.docx) 다운로드",
                data=dl["docx_bytes"],
                file_name=dl["file_name"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
                key=f"dl_btn_{dl['ptw_id']}",
            )
            st.markdown("---")

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
        st.success("✅ 안전관리자 검토(또는 중·하 자동 통과) 완료! 아래 내용을 확인하고 3~4단계를 진행하세요.")
        if rec.get("before_marked_url"):
            st.image(rec["before_marked_url"], caption="최종 확정된 1차 마킹 사진", width=480)
        hazards = rec.get("hazards_json") or []
        active_hazards = [h for h in hazards if not h.get("excluded", False)]
        render_hazard_cards(hazards)

        st.markdown("---")
        st.subheader("🔍 3단계: 안전조치 이행 확인 (개선 후 검증 / 대체조치)")

        ver_key = f"verification_{rec['id']}"
        after_key = f"after_img_{rec['id']}"

        has_high_or_mid = any(h.get("grade") in ("상", "중") for h in active_hazards)

        # (기능 3) 하 등급 전용 패스트트랙: 개선 후 사진 생략
        if not has_high_or_mid:
            st.success(
                "🟢 **하 등급 전용(허용 가능 수준) 패스트트랙 적용**\n\n"
                "고위험(상) 및 중위험(중) 항목이 없으므로 개선 후 사진 검증을 생략하고, "
                "아래 4단계에서 TBM 안전수칙 확인 및 자필 서명만으로 즉시 작업승인(PTW)을 발급할 수 있습니다."
            )
            st.session_state[ver_key] = build_non_photo_verification(
                active_hazards=active_hazards,
                mode="LOW_ONLY_FAST_TRACK",
            )
            st.session_state[after_key] = None
        else:
            # (기능 2) 중·하 등급 항목별 '타인 재산·시설물 제거 불가 → 대체조치' 선택
            alt_measures: dict[int, str] = {}
            mid_low_hazards = [h for h in active_hazards if h.get("grade") in ("중", "하")]

            if mid_low_hazards:
                with st.expander(
                    "🏠 타인 재산·시설물로 물리적 제거가 불가한 항목이 있나요? (중·하 등급 대체 안전조치 지정)",
                    expanded=False,
                ):
                    st.caption(
                        "고객 사유지 적치물·노후 구조물 등 작업자가 임의로 치우거나 수리할 수 없는 항목을 체크하고 "
                        "대체 안전조치를 선택하면, 해당 항목은 사진상 제거 여부 대신 대체조치 이행으로 기록·통과됩니다. "
                        "(※ 고위험 '상' 등급은 반드시 안전조치 후 사진 검증이 필요합니다.)"
                    )
                    for h in mid_low_hazards:
                        h_id = int(h["id"])
                        use_alt = st.checkbox(
                            f"#{h_id} [{h['grade']}] {h['hazard_name']} — 타인 재산·시설물로 제거 불가 (대체조치 적용)",
                            key=f"alt_chk_{rec['id']}_{h_id}",
                        )
                        if use_alt:
                            preset = st.selectbox(
                                f"#{h_id} 적용할 대체 안전조치 선택",
                                ALTERNATIVE_MEASURE_PRESETS,
                                key=f"alt_sel_{rec['id']}_{h_id}",
                            )
                            if preset == "직접 입력":
                                custom_txt = st.text_input(
                                    f"#{h_id} 대체 안전조치 내용 직접 입력",
                                    placeholder="예: 안전모 착용 및 적치물 우회 접근하여 검침 완료",
                                    key=f"alt_txt_{rec['id']}_{h_id}",
                                )
                                alt_measures[h_id] = custom_txt.strip() or ALTERNATIVE_MEASURE_PRESETS[0]
                            else:
                                alt_measures[h_id] = preset

            remaining_photo_targets = [
                h for h in active_hazards
                if not (h["grade"] in ("중", "하") and int(h["id"]) in alt_measures)
            ]

            if not remaining_photo_targets:
                st.info(
                    "🛡️ **모든 중·하 항목이 '타인 재산·시설물 제거 불가 → 대체 안전조치'로 지정되었습니다.**\n\n"
                    "개선 후 사진 없이 아래 버튼으로 대체 안전조치 이행을 확정하고 4단계(서명·PTW 발급)로 넘어갈 수 있습니다."
                )
                if st.button("✅ 대체 안전조치 이행 확정 (4단계로 이동)", type="primary", use_container_width=True):
                    st.session_state[ver_key] = build_non_photo_verification(
                        active_hazards=active_hazards,
                        mode="ALTERNATIVE_ONLY",
                        alt_measures=alt_measures,
                    )
                    st.session_state[after_key] = None
            else:
                if alt_measures:
                    st.info(
                        f"대체조치 지정 {len(alt_measures)}건은 자동 해소 처리되며, "
                        f"나머지 {len(remaining_photo_targets)}건(고위험 포함)에 대해 개선 후 사진을 검증합니다."
                    )
                after_img = capture_image_widget(f"after_{rec['id']}", "개선 후 현장 사진")

                if after_img is not None:
                    st.image(after_img, caption="업로드된 개선 후 사진", width=450)
                    if st.button("🔎 개선 전·후 AI 대조 검증 실행", type="primary", use_container_width=True):
                        with st.spinner("AI가 개선 전·후 사진 및 대체 안전조치를 검증 중입니다..."):
                            before_orig = db.download_image_from_url(rec.get("before_img_url", ""))
                            if before_orig is None:
                                st.error("개선 전 원본 사진을 불러오는 데 실패했습니다.")
                            else:
                                verification = verify_before_after_images(
                                    before_orig,
                                    after_img,
                                    active_hazards,
                                    alt_measures=alt_measures,
                                )
                                st.session_state[ver_key] = verification
                                st.session_state[after_key] = after_img

        if ver_key in st.session_state:
            ver = st.session_state[ver_key]
            saved_after = st.session_state.get(after_key)
            st.markdown(f"#### 📊 안전조치 검증 판정 결과 ({ver.get('mode_label', '')})")
            m1, m2, m3 = st.columns(3)
            m1.metric("전체 해소율 (기준 80%)", f"{ver['resolution_rate']}%")
            m2.metric("고위험(상) 해소 (기준 100%)", f"{ver['high_resolved']} / {ver['high_total']}")
            m3.metric("PTW 승인 게이트", "🟢 통과" if ver["passed"] else "🔴 보완 필요")

            for item in ver["items"]:
                icon = "🛡️" if item.get("is_alternative") else ("✅" if item["resolved"] else "❌")
                st.write(f"{icon} **#{item['id']} [{item['grade']}] {item['hazard_name']}** — {item['reason']}")
            if ver.get("overall_comment"):
                st.info(f"💬 종합 소견: {ver['overall_comment']}")

            if not ver["passed"]:
                st.error(
                    "🚫 고위험(상) 100% 해소 및 전체 80% 이상 해소 조건을 충족하지 못했습니다. "
                    "미조치 항목을 보완해 다시 촬영하거나, 타인 소유물로 제거 불가한 중·하 항목은 위 대체 안전조치를 지정해 주세요."
                )
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
            canvas_result = None
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
                if canvas_result is not None:
                    json_data = getattr(canvas_result, "json_data", None)
                    if isinstance(json_data, dict) and len(json_data.get("objects", [])) > 0:
                        sig_confirmed = True
                    else:
                        try:
                            if canvas_result.image_data is not None and (canvas_result.image_data[:, :, :3] < 235).any():
                                sig_confirmed = True
                        except Exception:
                            pass

            sig_check = st.checkbox("위 유해·위험요인, 대체 안전조치 및 TBM 내용을 전원 숙지하고 안전수칙을 준수할 것을 확인합니다.")

            if st.button("🛡️ 최종 작업승인(PTW) 발급 및 법정 양식 생성", type="primary", use_container_width=True):
                signers = [s.strip() for s in signers_raw.split(",") if s.strip()]
                if not signers:
                    st.warning("참여 근로자 성명을 1명 이상 입력해 주세요.")
                elif not sig_check:
                    st.warning("안전조치 완료 확인 체크박스에 동의해 주세요.")
                else:
                    eval_now = now_kst()
                    ptw_id = f"PTW-{eval_now.strftime('%Y%m%d-%H%M%S')}"
                    sig_png_bytes = extract_signature_png_bytes(canvas_result, width=380, height=140)
                    hash_payload = json.dumps(
                        {
                            "ptw_id": ptw_id,
                            "branch": branch_name,
                            "zone": rec["zone_code"],
                            "work_type": rec["work_type"],
                            "gps": rec.get("gps_coords", ""),
                            "signers": signers,
                            "sig_confirmed": sig_confirmed or (sig_png_bytes is not None),
                            "verification": ver,
                        },
                        ensure_ascii=False,
                        sort_keys=True,
                    )
                    integrity_hash = hashlib.sha256(hash_payload.encode("utf-8")).hexdigest()

                    with st.spinner("최종 PTW 기록을 중앙 DB에 보존하고 현장 위험성 평가표를 생성 중입니다..."):
                        before_marked = db.download_image_from_url(rec.get("before_marked_url", ""))
                        db_after_img = (
                            saved_after
                            if saved_after is not None
                            else _get_fallback_after_img(rec, before_marked)
                        )
                        db.complete_ptw_approval(
                            ptw_id=ptw_id,
                            inspection_id=rec["id"],
                            branch_name=branch_name,
                            zone_code=rec["zone_code"],
                            work_type=rec["work_type"],
                            gps_coords=rec.get("gps_coords", ""),
                            address=rec.get("address", ""),
                            after_img=db_after_img,
                            verification=ver,
                            signers=signers,
                            integrity_hash=integrity_hash,
                        )
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
                            sig_png_bytes=sig_png_bytes,
                            eval_dt=eval_now,
                        )
                        st.session_state["last_ptw_download"] = {
                            "ptw_id": ptw_id,
                            "integrity_hash": integrity_hash,
                            "docx_bytes": docx_bytes,
                            "file_name": f"현장_위험성_평가표_{eval_now.strftime('%Y%m%d_%H%M%S')}.docx",
                        }
                    st.rerun()


# ==========================================
# 화면 2: 지점 안전관리자 모드 (1:1 지점 격리)
# ==========================================
def render_admin_view(branch_name: str, branch_info: dict[str, Any]) -> None:
    st.title(f"🛡️ 지점 안전관리자 모드 — [{branch_name}]")
    tab_review, tab_cleanup, tab_history, tab_settings = st.tabs([
        "🚨 검토 대기함 (오탐·점수 보정)",
        "🗑️ 진행 중 목록 정리(삭제)",
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
                    if h.get("current_measure"):
                        st.write(f"**현재 조치:** {h['current_measure']}")
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

    # --- 탭 2: 진행 중 목록 정리(삭제) ---
    with tab_cleanup:
        col_ch, col_cb = st.columns([4, 1])
        col_ch.subheader(f"🗑️ [{branch_name}] 진행 중인 점검 건 정리 (중복·재평가 삭제)")
        if col_cb.button("🔄 목록 새로고침", key="refresh_cleanup", use_container_width=True):
            st.rerun()

        st.caption(
            "작업자가 재평가를 진행해 중복으로 남은 '검토 대기' 및 '조치 가능' 건을 삭제할 수 있습니다. "
            "(※ 최종 승인 완료된 PTW 기록은 삭제 대상에서 제외됩니다.)"
        )

        in_progress = db.fetch_inspections_by_branch(
            branch_name=branch_name,
            statuses=["PENDING_REVIEW", "READY_FOR_ACTION"],
            limit=100,
        )
        if not in_progress:
            st.success(f"현재 [{branch_name}]에 정리할 진행 중 점검 건이 없습니다.")
        else:
            st.dataframe(
                [
                    {
                        "상태": STATUS_LABELS.get(r["status"], r["status"]),
                        "현장명": r["zone_code"],
                        "공종": r["work_type"],
                        "작업자": r.get("worker_name") or "미입력",
                        "등록시각": r["created_at"][:16].replace("T", " "),
                        "ID(앞 8자리)": r["id"][:8],
                    }
                    for r in in_progress
                ],
                use_container_width=True,
            )

            del_map = {
                f"{STATUS_LABELS.get(r['status'], r['status'])} | {r['zone_code']} ({r['work_type']}) | "
                f"{r.get('worker_name') or '미입력'} | {r['created_at'][:16].replace('T', ' ')} [{r['id'][:8]}]": r
                for r in in_progress
            }
            selected_labels = st.multiselect(
                "삭제할 중복·취소 점검 건을 선택하세요 (복수 선택 가능)",
                options=list(del_map.keys()),
                key="admin_delete_targets",
            )

            if selected_labels:
                if len(selected_labels) == 1:
                    preview_rec = del_map[selected_labels[0]]
                    if preview_rec.get("before_marked_url"):
                        st.image(
                            preview_rec["before_marked_url"],
                            caption=f"삭제 대상 미리보기: {preview_rec['zone_code']}",
                            width=360,
                        )

                confirm_del = st.checkbox(
                    f"선택한 {len(selected_labels)}건의 진행 중 점검 기록을 영구 삭제합니다. (복구 불가)",
                    key="confirm_admin_delete",
                )
                if st.button(
                    f"🗑️ 선택한 {len(selected_labels)}건 영구 삭제",
                    type="primary",
                    disabled=not confirm_del,
                    use_container_width=True,
                ):
                    deleted_cnt = 0
                    errors: list[str] = []
                    with st.spinner("선택한 점검 기록을 삭제 중입니다..."):
                        for label in selected_labels:
                            target = del_map[label]
                            ok, msg = db.delete_inspection_record(target["id"], branch_name)
                            if ok:
                                deleted_cnt += 1
                            else:
                                errors.append(f"[{target['zone_code']}] {msg}")

                    if deleted_cnt:
                        st.success(f"{deleted_cnt}건의 진행 중 점검 기록을 삭제했습니다.")
                    for err in errors:
                        st.error(err)
                    if deleted_cnt:
                        st.rerun()

    # --- 탭 3: 우리 지점 승인 이력 ---
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

    # --- 탭 4: 담당자 인수인계 설정 ---
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