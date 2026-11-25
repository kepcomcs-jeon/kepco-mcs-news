"""전국 지점 1:1 매칭 Supabase(중앙 DB + Storage) 및 웹훅/이메일 알림 모듈."""

from __future__ import annotations

import datetime
from email.mime.text import MIMEText
import io
import os
import smtplib
from typing import Any
import uuid

from PIL import Image
import requests
import streamlit as st
from supabase import Client, create_client

BUCKET_NAME = "safety-images"


def _get_secret(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.environ.get(key, default)


@st.cache_resource
def get_supabase_client() -> Client:
    url = _get_secret("SUPABASE_URL")
    key = _get_secret("SUPABASE_KEY")
    if not url or not key:
        raise ValueError("Streamlit Secrets에 SUPABASE_URL과 SUPABASE_KEY를 설정해 주세요.")
    return create_client(url, key)


def upload_image_to_storage(img: Image.Image, prefix: str = "before") -> str:
    """PIL 이미지를 JPEG로 압축해 Supabase Storage에 업로드하고 공개 URL을 반환합니다."""
    sb = get_supabase_client()
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=82, optimize=True)
    file_bytes = buf.getvalue()

    date_prefix = datetime.datetime.now().strftime("%Y%m%d")
    file_path = f"{date_prefix}/{prefix}_{uuid.uuid4().hex[:10]}.jpg"

    sb.storage.from_(BUCKET_NAME).upload(
        path=file_path,
        file=file_bytes,
        file_options={"content-type": "image/jpeg"},
    )
    return sb.storage.from_(BUCKET_NAME).get_public_url(file_path)


def download_image_from_url(url: str) -> Image.Image | None:
    """Storage 공개 URL에서 PIL 이미지를 불러옵니다."""
    if not url:
        return None
    try:
        res = requests.get(url, timeout=10)
        res.raise_for_status()
        return Image.open(io.BytesIO(res.content)).convert("RGB")
    except Exception:
        return None


def fetch_all_branches() -> list[dict[str, Any]]:
    """등록된 전국 지점 목록을 본부·지점명 순으로 조회합니다."""
    sb = get_supabase_client()
    res = sb.table("branches").select("*").order("hq_name").order("branch_name").execute()
    return res.data or []


def get_branch_info(branch_name: str) -> dict[str, Any] | None:
    """지점명으로 해당 지점의 PIN 및 알림 수신처 정보를 조회합니다."""
    sb = get_supabase_client()
    res = sb.table("branches").select("*").eq("branch_name", branch_name).limit(1).execute()
    return res.data[0] if res.data else None


def update_branch_manager_settings(
    branch_name: str,
    manager_name: str,
    manager_email: str,
    webhook_url: str,
    new_admin_pin: str,
) -> None:
    """[지점 인수인계] 해당 지점의 안전관리자 연락처와 관리자 PIN을 갱신합니다."""
    sb = get_supabase_client()
    payload: dict[str, Any] = {
        "manager_name": manager_name.strip(),
        "manager_email": manager_email.strip(),
        "webhook_url": webhook_url.strip(),
    }
    if new_admin_pin.strip():
        payload["admin_pin"] = new_admin_pin.strip()
    sb.table("branches").update(payload).eq("branch_name", branch_name).execute()


def send_branch_admin_alert(
    branch_name: str,
    zone_code: str,
    work_type: str,
    worker_name: str,
    high_risk_count: int,
    inspection_id: str,
) -> tuple[bool, str]:
    """해당 지점 안전관리자 1명에게만 웹훅 또는 이메일로 실시간 검토 알림을 발송합니다."""
    branch = get_branch_info(branch_name)
    if not branch:
        return False, "지점 정보 없음"

    subject = f"[한전MCS 위험성평가] {branch_name}({zone_code}) 고위험 {high_risk_count}건 검토 요청"
    body = (
        f"🚨 [{branch_name}] 고위험 작업 검토 요청\n"
        f"- 현장/전주번호: {zone_code}\n"
        f"- 공종: {work_type}\n"
        f"- 등록 작업자: {worker_name or '미입력'}\n"
        f"- 고위험(상) 탐지: {high_risk_count}건\n"
        f"- 점검 ID: {inspection_id[:8]}\n\n"
        f"👉 앱의 [지점 안전관리자 모드]에 접속해 오탐 제외 및 위험성 점수 보정을 진행해 주세요."
    )

    sent_channels: list[str] = []

    # 1) 지점 전용 메신저 웹훅 발송 (잔디/슬랙/팀즈 호환)
    webhook_url = (branch.get("webhook_url") or "").strip()
    if webhook_url:
        try:
            res = requests.post(webhook_url, json={"body": body, "text": body}, timeout=5)
            if res.status_code in (200, 201, 204):
                sent_channels.append("메신저 웹훅")
        except Exception:
            pass

    # 2) 지점 안전관리자 이메일 발송 (Secrets에 SMTP 설정 시)
    manager_email = (branch.get("manager_email") or "").strip()
    smtp_user = _get_secret("SMTP_USER")
    smtp_password = _get_secret("SMTP_PASSWORD")
    if manager_email and smtp_user and smtp_password:
        try:
            msg = MIMEText(body, _charset="utf-8")
            msg["Subject"] = subject
            msg["From"] = smtp_user
            msg["To"] = manager_email
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=8) as server:
                server.login(smtp_user, smtp_password)
                server.sendmail(smtp_user, [manager_email], msg.as_string())
            sent_channels.append(f"이메일({manager_email})")
        except Exception:
            pass

    if sent_channels:
        return True, ", ".join(sent_channels)
    return False, "수신처 미등록 (관리자 모드 대기함에서 직접 확인 가능)"


def get_zone_weight_multiplier(
    branch_name: str, zone_code: str, work_type: str
) -> tuple[float, str]:
    """2단계 계층형 가중치 산출 (1순위: 지점+구역+공종 -> 2순위: 지점+공종 평균, 0.7~1.3 클램핑)."""
    sb = get_supabase_client()
    clean_zone = zone_code.strip()

    rows: list[dict[str, Any]] = []
    basis = "초기 기본값 (보정 이력 누적 시 자동 갱신)"

    if clean_zone:
        res_exact = (
            sb.table("hazard_feedback")
            .select("ai_score, adjusted_score")
            .eq("branch_name", branch_name)
            .eq("zone_code", clean_zone)
            .eq("work_type", work_type)
            .eq("is_false_positive", False)
            .execute()
        )
        rows = res_exact.data or []
        if rows:
            basis = f"동일 현장·공종 이력 {len(rows)}건 반영"

    if not rows:
        res_branch = (
            sb.table("hazard_feedback")
            .select("ai_score, adjusted_score")
            .eq("branch_name", branch_name)
            .eq("work_type", work_type)
            .eq("is_false_positive", False)
            .execute()
        )
        rows = res_branch.data or []
        if rows:
            basis = f"{branch_name} 동일 공종 평균 이력 {len(rows)}건 반영"

    ratios = [
        float(r["adjusted_score"]) / float(r["ai_score"])
        for r in rows
        if r.get("ai_score", 0) > 0
    ]
    if not ratios:
        return 1.0, basis

    avg_ratio = sum(ratios) / len(ratios)
    clamped = round(max(0.7, min(1.3, avg_ratio)), 2)
    return clamped, basis


def create_inspection_record(
    branch_name: str,
    worker_name: str,
    zone_code: str,
    work_type: str,
    gps_coords: str,
    address: str,
    before_img: Image.Image,
    before_marked_img: Image.Image,
    hazards: list[dict[str, Any]],
    tbm_list: list[str],
) -> tuple[dict[str, Any], str]:
    """[현장 작업자] 1차 AI 분석 결과 등록 및 고위험(상) 포함 시 지점 관리자 1:1 알림."""
    sb = get_supabase_client()
    before_url = upload_image_to_storage(before_img, prefix=f"{branch_name}_before")
    marked_url = upload_image_to_storage(before_marked_img, prefix=f"{branch_name}_marked")

    high_risk_count = sum(
        1 for h in hazards if not h.get("excluded", False) and h.get("grade") == "상"
    )
    initial_status = "PENDING_REVIEW" if high_risk_count > 0 else "READY_FOR_ACTION"

    res = (
        sb.table("inspections")
        .insert({
            "branch_name": branch_name,
            "worker_name": worker_name.strip(),
            "zone_code": zone_code.strip() or f"{branch_name}-미지정현장",
            "work_type": work_type,
            "gps_coords": gps_coords,
            "address": address,
            "status": initial_status,
            "before_img_url": before_url,
            "before_marked_url": marked_url,
            "hazards_json": hazards,
            "tbm_json": tbm_list,
        })
        .execute()
    )
    record = res.data[0]

    alert_msg = "중·하 등급 전용 건으로 즉시 현장 조치 단계로 전환되었습니다."
    if initial_status == "PENDING_REVIEW":
        ok, channel = send_branch_admin_alert(
            branch_name=branch_name,
            zone_code=record["zone_code"],
            work_type=work_type,
            worker_name=worker_name,
            high_risk_count=high_risk_count,
            inspection_id=record["id"],
        )
        alert_msg = (
            f"고위험(상) {high_risk_count}건 탐지 — {branch_name} 안전관리자에게 알림 발송 완료 ({channel})"
            if ok
            else f"고위험(상) {high_risk_count}건 탐지 — {branch_name} 관리자 검토 대기함에 등록됨 ({channel})"
        )

    return record, alert_msg


def fetch_inspections_by_branch(
    branch_name: str | None = None,
    statuses: list[str] | None = None,
    limit: int = 40,
) -> list[dict[str, Any]]:
    """지점별(또는 전국 본사용) 점검 목록을 최신순으로 조회합니다."""
    sb = get_supabase_client()
    query = sb.table("inspections").select("*")
    if branch_name:
        query = query.eq("branch_name", branch_name)
    if statuses:
        query = query.in_("status", statuses)
    res = query.order("created_at", desc=True).limit(limit).execute()
    return res.data or []


def admin_confirm_feedback(
    inspection_id: str,
    branch_name: str,
    zone_code: str,
    work_type: str,
    updated_hazards: list[dict[str, Any]],
    updated_marked_img: Image.Image,
) -> None:
    """[지점 안전관리자] 오탐 제외 및 점수 보정을 확정해 학습 DB에 누적하고 '조치대기'로 전환합니다."""
    sb = get_supabase_client()
    new_marked_url = upload_image_to_storage(
        updated_marked_img, prefix=f"{branch_name}_confirmed"
    )

    feedback_rows = [
        {
            "inspection_id": inspection_id,
            "branch_name": branch_name,
            "zone_code": zone_code.strip() or "DEFAULT",
            "work_type": work_type,
            "hazard_title": h["hazard_name"],
            "ai_score": int(h["ai_raw_score"]),
            "adjusted_score": int(h["final_score"]),
            "is_false_positive": bool(h.get("excluded", False)),
        }
        for h in updated_hazards
    ]
    if feedback_rows:
        sb.table("hazard_feedback").insert(feedback_rows).execute()

    sb.table("inspections").update({
        "hazards_json": updated_hazards,
        "before_marked_url": new_marked_url,
        "status": "READY_FOR_ACTION",
        "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }).eq("id", inspection_id).execute()


def complete_ptw_approval(
    ptw_id: str,
    inspection_id: str,
    branch_name: str,
    zone_code: str,
    work_type: str,
    gps_coords: str,
    address: str,
    after_img: Image.Image,
    verification: dict[str, Any],
    signers: list[str],
    integrity_hash: str,
) -> None:
    """[최종 작업승인] 개선 후 사진, AI 검증 결과, 참여자 및 SHA-256 무결성 해시를 영구 보존합니다."""
    sb = get_supabase_client()
    after_url = upload_image_to_storage(after_img, prefix=f"{branch_name}_after")

    sb.table("inspections").update({
        "after_img_url": after_url,
        "verification_json": verification,
        "status": "APPROVED",
        "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }).eq("id", inspection_id).execute()

    sb.table("ptw_records").insert({
        "ptw_id": ptw_id,
        "inspection_id": inspection_id,
        "branch_name": branch_name,
        "zone_code": zone_code,
        "work_type": work_type,
        "gps_coords": gps_coords,
        "address": address,
        "verification_status": "PASSED",
        "signer_count": len(signers),
        "signers_json": signers,
        "integrity_hash": integrity_hash,
    }).execute()


def fetch_ptw_records(branch_name: str | None = None, limit: int = 50) -> list[dict[str, Any]]:
    """지점별(또는 본사 전국) 최종 승인된 PTW 이력을 조회합니다."""
    sb = get_supabase_client()
    query = sb.table("ptw_records").select("*")
    if branch_name:
        query = query.eq("branch_name", branch_name)
    res = query.order("approved_at", desc=True).limit(limit).execute()
    return res.data or []