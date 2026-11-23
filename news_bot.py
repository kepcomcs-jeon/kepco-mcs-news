import os
import time
import html
import re
import json
from datetime import datetime, timezone, timedelta
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# [1] 사용자 설정 구역
# ==========================================
KEYWORD = "한전MCS"

# 깃허브 시크릿(Environment Variables)에서 토큰과 API 키를 안전하게 불러옵니다.
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "sent_links.txt")
EXCEL_FILE = os.path.join(BASE_DIR, "한전MCS_뉴스모니터링_대장.xlsx")
# ==========================================

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
}

def clean_text(raw_text):
    if not raw_text:
        return ""
    text = re.sub(r"<[^>]+>", "", raw_text)
    text = html.unescape(text)
    text = re.sub(r"\s+", " ", text)
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return text.strip()

def clean_for_excel(text):
    return text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")

def load_sent_links():
    if not os.path.exists(DB_FILE) or os.path.isdir(DB_FILE):
        return set()
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return set(line.strip() for line in f if line.strip())
    except Exception as e:
        print(f"⚠️ 링크 파일 읽기 오류: {e}")
        return set()

def save_sent_link(link):
    try:
        with open(DB_FILE, "a", encoding="utf-8") as f:
            f.write(link + "\n")
    except Exception as e:
        print(f"⚠️ 링크 저장 오류: {e}")

# 1. 네이버 뉴스 크롤링
def fetch_naver_news(keyword):
    url = "https://search.naver.com/search.naver"
    params = {"where": "news", "query": keyword, "sort": "1"}
    news_list = []
    try:
        res = requests.get(url, params=params, headers=HEADERS, timeout=10)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            items = soup.select("div.news_wrap, li.bx, div.sds-comps-vertical-layout")
            for item in items:
                title_el = item.select_one("a.news_tit, a.sds-comps-text-type-headline1, a[data-heatmap-target='.tit']")
                if not title_el:
                    continue
                title = clean_text(title_el.get_text())
                link = title_el.get("href", "")
                if not link or not link.startswith("http"):
                    continue

                desc_el = item.select_one("div.news_dsc, a.api_txt_lines.dsc_txt_wrap, a.sds-comps-text-type-body1, a[data-heatmap-target='.body']")
                desc = clean_text(desc_el.get_text()) if desc_el else "본문 요약 없음"

                press_el = item.select_one("a.info.press, span.sds-comps-profile-info-title-text, a[data-heatmap-target='.prof']")
                press = clean_text(press_el.get_text()).replace("언론사 선정", "").strip() if press_el else "네이버 뉴스"

                pub_date = "최근"
                for sp in item.select("span.info, span.sds-comps-profile-info-subtext"):
                    txt = clean_text(sp.get_text())
                    if any(k in txt for k in ["전", "분", "시간", "일", "202"]):
                        pub_date = txt
                        break

                news_list.append({"portal": "네이버 뉴스", "press": press, "pub_date": pub_date, "title": title, "summary": desc, "link": link})
                if len(news_list) >= 10:
                    break
    except Exception as e:
        print(f"[네이버 크롤링 오류] {e}")
    return news_list

# 2. 다음 뉴스 크롤링
def fetch_daum_news(keyword):
    url = "https://search.daum.net/search"
    params = {"w": "news", "q": keyword, "sort": "recency"}
    daum_headers = {**HEADERS, "Referer": "https://www.daum.net/"}
    news_list = []
    try:
        res = requests.get(url, params=params, headers=daum_headers, timeout=10)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            items = soup.select("div.wrap_cont, ul.c-list-basic > li")[:10]
            for item in items:
                title_el = item.select_one("a.tit_main, div.item-title a")
                if not title_el:
                    continue
                title = clean_text(title_el.get_text())
                link = title_el.get("href", "")
                if not link:
                    continue

                desc_el = item.select_one("a.desc, p.desc, p.conts-desc")
                desc = clean_text(desc_el.get_text()) if desc_el else "요약 내용 없음"

                press_el = item.select_one("a.f_nb, span.f_nb, div.c-tit-doc span.tit_item, span.txt_info")
                press = clean_text(press_el.get_text()) if press_el else "다음 뉴스"

                pub_date = "최근"
                for sp in item.select("span.f_nb, span.gem-subinfo, span.txt_info"):
                    txt = clean_text(sp.get_text())
                    if any(k in txt for k in ["전", "분", "시간", "일", "202", "."]):
                        pub_date = txt
                        break

                news_list.append({
                    "portal": "다음 뉴스",
                    "press": press,
                    "pub_date": pub_date,
                    "title": title,
                    "summary": desc,
                    "link": link
                })
    except Exception as e:
        print(f"[다음 크롤링 오류] {e}")
    return news_list

# 키워드 기반 예비 판별 함수 (API 오류 시 안전장치)
def analyze_sentiment_fallback(title, summary):
    positive_words = ["혁신", "성공", "표창", "수상", "협약", "봉사", "기부", "나눔", "강화", "개선", "최우수", "선도", "확대", "호평"]
    negative_words = ["논란", "사고", "비판", "갈등", "파업", "적자", "우려", "지적", "처벌", "위반", "고발", "중단", "단독", "재해", "부실", "해고", "해이"]
    combined = title + " " + summary
    pos_score = sum(1 for w in positive_words if w in combined)
    neg_score = sum(1 for w in negative_words if w in combined)
    if pos_score > neg_score:
        return "긍정 (우호적 보도)", "키워드 기반 자동 분류"
    elif neg_score > pos_score:
        return "부정 (이슈 점검)", "키워드 기반 자동 분류"
    else:
        return "중립 (스트레이트)", "키워드 기반 자동 분류"

# AI 기반 문맥 논조 분석 (Gemini API)
def analyze_sentiment(title, summary):
    if not GEMINI_API_KEY:
        print("⚠️ GEMINI_API_KEY가 설정되지 않아 키워드 방식으로 분석합니다.")
        return analyze_sentiment_fallback(title, summary)

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
    prompt = (
        f"당신은 공기업 '{KEYWORD}'의 홍보 담당자입니다.\n"
        f"아래 뉴스 기사의 제목과 요약을 읽고, '{KEYWORD}' 입장에서 이 기사가 긍정적인지, 부정적인지, 중립적인지 판단하세요.\n"
        f"- 긍정 (우호적 보도): 사회공헌, 봉사, 수상, 협약, 혁신, 실적 개선, 안전 예방 활동 등 기관 이미지에 도움이 되는 기사\n"
        f"- 부정 (이슈 점검): 사건/사고 발생, 노사 갈등, 비리, 민원, 비판, 감사 지적, 경영 부실 등 기관에 리스크가 되는 기사\n"
        f"- 중립 (스트레이트): 단순 인사, 행사 안내, 일반 업계 동향, 단순 사실 전달 위주의 기사\n\n"
        f"기사 제목: {clean_for_excel(title)}\n"
        f"기사 요약: {clean_for_excel(summary)}\n\n"
        f'반드시 아래 JSON 형식으로만 응답하세요:\n'
        f'{{"sentiment": "긍정 (우호적 보도)" 또는 "부정 (이슈 점검)" 또는 "중립 (스트레이트)", "reason": "{KEYWORD} 입장에서의 판단 근거를 1문장(40자 이내)으로 요약"}}'
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.1
        }
    }

    try:
        res = requests.post(url, json=payload, timeout=15)
        if res.status_code == 200:
            data = res.json()
            text = data["candidates"][0]["content"]["parts"][0]["text"]
            result = json.loads(text)
            sentiment = result.get("sentiment", "중립 (스트레이트)")
            reason = result.get("reason", "AI 문맥 분석 완료")
            if sentiment not in ["긍정 (우호적 보도)", "부정 (이슈 점검)", "중립 (스트레이트)"]:
                if "긍정" in sentiment:
                    sentiment = "긍정 (우호적 보도)"
                elif "부정" in sentiment:
                    sentiment = "부정 (이슈 점검)"
                else:
                    sentiment = "중립 (스트레이트)"
            return sentiment, clean_text(reason)
        else:
            print(f"⚠️ Gemini API 응답 오류 ({res.status_code}): {res.text[:100]}")
    except Exception as e:
        print(f"⚠️ Gemini API 호출 오류: {e}")

    return analyze_sentiment_fallback(title, summary)

def save_to_excel(article, sentiment, reason):
    now_str = time.strftime("%Y-%m-%d %H:%M:%S")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
    )
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")

    try:
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            ws = wb.active
            ws.title = "뉴스모니터링_내역"
            headers = ["수집일시", "포털 구분", "언론사명", "기사 발행일시", "논조 분류", "기사 제목", "핵심 요약", "원문 링크", "AI 판단 근거"]
            ws.append(headers)
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            col_widths = {"A": 20, "B": 13, "C": 18, "D": 16, "E": 18, "F": 45, "G": 55, "H": 40, "I": 35}
            for col_letter, width in col_widths.items():
                ws.column_dimensions[col_letter].width = width
            ws.freeze_panes = "A2"
        else:
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            # 기존 엑셀 파일에 I열 헤더가 없으면 자동 추가
            if not ws.cell(row=1, column=9).value:
                cell = ws.cell(row=1, column=9, value="AI 판단 근거")
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.column_dimensions["I"].width = 35

        row_data = [
            now_str, article["portal"], clean_for_excel(article["press"]),
            clean_for_excel(article["pub_date"]), sentiment,
            clean_for_excel(article["title"]), clean_for_excel(article["summary"]),
            article["link"], clean_for_excel(reason)
        ]
        ws.append(row_data)
        curr_row = ws.max_row

        for col_idx in range(1, 10):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = Font(name="맑은 고딕", size=10)
            cell.border = thin_border
            if col_idx in [1, 2, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 8:
                cell.hyperlink = article["link"]
                cell.font = Font(name="맑은 고딕", size=10, color="0563C1", underline="single")
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        wb.save(EXCEL_FILE)
    except Exception as e:
        print(f"[엑셀 저장 오류] {e}")

def send_telegram_alert(article, sentiment, reason):
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️ 텔레그램 토큰 또는 채팅방 ID가 설정되지 않아 알림을 건너뜁니다.")
        return False

    portal = article["portal"]
    press = article["press"]
    pub_date = article["pub_date"]
    title = article["title"]
    summary = article["summary"]
    link = article["link"]
    
    if "긍정" in sentiment:
        status_header = "🟢 긍정 (우호적 보도)"
    elif "부정" in sentiment:
        status_header = "🔴 부정 (이슈점검)"
    else:
        status_header = "🔵 중립 (스트레이트)"
    
    message = (
        f"🚨 <b>[{status_header}] {KEYWORD} 뉴스 알림</b>\n"
        f"━━━━━━━━━━━━━━━━━━\n"
        f"📰 <b>제목:</b> <a href='{link}'>{title}</a>\n"
        f"🏢 <b>매체 및 시간:</b> {press} ({pub_date}) | {portal}\n"
        f"🤖 <b>AI 판단 근거:</b> {reason}\n\n"
        f"💡 <b>핵심 내용 요약:</b>\n{summary}\n\n"
        f"🔗 <a href='{link}'>👉 기사 원문 바로가기 (클릭)</a>"
    )
    
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": message, "parse_mode": "HTML", "disable_web_page_preview": False}
    try:
        res = requests.post(url, json=payload, timeout=10)
        return res.status_code == 200
    except Exception as e:
        print(f"[텔레그램 전송 오류] {e}")
        return False

# [메인 실행 함수] 무한 루프 없이 한 번 실행되고 종료됨 (GitHub Actions 전용)
def main():
    # 🌙 한국 시간(KST) 기준 방해금지 시간(23시~07시) 설정
    kst = timezone(timedelta(hours=9))
    now_hour = datetime.now(kst).hour
    if now_hour >= 23 or now_hour < 7:
        print(f"🌙 방해금지 시간({now_hour}시)입니다. 아침 7시까지 수집을 일시 정지합니다.")
        return
    
    print(f"🚀 [{KEYWORD}] 1회성 뉴스 클리핑 및 AI 논조 분석을 시작합니다...")
    sent_links = load_sent_links()
    seen_titles = set()
    
    naver_articles = fetch_naver_news(KEYWORD)
    daum_articles = fetch_daum_news(KEYWORD)
    
    all_articles = naver_articles + daum_articles
    new_count = 0
    
    for article in reversed(all_articles):
        link = article["link"]
        title = article["title"]
        
        if link not in sent_links and title not in seen_titles:
            sent_links.add(link)
            seen_titles.add(title)
            save_sent_link(link)
            
            sentiment, reason = analyze_sentiment(title, article["summary"])
            save_to_excel(article, sentiment, reason)
            
            success = send_telegram_alert(article, sentiment, reason)
            if success:
                new_count += 1
                print(f"✅ [전송 완료] ({article['portal']}) [{sentiment}] {title}")
            else:
                print(f"📝 [대장 기록 완료 / 텔레그램 미전송] ({article['portal']}) {title}")
            time.sleep(1)
            
    print(f"✨ 작업 완료! 신규 기사 {new_count}건 알림 전송 후 정상 종료합니다.")

if __name__ == "__main__":
    main()