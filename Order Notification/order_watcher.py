# -*- coding: utf-8 -*-
"""
수주통보 게시판 새 글 감시 -> 메일 알림 (SMTP 직접 발송)
(HMT/한발 WEHAGO 그룹웨어)

동작 방식
  실제 크롬(Playwright)을 로그인된 세션으로 띄워두고, 1분마다 게시판
  목록 API(ViewBoardArtList) 응답을 가로채서 새 글이 있으면 SMTP로
  메일을 보냅니다. 서명(wehago-sign)/토큰/만료는 브라우저가 알아서 처리.
  Outlook 을 거치지 않으므로 보안 팝업이 뜨지 않습니다.

사용법 (venv 파이썬으로 실행)
  1) 최초 1회 로그인:   python order_watcher.py login
     -> 크롬 창이 뜨면 사내 계정으로 로그인. 게시판이 보이면 창을 닫음.
  2) 감시 시작:         python order_watcher.py watch
  3) 메일 테스트:       python order_watcher.py test

  * 세션이 만료되면(재로그인 필요) 감시를 멈추고 다시 1) 을 실행하세요.
"""

import sys
import os
import json
import time
import datetime as dt

# ----------------------------- 설정 -----------------------------
BOARD_URL   = ("https://hmt.hanbal.kr/#/UF/UFA/UFA0000?specialLnb=Y"
               "&moduleCode=UF&menuCode=3000300_001003&pageCode=UFA3000")
POLL_SECONDS = 600         # 확인 주기(초). 600=10분
                           # ※ 주기가 길면 그 사이 로그인 세션이 만료될 수 있음.
TITLE_FILTER = ""          # 제목에 이 문자열이 포함된 글만 알림(빈칸이면 전체)

# --- 첨부파일 설정 ---
ATTACH_ENABLE = True                # 새 글의 첨부를 받아 메일에 첨부할지
ATTACH_ALWAYS = ["수주통보서"]        # 이 키워드가 파일명에 있으면 항상 받음(대소문자 무시)
# 견적서: 1순위 키워드가 있는 파일을 받고, 하나도 없으면 대체 키워드로 찾아 받음.
#   (대체 키워드 "Priced" 는 부분일치 → "Priced" 와 "Unpriced" 파일 모두 포함)
ATTACH_QUOTE_PRIMARY  = "Quotation"  # 견적서 1순위
ATTACH_QUOTE_FALLBACK = "Priced"     # Quotation 파일이 없을 때 대체(Unpriced 포함)

# --------------------------- 메일(SMTP) 설정 ---------------------------
MAIL_TO   = "sjlee@hanbalmasstech.com"     # 받는 사람
MAIL_FROM = "sjlee@hanbalmasstech.com"     # 보내는 사람

# [방법 A] 사내 메일서버로 발송 (팝업 없음)
#   서버 주소/포트는 Outlook 계정설정의 '보내는 메일 서버(SMTP)'에서 확인하세요.
#   인증이 필요 없는 사내 릴레이면 SMTP_USER / SMTP_PASS 를 비워두세요.
SMTP_HOST     = "wblock.hanbalmasstech.com"
SMTP_PORT     = 25                 # 사내 게이트웨이는 25번만 열려있음(587/465 차단)
SMTP_SECURITY = "none"             # "starttls"(587) / "ssl"(465) / "none"(25)
SMTP_USER     = ""                 # 사내 내부주소 발송은 인증 불필요(비움). 필요 시 아이디
SMTP_PASS     = ""                 # 인증 필요 시 비밀번호. 없으면 비움
#
# [방법 B] 개인 Gmail 로 발송하려면 위 값을 아래처럼 바꾸세요:
#   SMTP_HOST = "smtp.gmail.com"
#   SMTP_PORT = 587
#   SMTP_SECURITY = "starttls"
#   SMTP_USER = "내주소@gmail.com"
#   SMTP_PASS = "앱비밀번호16자리"     (구글계정>보안>2단계인증>앱 비밀번호 에서 발급)
#   MAIL_FROM = "내주소@gmail.com"
# ----------------------------------------------------------------------

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
ATTACH_DIR  = os.path.join(BASE_DIR, "수주통보서_다운로드")  # 첨부 저장 전용 폴더
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile")   # 로그인 세션 저장 폴더
STATE_FILE  = os.path.join(BASE_DIR, "state.json")
LOG_FILE    = os.path.join(BASE_DIR, "order_watcher.log")
# ----------------------------------------------------------------


def log(msg: str):
    line = f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def load_state() -> dict:
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_state(state: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# --------------------------- 메일 발송 (SMTP) ---------------------------
# 메일 게이트웨이(WBlock)가 한글 첨부 파일명을 지워버리므로,
# 첨부 이름은 영문으로 변환해서 보낸다. (PC 저장본은 원본 한글명 유지)
FILENAME_MAP = {
    "수주통보서": "OrderNotice",
    "관련메일": "RelatedMail",
    "견적서": "Quotation",
}


def ascii_filename(fname: str) -> str:
    import re
    for k, v in FILENAME_MAP.items():
        fname = fname.replace(k, v)
    fname = "".join(ch if ord(ch) < 128 else "_" for ch in fname)
    fname = re.sub(r"_{2,}", "_", fname)          # 연속 '_' 정리
    return fname.strip("_ ") or "attachment.pdf"


def send_mail(subject: str, html_body: str, attachments=None):
    """SMTP 로 메일 발송. 설정(SMTP_*) 값을 사용. Outlook 을 거치지 않음.
       attachments: 첨부할 파일 경로 리스트(선택)."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    from email.header import Header
    from email.utils import formatdate

    msg = MIMEMultipart()
    msg["Subject"] = str(Header(subject, "utf-8"))
    msg["From"] = MAIL_FROM
    msg["To"] = MAIL_TO
    msg["Date"] = formatdate(localtime=True)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    import mimetypes
    for path in (attachments or []):
        try:
            fname = ascii_filename(os.path.basename(path))   # 영문 이름으로 변환
            ctype = mimetypes.guess_type(fname)[0] or "application/octet-stream"
            subtype = ctype.split("/", 1)[1]
            with open(path, "rb") as f:
                part = MIMEApplication(f.read(), _subtype=subtype, Name=fname)
            part.add_header("Content-Disposition", "attachment", filename=fname)
            msg.attach(part)
        except Exception as e:
            log(f"첨부 추가 실패({path}): {e}")

    if SMTP_SECURITY == "ssl":
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30)
    else:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
    try:
        server.ehlo()
        if SMTP_SECURITY == "starttls":
            server.starttls()
            server.ehlo()
        if SMTP_USER:                      # 아이디가 설정돼 있으면 인증
            server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(MAIL_FROM, [MAIL_TO], msg.as_string())
    finally:
        try:
            server.quit()
        except Exception:
            pass


def build_mail_html(article: dict) -> str:
    title  = article.get("art_title") or article.get("art_content") or "(제목 없음)"
    writer = article.get("mbr_nick") or ""
    dept   = article.get("deptname") or article.get("dept_name") or ""
    when   = article.get("write_date_origin") or article.get("write_date") or ""
    files  = article.get("file_cnt") or "0"
    seq    = article.get("art_seq_no") or ""
    return f"""
    <div style="font-family:'맑은 고딕',Malgun Gothic,sans-serif;font-size:14px;color:#222">
      <p style="font-size:16px;margin:0 0 12px"><b>새 수주통보 게시글이 등록되었습니다.</b></p>
      <table cellpadding="6" style="border-collapse:collapse;font-size:14px">
        <tr><td style="color:#888">제목</td><td><b>{title}</b></td></tr>
        <tr><td style="color:#888">작성자</td><td>{writer} ({dept})</td></tr>
        <tr><td style="color:#888">등록일시</td><td>{when}</td></tr>
        <tr><td style="color:#888">첨부</td><td>{files} 개</td></tr>
        <tr><td style="color:#888">글번호</td><td>{seq}</td></tr>
      </table>
      <p style="margin-top:16px">
        <a href="{BOARD_URL}">▶ 게시판에서 열기</a>
      </p>
      <p style="color:#aaa;font-size:12px;margin-top:20px">
        자동 알림 · 수주통보 게시판 감시 프로그램
      </p>
    </div>
    """


# ----------------------- 게시판 목록 가져오기 ----------------------
def _matches_list_api(r):
    return "ViewBoardArtList" in r.url and r.request.method == "POST"


def fetch_article_data(page):
    """게시판 목록 API(ViewBoardArtList) 응답 JSON 을 가로챈다.
       1순위: 게시판의 '새로고침 버튼(↻)'만 클릭(페이지 이동 없음, 깜빡임 없음).
       2순위: 버튼을 못 찾으면 게시판으로 진짜 이동해서 목록을 재요청.
       실패하면 None 반환."""
    from playwright.sync_api import TimeoutError as PWTimeout

    # 1순위: 새로고침 버튼 클릭 (class 'titleRefresh')
    try:
        btn = page.query_selector("button.titleRefresh")
    except Exception:
        btn = None
    if btn:
        try:
            with page.expect_response(_matches_list_api, timeout=15000) as resp_info:
                btn.click(timeout=5000)
            return resp_info.value.json()
        except PWTimeout:
            pass                                  # 버튼으로 안 되면 아래로 폴백
        except Exception as e:
            log(f"새로고침 버튼 클릭 실패(폴백 진행): {e}")

    # 2순위(폴백): about:blank 를 거쳐 게시판으로 진짜 이동
    try:
        try:
            page.goto("about:blank")
        except Exception:
            pass
        with page.expect_response(_matches_list_api, timeout=30000) as resp_info:
            page.goto(BOARD_URL, wait_until="domcontentloaded")
        return resp_info.value.json()
    except PWTimeout:
        return None
    except Exception as e:
        log(f"목록 요청 중 예외: {e}")
        return None


def extract_articles(data: dict):
    if not data:
        return []
    return (data.get("resultData") or {}).get("articleList") or []


def order_folder_for(post_title: str) -> str:
    """글 제목에서 수주통보 번호(예: HMT26-142)를 뽑아 저장 폴더 경로를 만든다.
       Rev.1 글도 기본 번호 폴더에 함께 저장(원본+개정본이 한 폴더에 모임).
       번호가 없으면 제목을 폴더명으로 사용."""
    import re
    m = re.search(r"HMT\d{2}-\d{3}[A-Za-z]?", post_title, re.I)
    name = m.group(0).upper() if m else post_title
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "기타"
    return os.path.join(ATTACH_DIR, name)


def select_attachment_lis(files):
    """files: [(li_handle, 파일명텍스트)]. 다운로드할 li 핸들 리스트를 규칙대로 고른다.
       - ATTACH_ALWAYS 키워드 포함 파일은 항상
       - 견적서: ATTACH_QUOTE_PRIMARY(Quotation) 포함 파일, 하나도 없으면
         ATTACH_QUOTE_FALLBACK(Priced) 부분일치 파일(Unpriced 도 포함)"""
    chosen = []

    def add(li):
        if li not in chosen:
            chosen.append(li)

    # 1) 항상 받는 키워드
    for li, txt in files:
        low = txt.lower()
        if any(k.lower() in low for k in ATTACH_ALWAYS):
            add(li)

    # 2) 견적서: Quotation 우선, 없으면 Priced(부분일치 → Unpriced 포함)
    quote = [li for li, txt in files
             if ATTACH_QUOTE_PRIMARY.lower() in txt.lower()]
    if not quote:
        fb = ATTACH_QUOTE_FALLBACK.lower()
        quote = [li for li, txt in files if fb in txt.lower()]
    for li in quote:
        add(li)

    return chosen


def download_matching_attachments(page, post_title, save_dir=None):
    """게시판 목록에서 post_title 글을 열어, 규칙(select_attachment_lis)에 맞는 첨부를
       개별 다운로드 아이콘(div.downIco)으로 내려받아 save_dir 에 저장한다.
       저장된 파일 경로 리스트를 반환. (글은 목록에서 제목 클릭으로 연다)"""
    save_dir = save_dir or order_folder_for(post_title)
    saved = []
    try:
        os.makedirs(save_dir, exist_ok=True)
    except Exception:
        pass

    # 1) 글 열기 — '화면에 보이면서' 제목과 거의 일치하는 가장 작은 요소를 클릭
    #    (숨겨진 미리보기 요소, 제목을 포함한 큰 컨테이너 오클릭 방지)
    open_js = """
    (title) => {
      const cands = Array.from(document.querySelectorAll('*')).filter(el => {
        if (!el.offsetParent) return false;                 // 숨겨진 요소 제외
        const t = (el.textContent || '').trim();
        if (!t.includes(title)) return false;
        return t.length < title.length + 60;                // 큰 컨테이너 제외
      });
      if (!cands.length) return false;
      cands.sort((a, b) => a.textContent.length - b.textContent.length);
      const target = cands[0];
      target.scrollIntoView({block: 'center'});
      target.click();
      return true;
    }
    """
    try:
        opened = page.evaluate(open_js, post_title)
    except Exception as e:
        log(f"첨부: 글 열기 중 예외 '{post_title}': {e}")
        return saved
    if not opened:
        log(f"첨부: 목록에서 '{post_title}' 글을 찾지 못했습니다.")
        return saved

    # 2) 첨부 목록(li)이 나타날 때까지 대기
    try:
        page.wait_for_selector("ul.fb_div li", timeout=10000)
    except Exception:
        log("첨부: 파일 목록을 찾지 못함(첨부 없음 또는 로딩 지연).")
        return saved

    # 보이는 파일들의 (핸들, 파일명) 수집 후 규칙대로 선별
    files = []
    for li in page.query_selector_all("ul.fb_div li"):
        try:
            if not li.is_visible():                    # 뒤에 숨은 목록 제외
                continue
            txt = li.inner_text()
        except Exception:
            continue
        files.append((li, txt))

    for li in select_attachment_lis(files):
        downico = li.query_selector("div.downIco")     # 파일별 다운로드 아이콘
        if not downico:
            continue
        try:
            # 아이콘 클릭 → 'PC저장/ONECHAMBER 저장' 메뉴가 뜸 → 'PC저장' 클릭
            downico.click()
            with page.expect_download(timeout=30000) as dl_info:
                pc = page.wait_for_selector("text=PC저장", timeout=5000)
                pc.click()
            dl = dl_info.value
            fname = dl.suggested_filename or "attachment"
            path = os.path.join(save_dir, fname)
            dl.save_as(path)
            saved.append(path)
            log(f"첨부 저장: {os.path.relpath(path, ATTACH_DIR)}")
        except Exception as e:
            log(f"첨부 다운로드 실패: {e}")
            try:                                       # 열린 메뉴가 남았으면 닫기
                page.keyboard.press("Escape")
            except Exception:
                pass
    return saved


# ------------------------- 새 글 처리 로직 -------------------------
def process(data: dict, state: dict, page=None) -> bool:
    """새 글이 있으면 (첨부 받아서) 메일 발송. 상태가 바뀌면 True."""
    articles = extract_articles(data)
    if not articles:
        log("목록이 비어있음(로그인 상태 확인 필요).")
        return False

    def seq_of(a):
        try:
            return int(a.get("art_seq_no") or 0)
        except (TypeError, ValueError):
            return 0

    max_seq = max(seq_of(a) for a in articles)
    last = state.get("last_seq")

    # 최초 실행: 기존 글을 다 메일로 보내지 않도록 현재 상태를 기준선으로 저장
    if last is None:
        state["last_seq"] = max_seq
        save_state(state)
        log(f"기준선 설정 완료 (최신 글번호 {max_seq}). 이제부터 새 글을 감시합니다.")
        return True

    new_articles = sorted(
        [a for a in articles if seq_of(a) > last],
        key=seq_of,
    )

    # 제목 필터 적용
    if TITLE_FILTER:
        new_articles = [
            a for a in new_articles
            if TITLE_FILTER in (a.get("art_title") or a.get("art_content") or "")
        ]

    if not new_articles:
        return False

    for a in new_articles:
        title = a.get("art_title") or a.get("art_content") or "(제목 없음)"

        # 첨부가 있는 글이면 필터에 맞는 첨부를 받아둔다
        attachments = []
        if ATTACH_ENABLE and page is not None:
            try:
                file_cnt = int(a.get("file_cnt") or 0)
            except (TypeError, ValueError):
                file_cnt = 0
            if file_cnt > 0:
                try:
                    attachments = download_matching_attachments(page, title)
                except Exception as e:
                    log(f"첨부 처리 실패(메일은 첨부 없이 발송): {e}")
                try:
                    page.keyboard.press("Escape")   # 열린 글 팝업 닫기
                except Exception:
                    pass

        try:
            send_mail(f"[수주통보] {title}", build_mail_html(a), attachments=attachments)
            log(f"메일 발송: {title} (글번호 {seq_of(a)}, 첨부 {len(attachments)}개)")
        except Exception as e:
            log(f"메일 발송 실패: {e} / 글: {title}")

    state["last_seq"] = max(last, max_seq)
    save_state(state)
    return True


# ----------------------------- 모드 -----------------------------
def run_login():
    """로그인용: 크롬 창을 띄우고 사용자가 직접 로그인하도록 한다."""
    from playwright.sync_api import sync_playwright
    log("로그인 모드: 크롬 창이 뜹니다. 사내 계정으로 로그인 후 게시판이 보이면 창을 닫으세요.")
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, args=["--start-maximized"],
            no_viewport=True,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(BOARD_URL)
        log("로그인이 끝나면 이 창(콘솔)에서 Enter 를 누르거나, 그냥 크롬 창을 닫으세요.")
        try:
            # 브라우저가 닫힐 때까지 대기
            page.wait_for_event("close", timeout=0)
        except Exception:
            pass
        try:
            ctx.close()
        except Exception:
            pass
    log("로그인 세션이 저장되었습니다. 이제 'python order_watcher.py watch' 로 감시를 시작하세요.")


def _is_login_page(page) -> bool:
    """현재 페이지가 로그인 화면인지 판별."""
    try:
        if "/login" in (page.url or ""):
            return True
        return bool(page.query_selector("input[type=password]"))
    except Exception:
        return False


def run_watch():
    """감시 모드: 크롬 창을 계속 띄워둔 채 1분마다 목록을 확인.
       로그인 화면이면 그 창에서 직접 로그인하면 자동으로 이어진다.
       (이 그룹웨어는 세션 쿠키를 써서, 창을 닫으면 로그인이 풀림)"""
    from playwright.sync_api import sync_playwright
    state = load_state()
    log(f"감시 시작 (주기 {POLL_SECONDS}초, 수신자 {MAIL_TO}).")
    log("※ 크롬 창을 '닫지 말고' 열어두세요(최소화는 OK). 닫으면 로그인이 풀립니다.")

    login_alerted = False
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, accept_downloads=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            page.goto(BOARD_URL, wait_until="domcontentloaded")
        except Exception:
            pass

        while True:
            # 1) 로그인 화면이면: 네비게이션하지 말고(입력 방해 방지) 로그인 완료를 기다림
            if _is_login_page(page):
                log("로그인이 필요합니다 → 열려있는 크롬 창에서 사내 계정으로 로그인하세요. (되면 자동으로 이어집니다)")
                if not login_alerted:
                    try:
                        send_mail("[수주통보 감시] 로그인 필요",
                                  "<p>감시 프로그램의 크롬 창에서 다시 로그인해 주세요. "
                                  "로그인하면 자동으로 감시가 재개됩니다.</p>")
                    except Exception as e:
                        log(f"알림 메일 실패: {e}")
                    login_alerted = True
                time.sleep(5)          # 로그인 대기 중에는 짧게 재확인
                continue

            login_alerted = False

            # 2) 로그인 상태 → 게시판을 새로 로드해 목록 API 응답을 가로챔
            data = fetch_article_data(page)
            articles = extract_articles(data)
            if data is None or not articles:
                # 방금 로그아웃됐을 수 있으니 다음 루프에서 로그인 화면 판별
                if _is_login_page(page):
                    continue
                log("목록을 가져오지 못함(일시적 오류일 수 있음). 다음 주기에 재시도.")
            else:
                if not process(data, state, page):
                    log(f"확인 완료 — 새 글 없음 (최신 글번호 {state.get('last_seq')}).")

            time.sleep(POLL_SECONDS)


def run_test():
    log(f"테스트 메일 발송 시도... (SMTP {SMTP_HOST}:{SMTP_PORT}, 보안={SMTP_SECURITY})")
    try:
        send_mail(
            "[수주통보 감시] 테스트 메일",
            "<p>이 메일이 보이면 SMTP 발송이 정상 동작합니다.</p>",
        )
    except Exception as e:
        log(f"발송 실패: {e}")
        log("→ SMTP_HOST/PORT/SECURITY 또는 아이디/비밀번호 설정을 확인하세요.")
        return
    log(f"테스트 메일을 {MAIL_TO} 로 보냈습니다. 받은편지함을 확인하세요.")


def run_diag():
    """진단: 크롬을 '보이게' 띄워 로그인 상태와 목록 API 발생 여부를 확인."""
    from playwright.sync_api import sync_playwright
    seen = []
    log("진단 모드: 크롬 창을 띄워 게시판을 열어봅니다. 무슨 일이 일어나는지 관찰하세요.")
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(PROFILE_DIR, headless=False)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        def on_resp(r):
            u = r.url
            if ("APIHandler" in u) or ("ViewBoardArtList" in u) or ("login" in u.lower()):
                seen.append(u)
                log(f"  요청감지: {r.request.method} {r.status} {u[:110]}")

        page.on("response", on_resp)
        try:
            page.goto(BOARD_URL, wait_until="domcontentloaded")
        except Exception as e:
            log(f"goto 예외: {e}")
        page.wait_for_timeout(20000)      # 20초 관찰

        log(f"최종 URL   : {page.url}")
        try:
            log(f"페이지 제목: {page.title()}")
        except Exception:
            pass
        has_pw = False
        try:
            has_pw = bool(page.query_selector("input[type=password]"))
        except Exception:
            pass
        log(f"비밀번호 입력창 존재(=로그인 안 됨): {has_pw}")
        log(f"ViewBoardArtList 감지 횟수: {sum('ViewBoardArtList' in u for u in seen)}")
        log("확인이 끝나면 크롬 창을 닫으세요.")
        try:
            page.wait_for_event("close", timeout=0)
        except Exception:
            pass
        try:
            ctx.close()
        except Exception:
            pass


def run_explore():
    """첨부 탐색: 크롬을 띄우고, 사용자가 첨부 있는 글을 열면
       (1) 첨부목록 API(ecm001A09) 응답 JSON 과
       (2) 화면의 '저장/다운로드' 버튼 후보 요소를 자동으로 찾아 출력한다."""
    from playwright.sync_api import sync_playwright
    log("첨부 탐색 모드: 크롬이 뜨면 첨부가 있는 글(예: 107615 UOP LLC)을 직접 여세요.")
    log("글을 열면 첨부목록/저장버튼 정보를 자동 출력합니다. 다 되면 크롬 창을 닫으세요.")
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(PROFILE_DIR, headless=False)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        def on_resp(r):
            u = r.url
            if ("ecm001A09" in u) or ("ecm001A03" in u):
                try:
                    body = json.dumps(r.json(), ensure_ascii=False)
                    log(f"[API] {u.split('/')[-1]} 응답(JSON, 앞부분):")
                    log("  " + body[:1800])
                except Exception:
                    log(f"[API] {u.split('/')[-1]} 응답(비 JSON, 다운로드 요청으로 추정)")

        page.on("response", on_resp)
        try:
            page.goto(BOARD_URL, wait_until="domcontentloaded")
        except Exception:
            pass

        # 저장 아이콘/다운로드 버튼을 iframe 까지 포함해 정밀 탐색
        selectors = [
            'img[src*="save"]', 'img[src*="ic_pc"]', 'img[src*="down"]',
            '[title*="저장"]', '[title*="다운"]', '[title*="첨부"]',
            '[class*="download"]', '[class*="Download"]', '[class*="attach"]',
            '[class*="titleDown"]', '[class*="titleSave"]', '[class*="File"]',
        ]
        printed = set()
        dumped_detail = False
        deadline = time.time() + 300     # 최대 5분 관찰
        log("... 첨부가 보이게 글을 열어 주세요. (저장 아이콘을 계속 스캔합니다) ...")
        while time.time() < deadline:
            # 첨부 목록 영역(detail-file) 전체 구조를 한 번 통째로 출력
            if not dumped_detail:
                try:
                    frames_now = page.frames
                except Exception:
                    frames_now = [page]
                for fr in frames_now:
                    try:
                        node = fr.query_selector("div.detail-file .UpDownLoader") or fr.query_selector("div.detail-file")
                    except Exception:
                        node = None
                    if node:
                        try:
                            html = node.evaluate("e => e.outerHTML") or ""
                        except Exception:
                            html = ""
                        if html:
                            dump_path = os.path.join(BASE_DIR, "_detail_dump.html")
                            try:
                                with open(dump_path, "w", encoding="utf-8") as f:
                                    f.write(html)
                                log(f"첨부영역 HTML을 파일로 저장했습니다: {dump_path}")
                            except Exception as e:
                                log(f"HTML 저장 실패: {e}")
                            dumped_detail = True
                            break
            try:
                if page.is_closed():
                    break
            except Exception:
                break
            try:
                frames = page.frames
            except Exception:
                frames = [page]
            for fr in frames:
                for sel in selectors:
                    try:
                        els = fr.query_selector_all(sel)
                    except Exception:
                        continue
                    for el in els:
                        try:
                            # 실제로 클릭할 수 있는 조상 요소(버튼/링크/onclick)를 찾아 출력
                            info = el.evaluate(
                                "e => { const c = e.closest('button,a,[onclick],[role=button]') || e.parentElement || e;"
                                " return {tag:c.tagName, cls:c.className, title:c.getAttribute('title')||'', html:c.outerHTML.slice(0,240)}; }"
                            )
                        except Exception:
                            continue
                        key = info.get("html", "")
                        if not key or key in printed:
                            continue
                        printed.add(key)
                        furl = ""
                        try:
                            furl = fr.url
                        except Exception:
                            pass
                        inframe = "" if fr == page.main_frame else "  [iframe]"
                        log(f"후보<{info.get('tag')}>{inframe} class='{info.get('cls')}' title='{info.get('title')}'")
                        log(f"   HTML: {info.get('html')}")
                        log(f"   (frame: {furl[:70]})")
            time.sleep(3)
        try:
            ctx.close()
        except Exception:
            pass
    log("탐색 종료. 위에 나온 '후보버튼'과 'ecm001A09 응답'을 붙여주세요.")


def run_attachtest():
    """첨부 다운로드 테스트: 지정한 글을 열어 '수주통보서' 첨부를 받아
       전용 폴더에 저장하고, 그 파일을 첨부한 메일을 보낸다."""
    from playwright.sync_api import sync_playwright
    title = sys.argv[2] if len(sys.argv) > 2 else "수주통보서(HMT26-146) 현대중공업파워시스템"
    log(f"첨부 테스트: '{title}' — 항상:{ATTACH_ALWAYS}, "
        f"견적서:{ATTACH_QUOTE_PRIMARY}(없으면 {ATTACH_QUOTE_FALLBACK}) 받습니다.")
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, accept_downloads=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        data = fetch_article_data(page)
        if not extract_articles(data):
            log("목록 로드 실패 — 크롬 창에서 로그인 후, 로그인되면 자동으로 진행합니다.")
            for _ in range(60):        # 최대 5분 로그인 대기
                if _is_login_page(page):
                    time.sleep(5)
                    continue
                data = fetch_article_data(page)
                if extract_articles(data):
                    break
                time.sleep(5)

        # 목록에서 실제 게시글 정보를 찾아 진짜 알림과 동일하게 구성
        art = None
        for a in extract_articles(data):
            if title in (a.get("art_title") or ""):
                art = a
                break
        if art is None:
            art = {"art_title": title, "mbr_nick": "(테스트)", "file_cnt": "0"}

        saved = download_matching_attachments(page, title)
        log(f"저장 완료: {len(saved)}개 -> {ATTACH_DIR}")

        try:
            send_mail(f"[수주통보] {title}", build_mail_html(art), attachments=saved)
            log(f"첨부 메일 발송 완료 ({len(saved)}개 첨부).")
        except Exception as e:
            log(f"첨부 메일 발송 실패: {e}")
        try:
            ctx.close()
        except Exception:
            pass


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    if mode == "login":
        run_login()
    elif mode == "attachtest":
        run_attachtest()
    elif mode == "watch":
        run_watch()
    elif mode == "test":
        run_test()
    elif mode == "diag":
        run_diag()
    elif mode == "explore":
        run_explore()
    else:
        print(__doc__)
        print("사용법: python order_watcher.py [login|watch|test]")


if __name__ == "__main__":
    main()
