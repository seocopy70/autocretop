import json
import re
import random
import time
import argparse
from pathlib import Path
from datetime import datetime

from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# CRETOP 최종 자동 수집
#
# 기능
# 1. 현재 CRETOP 검색결과 화면에서 시작
# 2. 검색결과 전체 회사 수 확인
# 3. 현재 페이지부터 마지막 페이지까지 자동 수집
#
# 검색결과 카드
# - 회사명
# - 대표자명
# - 산업분류
# - 전화번호
# - 주소
#
# 산업분류 제외
# - 투자
# - 금융
# - 컨설팅
# - 보험
# - 법무
# - 자문
# - 세무
# - 회계
#
# 일반 화면
# - 설립년도
# - 종업원수
# - 이메일
#
# 재무 화면
# - 최근 5개년
# - 매출액
# - 영업이익
# - 순이익
#
# AI 판정 (법인 보험영업용)
# - A_우수 : 3년 영업흑자 + 규모(자산100억/자본30억) + 점수≥9
# - B_양호 : 최근 흑자+자본양수 또는 점수≥6
# - C_보통 / D_주의 / E_위험 / F_판단보류
#
# Excel
# - 종합결과
# - 재무정보
# - 검색정보
# - E~F등급
#
# JSON
# - 전체 분석 결과 저장
#
# 중요
# - 상세 화면에서 page.go_back() 사용
# - 검증된 페이지 이동 코드 유지
# - Chrome 종료하지 않음
# - 전체 검색결과 수를 처리 상한으로 사용
# - 페이지당 회사 수는 동적으로 대응
# ============================================================


# ============================================================
# 설정
# ============================================================

EXCLUDED_INDUSTRY_KEYWORDS = [
    "투자",
    "금융",
    "컨설팅",
    "보험",
    "법무",
    "자문",
    "세무",
    "회계",
]

FINANCIAL_YEAR_COUNT = 5

# 재무정보 Excel의 고정 연도
FINANCIAL_YEARS = [
    2025,
    2024,
    2023,
    2022,
    2021,
]

DEFAULT_FINANCIAL_UNIT = "백만원"

DETAIL_WAIT_MS = 150
MANUAL_PAGE_CHANGE_TIMEOUT_SECONDS = 60


class ManualInterventionTimeout(Exception):
    """팝업 수동 처리 시간 초과 — 수집 중단"""
    pass


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def human_delay(min_ms=20, max_ms=60):
    """짧은 랜덤 대기 (속도 우선, 최소만)"""

    delay = random.uniform(min_ms, max_ms) / 1000.0
    time.sleep(delay)


# ============================================================
# 정보제공 거부 등 차단 팝업 닫기
# ============================================================

def is_page_expired(page):

    try:

        text = normalize_text(
            page.locator("body").inner_text()
        )

        keywords = [
            "웹페이지가 만료",
            "페이지가 만료",
            "세션이 만료",
            "로그인 세션",
            "다시 로그인",
            "중복 로그인",
        ]

        return any(k in text for k in keywords)

    except Exception:
        return False


def dismiss_blocking_popups(page, context=""):

    """
    알려진 차단 팝업만 닫는다.
    정보제공중지 → 팝업 내부 확인 버튼 일반 클릭
    동시접속 초과 → 자동 클릭 안 함
    """

    try:

        # Playwright로 정보제공중지 확인 버튼 직접 클릭 (가장 확실)
        try:

            modal = page.locator(
                ".pop-area.PLCM910P1, "
                ".modals-container [role='dialog'], "
                "[role='dialog'][aria-modal='true']"
            )

            for i in range(min(modal.count(), 5)):

                root = modal.nth(i)

                try:
                    if not root.is_visible():
                        continue
                except Exception:
                    continue

                t = ""
                try:
                    t = normalize_text(root.inner_text())
                except Exception:
                    continue

                if not any(
                    k in t
                    for k in [
                        "정보제공중지",
                        "정보가 공개되지",
                        "제공중지",
                        "제공 중지",
                        "제공거부",
                    ]
                ):
                    continue

                for label in ("확인",):

                    btn = root.locator(
                        "button",
                        has_text=label,
                    )

                    for j in range(min(btn.count(), 3)):

                        b = btn.nth(j)

                        try:
                            if normalize_text(b.inner_text()) != label:
                                continue
                            if not b.is_visible():
                                continue
                            b.click(timeout=3000)
                        except Exception:
                            continue

                        msg = f"정보제공중지 닫음: '{label}'"
                        if context:
                            msg = f"{context} {msg}"
                        print(f"    ✓ {msg}")

                        page.wait_for_timeout(
                            random.randint(200, 350)
                        )
                        return True

        except Exception:
            pass

        # JS 백업
        result = page.evaluate(
            """() => {
                function norm(s) {
                    return (s || '').replace(/\\s+/g, ' ').trim();
                }
                const roots = document.querySelectorAll(
                    '.pop-area.PLCM910P1, .modals-container [role=dialog], [role=dialog][aria-modal=true], .pop-alert-close'
                );
                for (const root of roots) {
                    if (root.closest('.quick-wrap, .quick-view')) continue;
                    const style = window.getComputedStyle(root);
                    const r = root.getBoundingClientRect();
                    if (style.display === 'none' || r.width < 40) continue;
                    const text = norm(root.innerText);
                    if (/동시접속자|기존 이용자를 종료/.test(text)) {
                        return {ok: false, kind: 'concurrent_login', text: text.slice(0, 80)};
                    }
                }
                return {ok: false, kind: null};
            }"""
        )

        if result and result.get("kind") == "concurrent_login":

            print(
                f"    ⚠ {(context + ' ') if context else ''}"
                f"동시접속 팝업 감지 (자동 닫기 안 함)"
            )
            return False

    except Exception:
        pass

    return False



def normalize_text(text):
    if text is None:
        return ""

    text = str(text)

    text = (
        text
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .replace("\r", " ")
        .replace("\n", " ")
    )

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def safe_number(value):
    """
    숫자 문자열을 숫자로 변환한다.

    예:
    19,196 -> 19196
    -2,493 -> -2493
    8,051 -> 8051
    - -> None
    """

    if value is None:
        return None

    text = normalize_text(value)

    if not text:
        return None

    if text in [
        "-",
        "–",
        "—",
        "없음",
        "N/A",
        "n/a",
    ]:
        return None

    negative = False

    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    text = text.replace(",", "")

    text = re.sub(
        r"[^0-9.\-]",
        "",
        text,
    )

    if not text:
        return None

    try:
        number = float(text)

        if negative:
            number = -abs(number)

        if number.is_integer():
            return int(number)

        return number

    except Exception:
        return None


def format_number(value):
    if value is None:
        return ""

    try:
        if isinstance(value, float) and value.is_integer():
            value = int(value)

        return f"{value:,}"

    except Exception:
        return str(value)


def extract_year(text):
    text = normalize_text(text)

    match = re.search(
        r"(19|20)\d{2}",
        text,
    )

    if match:
        return int(match.group(0))

    return None


# ============================================================
# 로그인 상태 확인
# ============================================================

def is_login_page(page):

    try:
        url = page.url.lower()

        if "login" in url:
            return True

        body_text = normalize_text(
            page.locator("body").inner_text()
        )

        login_keywords = [
            "로그인",
            "아이디",
            "비밀번호",
        ]

        if "로그인" in body_text and (
            "아이디" in body_text
            or "비밀번호" in body_text
        ):
            return True

    except Exception:
        pass

    return False


def check_login(page):

    if is_login_page(page):
        print()
        print("❌ CRETOP 로그인 화면입니다.")
        print(
            "Chrome에서 CRETOP에 로그인한 뒤 "
            "다시 실행해 주세요."
        )

        return False

    return True


# ============================================================
# 검색결과 목록 대기
# ============================================================

def wait_for_search_list(page, timeout=15000):

    page.wait_for_selector(
        "ul.search-result__list > li",
        timeout=timeout,
    )


# ============================================================
# 검색결과 카드
# ============================================================

def get_search_cards(page):

    return page.locator(
        "ul.search-result__list > li"
    )


# ============================================================
# 카드 안의 특정 항목
# ============================================================

def get_value(card, label):

    rows = card.locator(
        "span.list-tit"
    )

    for i in range(rows.count()):

        try:
            title = normalize_text(
                rows.nth(i).inner_text()
            )

        except Exception:
            continue

        if title != label:
            continue

        try:

            row = rows.nth(i).locator(
                ".."
            )

            values = row.locator(
                "span.list-info"
            )

            parts = []

            for j in range(values.count()):

                try:

                    text = normalize_text(
                        values.nth(j).inner_text()
                    )

                except Exception:
                    continue

                if text and text != "·":
                    parts.append(text)

            if label == "기업유형/형태":
                return " · ".join(parts)

            return " ".join(parts)

        except Exception:
            return ""

    return ""


# ============================================================
# 회사명
# ============================================================

def get_company_name(card, index=0):

    name_button = card.locator(
        "button.result-layer-open"
    )

    if name_button.count() > 0:

        try:

            name = normalize_text(
                name_button.first.inner_text()
            )

            if name:
                return name

        except Exception:
            pass

    buttons = card.locator("button")

    for j in range(buttons.count()):

        try:

            button = buttons.nth(j)

            button_class = (
                button.get_attribute("class")
                or ""
            )

            if "result-layer-open" not in button_class:
                continue

            name = normalize_text(
                button.inner_text()
            )

            if name:
                return name

        except Exception:
            continue

    fallback_button = card.locator(
        "div.result-txt-wrap button"
    )

    if fallback_button.count() > 0:

        try:

            name = normalize_text(
                fallback_button.first.inner_text()
            )

            if name:
                return name

        except Exception:
            pass

    spans = card.locator("span")

    excluded = [
        "대표자명",
        "기업상태",
        "기업유형/형태",
        "사업자번호",
        "법인번호",
        "산업분류",
        "주소",
        "전화번호",
        "최근 재무년도",
    ]

    for j in range(spans.count()):

        try:

            text = normalize_text(
                spans.nth(j).inner_text()
            )

            if not text:
                continue

            if len(text) > 100:
                continue

            if text in excluded:
                continue

            if "(주)" in text or "㈜" in text:
                return text

        except Exception:
            continue

    print(
        f"⚠ 회사명을 찾지 못했습니다. "
        f"→ {index + 1}번째 카드"
    )

    return ""


# ============================================================
# 검색결과 기본정보
# ============================================================

def collect_search_basic_info(card, index):

    company_name = get_company_name(
        card,
        index,
    )

    representative = get_value(
        card,
        "대표자명",
    )

    industry = get_value(
        card,
        "산업분류",
    )

    phone = get_value(
        card,
        "전화번호",
    )

    address = get_value(
        card,
        "주소",
    )

    return {
        "회사명": company_name,
        "대표자명": representative,
        "산업분류": industry,
        "전화번호": phone,
        "주소": address,
    }


# ============================================================
# 산업분류 제외 여부
# ============================================================

def is_excluded_industry(industry):

    industry = normalize_text(industry)

    if not industry:
        return False

    for keyword in EXCLUDED_INDUSTRY_KEYWORDS:

        if keyword in industry:
            return True

    return False


# ============================================================
# 현재 페이지 식별값
# ============================================================

def get_current_page_signature(page):

    cards = get_search_cards(page)

    count = cards.count()

    if count == 0:
        return ""

    try:
        first_name = get_company_name(
            cards.nth(0),
            0,
        )

    except Exception:
        first_name = ""

    try:
        last_name = get_company_name(
            cards.nth(count - 1),
            count - 1,
        )

    except Exception:
        last_name = ""

    return (
        f"{first_name}|"
        f"{last_name}|"
        f"{count}"
    )


# ============================================================
# 전체 검색결과 개수
#
# 가장 먼저
# "검색결과 총54건"
# 형태를 직접 찾는다.
#
# CRETOP 화면의 다른 "200건" 같은 숫자를
# 잘못 가져오는 문제를 방지한다.
# ============================================================

def get_total_result_count(page):

    # --------------------------------------------------------
    # 1순위: body 전체에서 "검색결과 총54건"
    # --------------------------------------------------------

    try:

        body_text = normalize_text(
            page.locator("body").inner_text()
        )

        patterns = [
            r"검색결과\s*총\s*([\d,]+)\s*건",
            r"검색결과\s*총\s*([\d,]+)\s*개",
            r"검색결과\s*([\d,]+)\s*건",
            r"검색결과\s*([\d,]+)\s*개",
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                body_text,
            )

            if match:

                number = (
                    match.group(1)
                    .replace(",", "")
                )

                return int(number)

    except Exception:
        pass

    # --------------------------------------------------------
    # 2순위: 검색결과 관련 영역만 탐색
    # --------------------------------------------------------

    selectors = [
        "[class*='search-result']",
        "[class*='result-count']",
    ]

    for selector in selectors:

        locator = page.locator(selector)

        count = locator.count()

        for i in range(
            min(count, 30)
        ):

            try:

                element = locator.nth(i)

                if not element.is_visible():
                    continue

                text = normalize_text(
                    element.inner_text()
                )

                if not text:
                    continue

                patterns = [
                    r"검색결과\s*총\s*([\d,]+)\s*건",
                    r"검색결과\s*총\s*([\d,]+)\s*개",
                    r"총\s*([\d,]+)\s*건",
                    r"총\s*([\d,]+)\s*개",
                ]

                for pattern in patterns:

                    match = re.search(
                        pattern,
                        text,
                    )

                    if match:

                        number = (
                            match.group(1)
                            .replace(",", "")
                        )

                        return int(number)

            except Exception:
                continue

    return 0


# ============================================================
# 검색 제목 추출
#
# 예: "상속컨설팅 거래처 탐색 / 인천 강화군검색결과 총6건"
#  → "상속컨설팅 거래처_인천 강화군"
# ============================================================

def get_search_title(page):

    texts = []

    # 헤더성 영역 우선 수집
    header_selectors = [
        ".search-result__header",
        ".search-result-header",
        ".result-header",
        ".search-top",
        "[class*='search-result']",
        "[class*='result-count']",
        ".title__main",
        ".banner-data-header",
        "h2",
        "h3",
    ]

    for selector in header_selectors:

        try:

            loc = page.locator(selector)
            count = min(loc.count(), 15)

            for i in range(count):

                try:
                    t = normalize_text(
                        loc.nth(i).inner_text()
                    )
                except Exception:
                    continue

                if t and "검색결과" in t:
                    texts.append(t)

        except Exception:
            continue

    # body 전체 + 앞부분
    try:

        body_text = normalize_text(
            page.locator("body").inner_text()
        )
        texts.append(body_text[:2000])
        texts.append(body_text)

    except Exception:
        pass

    def parse_title(text):

        if not text or "검색결과" not in text:
            return ""

        patterns = [
            # 상속컨설팅 거래처 탐색 / 인천 강화군검색결과 총6건
            r"([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ·_\-]{0,30}?)\s*탐색\s*/\s*([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ·_\-]{0,30}?)검색결과\s*총\s*[\d,]+\s*건",
            r"([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ·_\-]{0,30}?)\s*탐색\s*/\s*([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ·_\-]{0,30}?)\s+검색결과\s*총\s*[\d,]+\s*건",
            # ○○ / △△ 검색결과 총N건
            r"([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ·_\-]{0,30}?)\s*/\s*([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ·_\-]{0,30}?)\s*검색결과\s*총\s*[\d,]+\s*건",
            # 직전 40자 + 검색결과 총N건 (마지막 매치 사용)
            r"([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ·_\-/]{1,40}?)\s*검색결과\s*총\s*[\d,]+\s*건",
        ]

        for pattern in patterns:

            matches = list(
                re.finditer(pattern, text)
            )

            if not matches:
                continue

            # 가장 마지막 매치가 실제 결과 헤더일 가능성 높음
            match = matches[-1]
            parts = []

            for g in match.groups():

                g = normalize_text(g or "")
                g = re.sub(r"\s*탐색\s*", " ", g)
                g = normalize_text(g).strip(" /_-")

                if not g:
                    continue

                # 메뉴 단어 제거 후 뒤쪽 의미 구간만 사용
                noise_words = [
                    "메뉴",
                    "전체메뉴",
                    "통합",
                    "기업",
                    "스마트검색",
                    "브리핑",
                    "일반",
                    "현황",
                    "재무",
                    "신용",
                    "로그인",
                    "관심기업",
                    "최근조회내역",
                    "고객센터",
                    "검색",
                    "결과",
                ]

                tokens = g.split()
                filtered = [
                    t
                    for t in tokens
                    if t not in noise_words
                ]

                if filtered:
                    # 앞쪽 메뉴 잔여를 줄이기 위해 뒤에서 최대 4토큰
                    g = " ".join(filtered[-4:])

                g = g.strip(" /_-")

                if not g or g in noise_words:
                    continue

                if g not in parts:
                    parts.append(g)

            if parts:
                return "_".join(parts)

        return ""

    for text in texts:

        title = parse_title(text)

        if title:
            return title

    return ""


def unique_output_path(path):

    """
    동일 파일명이 있으면 _2, _3 ... 을 붙여 덮어쓰지 않음.
    """

    path = Path(path)

    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    n = 2

    while True:

        candidate = parent / f"{stem}_{n}{suffix}"

        if not candidate.exists():
            return candidate

        n += 1


def sanitize_filename(name):

    name = normalize_text(name)

    if not name:
        return "수집결과"

    # 앞에 붙은 "검색결과" 제거
    name = re.sub(
        r"^검색결과[\s_\-]*",
        "",
        name,
    )
    name = re.sub(
        r"[\s_\-]*검색결과[\s_\-]*",
        " ",
        name,
    )
    name = normalize_text(name).strip(" _-")

    # Windows 파일명 금지 문자 제거
    name = re.sub(
        r'[\\/:*?"<>|]',
        "",
        name,
    )

    name = re.sub(
        r"\s+",
        " ",
        name,
    ).strip()

    # 너무 길면 자름
    if len(name) > 80:
        name = name[:80].strip()

    return name or "수집결과"


# ============================================================
# 일반 메뉴 찾기
# ============================================================

def find_general_link_in_card(card):

    links = card.locator(
        "ul.btn__list "
        "a[title='일반 페이지로 이동하기']"
    )

    count = links.count()

    if count == 0:
        return None

    # 일반 링크가 여러 개 존재하는 경우
    # 마지막 링크를 사용
    return links.last


# ============================================================
# 일반 페이지 진입
# ============================================================

def click_general_from_card(
    page,
    card,
    company_name,
):

    link = find_general_link_in_card(
        card
    )

    if link is None:

        raise RuntimeError(
            "'일반 페이지로 이동하기' "
            "링크를 찾지 못했습니다. "
            f"회사={company_name}"
        )

    try:
        link.scroll_into_view_if_needed()

    except Exception:
        pass

    try:
        link.click()

    except Exception:
        link.click(force=True)

    for _ in range(20):

        page.wait_for_timeout(
            random.randint(50, 100)
        )

        try:

            if "ETGN" in page.url:
                break

        except Exception:
            pass

    page.wait_for_timeout(
        DETAIL_WAIT_MS + random.randint(0, 80)
    )

    if is_info_provision_blocked(page):
        print("    정보제공중지 안내 감지: 팝업 처리 경로로 넘깁니다.")
        return

    if "ETGN" not in page.url:

        try:

            body_text = normalize_text(
                page.locator("body").inner_text()
            )

            if "설립년월" not in body_text:

                raise RuntimeError(
                    "일반 화면으로 이동하지 못했습니다."
                )

        except Exception:

            raise RuntimeError(
                f"일반 화면 이동 실패: {page.url}"
            )


# ============================================================
# 일반 페이지 표 추출
# ============================================================

def extract_general_table(page):

    tables = page.locator("table")

    for i in range(tables.count()):

        table = tables.nth(i)

        try:

            text = normalize_text(
                table.inner_text()
            )

            if (
                "대표자명" in text
                and "설립년월" in text
                and "종업원수" in text
            ):

                result = {}

                rows = table.locator("tr")

                for r in range(rows.count()):

                    row = rows.nth(r)

                    ths = row.locator("th")
                    tds = row.locator("td")

                    pair_count = min(
                        ths.count(),
                        tds.count(),
                    )

                    for j in range(
                        pair_count
                    ):

                        try:

                            key = normalize_text(
                                ths.nth(j).inner_text()
                            )

                            value = normalize_text(
                                tds.nth(j).inner_text()
                            )

                            if key:
                                result[key] = value

                        except Exception:
                            continue

                if result:
                    return result

        except Exception:
            continue

    return {}


# ============================================================
# 일반 기본정보
# ============================================================

def extract_general_basic_info(page):

    table_data = extract_general_table(
        page
    )

    establishment = table_data.get(
        "설립년월",
        "",
    )

    establishment_year = extract_year(
        establishment
    )

    if establishment_year is None:

        establishment_year = extract_year(
            table_data.get(
                "설립년도",
                "",
            )
        )

    employee = table_data.get(
        "종업원수",
        "",
    )

    email = table_data.get(
        "이메일",
        "",
    )

    return {
        "설립년도": (
            establishment_year
            if establishment_year is not None
            else ""
        ),
        "종업원수": employee,
        "이메일": email,
    }


# ============================================================
# 재무 테이블 찾기
# ============================================================

def find_financial_table(page):

    tables = page.locator("table")

    required = [
        "결산기준일자",
        "매출액",
        "영업이익",
        "순이익",
    ]

    for i in range(tables.count()):

        table = tables.nth(i)

        try:

            text = normalize_text(
                table.inner_text()
            )

            if all(
                keyword in text
                for keyword in required
            ):
                return table

        except Exception:
            continue

    return None


# ============================================================
# 재무 단위
# ============================================================

def extract_financial_unit(page):

    try:

        table = find_financial_table(
            page
        )

        if table is not None:

            for locator in [
                table.locator(".."),
                table.locator("../.."),
                table.locator("../../.."),
            ]:

                try:

                    text = normalize_text(
                        locator.inner_text()
                    )

                    match = re.search(
                        r"\(([^)]*(?:원|천원|백만원|억원)[^)]*)\)",
                        text,
                    )

                    if match:

                        return normalize_text(
                            match.group(1)
                        )

                    match = re.search(
                        r"(억원|백만원|천원|원)",
                        text,
                    )

                    if match:
                        return match.group(1)

                except Exception:
                    continue

    except Exception:
        pass

    return DEFAULT_FINANCIAL_UNIT


# ============================================================
# 재무정보 추출
# ============================================================

def extract_financial_data(page):

    table = find_financial_table(
        page
    )

    if table is None:

        return (
            [],
            DEFAULT_FINANCIAL_UNIT,
        )

    unit = extract_financial_unit(
        page
    )

    financial_rows = []

    rows = table.locator("tr")

    for r in range(rows.count()):

        row = rows.nth(r)

        cells = row.locator(
            "th, td"
        )

        count = cells.count()

        if count < 7:
            continue

        try:

            values = []

            for c in range(count):

                values.append(
                    normalize_text(
                        cells.nth(c).inner_text()
                    )
                )

            if (
                "결산기준일자" in values
                or "매출액" in values
                or "영업이익" in values
            ):
                continue

            year = extract_year(
                values[0]
            )

            if year is None:
                continue

            # 0 결산기준일자
            # 1 총자산
            # 2 납입자본금
            # 3 자본총계
            # 4 매출액
            # 5 영업이익
            # 6 순이익

            total_assets = safe_number(
                values[1]
            )

            equity = safe_number(
                values[3]
            )

            revenue = safe_number(
                values[4]
            )

            operating_profit = safe_number(
                values[5]
            )

            net_income = safe_number(
                values[6]
            )

            financial_rows.append({
                "연도": year,
                "총자산": total_assets,
                "자본총계": equity,
                "매출액": revenue,
                "영업이익": operating_profit,
                "순이익": net_income,
            })

        except Exception:
            continue

    # 동일 연도 중복 제거
    unique = {}

    for row in financial_rows:

        unique[
            row["연도"]
        ] = row

    financial_rows = list(
        unique.values()
    )

    financial_rows.sort(
        key=lambda x: x["연도"],
        reverse=True,
    )

    # 5개년
    financial_rows = financial_rows[
        :FINANCIAL_YEAR_COUNT
    ]

    return (
        financial_rows,
        unit,
    )


# ============================================================
# 추세 계산
# ============================================================

def calculate_trend(values):

    values = [
        v
        for v in values
        if v is not None
    ]

    if len(values) < 2:
        return "판단불가"

    first = values[0]
    last = values[-1]

    if first == 0:

        if last > 0:
            return "상승"

        if last < 0:
            return "하락"

        return "보합"

    change_ratio = (
        (last - first)
        / abs(first)
    )

    if change_ratio >= 0.10:
        return "상승"

    if change_ratio <= -0.10:
        return "하락"

    return "보합"


# ============================================================
# 재무판단
#
# 목적: 법인 보험영업용 — 쓸 여력이 있는 기업 선별
#
# A_우수 : 3년 영업흑자 + 규모(자산100억/자본30억) + 점수≥9
# B_양호 : 최근 흑자+자본양수 또는 점수≥6
# C_보통 / D_주의 / E_위험 / F_판단보류
# ============================================================

def judge_financial(financial_rows):

    if not financial_rows:

        return (
            "F_판단보류",
            "재무정보가 없어 재무상태를 판단하기 어렵습니다.",
        )

    rows = sorted(
        financial_rows,
        key=lambda x: x["연도"],
        reverse=True,
    )

    recent = rows[0]

    net_values_desc = [
        row["순이익"]
        for row in rows
        if row.get("순이익") is not None
    ]

    op_values_desc = [
        row["영업이익"]
        for row in rows
        if row.get("영업이익") is not None
    ]

    revenue_values_desc = [
        row["매출액"]
        for row in rows
        if row.get("매출액") is not None
    ]

    recent_net = recent.get("순이익")
    recent_op = recent.get("영업이익")
    recent_revenue = recent.get("매출액")
    recent_assets = recent.get("총자산")
    recent_equity = recent.get("자본총계")

    if (
        recent_net is None
        and recent_op is None
        and recent_revenue is None
        and recent_assets is None
        and recent_equity is None
    ):

        return (
            "F_판단보류",
            "매출·이익·자산 자료가 없어 재무상태를 판단하기 어렵습니다.",
        )

    score = 0
    revenue_trend = "판단불가"

    # --------------------------------------------------------
    # 최근 영업이익 (본업 수익성 — 가중치 상향)
    # --------------------------------------------------------

    if recent_op is not None:

        if recent_op > 0:
            score += 3

        elif recent_op < 0:
            score -= 3

    # --------------------------------------------------------
    # 최근 순이익
    # --------------------------------------------------------

    if recent_net is not None:

        if recent_net > 0:
            score += 2

        elif recent_net < 0:
            score -= 2

    # --------------------------------------------------------
    # 규모·자본 여력 (총자산 / 자본총계)
    # 단위는 보통 백만원 기준
    # --------------------------------------------------------

    if recent_equity is not None:

        if recent_equity > 0:
            score += 1

        else:
            # 자본총계 음수 = 자본 잠식
            score -= 3

    if recent_assets is not None:

        # 총자산 50억 이상
        if recent_assets >= 5000:
            score += 1

        # 총자산 100억 이상
        if recent_assets >= 10000:
            score += 1

        # 총자산 300억 이상
        if recent_assets >= 30000:
            score += 1

    # --------------------------------------------------------
    # 영업이익 지속성
    # --------------------------------------------------------

    if len(op_values_desc) >= 3:

        positive_count = sum(
            1 for v in op_values_desc if v > 0
        )
        negative_count = sum(
            1 for v in op_values_desc if v < 0
        )

        if positive_count >= 4:
            score += 2

        elif negative_count >= 4:
            score -= 3

        elif positive_count >= 3:
            score += 1

    # --------------------------------------------------------
    # 순이익 지속성
    # --------------------------------------------------------

    if len(net_values_desc) >= 3:

        positive_count = sum(
            1 for v in net_values_desc if v > 0
        )
        negative_count = sum(
            1 for v in net_values_desc if v < 0
        )

        if positive_count >= 4:
            score += 1

        elif negative_count >= 4:
            score -= 2

    # --------------------------------------------------------
    # 매출 추세 (성장은 보너스, 급락은 감점)
    # --------------------------------------------------------

    if len(revenue_values_desc) >= 2:

        revenue_trend = calculate_trend(
            list(reversed(revenue_values_desc))
        )

        if revenue_trend == "상승":
            score += 1

        elif revenue_trend == "하락":
            score -= 1

    # --------------------------------------------------------
    # 흑자/적자 전환 및 최근 악화
    # --------------------------------------------------------

    net_transition = None
    op_transition = None
    net_decline = False

    if len(rows) >= 2:

        previous = rows[1]
        previous_net = previous.get("순이익")
        previous_op = previous.get("영업이익")

        if (
            previous_net is not None
            and recent_net is not None
        ):

            if previous_net > 0 and recent_net < 0:
                score -= 2
                net_transition = "흑자에서 적자로 전환"

            elif previous_net < 0 and recent_net > 0:
                score += 2
                net_transition = "적자에서 흑자로 전환"

            elif (
                previous_net > 0
                and recent_net > 0
                and recent_net < previous_net * 0.5
            ):
                score -= 1
                net_decline = True

        if (
            previous_op is not None
            and recent_op is not None
        ):

            if previous_op > 0 and recent_op < 0:
                score -= 3
                op_transition = "흑자에서 적자로 전환"

            elif previous_op < 0 and recent_op > 0:
                score += 2
                op_transition = "적자에서 흑자로 전환"

    # --------------------------------------------------------
    # 최근 3개년 연속 적자
    # --------------------------------------------------------

    recent_three_net_loss = False
    recent_three_op_loss = False

    if len(net_values_desc) >= 3:

        if all(v < 0 for v in net_values_desc[:3]):
            score -= 2
            recent_three_net_loss = True

    if len(op_values_desc) >= 3:

        if all(v < 0 for v in op_values_desc[:3]):
            score -= 3
            recent_three_op_loss = True

    # ========================================================
    # 등급 (보험영업 목적 — A는 소수 우량만)
    #
    # A_우수: 아래를 모두 충족
    #   1) 점수 ≥ 9
    #   2) 최근 영업이익·순이익 흑자
    #   3) 최근 3개년 영업이익 모두 흑자
    #   4) 자본총계 > 0
    #   5) 규모: 총자산 ≥ 100억 또는 자본총계 ≥ 30억
    #
    # B_양호: 최근 영업·순이익 흑자 + 자본 양수, 또는 점수 ≥ 6
    # ========================================================

    # 3개년 영업이익 연속 흑자
    recent_three_op_profit = (
        len(op_values_desc) >= 3
        and all(v > 0 for v in op_values_desc[:3])
    )

    # 규모: 총자산 100억(10000백만) 또는 자본 30억(3000백만)
    has_scale_a = (
        recent_assets is not None
        and recent_assets >= 10000
    ) or (
        recent_equity is not None
        and recent_equity >= 3000
    )

    recent_both_profit = (
        recent_op is not None
        and recent_op > 0
        and recent_net is not None
        and recent_net > 0
    )

    equity_ok = (
        recent_equity is not None
        and recent_equity > 0
    )

    if (
        score >= 9
        and recent_both_profit
        and recent_three_op_profit
        and equity_ok
        and has_scale_a
    ):

        judgment = "A_우수"

    elif (
        (
            recent_both_profit
            and equity_ok
        )
        or score >= 6
    ):

        judgment = "B_양호"

    elif score >= 1:

        judgment = "C_보통"

    elif score >= -3:

        judgment = "D_주의"

    else:

        judgment = "E_위험"

    # ========================================================
    # 코멘트 생성
    # ========================================================

    def to_eok(value):
        """백만원 단위 → 억 단위 문자열"""

        if value is None:
            return ""

        try:
            eok = float(value) / 100.0

            if abs(eok) >= 10:
                return f"{eok:.0f}억"

            if abs(eok) >= 1:
                return f"{eok:.1f}억".rstrip("0").rstrip(".")

            # 1억 미만은 백만 단위 유지
            return f"{format_number(value)}백만"

        except Exception:
            return str(value)

    comments = []
    year_count = len(rows)

    # 규모 (억 단위, 간결)
    size_bits = []

    if recent_assets is not None:
        size_bits.append(
            f"자산 {to_eok(recent_assets)}"
        )

    if recent_equity is not None:

        if recent_equity <= 0:
            size_bits.append("자본 잠식")
        else:
            size_bits.append(
                f"자본 {to_eok(recent_equity)}"
            )

    if size_bits:
        comments.append(
            ", ".join(size_bits)
        )

    # 매출 추세
    if revenue_trend == "상승":
        comments.append(
            f"{year_count}개년 매출 증가"
        )

    elif revenue_trend == "하락":
        comments.append(
            f"{year_count}개년 매출 감소"
        )

    # 영업이익
    if op_transition == "흑자에서 적자로 전환":
        comments.append("영업이익 흑자→적자 전환")

    elif op_transition == "적자에서 흑자로 전환":
        comments.append("영업이익 적자→흑자 전환")

    elif recent_three_op_loss:
        comments.append("영업이익 3년 연속 적자")

    elif recent_op is not None:

        if recent_op > 0:

            if (
                len(op_values_desc) >= 3
                and all(v > 0 for v in op_values_desc[:3])
            ):
                comments.append("영업이익 흑자 유지")
            else:
                comments.append("영업이익 흑자")

        elif recent_op < 0:
            comments.append("영업이익 적자")

    # 순이익
    if net_transition == "흑자에서 적자로 전환":
        comments.append("순이익 흑자→적자 전환")

    elif recent_three_net_loss:
        comments.append("순이익 3년 연속 적자")

    elif net_decline and net_transition is None:
        comments.append("순이익 전년 대비 급감")

    elif recent_net is not None:

        if recent_net > 0:
            comments.append("순이익 흑자")

        elif recent_net < 0:
            comments.append("순이익 적자")

    if not comments:
        comments.append("재무 추세 판단 어려움")

    # 보험료 지불여력 (등급 기준 간결 표현)
    if judgment == "A_우수":
        pay_ability = "보험료 지불여력 양호"
    elif judgment == "B_양호":
        pay_ability = "보험료 지불여력 보통 이상"
    elif judgment == "C_보통":
        pay_ability = "보험료 지불여력 제한적"
    elif judgment == "D_주의":
        pay_ability = "보험료 지불여력 주의"
    elif judgment == "E_위험":
        pay_ability = "보험료 지불여력 부족 가능"
    else:
        pay_ability = "보험료 지불여력 판단 보류"

    body = ". ".join(comments[:3])

    if body and not body.endswith("."):
        body += "."

    comment = f"{body} {pay_ability}."

    if len(comment) > 180:
        comment = comment[:177] + "..."

    return (
        judgment,
        comment,
    )


# ============================================================
# 상세 페이지 → 검색 페이지 복귀
# ============================================================

def is_info_provision_blocked(page):

    """정보제공중지/거부 안내 여부"""

    try:

        text = normalize_text(
            page.locator("body").inner_text()
        )

        keywords = [
            "정보제공중지 안내",
            "정보제공중지",
            "정보가 공개되지",
            "정보제공 중지",
            "정보 제공 중지",
            "정보제공거부",
            "정보제공 거부",
            "정보 제공 거부",
            "제공할 수 없",
            "열람할 수 없",
            "제공이 중지",
            "제공중지",
        ]

        return any(k in text for k in keywords)

    except Exception:
        return False


def has_search_result_list(page, timeout_ms=3000):

    try:

        page.wait_for_selector(
            "ul.search-result__list > li",
            timeout=timeout_ms,
        )
        return True

    except Exception:
        return False


def force_restore_search_list(
    page,
    search_url="",
    target_page=1,
):

    """검색 URL + 페이지로 강제 복구 (빠른 경로)"""

    if not search_url:
        print("    ⚠ 검색 URL이 없어 강제 복구 불가")
        return False

    print(
        f"    → 검색목록 강제 복구 "
        f"(URL + {target_page}페이지)"
    )

    try:
        page.goto(
            search_url,
            wait_until="domcontentloaded",
            timeout=20000,
        )
    except Exception as e:
        print(f"    검색 URL 이동 실패: {e}")
        return False

    try:
        page.wait_for_selector(
            "ul.search-result__list > li",
            timeout=8000,
        )
    except Exception:
        print("    ⚠ 강제 복구 후에도 검색목록 없음")
        return False

    page.wait_for_timeout(random.randint(80, 150))

    # 정보제공중지 등만 처리 (일반 닫기 안 함)
    dismiss_blocking_popups(page, context="강제복구")

    if target_page and target_page > 1:
        ok = _goto_page_number(page, target_page)
        if ok:
            print(f"    ✓ {target_page}페이지 복구 완료")
        else:
            print(
                f"    ⚠ {target_page}페이지 맞춤 실패 "
                f"(현재 목록에서 계속)"
            )

    return True


def ensure_search_list(page, search_url="", target_page=1):

    """목록 없으면 강제 복구. 정상 목록이면 즉시 통과."""

    # 목록이 보이면 바로 OK (불필요한 dismiss/대기 제거)
    if has_search_result_list(page, timeout_ms=800):

        active = _get_active_page_number(page)

        if (
            target_page
            and active
            and active != target_page
        ):
            _goto_page_number(page, target_page)

        return True

    if is_info_provision_blocked(page):
        dismiss_blocking_popups(
            page, context="정보제공중지"
        )

    return force_restore_search_list(
        page,
        search_url=search_url,
        target_page=target_page,
    )


def return_to_search(
    page,
    search_url="",
    target_page=1,
    force=False,
):

    """
    상세 → 검색목록 복귀 (빠른 경로).
    정상 시 go_back 후 목록만 확인. 팝업 닫기 호출 안 함.
    """

    if force:
        ok = force_restore_search_list(
            page,
            search_url=search_url,
            target_page=target_page,
        )
        if ok:
            print("    ✓ 검색목록 복귀 (강제)")
        return ok

    try:
        page.go_back(
            wait_until="domcontentloaded",
            timeout=10000,
        )
    except Exception as e:
        print(f"    go_back 실패: {e}")
        return force_restore_search_list(
            page,
            search_url=search_url,
            target_page=target_page,
        )

    # 목록이 빨리 보이면 끝
    if has_search_result_list(page, timeout_ms=2500):
        return True

    # 정보제공중지 모달만 닫기 시도
    dismiss_blocking_popups(page, context="복귀")

    if has_search_result_list(page, timeout_ms=1000):
        return True

    return force_restore_search_list(
        page,
        search_url=search_url,
        target_page=target_page,
    )


def _goto_page_number(page, target_page):

    """검색목록에서 특정 페이지 번호 버튼 클릭"""

    target_text = str(target_page)
    active = _get_active_page_number(page)

    if active == target_page:
        return True

    # 클릭 전 짧은 안정화만
    _wait_list_stable(page, checks=2, interval_ms=50)

    old_sig = get_current_page_signature(page)

    _scroll_to_pagination(page)
    human_delay(20, 50)

    # 1) Vue 이벤트 클릭 우선
    try:

        result = _vue_click_page_button(
            page,
            target_page,
        )

        print(
            f"    페이지 {target_text} JS 클릭: {result}"
        )

        if result and result.get("ok"):

            if _wait_for_page_number(
                page,
                target_page,
                timeout_ms=6000,
                old_signature=old_sig,
            ):
                return True

    except Exception as e:

        print(
            f"    JS 페이지 이동 오류: {e}"
        )

    # 2) Playwright
    try:

        buttons = page.locator(
            "div.pagination ul.paging button.num"
        )

        for i in range(buttons.count()):

            btn = buttons.nth(i)
            label = normalize_text(
                btn.inner_text()
            )

            if label != target_text:
                continue

            try:
                btn.click(timeout=5000)
            except Exception:
                btn.click(force=True, timeout=5000)

            if _wait_for_page_number(
                page,
                target_page,
                timeout_ms=6000,
                old_signature=old_sig,
            ):
                return True

    except Exception:
        pass

    return False


# ============================================================
# 회사 1개 처리
# ============================================================

def process_company(
    page,
    card_index,
    page_number,
    search_url="",
):

    cards = get_search_cards(page)

    if card_index >= cards.count():

        raise RuntimeError(
            f"{card_index + 1}번째 카드를 찾을 수 없습니다."
        )

    card = cards.nth(
        card_index
    )

    basic = collect_search_basic_info(
        card,
        card_index,
    )

    company_name = basic[
        "회사명"
    ]

    print(
        f"{company_name or '(회사명 미확인)'}"
    )

    # --------------------------------------------------------
    # 산업분류 제외
    # --------------------------------------------------------

    industry = basic[
        "산업분류"
    ]

    if is_excluded_industry(
        industry
    ):

        print(
            f"    ⛔ 산업분류 제외: {industry}"
        )

        return {
            **basic,
            "설립년도": "",
            "종업원수": "",
            "이메일": "",
            "재무정보": [],
            "단위": "",
            "AI_판정": "F_판단보류",
            "코멘트": (
                "산업분류에 제외 키워드가 포함되어 "
                "상세 재무분석에서 제외했습니다. "
                f"({industry})"
            ),
            "처리상태": "제외",
        }

    # --------------------------------------------------------
    # 기본값
    # --------------------------------------------------------

    result = {
        **basic,
        "설립년도": "",
        "종업원수": "",
        "이메일": "",
        "재무정보": [],
        "단위": DEFAULT_FINANCIAL_UNIT,
        "AI_판정": "F_판단보류",
        "코멘트": "",
        "처리상태": "분석완료",
    }

    try:

        # ----------------------------------------------------
        # 일반 화면
        # ----------------------------------------------------

        print(
            "    일반 화면 진입..."
        )

        click_general_from_card(
            page,
            card,
            company_name,
        )

        print(
            "    ✓ 일반 화면 진입 완료"
        )

        # ----------------------------------------------------
        # 정보제공중지 안내는 '진입 직후'에만 뜨는 경우가 많음
        # → 여기서 먼저 감지·닫고, 바로 검색리스트로 복귀
        # ----------------------------------------------------

        blocked = False

        for _ in range(3):

            page.wait_for_timeout(
                random.randint(50, 90)
            )

            if is_info_provision_blocked(page):

                blocked = True
                break

            # 모달이 보이는데 본문에 중지/거부 문구가 있으면 동일 취급
            try:

                modal_text = ""

                for sel in [
                    ".modals-container",
                    ".layer_popup",
                    ".popup",
                    "[role='dialog']",
                    ".modal",
                ]:

                    loc = page.locator(sel)

                    if loc.count() == 0:
                        continue

                    el = loc.first

                    if not el.is_visible():
                        continue

                    modal_text = normalize_text(
                        el.inner_text()
                    )

                    if modal_text:
                        break

                if modal_text and any(
                    k in modal_text
                    for k in [
                        "정보제공중지",
                        "정보제공 중지",
                        "정보 제공 중지",
                        "정보제공거부",
                        "정보제공 거부",
                        "제공 중지",
                        "제공중지",
                        "열람할 수 없",
                        "제공할 수 없",
                    ]
                ):

                    blocked = True
                    break

            except Exception:
                pass

        if blocked:

            print()
            print(
                "    ⚠ 정보제공중지 안내 팝업"
            )
            print(
                "    → 팝업 안의 [확인] 버튼을 클릭합니다."
            )

            result["AI_판정"] = "F_판단보류"
            result["코멘트"] = (
                "정보제공중지 안내로 "
                "상세 정보를 확인할 수 없습니다."
            )
            result["처리상태"] = "정보제공중지"

            if not dismiss_blocking_popups(
                page,
                context="정보제공중지",
            ):
                raise ManualInterventionTimeout(
                    "정보제공중지 팝업의 확인 버튼을 클릭하지 못했습니다."
                )

            for _ in range(30):
                if (
                    not is_info_provision_blocked(page)
                    and has_search_result_list(page, timeout_ms=100)
                ):
                    print("    ✓ 검색목록 복귀 확인")
                    return result

                page.wait_for_timeout(150)

            raise ManualInterventionTimeout(
                "정보제공중지 확인 후 검색목록 복귀를 확인하지 못했습니다."
            )

        # 중지 안내가 없을 때만 일반·재무 수집
        dismiss_blocking_popups(
            page,
            context="일반화면",
        )

        general_info = (
            extract_general_basic_info(
                page
            )
        )

        result.update(
            general_info
        )

        print(
            f"    설립년도: "
            f"{result['설립년도']}"
        )

        print(
            f"    종업원수: "
            f"{result['종업원수']}"
        )

        print(
            f"    이메일: "
            f"{result['이메일']}"
        )

        # ----------------------------------------------------
        # 재무정보
        # ----------------------------------------------------

        print(
            "    재무정보 추출 중..."
        )

        dismiss_blocking_popups(
            page,
            context="재무추출전",
        )

        (
            financial_rows,
            unit,
        ) = extract_financial_data(
            page
        )

        result[
            "재무정보"
        ] = financial_rows

        result[
            "단위"
        ] = unit

        print(
            f"    최근 재무자료: "
            f"{len(financial_rows)}개년"
        )

        for row in financial_rows:

            print(
                f"      {row['연도']} | "
                f"자산 {format_number(row.get('총자산'))} | "
                f"자본 {format_number(row.get('자본총계'))} | "
                f"매출 {format_number(row.get('매출액'))} | "
                f"영업 {format_number(row.get('영업이익'))} | "
                f"순익 {format_number(row.get('순이익'))}"
            )

        # ----------------------------------------------------
        # AI 재무판정
        # ----------------------------------------------------

        (
            judgment,
            comment,
        ) = judge_financial(
            financial_rows
        )

        result[
            "AI_판정"
        ] = judgment

        result[
            "코멘트"
        ] = comment

        print(
            f"    AI_판정: {judgment}"
        )

        print(
            f"    코멘트: {comment}"
        )

        # ----------------------------------------------------
        # 검색결과 복귀
        # ----------------------------------------------------

        return_to_search(
            page,
            search_url=search_url,
            target_page=page_number,
        )

        return result

    except Exception as e:

        print(
            f"    ⚠ 상세 수집 오류: {e}"
        )

        result[
            "AI_판정"
        ] = "F_판단보류"

        result[
            "코멘트"
        ] = (
            "일반정보 또는 재무정보 수집 과정에서 "
            "오류가 발생하여 재무상태를 판단할 수 없습니다."
        )

        result[
            "처리상태"
        ] = "오류"

        try:

            # 오류 시 go_back 대신 검색 URL로 강제 복구
            # (정보제공중지 닫으면 스마트검색 홈으로 떨어지는 문제 방지)
            return_to_search(
                page,
                search_url=search_url,
                target_page=page_number,
                force=True,
            )

        except Exception as return_error:

            print(
                f"    ⚠ 검색결과 복귀도 실패: "
                f"{return_error}"
            )

            try:

                force_restore_search_list(
                    page,
                    search_url=search_url,
                    target_page=page_number,
                )

            except Exception:
                pass

        return result


# ============================================================
# 현재 페이지 회사 처리
#
# 핵심:
# total_result_count - processed_count 만큼만 처리한다.
#
# 예:
# 전체 54개
# 페이지당 10개
# → 10 + 10 + 10 + 10 + 10 + 4
#
# 페이지당 50개
# → 50 + 4
#
# 페이지당 100개
# → 54
# ============================================================

def collect_current_page(
    page,
    page_number,
    processed_count,
    total_result_count,
    processed_companies,
    start_card_index=1,
    search_url="",
):

    cards = get_search_cards(page)

    card_count = cards.count()

    print()
    print("=" * 70)
    print(
        f"{page_number}페이지 처리"
    )
    print("=" * 70)

    print(
        f"현재 페이지 회사 카드: "
        f"{card_count}개"
    )

    # start_card_index: 1부터 시작. 이전 카드는 스킵
    start_i = max(
        0,
        int(start_card_index) - 1,
    )

    if start_i > 0:

        print(
            f"이어하기: "
            f"{start_card_index}번째 카드부터 시작 "
            f"(앞 {start_i}개 스킵)"
        )

    # --------------------------------------------------------
    # 전체 결과 대비 남은 처리량
    # --------------------------------------------------------

    remaining_count = (
        total_result_count
        - processed_count
    )

    if remaining_count <= 0:

        return (
            [],
            processed_count,
        )

    # 현재 페이지에서 실제 처리할 수 있는 회사 수
    process_count = min(
        card_count,
        remaining_count + start_i,
    )

    print(
        f"이번 페이지 실제 처리: "
        f"{max(0, process_count - start_i)}개 "
        f"(인덱스 {start_i + 1}~{process_count})"
    )

    if process_count < card_count:

        print(
            f"⚠ 전체 검색결과 상한 때문에 "
            f"{card_count - process_count}개는 처리하지 않습니다."
        )

    page_results = []

    for i in range(
        process_count
    ):

        if i < start_i:
            continue

        print()
        print(
            f"[{page_number}페이지 "
            f"{i + 1}/{process_count}] "
            f"전체 처리 "
            f"{processed_count + 1}/"
            f"{total_result_count}"
        )

        # ----------------------------------------------------
        # 중복 회사 방지
        # ----------------------------------------------------

        company_key = ""

        try:

            current_cards = (
                get_search_cards(page)
            )

            if i < current_cards.count():

                company_key = normalize_text(
                    get_company_name(
                        current_cards.nth(i),
                        i,
                    )
                )

        except Exception:
            pass

        # 회사명이 있고 이미 처리한 회사라면
        # 다시 처리하지 않는다.
        #
        # 단, 전체 처리 건수 자체는 증가시킨다.
        # 검색결과 전체 건수와 처리 위치를 맞추기 위함이다.
        if (
            company_key
            and company_key in processed_companies
        ):

            print(
                f"    ⚠ 중복 회사 발견: "
                f"{company_key}"
            )

            print(
                "    → 중복 처리를 건너뜁니다."
            )

            duplicate_result = {
                "회사명": company_key,
                "대표자명": "",
                "산업분류": "",
                "전화번호": "",
                "주소": "",
                "설립년도": "",
                "종업원수": "",
                "이메일": "",
                "재무정보": [],
                "단위": DEFAULT_FINANCIAL_UNIT,
                "AI_판정": "F_판단보류",
                "코멘트": (
                    "검색결과에서 중복 회사가 확인되어 "
                    "중복 수집을 방지했습니다."
                ),
                "처리상태": "중복",
            }

            page_results.append(
                duplicate_result
            )

            processed_count += 1

            continue

        try:

            result = process_company(
                page,
                i,
                page_number,
                search_url=search_url,
            )

            page_results.append(
                result
            )

            status = result.get(
                "처리상태",
                "",
            )

            if status == "제외":

                print(
                    "    → 제외 처리 완료"
                )

            elif status == "정보제공중지":

                print(
                    "    → 정보제공중지로 상세정보 건너뜀"
                )

            elif status == "오류":

                print(
                    "    → 오류 처리 완료"
                )

            else:

                print(
                    "    → 분석 완료"
                )

            # 회사명이 있으면 처리 목록에 등록
            result_company_name = normalize_text(
                result.get(
                    "회사명",
                    "",
                )
            )

            if result_company_name:

                processed_companies.add(
                    result_company_name
                )

        except (KeyboardInterrupt, ManualInterventionTimeout):
            raise

        except Exception as e:

            print(
                f"    ❌ 회사 처리 실패: {e}"
            )

            company_name = ""

            try:

                cards_retry = (
                    get_search_cards(page)
                )

                if i < cards_retry.count():

                    company_name = (
                        get_company_name(
                            cards_retry.nth(i),
                            i,
                        )
                    )

            except Exception:
                pass

            page_results.append({
                "회사명": company_name,
                "대표자명": "",
                "산업분류": "",
                "전화번호": "",
                "주소": "",
                "설립년도": "",
                "종업원수": "",
                "이메일": "",
                "재무정보": [],
                "단위": DEFAULT_FINANCIAL_UNIT,
                "AI_판정": "F_판단보류",
                "코멘트": (
                    "회사 처리 중 오류가 발생하여 "
                    "재무상태를 판단할 수 없습니다."
                ),
                "처리상태": "오류",
            })

            if company_name:

                processed_companies.add(
                    normalize_text(
                        company_name
                    )
                )

        processed_count += 1

        # 회사 간 짧은 랜덤 간격 (패턴 완화)
        human_delay(70, 220)

        print()
        print(
            f"현재까지 전체 처리: "
            f"{processed_count} / "
            f"{total_result_count}"
        )

        # ----------------------------------------------------
        # 전체 처리 완료
        # ----------------------------------------------------

        if (
            processed_count
            >= total_result_count
        ):

            break

    return (
        page_results,
        processed_count,
    )


# ============================================================
# 다음 페이지 이동
#
# CRETOP HTML:
#   div.pagination > ul.paging > li > button.num > span
#   현재 페이지: button.num.on
#   next/prev: button.next / button.prev  (텍스트 "다음그룹" 등)
# ============================================================

def _scroll_to_pagination(page):

    try:

        page.evaluate(
            """() => {
                const el = document.querySelector('div.pagination');
                if (el) {
                    el.scrollIntoView({block: 'center', behavior: 'instant'});
                    return true;
                }
                window.scrollTo(0, document.body.scrollHeight);
                return false;
            }"""
        )
        page.wait_for_timeout(
            random.randint(250, 450)
        )

    except Exception:
        pass


def _get_active_page_number(page):

    try:

        active = page.locator(
            "div.pagination ul.paging button.num.on"
        )

        if active.count() == 0:
            return None

        text = normalize_text(
            active.first.inner_text()
        )

        if text.isdigit():
            return int(text)

    except Exception:
        pass

    return None


def _wait_list_stable(page, checks=2, interval_ms=50):

    """
    목록 시그니처가 연속 동일하면 렌더 완료로 판단.
    기본 짧게 (약 0.2~0.4초). 페이지 클릭 직전에만 사용.
    """

    prev = None
    same = 0

    for _ in range(checks + 3):

        try:

            page.wait_for_selector(
                "ul.search-result__list > li",
                timeout=2000,
            )

        except Exception:
            pass

        sig = get_current_page_signature(page)

        if sig and sig == prev:
            same += 1
            if same >= checks:
                return True
        else:
            same = 0
            prev = sig

        try:
            page.wait_for_timeout(interval_ms)
        except Exception:
            return False

    return prev is not None


def _wait_for_page_number(
    page,
    target_page,
    timeout_ms=10000,
    old_signature=None,
):

    """
    이동 성공 판정:
    1) button.num.on == 목표 페이지
    2) old_signature 가 있으면 목록 시그니처 변경
    성공 후 짧은 정착만 (풀 안정화 중복 제거)
    """

    target_text = str(target_page)
    deadline = time.time() + (timeout_ms / 1000.0)

    while time.time() < deadline:

        try:

            active = page.locator(
                "div.pagination ul.paging button.num.on"
            )

            active_ok = False

            if active.count() > 0:

                label = normalize_text(
                    active.first.inner_text()
                )
                active_ok = (label == target_text)

            if not active_ok:

                try:
                    page.wait_for_timeout(150)
                except Exception:
                    return False

                continue

            if old_signature:

                new_sig = get_current_page_signature(
                    page
                )

                if (
                    not new_sig
                    or new_sig == old_signature
                ):

                    try:
                        page.wait_for_timeout(150)
                    except Exception:
                        return False

                    continue

            # 성공 후 짧은 정착만
            try:
                page.wait_for_timeout(
                    random.randint(100, 200)
                )
            except Exception:
                pass

            return True

        except Exception:
            return False

    return False


def _vue_click_page_button(page, target_page):

    """
    Vue 대응: MouseEvent/PointerEvent 버블링 클릭
    반환: 클릭 성공 여부 + 방법
    """

    return page.evaluate(
        """(targetPage) => {
            const target = String(targetPage);
            const root = document.querySelector('div.pagination');
            if (!root) {
                return {ok: false, reason: 'no_pagination'};
            }

            const buttons = root.querySelectorAll('ul.paging button.num');
            let targetBtn = null;

            for (const btn of buttons) {
                const t = (btn.innerText || '').replace(/\\s+/g, '').trim();
                if (t === target) {
                    targetBtn = btn;
                    break;
                }
            }

            if (!targetBtn) {
                return {
                    ok: false,
                    reason: 'no_button',
                    found: Array.from(buttons).map(b => (b.innerText || '').trim())
                };
            }

            const tokens = (targetBtn.className || '').toString().split(/\\s+/);
            if (tokens.includes('on')) {
                return {ok: false, reason: 'already_on'};
            }
            if (targetBtn.disabled) {
                return {ok: false, reason: 'disabled'};
            }

            targetBtn.scrollIntoView({block: 'center', behavior: 'instant'});

            function fire(el, type, init) {
                el.dispatchEvent(new MouseEvent(type, Object.assign({
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    buttons: 1
                }, init || {})));
            }

            // pointer + mouse 시퀀스 (Vue/모던 UI)
            try {
                targetBtn.dispatchEvent(new PointerEvent('pointerdown', {
                    bubbles: true, cancelable: true, view: window, pointerId: 1, pointerType: 'mouse', isPrimary: true
                }));
            } catch (e) {}

            fire(targetBtn, 'mousedown');
            fire(targetBtn, 'mouseup');
            fire(targetBtn, 'click');

            try {
                targetBtn.click();
            } catch (e) {}

            return {ok: true, reason: 'clicked', text: target};
        }""",
        target_page,
    )


def go_to_next_page(
    page,
    current_page_number,
    search_url="",
):

    target_page = current_page_number + 1
    target_text = str(target_page)

    print()
    print(
        f"{current_page_number}페이지 "
        f"→ {target_page}페이지 이동"
    )

    # 만료/목록 깨짐 복구
    if is_page_expired(page) or page.locator(
        "div.pagination"
    ).count() == 0:

        print(
            "    이동 전 목록 상태 복구 시도..."
        )
        ensure_search_list(
            page,
            search_url=search_url,
            target_page=current_page_number,
        )

    dismiss_blocking_popups(
        page,
        context="페이지이동전",
    )

    # 클릭 전 짧은 안정화만 (2→3 무시 방지)
    _wait_list_stable(page, checks=2, interval_ms=50)

    _scroll_to_pagination(page)

    old_sig = get_current_page_signature(
        page
    )

    # 현재 활성 페이지 로그
    active_now = _get_active_page_number(page)
    print(
        f"    현재 활성 페이지 표시: {active_now}"
    )
    print(
        f"    이동 전 시그니처: {old_sig}"
    )

    # pagination 존재 여부
    try:

        pag_count = page.locator(
            "div.pagination"
        ).count()
        num_count = page.locator(
            "div.pagination ul.paging button.num"
        ).count()
        print(
            f"    pagination={pag_count}, "
            f"num버튼={num_count}"
        )

        if num_count > 0:

            labels = []

            for i in range(num_count):

                try:
                    labels.append(
                        normalize_text(
                            page.locator(
                                "div.pagination ul.paging button.num"
                            ).nth(i).inner_text()
                        )
                    )
                except Exception:
                    labels.append("?")

            print(
                f"    번호 목록: {labels}"
            )

    except Exception as e:

        print(
            f"    pagination 상태 확인 오류: {e}"
        )

    # --------------------------------------------------------
    # 1) Playwright 로케이터 클릭
    # --------------------------------------------------------

    print(
        f"    [1] Playwright button.num '{target_text}' 클릭..."
    )

    try:

        buttons = page.locator(
            "div.pagination ul.paging button.num"
        )

        for i in range(buttons.count()):

            btn = buttons.nth(i)
            label = normalize_text(
                btn.inner_text()
            )

            if label != target_text:
                continue

            cls_tokens = (
                btn.get_attribute("class")
                or ""
            ).split()

            if "on" in cls_tokens:
                print(
                    "    이미 현재 페이지로 표시됨"
                )
                return True

            human_delay(50, 150)

            try:
                btn.scroll_into_view_if_needed()
            except Exception:
                pass

            human_delay(40, 100)

            # 일반 클릭 → force → span 클릭
            clicked = False

            for attempt in range(3):

                try:

                    if attempt == 0:
                        btn.click(timeout=5000)
                    elif attempt == 1:
                        btn.click(
                            force=True,
                            timeout=5000,
                        )
                    else:
                        span = btn.locator("span")
                        if span.count() > 0:
                            span.first.click(
                                force=True,
                                timeout=5000,
                            )
                        else:
                            btn.click(
                                force=True,
                                timeout=5000,
                            )

                    clicked = True
                    break

                except Exception as e:

                    print(
                        f"    클릭 시도 {attempt + 1} 실패: {e}"
                    )

            if not clicked:
                continue

            print(
                f"    클릭 완료 → 활성 페이지 {target_text} 대기"
            )

            if _wait_for_page_number(
                page,
                target_page,
                timeout_ms=6000,
                old_signature=old_sig,
            ):

                print(
                    f"✓ {target_page}페이지 이동 완료 (Playwright)"
                )
                return True

            print(
                "    → 활성 페이지 미변경"
            )

    except Exception as e:

        print(
            f"    Playwright 클릭 오류: {e}"
        )

    # --------------------------------------------------------
    # 2) Vue 이벤트 시퀀스 JS 클릭
    # --------------------------------------------------------

    print(
        "    [2] Vue MouseEvent 클릭..."
    )

    _scroll_to_pagination(page)
    human_delay(50, 120)

    try:

        result = _vue_click_page_button(
            page,
            target_page,
        )
        print(
            f"    JS 결과: {result}"
        )

        if result and result.get("ok"):

            if _wait_for_page_number(
                page,
                target_page,
                timeout_ms=6000,
                old_signature=old_sig,
            ):

                print(
                    f"✓ {target_page}페이지 이동 완료 (Vue event)"
                )
                return True

            print(
                "    → Vue 클릭 후 활성 페이지 미변경"
            )

        elif result and result.get("reason") == "already_on":

            print(
                f"✓ 이미 {target_page}페이지"
            )
            return True

    except Exception as e:

        print(
            f"    Vue 클릭 오류: {e}"
        )

    # --------------------------------------------------------
    # 3) 한 번 더 force + 긴 대기
    # --------------------------------------------------------

    print(
        "    [3] 최종 재시도..."
    )

    _scroll_to_pagination(page)
    human_delay(80, 160)

    try:

        btn = page.locator(
            "div.pagination ul.paging button.num",
            has_text=target_text,
        )

        # exact: has_text can match 12 when looking for 1 — filter
        matched = None

        for i in range(btn.count()):

            el = btn.nth(i)
            if normalize_text(el.inner_text()) == target_text:
                matched = el
                break

        if matched is not None:

            matched.click(force=True, timeout=5000)

            if _wait_for_page_number(
                page,
                target_page,
                timeout_ms=8000,
                old_signature=old_sig,
            ):

                print(
                    f"✓ {target_page}페이지 이동 완료 (최종)"
                )
                return True

    except Exception as e:

        print(
            f"    최종 재시도 오류: {e}"
        )

    active_after = _get_active_page_number(page)
    print(
        f"    실패 시점 활성 페이지: {active_after}"
    )
    print(
        f"❌ {target_page}페이지로 이동하지 못했습니다."
    )

    return False


# ============================================================
# Excel 공통 스타일
# ============================================================

def style_worksheet(ws):

    if ws.max_row == 0:
        return

    for cell in ws[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    for row in ws.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    if ws.max_row > 1:

        ws.auto_filter.ref = (
            ws.dimensions
        )

    ws.freeze_panes = "A2"

    for col in range(
        1,
        ws.max_column + 1,
    ):

        max_length = 0

        for row in range(
            1,
            ws.max_row + 1,
        ):

            value = ws.cell(
                row,
                col,
            ).value

            if value is None:
                continue

            length = len(
                str(value)
            )

            if length > max_length:
                max_length = length

        width = min(
            max(
                max_length + 2,
                10,
            ),
            60,
        )

        ws.column_dimensions[
            get_column_letter(col)
        ].width = width


# ============================================================
# Excel 저장
# ============================================================

def save_to_excel(
    data,
    search_url,
    page_count,
    total_result_count,
    processed_count,
    excluded_count,
    analyzed_count,
    search_title="",
):

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    safe_title = sanitize_filename(
        search_title
    )

    file_path = unique_output_path(
        OUTPUT_DIR
        / f"{safe_title}_{timestamp}.xlsx"
    )

    wb = Workbook()

    # ========================================================
    # 종합결과
    # ========================================================

    ws = wb.active

    ws.title = "종합결과"

    # AI 판정 점수 (높을수록 우선)
    JUDGMENT_RANK = {
        "A_우수": 1,
        "B_양호": 2,
        "C_보통": 3,
        "D_주의": 4,
        "E_위험": 5,
        "F_판단보류": 6,
    }

    headers = [
        "회사명",
        "AI_판정",
        "코멘트",
        "2025_매출",
        "2025_영업이익",
        "2025_순이익",
        "설립년도",
        "대표자명",
        "종업원수",
        "산업분류",
        "전화번호",
        "이메일",
        "주소",
    ]

    ws.append(headers)

    # 제외·중복 제외 후 AI 판정 높은 순 정렬
    export_rows = []

    for item in data:

        if item.get("처리상태") in ("제외", "중복"):
            continue

        financial_rows = item.get(
            "재무정보",
            [],
        )

        financial_by_year = {
            row.get("연도"): row
            for row in financial_rows
        }

        latest = financial_rows[0] if financial_rows else {}

        latest_2025 = financial_by_year.get(
            2025,
            {},
        )

        judgment = item.get("AI_판정", "F_판단보류")

        export_rows.append({
            "rank": JUDGMENT_RANK.get(judgment, 99),
            "values": [
                item.get("회사명", ""),
                judgment,
                item.get("코멘트", ""),
                latest_2025.get("매출액", latest.get("매출액", "")),
                latest_2025.get("영업이익", latest.get("영업이익", "")),
                latest_2025.get("순이익", latest.get("순이익", "")),
                item.get("설립년도", ""),
                item.get("대표자명", ""),
                item.get("종업원수", ""),
                item.get("산업분류", ""),
                item.get("전화번호", ""),
                item.get("이메일", ""),
                item.get("주소", ""),
            ],
        })

    export_rows.sort(key=lambda x: x["rank"])

    low_grade_rows = [
        row
        for row in export_rows
        if row["values"][1] in {
            "E_위험",
            "F_판단보류",
        }
    ]

    for row in export_rows:

        if row["values"][1] not in {
            "E_위험",
            "F_판단보류",
        }:
            ws.append(row["values"])

    style_worksheet(ws)

    # ========================================================
    # 재무정보
    #
    # 회사 1개 = 1행
    # 2025~2021년 가로 배치
    # ========================================================

    finance_ws = wb.create_sheet(
        "재무정보"
    )

    finance_headers = [
        "회사명",
        "단위",
        "AI_판정",
    ]

    for year in FINANCIAL_YEARS:

        finance_headers.extend([
            f"{year}_총자산",
            f"{year}_자본총계",
            f"{year}_매출",
            f"{year}_영업이익",
            f"{year}_순이익",
        ])

    finance_ws.append(
        finance_headers
    )

    finance_export = []

    for item in data:

        if item.get("처리상태") in ("제외", "중복"):
            continue

        company_name = item.get(
            "회사명",
            "",
        )

        unit = item.get(
            "단위",
            DEFAULT_FINANCIAL_UNIT,
        )

        judgment = item.get("AI_판정", "F_판단보류")

        if judgment in {
            "E_위험",
            "F_판단보류",
        }:
            continue

        financial_rows = item.get(
            "재무정보",
            [],
        )

        by_year = {
            row.get("연도"): row
            for row in financial_rows
        }

        row_values = [
            company_name,
            unit,
            judgment,
        ]

        for year in FINANCIAL_YEARS:

            year_data = by_year.get(
                year,
                {},
            )

            row_values.extend([
                year_data.get("총자산", ""),
                year_data.get("자본총계", ""),
                year_data.get("매출액", ""),
                year_data.get("영업이익", ""),
                year_data.get("순이익", ""),
            ])

        finance_export.append({
            "rank": JUDGMENT_RANK.get(judgment, 99),
            "values": row_values,
        })

    finance_export.sort(key=lambda x: x["rank"])

    for row in finance_export:
        finance_ws.append(row["values"])

    style_worksheet(
        finance_ws
    )

    # ========================================================
    # 검색정보
    # ========================================================

    info_ws = wb.create_sheet(
        "검색정보"
    )

    info_ws["A1"] = "항목"
    info_ws["B1"] = "내용"

    info_ws["A2"] = "검색 제목"
    info_ws["B2"] = search_title or safe_title

    info_ws["A3"] = "검색 실행 시간"

    info_ws["B3"] = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    info_ws["A4"] = "검색결과 URL"
    info_ws["B4"] = search_url

    info_ws["A5"] = "CRETOP 전체 검색결과"
    info_ws["B5"] = total_result_count

    info_ws["A6"] = "수집 페이지 수"
    info_ws["B6"] = page_count

    info_ws["A7"] = "전체 처리 회사 수"
    info_ws["B7"] = processed_count

    info_ws["A8"] = "재무분석 회사 수"
    info_ws["B8"] = analyzed_count

    info_ws["A9"] = "산업분류 제외 회사 수"
    info_ws["B9"] = excluded_count

    info_ws["A10"] = "오류 회사 수"

    info_ws["B10"] = sum(
        1
        for item in data
        if item.get("처리상태") == "오류"
    )

    info_ws["A11"] = "중복 회사 수"

    info_ws["B11"] = sum(
        1
        for item in data
        if item.get("처리상태") == "중복"
    )

    for cell in info_ws[1]:

        cell.font = Font(
            bold=True
        )

    info_ws.column_dimensions[
        "A"
    ].width = 30

    info_ws.column_dimensions[
        "B"
    ].width = 100

    # ========================================================
    # E~F등급 목록
    # ========================================================

    low_grade_ws = wb.create_sheet(
        "E~F등급"
    )

    low_grade_ws.append(headers)

    for row in low_grade_rows:
        low_grade_ws.append(row["values"])

    style_worksheet(
        low_grade_ws
    )

    # ========================================================
    # 전화번호 등 텍스트 형식
    # ========================================================

    for target_ws in [
        ws,
        finance_ws,
        low_grade_ws,
    ]:

        header_index = {}

        for col in range(
            1,
            target_ws.max_column + 1,
        ):

            header_index[
                target_ws.cell(
                    1,
                    col,
                ).value
            ] = col

        for column_name in [
            "전화번호",
            "사업자번호",
            "법인번호",
        ]:

            col = header_index.get(
                column_name
            )

            if col:

                for row in range(
                    2,
                    target_ws.max_row + 1,
                ):

                    target_ws.cell(
                        row,
                        col,
                    ).number_format = "@"

    # ========================================================
    # 숫자 표시 형식
    # ========================================================

    # 종합결과 숫자 컬럼

    for target_ws in [
        ws,
        low_grade_ws,
    ]:

        result_header_index = {}

        for col in range(
            1,
            target_ws.max_column + 1,
        ):

            result_header_index[
                target_ws.cell(
                    1,
                    col,
                ).value
            ] = col

        for column_name in [
            "2025_매출",
            "2025_영업이익",
            "2025_순이익",
        ]:

            col = result_header_index.get(
                column_name
            )

            if col:

                for row in range(
                    2,
                    target_ws.max_row + 1,
                ):

                    target_ws.cell(
                        row,
                        col,
                    ).number_format = "#,##0"

    # 재무정보 숫자 (AI_판정 제외)

    for col in range(
        4,
        finance_ws.max_column + 1,
    ):

        for row in range(
            2,
            finance_ws.max_row + 1,
        ):

            finance_ws.cell(
                row,
                col,
            ).number_format = "#,##0"

    # ========================================================
    # 저장
    # ========================================================

    wb.save(
        file_path
    )

    return file_path


# ============================================================
# JSON 저장
# ============================================================

def save_to_json(
    data,
    search_url,
    page_count,
    total_result_count,
    processed_count,
    excluded_count,
    analyzed_count,
    search_title="",
):

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    safe_title = sanitize_filename(
        search_title
    )

    file_path = unique_output_path(
        OUTPUT_DIR
        / f"{safe_title}_{timestamp}.json"
    )

    result = {
        "검색정보": {
            "검색제목": search_title or safe_title,
            "검색실행시간": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            "검색결과URL": search_url,
            "전체검색결과": total_result_count,
            "수집페이지수": page_count,
            "전체처리회사수": processed_count,
            "재무분석회사수": analyzed_count,
            "산업분류제외회사수": excluded_count,
            "오류회사수": sum(
                1
                for item in data
                if item.get("처리상태") == "오류"
            ),
            "중복회사수": sum(
                1
                for item in data
                if item.get("처리상태") == "중복"
            ),
        },
        "회사정보": data,
    }

    with open(
        file_path,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            result,
            f,
            ensure_ascii=False,
            indent=2,
        )

    return file_path


# ============================================================
# 메인
# ============================================================

def parse_args():

    parser = argparse.ArgumentParser(
        description=(
            "CRETOP 자동 수집 "
            "(이어서 실행: --start-page --start-card)"
        )
    )

    parser.add_argument(
        "--start-page",
        type=int,
        default=1,
        help="시작 페이지 번호 (기본 1)",
    )

    parser.add_argument(
        "--start-card",
        type=int,
        default=1,
        help="해당 페이지에서 시작할 카드 순번 1부터 (기본 1)",
    )

    parser.add_argument(
        "--processed-count",
        type=int,
        default=None,
        help=(
            "이미 처리한 회사 수 "
            "(미지정 시 페이지·순번으로 추정)"
        ),
    )

    return parser.parse_args()


def main():

    args = parse_args()

    start_page = max(1, args.start_page)
    start_card = max(1, args.start_card)

    print()
    print("=" * 70)
    print(
        "CRETOP 기업정보 + 재무정보 + AI판정 자동 수집"
    )
    print("=" * 70)

    if start_page > 1 or start_card > 1:

        print(
            f"이어하기 모드: "
            f"{start_page}페이지 "
            f"{start_card}번째 카드부터"
        )

    print()

    with sync_playwright() as p:

        # ----------------------------------------------------
        # Chrome 연결
        # ----------------------------------------------------

        print(
            "Chrome 연결 중..."
        )

        browser = (
            p.chromium.connect_over_cdp(
                "http://127.0.0.1:9222"
            )
        )

        print(
            "✓ Chrome 연결 완료"
        )

        # ----------------------------------------------------
        # Context
        # ----------------------------------------------------

        if not browser.contexts:

            print(
                "❌ Chrome context를 찾지 못했습니다."
            )

            p.stop()
            return

        context = browser.contexts[0]

        pages = context.pages

        if not pages:

            print(
                "❌ Chrome 페이지를 찾지 못했습니다."
            )

            p.stop()
            return

        # ----------------------------------------------------
        # CRETOP 탭 찾기
        # ----------------------------------------------------

        page = None

        for current_page in pages:

            try:

                if "cretop.com" in (
                    current_page.url
                ):

                    page = current_page
                    break

            except Exception:
                continue

        if page is None:
            page = pages[-1]

        print()
        print(
            "현재 페이지:"
        )

        print(
            page.url
        )

        # ----------------------------------------------------
        # 로그인 확인
        # ----------------------------------------------------

        if not check_login(page):

            p.stop()
            return

        # ----------------------------------------------------
        # 검색결과 대기
        # ----------------------------------------------------

        try:

            wait_for_search_list(
                page
            )

        except Exception:

            print()
            print(
                "❌ CRETOP 검색결과 목록을 찾지 못했습니다."
            )

            print(
                f"현재 URL: {page.url}"
            )

            p.stop()
            return

        # ----------------------------------------------------
        # 전체 결과 수
        # ----------------------------------------------------

        print()
        print(
            "전체 검색결과 개수 확인 중..."
        )

        total_result_count = (
            get_total_result_count(
                page
            )
        )

        if total_result_count <= 0:

            print()
            print(
                "❌ 전체 검색결과 개수를 찾지 못했습니다."
            )

            print(
                "안전을 위해 작업을 중단합니다."
            )

            p.stop()
            return

        print()
        print("=" * 70)
        print(
            f"CRETOP 전체 검색결과: "
            f"{total_result_count}개"
        )
        print("=" * 70)

        # 검색결과 URL 저장 (만료 시 복구용)
        search_url = page.url
        print(
            f"검색 URL 저장: {search_url}"
        )

        # ----------------------------------------------------
        # 검색 제목
        # ----------------------------------------------------

        search_title = get_search_title(
            page
        )

        if search_title:

            print(
                f"검색 제목: {search_title}"
            )

        else:

            print(
                "검색 제목을 찾지 못해 기본 파일명을 사용합니다."
            )

        # ----------------------------------------------------
        # 전체 결과
        # ----------------------------------------------------

        all_data = []

        excluded_count = 0
        analyzed_count = 0

        # 중복 회사 방지용
        processed_companies = set()

        current_page_number = 1

        # 이어하기: 시작 페이지로 한 번에 이동 (1→2→3 순차 X)
        if start_page > 1:

            print()
            print(
                f"시작 페이지 {start_page}로 이동 중..."
            )

            old_sig = get_current_page_signature(
                page
            )

            moved = _goto_page_number(
                page,
                start_page,
            )

            if not moved:

                # 순차 이동 폴백
                print(
                    "직접 이동 실패 → 순차 이동 시도"
                )

                while current_page_number < start_page:

                    moved = go_to_next_page(
                        page,
                        current_page_number,
                        search_url=search_url,
                    )

                    if not moved:

                        print(
                            f"❌ {start_page}페이지로 "
                            f"이동하지 못했습니다. "
                            f"(현재 {current_page_number})"
                        )
                        print(
                            "검색결과에서 해당 페이지를 "
                            "직접 연 뒤 다시 실행하세요."
                        )
                        return

                    current_page_number += 1

            else:

                current_page_number = start_page

            # 최종 확인
            active = _get_active_page_number(page)

            if active:
                current_page_number = active

            print(
                f"✓ {current_page_number}페이지 도착"
            )

        # 이미 처리한 건수 (상한 계산용)
        if args.processed_count is not None:

            processed_count = max(
                0,
                args.processed_count,
            )

        else:

            # 대략 추정: (페이지-1)*현재카드수 + (카드-1)
            cards_now = get_search_cards(
                page
            )
            per_page = max(
                cards_now.count(),
                1,
            )
            processed_count = (
                (current_page_number - 1)
                * per_page
            ) + (start_card - 1)

        print(
            f"처리 카운트 시작값: "
            f"{processed_count}"
        )

        # ----------------------------------------------------
        # 페이지 반복
        # ----------------------------------------------------

        interrupted = False
        excel_path = None
        json_path = None

        try:

          while True:

            # 로그인 만료 감지
            if not check_login(page):

                print()
                print(
                    "❌ 로그인이 풀린 것으로 보입니다."
                )
                print(
                    "다시 로그인한 뒤 같은 검색을 열고 "
                    "아래처럼 이어서 실행하세요:"
                )
                print(
                    f"  python ct.py "
                    f"--start-page {current_page_number} "
                    f"--start-card 1 "
                    f"--processed-count {processed_count}"
                )
                break

            dismiss_blocking_popups(
                page,
                context="목록",
            )

            cards = get_search_cards(
                page
            )

            card_count = cards.count()

            if card_count == 0:

                print(
                    "❌ 현재 페이지에 회사 카드가 없습니다."
                )

                break

            # ------------------------------------------------
            # 현재 페이지 처리
            # ------------------------------------------------

            # 이어하기 시작 페이지만 start_card 적용, 이후는 1부터
            page_start_card = (
                start_card
                if current_page_number == start_page
                else 1
            )

            (
                page_data,
                processed_count,
            ) = collect_current_page(
                page,
                current_page_number,
                processed_count,
                total_result_count,
                processed_companies,
                start_card_index=page_start_card,
                search_url=search_url,
            )

            all_data.extend(
                page_data
            )

            # ------------------------------------------------
            # 통계
            # ------------------------------------------------

            excluded_count = sum(
                1
                for item in all_data
                if item.get("처리상태")
                == "제외"
            )

            # 오류 회사는 재무분석 완료 회사로 세지 않음
            analyzed_count = sum(
                1
                for item in all_data
                if item.get("처리상태")
                == "분석완료"
            )

            print()
            print("=" * 70)

            print(
                f"{current_page_number}페이지 처리 완료"
            )

            print(
                f"전체 처리: "
                f"{processed_count} / "
                f"{total_result_count}"
            )

            print(
                f"재무분석: "
                f"{analyzed_count}개"
            )

            print(
                f"산업분류 제외: "
                f"{excluded_count}개"
            )

            print("=" * 70)

            # ------------------------------------------------
            # 전체 처리 완료
            # ------------------------------------------------

            if (
                processed_count
                >= total_result_count
            ):

                print()
                print(
                    "✓ CRETOP 전체 검색결과 처리가 완료되었습니다."
                )

                break

            target_page = current_page_number + 1
            old_signature = get_current_page_signature(page)
            deadline = time.monotonic() + MANUAL_PAGE_CHANGE_TIMEOUT_SECONDS

            print(
                f"현재 페이지 처리를 완료했습니다. Chrome에서 직접 "
                f"{target_page}페이지로 이동해 주세요. "
                f"최대 {MANUAL_PAGE_CHANGE_TIMEOUT_SECONDS}초 기다립니다."
            )

            moved_manually = False

            while time.monotonic() < deadline:

                if is_page_expired(page):
                    print("페이지 만료를 감지했습니다. 현재 결과를 저장하고 종료합니다.")
                    interrupted = True
                    break

                active_page = _get_active_page_number(page)
                page_signature = get_current_page_signature(page)

                if (
                    active_page == target_page
                    and page_signature
                    and page_signature != old_signature
                ):
                    if _wait_list_stable(page, checks=2, interval_ms=150):
                        current_page_number = target_page
                        moved_manually = True
                        print(f"✓ 수동으로 {target_page}페이지 이동 확인")
                        break

                page.wait_for_timeout(300)

            if interrupted:
                break

            if moved_manually:
                continue

            print(
                f"{MANUAL_PAGE_CHANGE_TIMEOUT_SECONDS}초 동안 "
                f"{target_page}페이지 이동이 확인되지 않아 저장 후 종료합니다."
            )
            interrupted = True
            break

        except ManualInterventionTimeout as pause_err:

            interrupted = True
            print()
            print(
                f"⚠ 수동 처리 필요: {pause_err}"
            )
            print(
                "지금까지 수집한 데이터를 저장합니다."
            )
            print(
                "팝업을 닫고 검색리스트를 연 뒤 "
                "이어하기 하세요."
            )
            print(
                f"  이어하기 예: "
                f"python ct.py "
                f"--start-page {current_page_number} "
                f"--start-card 1 "
                f"--processed-count {processed_count}"
            )

        except KeyboardInterrupt:

            interrupted = True
            print()
            print(
                "⚠ 사용자 중단 (Ctrl+C) — "
                "지금까지 수집한 데이터를 저장합니다."
            )
            print(
                f"  이어하기 예: "
                f"python ct.py "
                f"--start-page {current_page_number} "
                f"--start-card 1 "
                f"--processed-count {processed_count}"
            )

        except Exception as loop_error:

            interrupted = True
            print()
            print(
                f"⚠ 수집 중 오류: {loop_error}"
            )
            print(
                "지금까지 수집한 데이터를 저장합니다."
            )

        # ====================================================
        # 최종 저장 (정상 종료 / 중단 / 오류 모두)
        # ====================================================

        excluded_count = sum(
            1
            for item in all_data
            if item.get("처리상태") == "제외"
        )
        analyzed_count = sum(
            1
            for item in all_data
            if item.get("처리상태") == "분석완료"
        )

        if not all_data:

            print()
            print(
                "저장할 수집 데이터가 없습니다."
            )

        else:

            print()
            print("=" * 70)

            if interrupted:

                print(
                    "중간 결과 저장 중..."
                )

            else:

                print(
                    "최종 결과 저장 중..."
                )

            print("=" * 70)

            try:

                save_url = search_url

                try:
                    save_url = page.url
                except Exception:
                    pass

                excel_path = save_to_excel(
                    all_data,
                    save_url,
                    current_page_number,
                    total_result_count,
                    processed_count,
                    excluded_count,
                    analyzed_count,
                    search_title,
                )

                json_path = save_to_json(
                    all_data,
                    save_url,
                    current_page_number,
                    total_result_count,
                    processed_count,
                    excluded_count,
                    analyzed_count,
                    search_title,
                )

                print(
                    f"✓ Excel: {excel_path}"
                )
                print(
                    f"✓ JSON: {json_path}"
                )

            except Exception as save_error:

                print(
                    f"❌ 저장 실패: {save_error}"
                )

        # ====================================================
        # 최종 통계
        # ====================================================

        error_count = sum(
            1
            for item in all_data
            if item.get("처리상태")
            == "오류"
        )

        duplicate_count = sum(
            1
            for item in all_data
            if item.get("처리상태")
            == "중복"
        )

        print()
        print("=" * 70)
        print(
            "CRETOP 자동 수집 완료"
        )
        print("=" * 70)

        print()
        print(
            f"CRETOP 전체 검색결과 : "
            f"{total_result_count}개"
        )

        print(
            f"전체 처리             : "
            f"{processed_count}개"
        )

        print(
            f"재무분석 대상         : "
            f"{analyzed_count}개"
        )

        print(
            f"산업분류 제외         : "
            f"{excluded_count}개"
        )

        print(
            f"처리 오류             : "
            f"{error_count}개"
        )

        print(
            f"중복 발견             : "
            f"{duplicate_count}개"
        )

        print(
            f"수집 페이지           : "
            f"{current_page_number}페이지"
        )

        if excel_path:

            print()
            print("Excel:")
            print(excel_path)

        if json_path:

            print()
            print("JSON:")
            print(json_path)

        print()
        print("=" * 70)

        # ----------------------------------------------------
        # Chrome은 종료하지 않는다.
        # ----------------------------------------------------

        p.stop()


# ============================================================
# 실행
# ============================================================

if __name__ == "__main__":
    main()