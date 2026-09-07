"""
CRETOP 팝업/안내창 감지·닫기 테스트

대상
1) 정보제공중지 안내 팝업 (일반상세 진입 직후)
2) 재로그인 안내 팝업 (다른 사용자 로그인 시 세션 밀림)
3) 페이지 완료/만료 등 전체 화면 안내 (팝업이 아니라 화면 전환)

사용법
  1. start_chrome.ps1 로 Chrome 실행 후 CRETOP 로그인·검색
  2. 테스트할 화면으로 이동 (또는 이 스크립트 실행 후 수동으로 팝업 유도)
  3. python test_popups.py

옵션
  python test_popups.py --watch          # 계속 감시 (팝업 뜨면 덤프)
  python test_popups.py --once           # 한 번만 스캔
  python test_popups.py --try-close      # 감지되면 닫기 시도까지
  python test_popups.py --out dumps      # HTML 덤프 폴더
"""

from __future__ import annotations

import argparse
import json
import re
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright


CDP_URL = "http://127.0.0.1:9222"
BASE_DIR = Path(__file__).resolve().parent


# ---------------------------------------------------------------------------
# 키워드
# ---------------------------------------------------------------------------

INFO_STOP_KEYWORDS = [
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

RELOGIN_KEYWORDS = [
    "재로그인",
    "다시 로그인",
    "로그인 세션",
    "세션이 만료",
    "세션 만료",
    "중복 로그인",
    "다른 곳",
    "다른 기기",
    "다른 사용자",
    "강제 로그아웃",
    "로그인되어",
    "로그인이 필요",
]

PAGE_NOTICE_KEYWORDS = [
    "웹페이지가 만료",
    "페이지가 만료",
    "페이지 만료",
    "요청하신 페이지",
    "완료되었습니다",
    "작업이 완료",
    "더 이상 유효하지",
    "세션이 종료",
]


def normalize(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r"\s+", " ", str(text))
    return text.strip()


def contains_any(text: str, keywords: list[str]) -> list[str]:
    text = normalize(text)
    return [k for k in keywords if k in text]


# ---------------------------------------------------------------------------
# 스냅샷
# ---------------------------------------------------------------------------

def snapshot_page(page) -> dict:
    """현재 URL, 제목, 목록 여부, 본문 일부"""

    info = {
        "time": datetime.now().isoformat(timespec="seconds"),
        "url": "",
        "title": "",
        "has_search_list": False,
        "search_card_count": 0,
        "active_page": None,
        "body_preview": "",
        "body_len": 0,
    }

    try:
        info["url"] = page.url
    except Exception as e:
        info["url_error"] = str(e)

    try:
        info["title"] = page.title()
    except Exception:
        pass

    try:
        cards = page.locator("ul.search-result__list > li")
        info["search_card_count"] = cards.count()
        info["has_search_list"] = info["search_card_count"] > 0
    except Exception:
        pass

    try:
        active = page.locator(
            "div.pagination ul.paging button.num.on"
        )
        if active.count() > 0:
            info["active_page"] = normalize(active.first.inner_text())
    except Exception:
        pass

    try:
        body = normalize(page.locator("body").inner_text())
        info["body_len"] = len(body)
        info["body_preview"] = body[:800]
    except Exception:
        pass

    return info


def list_candidate_roots(page) -> list[dict]:
    """모달/레이어 후보 DOM 정보"""

    return page.evaluate(
        """() => {
            const sels = [
                '.modals-container',
                '[role=dialog]',
                '.layer_popup',
                '.popup',
                '.modal',
                '.info-toast',
                '.alert',
                '[class*=modal]',
                '[class*=popup]',
                '[class*=layer]'
            ];
            const seen = new Set();
            const out = [];

            for (const sel of sels) {
                for (const el of document.querySelectorAll(sel)) {
                    if (seen.has(el)) continue;
                    seen.add(el);

                    // 사이드 퀵메뉴 제외 (참고용으로 flag 만)
                    const inQuick = !!el.closest('.quick-wrap, .quick-view');

                    const style = window.getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    const visible = (
                        style.display !== 'none'
                        && style.visibility !== 'hidden'
                        && style.opacity !== '0'
                        && r.width >= 10
                        && r.height >= 10
                    );

                    const text = (el.innerText || '').replace(/\\s+/g, ' ').trim();
                    if (!visible && text.length === 0) continue;

                    const buttons = Array.from(el.querySelectorAll('button, a.btn, [role=button]'))
                        .slice(0, 20)
                        .map(b => ({
                            tag: b.tagName,
                            className: (b.className || '').toString().slice(0, 120),
                            text: (b.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 80),
                            aria: b.getAttribute('aria-label') || '',
                            title: b.getAttribute('title') || '',
                        }));

                    out.push({
                        selector_hint: sel,
                        in_quick_panel: inQuick,
                        visible,
                        display: style.display,
                        visibility: style.visibility,
                        opacity: style.opacity,
                        rect: {
                            x: Math.round(r.x),
                            y: Math.round(r.y),
                            w: Math.round(r.width),
                            h: Math.round(r.height),
                        },
                        text_len: text.length,
                        text_preview: text.slice(0, 400),
                        html_preview: (el.outerHTML || '').slice(0, 1500),
                        buttons,
                    });
                }
            }
            return out;
        }"""
    )


def classify_state(snap: dict, roots: list[dict]) -> dict:
    """어떤 종류의 안내인지 분류"""

    body = snap.get("body_preview", "")
    full_hits = {
        "info_stop": contains_any(body, INFO_STOP_KEYWORDS),
        "relogin": contains_any(body, RELOGIN_KEYWORDS),
        "page_notice": contains_any(body, PAGE_NOTICE_KEYWORDS),
    }

    modal_hits = []

    for root in roots:
        if root.get("in_quick_panel"):
            continue
        if not root.get("visible"):
            continue

        t = root.get("text_preview", "")
        hits = {
            "info_stop": contains_any(t, INFO_STOP_KEYWORDS),
            "relogin": contains_any(t, RELOGIN_KEYWORDS),
            "page_notice": contains_any(t, PAGE_NOTICE_KEYWORDS),
        }
        if any(hits.values()):
            modal_hits.append(
                {
                    "rect": root.get("rect"),
                    "text_preview": t[:200],
                    "buttons": root.get("buttons"),
                    "hits": {k: v for k, v in hits.items() if v},
                }
            )

    # 전체 화면 안내: 검색목록 없고, 본문에 notice 키워드, 작은 모달 없음
    is_full_page_notice = (
        not snap.get("has_search_list")
        and bool(full_hits["page_notice"] or full_hits["relogin"])
        and len(modal_hits) == 0
        and "ETSS" not in (snap.get("url") or "")
    )

    kinds = []
    if any(m["hits"].get("info_stop") for m in modal_hits) or full_hits["info_stop"]:
        kinds.append("정보제공중지_팝업_가능")
    if any(m["hits"].get("relogin") for m in modal_hits) or full_hits["relogin"]:
        kinds.append("재로그인_안내_가능")
    if is_full_page_notice or full_hits["page_notice"]:
        kinds.append("전체화면_안내_가능")

    return {
        "kinds": kinds or ["해당없음_또는_미분류"],
        "body_keyword_hits": {k: v for k, v in full_hits.items() if v},
        "modal_matches": modal_hits,
        "is_full_page_notice": is_full_page_notice,
    }


# ---------------------------------------------------------------------------
# 닫기 시도 (테스트용 — 성공/실패와 전후 상태만 기록)
# ---------------------------------------------------------------------------

def try_close_strategies(page, out_dir: Path, tag: str) -> list[dict]:
    """여러 닫기 전략을 순차 시도하고 결과 기록"""

    results = []

    strategies = [
        (
            "modals-container_확인",
            "() => {\n"
            "  const root = document.querySelector('.modals-container');\n"
            "  if (!root) return {ok:false, reason:'no root'};\n"
            "  for (const b of root.querySelectorAll('button')) {\n"
            "    const t = (b.innerText||'').trim();\n"
            "    if (t === '확인' || t === '닫기' || t === 'OK') {\n"
            "      b.click(); return {ok:true, text:t};\n"
            "    }\n"
            "  }\n"
            "  return {ok:false, reason:'no button'};\n"
            "}",
        ),
        (
            "modals-container_X",
            "() => {\n"
            "  const root = document.querySelector('.modals-container');\n"
            "  if (!root) return {ok:false, reason:'no root'};\n"
            "  const el = root.querySelector('button.btn.ico, i.close, i.close-layer-24, [class*=close]');\n"
            "  if (!el) return {ok:false, reason:'no x'};\n"
            "  const btn = el.tagName === 'BUTTON' ? el : el.closest('button');\n"
            "  if (!btn) return {ok:false, reason:'no btn'};\n"
            "  btn.click(); return {ok:true, text:'X'};\n"
            "}",
        ),
        (
            "role_dialog_확인",
            "() => {\n"
            "  const root = document.querySelector('[role=dialog], .layer_popup');\n"
            "  if (!root) return {ok:false, reason:'no root'};\n"
            "  for (const b of root.querySelectorAll('button')) {\n"
            "    const t = (b.innerText||'').trim();\n"
            "    if (['확인','닫기','OK','로그인','재로그인'].includes(t)) {\n"
            "      b.click(); return {ok:true, text:t};\n"
            "    }\n"
            "  }\n"
            "  return {ok:false, reason:'no button'};\n"
            "}",
        ),
        (
            "본문_링크_로그인",
            "() => {\n"
            "  for (const a of document.querySelectorAll('a, button')) {\n"
            "    const t = (a.innerText||'').trim();\n"
            "    if (/로그인|재로그인|확인|닫기/.test(t)) {\n"
            "      const r = a.getBoundingClientRect();\n"
            "      if (r.width < 5 || r.height < 5) continue;\n"
            "      // 사이드 패널 제외\n"
            "      if (a.closest('.quick-wrap, .quick-view')) continue;\n"
            "      a.click(); return {ok:true, text:t};\n"
            "    }\n"
            "  }\n"
            "  return {ok:false, reason:'no match'};\n"
            "}",
        ),
    ]

    before = snapshot_page(page)

    for name, js in strategies:
        entry = {
            "strategy": name,
            "before_url": before.get("url"),
            "js_result": None,
            "after": None,
        }

        try:
            entry["js_result"] = page.evaluate(js)
        except Exception as e:
            entry["js_result"] = {"ok": False, "error": str(e)}

        time.sleep(0.8)
        after = snapshot_page(page)
        entry["after"] = {
            "url": after.get("url"),
            "has_search_list": after.get("has_search_list"),
            "search_card_count": after.get("search_card_count"),
            "active_page": after.get("active_page"),
            "title": after.get("title"),
            "body_preview": (after.get("body_preview") or "")[:300],
        }
        results.append(entry)

        # 하나라도 성공하면 이후 전략은 참고만 (계속 시도해도 됨)
        print(f"  strategy={name} -> {entry['js_result']}")
        print(
            f"    after url={entry['after']['url']} "
            f"list={entry['after']['has_search_list']} "
            f"cards={entry['after']['search_card_count']}"
        )

        before = after

    # 덤프 저장
    dump_path = out_dir / f"{tag}_close_attempts.json"
    dump_path.write_text(
        json.dumps(results, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"  닫기 시도 결과 저장: {dump_path}")
    return results


def dump_scan(page, out_dir: Path, tag: str, try_close: bool) -> dict:
    snap = snapshot_page(page)
    roots = list_candidate_roots(page)
    classified = classify_state(snap, roots)

    report = {
        "snapshot": snap,
        "classification": classified,
        "roots_count": len(roots),
        "visible_roots": [r for r in roots if r.get("visible") and not r.get("in_quick_panel")],
        "all_roots_brief": [
            {
                "sel": r.get("selector_hint"),
                "visible": r.get("visible"),
                "in_quick": r.get("in_quick_panel"),
                "rect": r.get("rect"),
                "text_preview": (r.get("text_preview") or "")[:160],
                "buttons": r.get("buttons"),
            }
            for r in roots
        ],
    }

    # HTML 전체도 필요하면 저장
    try:
        html_path = out_dir / f"{tag}_page.html"
        html_path.write_text(page.content(), encoding="utf-8")
        report["html_file"] = str(html_path)
    except Exception as e:
        report["html_error"] = str(e)

    json_path = out_dir / f"{tag}_scan.json"
    json_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print("=" * 70)
    print(f"[{tag}] URL: {snap.get('url')}")
    print(f"  title: {snap.get('title')}")
    print(
        f"  search_list: {snap.get('has_search_list')} "
        f"cards={snap.get('search_card_count')} "
        f"active_page={snap.get('active_page')}"
    )
    print(f"  분류: {classified.get('kinds')}")
    print(f"  본문 키워드: {classified.get('body_keyword_hits')}")
    print(f"  모달 매칭 수: {len(classified.get('modal_matches') or [])}")
    for i, m in enumerate(classified.get("modal_matches") or []):
        print(f"    modal[{i}] hits={m.get('hits')} text={m.get('text_preview')!r}")
        print(f"             buttons={m.get('buttons')}")
    print(f"  전체화면 안내 추정: {classified.get('is_full_page_notice')}")
    print(f"  저장: {json_path}")
    if report.get("html_file"):
        print(f"  HTML: {report['html_file']}")

    if try_close and (
        classified.get("kinds") != ["해당없음_또는_미분류"]
        or classified.get("modal_matches")
        or classified.get("is_full_page_notice")
    ):
        print("  --- 닫기 전략 테스트 ---")
        report["close_attempts"] = try_close_strategies(
            page,
            out_dir,
            tag,
        )

    return report


def main():
    parser = argparse.ArgumentParser(
        description="CRETOP 팝업/안내창 감지 테스트"
    )
    parser.add_argument(
        "--watch",
        action="store_true",
        help="계속 감시 (팝업 유도하며 관찰)",
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="한 번만 스캔 (기본)",
    )
    parser.add_argument(
        "--try-close",
        action="store_true",
        help="감지 시 닫기 전략 시도",
    )
    parser.add_argument(
        "--interval",
        type=float,
        default=1.5,
        help="watch 간격 초 (기본 1.5)",
    )
    parser.add_argument(
        "--out",
        type=str,
        default="popup_dumps",
        help="덤프 폴더 (기본 popup_dumps)",
    )
    args = parser.parse_args()

    out_dir = BASE_DIR / args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    print("Chrome CDP 연결:", CDP_URL)

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CDP_URL)

        if not browser.contexts:
            print("❌ context 없음")
            return

        context = browser.contexts[0]
        pages = context.pages

        if not pages:
            print("❌ page 없음")
            return

        page = None
        for pg in pages:
            try:
                if "cretop.com" in pg.url:
                    page = pg
                    break
            except Exception:
                continue

        if page is None:
            page = pages[-1]

        print("현재 탭:", page.url)
        print()
        print("안내")
        print("  - 정보제공중지: 일반상세로 직접 들어가 보세요")
        print("  - 재로그인: 다른 세션으로 로그인하면 뜰 수 있습니다")
        print("  - 전체화면 안내: 만료/완료 화면이 되면 감지합니다")
        print("  - Ctrl+C 로 종료")
        print()

        if args.watch:
            last_sig = None
            n = 0

            try:
                while True:
                    snap = snapshot_page(page)
                    roots = list_candidate_roots(page)
                    classified = classify_state(snap, roots)

                    # 변화 있을 때만 덤프
                    sig = (
                        snap.get("url"),
                        tuple(classified.get("kinds") or []),
                        bool(classified.get("modal_matches")),
                        snap.get("has_search_list"),
                        (snap.get("body_preview") or "")[:120],
                    )

                    interesting = (
                        classified.get("kinds") != ["해당없음_또는_미분류"]
                        or classified.get("modal_matches")
                        or classified.get("is_full_page_notice")
                    )

                    if interesting and sig != last_sig:
                        n += 1
                        tag = datetime.now().strftime("%H%M%S") + f"_{n}"
                        print()
                        dump_scan(
                            page,
                            out_dir,
                            tag,
                            try_close=args.try_close,
                        )
                        last_sig = sig
                    else:
                        print(
                            f"\r감시 중... url={snap.get('url','')[:60]} "
                            f"list={snap.get('has_search_list')} "
                            f"kinds={classified.get('kinds')}",
                            end="",
                            flush=True,
                        )

                    time.sleep(args.interval)

            except KeyboardInterrupt:
                print("\n종료")

        else:
            tag = datetime.now().strftime("%Y%m%d_%H%M%S")
            dump_scan(
                page,
                out_dir,
                tag,
                try_close=args.try_close,
            )

        # Chrome 유지
        try:
            p.stop()
        except Exception:
            pass


if __name__ == "__main__":
    main()
