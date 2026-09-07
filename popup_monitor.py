"""
CRETOP 팝업/안내 감시·덤프 (ct2.py 연동)

강화 포인트
- vfm / role=dialog / .pop-area.PLCM* 직접 탐지
- 정보제공중지: h2 '정보제공중지 안내', 버튼 '확인'·'팝업닫기'
- 재로그인·만료 키워드
- 전체화면 안내(목록 없고 본문만 안내)
- 사이드 퀵패널 '닫기'는 제외
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DUMP_DIR = BASE_DIR / "popup_dumps_ct2"
DUMP_DIR.mkdir(parents=True, exist_ok=True)

_seq = 0


def _norm(text) -> str:
    if not text:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


INFO_STOP = [
    "정보제공중지",
    "정보제공 중지",
    "정보 제공 중지",
    "정보가 공개되지",
    "정보제공거부",
    "정보제공 거부",
    "제공 중지",
    "제공중지",
]

RELOGIN = [
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
    "로그인이 필요",
    "로그인되어 있지",
]

PAGE_NOTICE = [
    "웹페이지가 만료",
    "페이지가 만료",
    "페이지 만료",
    "요청하신 페이지",
    "더 이상 유효하지",
    "세션이 종료",
    "작업이 완료",
]


def scan_overlays(page) -> dict:
    """현재 떠 있는 오버레이/안내 스캔"""

    return page.evaluate(
        """() => {
            function norm(s) {
                return (s || '').replace(/\\s+/g, ' ').trim();
            }

            const kinds = [];
            const overlays = [];

            // 1) Vue Final Modal / dialog
            const dialogs = document.querySelectorAll(
                '.modals-container .vfm__container[role=dialog], [role=dialog][aria-modal=true], .pop-area'
            );

            for (const el of dialogs) {
                if (el.closest('.quick-wrap, .quick-view, .header, footer')) continue;
                const style = window.getComputedStyle(el);
                const r = el.getBoundingClientRect();
                if (style.display === 'none' || style.visibility === 'hidden') continue;
                if (r.width < 30 || r.height < 30) continue;

                const text = norm(el.innerText);
                if (!text) continue;

                const buttons = Array.from(el.querySelectorAll('button')).map(b => ({
                    text: norm(b.innerText),
                    className: (b.className || '').toString().slice(0, 100),
                }));

                let kind = 'unknown_modal';
                if (/정보제공중지|정보가 공개되지|제공.?중지|제공.?거부/.test(text)) {
                    kind = 'info_stop';
                } else if (/재로그인|다시 로그인|세션|중복 로그인|다른 (곳|기기|사용자)/.test(text)) {
                    kind = 'relogin';
                } else if (/만료|유효하지|세션이 종료/.test(text)) {
                    kind = 'page_notice';
                }

                overlays.push({
                    kind,
                    text: text.slice(0, 500),
                    buttons,
                    rect: {x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height)},
                    className: (el.className || '').toString().slice(0, 120),
                });
                kinds.push(kind);
            }

            // 2) 본문만 바뀐 전체화면 안내
            const list = document.querySelector('ul.search-result__list > li');
            const body = norm(document.body.innerText).slice(0, 1500);
            const hasList = !!list;

            if (!hasList) {
                if (/재로그인|다시 로그인|세션이 만료|중복 로그인/.test(body)) {
                    kinds.push('relogin_fullpage');
                    overlays.push({kind: 'relogin_fullpage', text: body.slice(0, 400), buttons: []});
                }
                if (/웹페이지가 만료|페이지가 만료|더 이상 유효하지|세션이 종료/.test(body)) {
                    kinds.push('page_notice_fullpage');
                    overlays.push({kind: 'page_notice_fullpage', text: body.slice(0, 400), buttons: []});
                }
            }

            // 3) PLCM910P1 등 코드명 직접
            const plcm = document.querySelector('.pop-area.PLCM910P1');
            if (plcm) {
                const style = window.getComputedStyle(plcm);
                const r = plcm.getBoundingClientRect();
                if (style.display !== 'none' && r.width > 30) {
                    if (!kinds.includes('info_stop')) {
                        kinds.push('info_stop');
                        overlays.push({
                            kind: 'info_stop',
                            text: norm(plcm.innerText).slice(0, 500),
                            buttons: Array.from(plcm.querySelectorAll('button')).map(b => ({
                                text: norm(b.innerText),
                                className: (b.className || '').toString().slice(0, 100),
                            })),
                            source: 'PLCM910P1',
                        });
                    }
                }
            }

            return {
                kinds: Array.from(new Set(kinds)),
                overlays,
                url: location.href,
                hasList,
                title: document.title || '',
            };
        }"""
    )


def dump_event(page, context: str, scan: dict | None = None) -> Path | None:
    """감지 내용을 파일로 저장"""

    global _seq
    _seq += 1

    if scan is None:
        try:
            scan = scan_overlays(page)
        except Exception as e:
            scan = {"error": str(e)}

    if not scan.get("kinds"):
        return None

    tag = datetime.now().strftime("%H%M%S") + f"_{_seq}_{context}"
    path = DUMP_DIR / f"{tag}.json"

    payload = {
        "context": context,
        "time": datetime.now().isoformat(timespec="seconds"),
        "scan": scan,
    }

    try:
        html_path = DUMP_DIR / f"{tag}.html"
        html_path.write_text(page.content(), encoding="utf-8")
        payload["html"] = str(html_path)
    except Exception as e:
        payload["html_error"] = str(e)

    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(
        f"    [감시] {context} kinds={scan.get('kinds')} → {path.name}"
    )
    for ov in scan.get("overlays") or []:
        print(
            f"           · {ov.get('kind')}: "
            f"{(ov.get('text') or '')[:80]!r} "
            f"btns={[b.get('text') for b in (ov.get('buttons') or [])]}"
        )

    return path


def try_close_info_stop(page) -> bool:
    """정보제공중지: 확인 또는 팝업닫기"""

    try:
        closed = page.evaluate(
            """() => {
                function norm(s) {
                    return (s || '').replace(/\\s+/g, ' ').trim();
                }
                const roots = document.querySelectorAll(
                    '.pop-area.PLCM910P1, .modals-container [role=dialog], [role=dialog][aria-modal=true]'
                );
                for (const root of roots) {
                    const text = norm(root.innerText);
                    if (!/정보제공중지|정보가 공개되지|제공.?중지/.test(text)) continue;
                    // 확인 우선
                    for (const b of root.querySelectorAll('button')) {
                        const t = norm(b.innerText);
                        if (t === '확인') { b.click(); return {ok:true, text:t}; }
                    }
                    for (const b of root.querySelectorAll('button')) {
                        const t = norm(b.innerText);
                        if (t === '팝업닫기' || t.includes('팝업닫기')) {
                            b.click(); return {ok:true, text:t};
                        }
                    }
                }
                return {ok:false};
            }"""
        )
        if closed and closed.get("ok"):
            print(
                f"    [감시] 정보제공중지 닫음: {closed.get('text')}"
            )
            page.wait_for_timeout(300)
            return True
    except Exception as e:
        print(f"    [감시] 정보제공중지 닫기 실패: {e}")

    return False


def try_close_generic_modal(page) -> bool:
    """기타 모달: 확인/팝업닫기 (퀵패널 제외)"""

    try:
        closed = page.evaluate(
            """() => {
                function norm(s) {
                    return (s || '').replace(/\\s+/g, ' ').trim();
                }
                const roots = document.querySelectorAll(
                    '.modals-container [role=dialog], [role=dialog][aria-modal=true], .pop-area'
                );
                for (const root of roots) {
                    if (root.closest('.quick-wrap, .quick-view')) continue;
                    const style = window.getComputedStyle(root);
                    const r = root.getBoundingClientRect();
                    if (style.display === 'none' || r.width < 30) continue;
                    for (const prefer of ['확인', '팝업닫기', '닫기', 'OK']) {
                        for (const b of root.querySelectorAll('button')) {
                            if (norm(b.innerText) === prefer) {
                                b.click();
                                return {ok:true, text:prefer};
                            }
                        }
                    }
                }
                return {ok:false};
            }"""
        )
        if closed and closed.get("ok"):
            print(
                f"    [감시] 모달 닫음: {closed.get('text')}"
            )
            page.wait_for_timeout(300)
            return True
    except Exception:
        pass

    return False


def monitor_and_handle(page, context: str, auto_close: bool = True) -> dict:
    """스캔 → 덤프 → (옵션) 닫기"""

    try:
        scan = scan_overlays(page)
    except Exception as e:
        return {"error": str(e), "kinds": []}

    kinds = scan.get("kinds") or []

    if not kinds:
        return scan

    dump_event(page, context, scan)

    if not auto_close:
        return scan

    if any(
        k in ("info_stop",)
        or str(k).startswith("info_stop")
        for k in kinds
    ):
        try_close_info_stop(page)
    elif any("relogin" in str(k) for k in kinds):
        # 재로그인은 자동 클릭이 위험할 수 있어 덤프만 (원하면 확인 클릭)
        print(
            "    [감시] 재로그인 안내 감지 — 자동 닫기 보류 (덤프만)"
        )
        dump_event(page, context + "_relogin", scan)
    elif any("page_notice" in str(k) for k in kinds):
        print(
            "    [감시] 전체화면/만료 안내 감지 — 덤프만"
        )
        try_close_generic_modal(page)
    else:
        try_close_generic_modal(page)

    return scan


def install_hooks(ct_module, auto_close: bool = True):
    """
    ct 모듈의 주요 지점에 감시 삽입.
    - dismiss_blocking_popups 전후
    - click_general_from_card 직후 (process_company 내부에서 호출되는 경로)
    """

    orig_dismiss = ct_module.dismiss_blocking_popups
    orig_click = ct_module.click_general_from_card
    orig_return = ct_module.return_to_search

    def dismiss_wrapped(page, context=""):
        monitor_and_handle(
            page,
            f"dismiss전_{context or 'na'}",
            auto_close=False,
        )
        result = orig_dismiss(page, context=context)
        monitor_and_handle(
            page,
            f"dismiss후_{context or 'na'}",
            auto_close=auto_close,
        )
        return result

    def click_wrapped(page, card, company_name):
        result = orig_click(page, card, company_name)
        # 진입 직후 여러 번 폴링 (늦게 뜨는 팝업)
        for i in range(6):
            page.wait_for_timeout(150)
            scan = monitor_and_handle(
                page,
                f"일반진입_{i}_{company_name[:20]}",
                auto_close=auto_close,
            )
            if scan.get("kinds"):
                break
        return result

    def return_wrapped(page, search_url="", target_page=1, force=False):
        monitor_and_handle(
            page,
            "복귀전",
            auto_close=auto_close,
        )
        if force:
            result = orig_return(
                page,
                search_url=search_url,
                target_page=target_page,
                force=True,
            )
        else:
            # force 인자 없는 구버전 호환
            try:
                result = orig_return(
                    page,
                    search_url=search_url,
                    target_page=target_page,
                    force=force,
                )
            except TypeError:
                result = orig_return(
                    page,
                    search_url=search_url,
                    target_page=target_page,
                )
        monitor_and_handle(
            page,
            "복귀후",
            auto_close=auto_close,
        )
        return result

    ct_module.dismiss_blocking_popups = dismiss_wrapped
    ct_module.click_general_from_card = click_wrapped
    ct_module.return_to_search = return_wrapped

    print("=" * 70)
    print("ct2 감시 훅 설치 완료")
    print(f"  덤프 폴더: {DUMP_DIR}")
    print(f"  자동 닫기: {auto_close}")
    print("=" * 70)
