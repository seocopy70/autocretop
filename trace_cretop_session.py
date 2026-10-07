import hashlib
import hmac
import json
import re
import secrets
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from playwright.sync_api import sync_playwright


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_PATH = OUTPUT_DIR / "cretop_session_trace.jsonl"
CDP_URL = "http://127.0.0.1:9222"
HASH_KEY = secrets.token_bytes(32)

EXPIRY_MARKERS = (
    "페이지가 만료되었습니다",
    "페이지가 만료",
    "웹페이지가 만료",
    "세션이 만료",
    "로그인 세션",
    "8004",
    "-8002",
)

SENSITIVE_FIELD_PATTERN = re.compile(
    r"password|passwd|token|session|cookie|authorization|secret|signature|kipSgnVal|credential",
    re.IGNORECASE,
)

CLICK_TRACE_SCRIPT = r"""(() => {
    if (window.__cretopTraceInstalled) return;
    window.__cretopTraceInstalled = true;
    document.addEventListener('click', event => {
        const element = event.target instanceof Element
            ? event.target.closest('button, a, [role="button"], input[type="button"]')
            : null;
        if (!element) return;

        const text = (element.innerText || element.value || '').replace(/\s+/g, ' ').trim();
        const safeText = /^\d{1,5}$/.test(text)
            || ['확인', '일반', '재무', '다음그룹', '이전그룹'].includes(text)
            ? text
            : '';
        const href = element.href ? new URL(element.href, location.href).pathname : '';
        console.debug('__CRETOP_TRACE_CLICK__' + JSON.stringify({
            tag: element.tagName,
            className: String(element.className || '').slice(0, 120),
            role: element.getAttribute('role') || '',
            isTrusted: event.isTrusted,
            safeText,
            textHashSource: text.slice(0, 300),
            href
        }));
    }, true);

    for (const eventName of ['input', 'change']) {
        document.addEventListener(eventName, event => {
            const element = event.target;
            if (!(element instanceof HTMLInputElement)
                && !(element instanceof HTMLTextAreaElement)
                && !(element instanceof HTMLSelectElement)) return;

            const value = element instanceof HTMLInputElement
                && ['checkbox', 'radio'].includes(element.type)
                ? String(element.checked)
                : String(element.value || '');

            console.debug('__CRETOP_TRACE_INPUT__' + JSON.stringify({
                event: eventName,
                tag: element.tagName,
                type: element.type || '',
                id: element.id || '',
                name: element.name || '',
                isTrusted: event.isTrusted,
                valueLength: value.length,
                valueHashSource: value.slice(0, 2000)
            }));
        }, true);
    }

    for (const methodName of ['pushState', 'replaceState']) {
        const original = history[methodName].bind(history);
        history[methodName] = (...args) => {
            const result = original(...args);
            console.debug('__CRETOP_TRACE_HISTORY__' + JSON.stringify({
                method: methodName,
                path: location.pathname
            }));
            return result;
        };
    }

    window.addEventListener('popstate', () => {
        console.debug('__CRETOP_TRACE_HISTORY__' + JSON.stringify({
            method: 'popstate',
            path: location.pathname
        }));
    });
})()"""


def _hash(value):
    return hmac.new(
        HASH_KEY,
        str(value).encode("utf-8", errors="replace"),
        hashlib.sha256,
    ).hexdigest()


def _safe_url(url):
    parsed = urlsplit(url)
    path = parsed.path
    if "/dynaPath/" in path:
        path = re.sub(r"(/dynaPath/)[^/]+(?:/[^/]+)*", r"\1<redacted>", path)
    return f"{parsed.scheme}://{parsed.netloc}{path}"


def _payload_shape(raw, content_type="", hash_values=True):
    if not raw:
        return []

    parsed = None
    try:
        parsed = json.loads(raw)
    except Exception:
        if "form" in content_type.lower() or "=" in raw:
            try:
                parsed = parse_qs(raw, keep_blank_values=True)
            except Exception:
                parsed = None

    if parsed is None:
        return [{
            "path": "$",
            "type": "raw",
            "length": len(raw),
            "hash": _hash(raw),
        }]

    fields = []

    def visit(value, path, depth=0):
        if len(fields) >= 120 or depth > 6:
            return

        if isinstance(value, dict):
            for key, child in value.items():
                safe_key = re.sub(r"[^A-Za-z0-9_.-]", "_", str(key))[:80]
                visit(child, f"{path}.{safe_key}", depth + 1)
                if len(fields) >= 120:
                    break
            return

        if isinstance(value, list):
            fields.append({"path": path, "type": "array", "length": len(value)})
            for index, child in enumerate(value[:20]):
                visit(child, f"{path}[{index}]", depth + 1)
                if len(fields) >= 120:
                    break
            return

        if value is None:
            value_type = "null"
            text = ""
        elif isinstance(value, bool):
            value_type = "boolean"
            text = str(value).lower()
        elif isinstance(value, (int, float)):
            value_type = "number"
            text = str(value)
        else:
            value_type = "string"
            text = str(value)

        entry = {"path": path, "type": value_type, "length": len(text)}
        if SENSITIVE_FIELD_PATTERN.search(path):
            entry["value"] = "<redacted>"
        elif hash_values:
            entry["valueHash"] = _hash(text)
        fields.append(entry)

    visit(parsed, "$raw")
    return fields


def _write(event, **fields):
    record = {
        "time": datetime.now().astimezone().isoformat(timespec="milliseconds"),
        "event": event,
        **fields,
    }
    with LOG_PATH.open("a", encoding="utf-8") as log_file:
        log_file.write(json.dumps(record, ensure_ascii=False) + "\n")
        log_file.flush()

    visible_fields = (
        "reason",
        "url",
        "method",
        "path",
        "resourceType",
        "status",
        "safeText",
        "event",
        "type",
        "name",
        "id",
        "activePage",
        "resultCards",
        "markers",
        "messageHash",
        "isTrusted",
        "requestKey",
    )
    summary = {
        key: record[key]
        for key in visible_fields
        if key in record
    }
    print(f"[trace] {json.dumps(summary, ensure_ascii=False)}", flush=True)


def _is_cretop_url(url):
    try:
        hostname = (urlsplit(url).hostname or "").lower()
        return hostname == "cretop.com" or hostname.endswith(".cretop.com")
    except Exception:
        return False


def _snapshot(context, page, reason):
    if not _is_cretop_url(page.url):
        return

    try:
        page_state = page.evaluate(
            """() => ({
                title: document.title,
                activePage: document.querySelector('div.pagination button.num.on')?.innerText.trim() || null,
                resultCards: document.querySelectorAll('ul.search-result__list > li').length,
                localStorage: Object.fromEntries(
                    Object.keys(localStorage).sort().map(key => [key, localStorage.getItem(key)])
                ),
                sessionStorage: Object.fromEntries(
                    Object.keys(sessionStorage).sort().map(key => [key, sessionStorage.getItem(key)])
                ),
                expiryMarkers: ['페이지가 만료되었습니다', '페이지가 만료', '세션이 만료', '8004', '-8002']
                    .filter(marker => (document.body?.innerText || '').includes(marker))
            })"""
        )
    except Exception as error:
        page_state = {"stateError": type(error).__name__}

    storage_hashes = {}
    for storage_name in ("localStorage", "sessionStorage"):
        values = page_state.pop(storage_name, {})
        storage_hashes[storage_name] = {
            key: _hash(value) for key, value in values.items()
        }

    cookie_state = []
    try:
        for cookie in context.cookies(["https://www.cretop.com"]):
            cookie_state.append({
                "name": cookie.get("name"),
                "domain": cookie.get("domain"),
                "path": cookie.get("path"),
                "expires": cookie.get("expires"),
                "httpOnly": cookie.get("httpOnly"),
                "secure": cookie.get("secure"),
                "valueHash": _hash(cookie.get("value", "")),
            })
    except Exception as error:
        cookie_state = [{"snapshotError": type(error).__name__}]

    _write(
        "state_snapshot",
        reason=reason,
        url=_safe_url(page.url),
        page=page_state,
        storageHashes=storage_hashes,
        cookies=cookie_state,
    )


def _attach_page(context, page, attached_pages):
    if page in attached_pages:
        return
    attached_pages.add(page)

    try:
        page.add_init_script(CLICK_TRACE_SCRIPT)
        if _is_cretop_url(page.url):
            page.evaluate(CLICK_TRACE_SCRIPT)
            _snapshot(context, page, "observer_attached")
    except Exception:
        pass

    def on_request(request):
        if not _is_cretop_url(request.url):
            return
        post_data = request.post_data or ""
        content_type = request.headers.get("content-type", "")
        post_data_hash = _hash(post_data) if post_data else None
        request_key = _hash(
            f"{request.method}|{_safe_url(request.url)}|{post_data_hash or ''}"
        )
        _write(
            "request",
            requestKey=request_key,
            method=request.method,
            resourceType=request.resource_type,
            url=_safe_url(request.url),
            contentType=content_type,
            postDataHash=post_data_hash,
            postDataLength=len(post_data) if post_data else 0,
            postDataFields=_payload_shape(post_data, content_type),
        )

    def on_response(response):
        if not _is_cretop_url(response.url):
            return
        record = {
            "status": response.status,
            "resourceType": response.request.resource_type,
            "url": _safe_url(response.url),
        }
        request = response.request
        request_post_data = request.post_data or ""
        record["requestKey"] = _hash(
            f"{request.method}|{_safe_url(request.url)}|"
            f"{_hash(request_post_data) if request_post_data else ''}"
        )
        if "/dynaPath/" in urlsplit(response.url).path:
            try:
                body = response.text()
                content_type = response.headers.get("content-type", "")
                record["bodyHash"] = _hash(body)
                record["bodyLength"] = len(body)
                record["contentType"] = content_type
                record["markers"] = [marker for marker in EXPIRY_MARKERS if marker in body]
                record["bodyFields"] = _payload_shape(
                    body,
                    content_type,
                    hash_values=False,
                )
            except Exception as error:
                record["bodyReadError"] = type(error).__name__
        _write("response", **record)
        if "/dynaPath/" in urlsplit(response.url).path:
            _snapshot(context, page, "dynaPath_response")

    def on_request_failed(request):
        if _is_cretop_url(request.url):
            _write(
                "request_failed",
                method=request.method,
                resourceType=request.resource_type,
                url=_safe_url(request.url),
                failure=request.failure,
            )

    def on_dialog(dialog):
        message = dialog.message or ""
        _write(
            "browser_dialog",
            type=dialog.type,
            messageHash=_hash(message),
            markers=[marker for marker in EXPIRY_MARKERS if marker in message],
        )

    def on_console(message):
        text = message.text or ""
        click_prefix = "__CRETOP_TRACE_CLICK__"
        input_prefix = "__CRETOP_TRACE_INPUT__"
        history_prefix = "__CRETOP_TRACE_HISTORY__"
        if text.startswith(click_prefix):
            try:
                data = json.loads(text[len(click_prefix):])
                raw_text = data.pop("textHashSource", "")
                data["textHash"] = _hash(raw_text) if raw_text else None
                _write("click", **data)
            except Exception:
                _write("click", parseError=True)
        elif text.startswith(input_prefix):
            try:
                data = json.loads(text[len(input_prefix):])
                raw_value = data.pop("valueHashSource", "")
                data["valueHash"] = _hash(raw_value)
                _write("form_change", **data)
            except Exception:
                _write("form_change", parseError=True)
        elif text.startswith(history_prefix):
            try:
                data = json.loads(text[len(history_prefix):])
                _write("history_change", **data)
            except Exception:
                _write("history_change", parseError=True)

    def on_frame_navigated(frame):
        if frame != page.main_frame or not _is_cretop_url(page.url):
            return
        _write("navigation", url=_safe_url(page.url))
        _snapshot(context, page, "navigation")
        try:
            page.evaluate(CLICK_TRACE_SCRIPT)
        except Exception:
            pass

    def on_popup(popup):
        _write("popup", openerUrl=_safe_url(page.url), popupUrl=_safe_url(popup.url))
        _attach_page(context, popup, attached_pages)

    page.on("request", on_request)
    page.on("response", on_response)
    page.on("requestfailed", on_request_failed)
    page.on("dialog", on_dialog)
    page.on("console", on_console)
    page.on("framenavigated", on_frame_navigated)
    page.on("popup", on_popup)


def main():
    print(f"CRETOP 세션 관찰 중: {LOG_PATH}")
    print("검색·상세조회 동작을 진행하세요. 종료하려면 이 터미널에서 Ctrl+C를 누르세요.")
    print("쿠키와 저장소 값 원문, 요청/응답 본문은 기록하지 않습니다.")

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(CDP_URL)
        if not browser.contexts:
            raise RuntimeError("원격 디버깅 Chrome context를 찾지 못했습니다.")

        context = browser.contexts[0]
        attached_pages = set()
        context.on("page", lambda page: _attach_page(context, page, attached_pages))
        for page in context.pages:
            _attach_page(context, page, attached_pages)

        _write(
            "observer_started",
            pages=len(attached_pages),
            hashScheme="per-run HMAC-SHA256",
        )
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            for page in context.pages:
                _snapshot(context, page, "observer_stopped")
            _write("observer_stopped")
            print("관찰을 종료했습니다. Chrome은 닫지 않았습니다.")


if __name__ == "__main__":
    main()