"""
ct2.py — 실제 수집(ct.py) + 팝업/안내 감시

사용법 (ct.py 와 동일 인자):
  py ct2.py
  py ct2.py --start-page 3 --start-card 7
  py ct2.py --no-auto-close    # 감지·덤프만, 자동 닫기 안 함

덤프 위치:
  popup_dumps_ct2/
"""

from __future__ import annotations

import sys


def main():
    auto_close = True
    argv = []

    for a in sys.argv[1:]:
        if a == "--no-auto-close":
            auto_close = False
        else:
            argv.append(a)

    # ct.py 의 argparse 가 sys.argv 를 쓰므로 정리
    sys.argv = [sys.argv[0]] + argv

    import ct
    from popup_monitor import install_hooks

    install_hooks(ct, auto_close=auto_close)
    ct.main()


if __name__ == "__main__":
    main()
