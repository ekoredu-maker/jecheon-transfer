# -*- coding: utf-8 -*-
"""
제천교육지원청 교원 내신전보 시스템 - 포터블 실행기

- 파이썬 표준 라이브러리만 사용합니다(추가 설치 없음).
- 이 PC 안에서만 접속되는 작은 웹서버(127.0.0.1)를 띄우고,
  Edge(없으면 Chrome, 그것도 없으면 기본 브라우저)를 '앱 창'으로 엽니다.
- 작업 데이터는 포터블 폴더의 data/storage.json 에 저장되고,
  10분마다 data/backup 에 백업이 쌓입니다.
- 브라우저 창을 닫으면 몇 분 뒤 서버도 스스로 종료됩니다.
"""
import datetime
import json
import os
import re
import secrets
import shutil
import socket
import subprocess
import sys
import threading
import time
import traceback
import urllib.request
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

from notice_export import build_notice_xlsx
from proposal_export import build_proposal_xlsx
from urllib.parse import parse_qs, urlparse

APP_NAME = "제천 내신전보 시스템"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(APP_DIR)
WEB = os.path.join(APP_DIR, "web")
DATA = os.path.join(ROOT, "data")
BACKUP = os.path.join(DATA, "backup")
LOGS = os.path.join(DATA, "logs")
PROFILE = os.path.join(DATA, "browser-profile")
OUTPUT = os.path.join(ROOT, "output")
STORE = os.path.join(DATA, "storage.json")
RUNFILE = os.path.join(DATA, "running.json")
NOTICE_TEMPLATE = os.path.join(APP_DIR, "templates", "notice_template.xlsx")
PROPOSAL_TEMPLATE = os.path.join(APP_DIR, "templates", "proposal_template.xlsx")

PREFERRED_PORT = 18731
HEARTBEAT_TIMEOUT = 180      # 창이 닫힌 뒤 이 시간(초) 동안 신호가 없으면 종료
FIRST_CONTACT_TIMEOUT = 600  # 시작 후 한 번도 접속이 없으면 종료
BACKUP_INTERVAL = 600        # 자동 백업 간격(초)
BACKUP_KEEP = 60             # 보관할 백업 개수
MAX_BODY = 64 * 1024 * 1024
IS_WIN = os.name == "nt"

for d in (DATA, BACKUP, LOGS, OUTPUT):
    os.makedirs(d, exist_ok=True)

# pythonw.exe 로 실행하면 콘솔이 없으므로 로그 파일로 출력
_log_path = os.path.join(LOGS, "server.log")
try:
    if os.path.exists(_log_path) and os.path.getsize(_log_path) > 2 * 1024 * 1024:
        os.replace(_log_path, _log_path + ".old")
except OSError:
    pass
_log_file = open(_log_path, "a", encoding="utf-8")
if sys.stdout is None or not hasattr(sys.stdout, "write") or "pythonw" in os.path.basename(sys.executable).lower():
    sys.stdout = _log_file
    sys.stderr = _log_file


def log(msg):
    line = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    try:
        _log_file.write(line + "\n")
        _log_file.flush()
    except Exception:
        pass
    if sys.stdout is not _log_file:
        try:
            print(line, flush=True)
        except Exception:
            pass


def message_box(text, title=APP_NAME):
    log("MSG: " + text.replace("\n", " "))
    if IS_WIN:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, text, title, 0x40)
            return
        except Exception:
            pass


# ---------------------------------------------------------------- 저장소
_store_lock = threading.Lock()
_last_backup = 0.0


def read_store():
    with _store_lock:
        if not os.path.exists(STORE):
            return {"version": 1, "saved_at": None, "items": {}}
        try:
            with open(STORE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data.get("items"), dict):
                raise ValueError("items 형식 오류")
            return data
        except Exception as e:
            # 손상된 파일은 보존하고 마지막 백업으로 복구 시도
            bad = STORE + f".damaged-{datetime.datetime.now():%Y%m%d_%H%M%S}"
            try:
                shutil.copy2(STORE, bad)
            except OSError:
                pass
            log(f"storage.json 읽기 실패({e}). 손상본 보관: {bad}")
            for name in sorted(list_backups(), reverse=True):
                try:
                    with open(os.path.join(BACKUP, name), "r", encoding="utf-8") as f:
                        data = json.load(f)
                    if isinstance(data.get("items"), dict):
                        log(f"백업에서 복구: {name}")
                        return data
                except Exception:
                    continue
            raise


def write_store(items):
    global _last_backup
    data = {"version": 1, "saved_at": datetime.datetime.now().isoformat(timespec="seconds"), "items": items}
    with _store_lock:
        tmp = STORE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, STORE)  # 원자적 교체(저장 중 전원이 꺼져도 이전 파일 유지)
    if time.time() - _last_backup >= BACKUP_INTERVAL:
        make_backup()
    return data


def list_backups():
    try:
        return [n for n in os.listdir(BACKUP) if n.startswith("storage_") and n.endswith(".json")]
    except OSError:
        return []


def make_backup(tag=""):
    global _last_backup
    if not os.path.exists(STORE):
        return None
    now = datetime.datetime.now()
    name = f"storage_{now:%Y%m%d_%H%M%S}_{now.microsecond // 1000:03d}{tag}.json"
    with _store_lock:
        shutil.copy2(STORE, os.path.join(BACKUP, name))
    _last_backup = time.time()
    backups = sorted(list_backups())
    for old in backups[:-BACKUP_KEEP]:
        try:
            os.remove(os.path.join(BACKUP, old))
        except OSError:
            pass
    log(f"백업 생성: {name}")
    return name


def open_folder(path):
    os.makedirs(path, exist_ok=True)
    if IS_WIN:
        os.startfile(path)  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ---------------------------------------------------------------- 웹서버
class State:
    token = secrets.token_urlsafe(24)
    port = 0
    last_ping = 0.0
    started = time.time()
    server = None


class Handler(SimpleHTTPRequestHandler):
    server_version = "JHRSPortable/1.0"

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=WEB, **kw)

    def log_message(self, fmt, *args):
        if "/api/ping" in (self.path or ""):
            return
        log("HTTP " + (fmt % args).replace(State.token, "***"))

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        super().end_headers()

    # 다른 웹사이트가 이 서버에 접근하지 못하도록 Host/토큰 확인
    def _host_ok(self):
        host = (self.headers.get("Host") or "").lower()
        return host in (f"127.0.0.1:{State.port}", f"localhost:{State.port}")

    def _token_ok(self, qs):
        return secrets.compare_digest(self.headers.get("X-Token") or (qs.get("t") or [""])[0], State.token)

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _guard(self):
        if not self._host_ok():
            self.send_error(403, "Forbidden host")
            return None
        u = urlparse(self.path)
        qs = parse_qs(u.query)
        if u.path.startswith("/api/") and not self._token_ok(qs):
            self._json(403, {"error": "token"})
            return None
        State.last_ping = time.time()
        return u, qs

    def do_GET(self):
        g = self._guard()
        if not g:
            return
        u, qs = g
        if u.path == "/api/ping":
            return self._json(200, {"ok": True})
        if u.path == "/api/storage":
            try:
                return self._json(200, read_store())
            except Exception as e:
                return self._json(500, {"error": str(e)})
        if u.path == "/api/info":
            return self._json(200, {"data": DATA, "output": OUTPUT, "backups": len(list_backups()),
                                    "python": sys.version.split()[0]})
        if u.path.startswith("/api/"):
            return self._json(404, {"error": "not found"})
        if u.path in ("/", ""):
            self.path = "/index.html"
        return super().do_GET()

    def do_HEAD(self):
        if not self._host_ok():
            return self.send_error(403)
        return super().do_HEAD()

    def do_POST(self):
        g = self._guard()
        if not g:
            return
        u, qs = g
        if u.path == "/api/storage":
            try:
                n = int(self.headers.get("Content-Length") or 0)
                if n <= 0 or n > MAX_BODY:
                    return self._json(413, {"error": "size"})
                payload = json.loads(self.rfile.read(n).decode("utf-8"))
                items = payload.get("items")
                if not isinstance(items, dict) or not all(isinstance(k, str) and isinstance(v, str) for k, v in items.items()):
                    return self._json(400, {"error": "format"})
                data = write_store(items)
                return self._json(200, {"ok": True, "saved_at": data["saved_at"], "backups": len(list_backups())})
            except Exception as e:
                log("저장 오류: " + traceback.format_exc())
                return self._json(500, {"error": str(e)})
        if u.path == "/api/export-notice":
            try:
                n = int(self.headers.get("Content-Length") or 0)
                if n <= 0 or n > MAX_BODY:
                    return self._json(413, {"error": "size"})
                payload = json.loads(self.rfile.read(n).decode("utf-8"))
                records = payload.get("records")
                appoint_date = str(payload.get("appointDate") or "").strip()
                if not isinstance(records, list):
                    return self._json(400, {"error": "records 형식 오류"})
                if not os.path.isfile(NOTICE_TEMPLATE):
                    return self._json(500, {"error": "발령통지서 템플릿 파일이 없습니다."})
                data, meta = build_notice_xlsx(NOTICE_TEMPLATE, records, appoint_date)
                safe_date = re.sub(r"[^0-9]+", "", appoint_date) or datetime.datetime.now().strftime("%Y%m%d")
                stamp = datetime.datetime.now().strftime("%H%M%S")
                filename = f"인사발령통지서_{safe_date}_{stamp}.xlsx"
                out_path = os.path.join(OUTPUT, filename)
                tmp_path = out_path + ".tmp"
                with open(tmp_path, "wb") as f:
                    f.write(data)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp_path, out_path)
                log(f"발령통지서 생성: {filename} / {meta.get('total', 0)}명")
                return self._json(200, {"ok": True, "file": filename, "total": meta.get("total", 0),
                                        "counts": meta.get("counts", {}), "output": OUTPUT})
            except Exception as e:
                log("발령통지서 생성 오류: " + traceback.format_exc())
                return self._json(500, {"error": str(e)})
        if u.path == "/api/export-proposal":
            try:
                n = int(self.headers.get("Content-Length") or 0)
                if n <= 0 or n > MAX_BODY:
                    return self._json(413, {"error": "size"})
                payload = json.loads(self.rfile.read(n).decode("utf-8"))
                if not os.path.isfile(PROPOSAL_TEMPLATE):
                    return self._json(500, {"error": "전의안 템플릿 파일이 없습니다."})
                data, meta = build_proposal_xlsx(PROPOSAL_TEMPLATE, payload)
                appoint_date = str(payload.get("appointDate") or "").strip()
                safe_date = re.sub(r"[^0-9]+", "", appoint_date) or datetime.datetime.now().strftime("%Y%m%d")
                stamp = datetime.datetime.now().strftime("%H%M%S")
                filename = f"초등교사_전의안_{safe_date}_{stamp}.xlsx"
                out_path = os.path.join(OUTPUT, filename)
                tmp_path = out_path + ".tmp"
                with open(tmp_path, "wb") as f:
                    f.write(data)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp_path, out_path)
                log(f"공식 전의안 생성: {filename} / 결원 {meta.get('departures', 0)}명 / 충원 {meta.get('arrivals', 0)}명")
                return self._json(200, {"ok": True, "file": filename, "output": OUTPUT, **meta})
            except Exception as e:
                log("전의안 생성 오류: " + traceback.format_exc())
                return self._json(500, {"error": str(e)})
        if u.path == "/api/backup":
            try:
                name = make_backup("_manual")
                return self._json(200, {"ok": True, "file": name, "backups": len(list_backups())})
            except Exception as e:
                return self._json(500, {"error": str(e)})
        if u.path == "/api/open":
            what = (qs.get("what") or [""])[0]
            target = {"data": DATA, "backup": BACKUP, "output": OUTPUT}.get(what)
            if not target:
                return self._json(400, {"error": "what"})
            try:
                open_folder(target)
                return self._json(200, {"ok": True})
            except Exception as e:
                return self._json(500, {"error": str(e)})
        return self._json(404, {"error": "not found"})


def pick_port():
    for port in (PREFERRED_PORT, 0):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.bind(("127.0.0.1", port))
            p = s.getsockname()[1]
            s.close()
            return p
        except OSError:
            s.close()
    raise RuntimeError("사용할 수 있는 포트가 없습니다.")


# ---------------------------------------------------------------- 브라우저
def find_browser():
    if not IS_WIN:
        for name in ("microsoft-edge", "google-chrome", "chromium", "chromium-browser"):
            p = shutil.which(name)
            if p:
                return p
        return None
    cands = []
    for env in ("ProgramFiles(x86)", "ProgramFiles", "LocalAppData"):
        base = os.environ.get(env)
        if base:
            cands.append(os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"))
            cands.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
    try:
        import winreg
        for exe in ("msedge.exe", "chrome.exe"):
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths" + "\\" + exe) as k:
                        cands.append(winreg.QueryValue(k, None))
                except OSError:
                    pass
    except Exception:
        pass
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def prepare_profile():
    """포터블 전용 브라우저 프로필: 엑셀 저장 위치를 output 폴더로 지정"""
    pref_dir = os.path.join(PROFILE, "Default")
    os.makedirs(pref_dir, exist_ok=True)
    pref_path = os.path.join(pref_dir, "Preferences")
    prefs = {}
    try:
        if os.path.exists(pref_path):
            with open(pref_path, "r", encoding="utf-8") as f:
                prefs = json.load(f)
    except Exception:
        prefs = {}
    prefs.setdefault("download", {})
    prefs["download"]["default_directory"] = OUTPUT
    prefs["download"]["prompt_for_download"] = False
    prefs["download"]["directory_upgrade"] = True
    prefs.setdefault("savefile", {})["default_directory"] = OUTPUT
    try:
        with open(pref_path, "w", encoding="utf-8") as f:
            json.dump(prefs, f, ensure_ascii=False)
    except Exception as e:
        log(f"브라우저 설정 기록 실패: {e}")


def open_app_window(url):
    exe = find_browser()
    if exe:
        try:
            prepare_profile()
            subprocess.Popen([exe, f"--app={url}", f"--user-data-dir={PROFILE}", "--no-first-run",
                              "--no-default-browser-check", "--disable-features=Translate",
                              "--window-size=1600,960"], close_fds=True)
            log(f"브라우저 실행: {exe}")
            return True
        except Exception as e:
            log(f"브라우저 실행 실패({exe}): {e}")
    log("기본 브라우저로 엽니다.")
    return webbrowser.open(url)


# ---------------------------------------------------------------- 실행
def already_running():
    """이미 실행 중이면 그 창을 다시 열고 True 반환"""
    try:
        with open(RUNFILE, "r", encoding="utf-8") as f:
            info = json.load(f)
        url = f"http://127.0.0.1:{info['port']}/api/ping?t={info['token']}"
        with urllib.request.urlopen(url, timeout=2) as r:
            if r.status == 200:
                open_app_window(f"http://127.0.0.1:{info['port']}/?t={info['token']}")
                log("이미 실행 중인 서버에 연결했습니다.")
                return True
    except Exception:
        pass
    return False


def watchdog():
    while True:
        time.sleep(5)
        now = time.time()
        if State.last_ping and now - State.last_ping > HEARTBEAT_TIMEOUT:
            log("창이 닫힌 것으로 보고 종료합니다.")
            break
        if not State.last_ping and now - State.started > FIRST_CONTACT_TIMEOUT:
            log("접속이 없어 종료합니다.")
            break
    State.server.shutdown()


def check_writable():
    probe = os.path.join(DATA, ".write_test")
    try:
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(probe)
        return True
    except OSError:
        return False


def self_test():
    """배포 폴더의 핵심 구성과 쓰기/저장 상태를 빠르게 점검합니다."""
    checks = []
    def add(name, ok, detail=""):
        checks.append((name, bool(ok), detail))

    add("Python 3.8 이상", sys.version_info >= (3, 8), sys.version.split()[0])
    required = [
        ("화면 파일", os.path.join(WEB, "index.html")),
        ("엑셀 라이브러리", os.path.join(WEB, "vendor", "xlsx.full.min.js")),
        ("Tailwind CSS", os.path.join(WEB, "vendor", "tailwind.css")),
        ("Pretendard CSS", os.path.join(WEB, "vendor", "pretendard.css")),
        ("Font Awesome", os.path.join(WEB, "vendor", "fontawesome", "css", "all.min.css")),
        ("발령통지서 템플릿", NOTICE_TEMPLATE),
        ("전의안 템플릿", PROPOSAL_TEMPLATE),
    ]
    for label, path in required:
        add(label, os.path.isfile(path), os.path.relpath(path, ROOT))
    add("data 폴더 쓰기", check_writable(), DATA)
    try:
        if os.path.exists(STORE):
            data = read_store()
            add("storage.json 읽기", isinstance(data.get("items"), dict), STORE)
        else:
            add("storage.json 읽기", True, "아직 생성 전")
    except Exception as e:
        add("storage.json 읽기", False, str(e))

    print("=" * 58)
    print(f"{APP_NAME} 포터블 진단")
    print("=" * 58)
    for name, ok, detail in checks:
        print(f"[{'정상' if ok else '오류'}] {name}: {detail}")
    failed = [x for x in checks if not x[1]]
    print("-" * 58)
    print("진단 결과:", "정상" if not failed else f"오류 {len(failed)}건")
    return 0 if not failed else 1


def main():
    if "--self-test" in sys.argv:
        return self_test()
    no_browser = "--no-browser" in sys.argv
    log(f"=== 시작 (Python {sys.version.split()[0]}, 폴더 {ROOT}) ===")
    if not os.path.isfile(os.path.join(WEB, "index.html")):
        message_box("app\\web\\index.html 파일이 없습니다. 압축을 다시 풀어 주세요.")
        return 1
    if not check_writable():
        message_box("data 폴더에 쓸 수 없습니다.\n쓰기 금지된 USB이거나 읽기 전용 폴더인지 확인하세요.")
        return 1
    if not no_browser and already_running():
        return 0

    State.port = pick_port()
    State.server = ThreadingHTTPServer(("127.0.0.1", State.port), Handler)
    State.server.daemon_threads = True
    url = f"http://127.0.0.1:{State.port}/?t={State.token}"
    with open(RUNFILE, "w", encoding="utf-8") as f:
        json.dump({"port": State.port, "token": State.token, "pid": os.getpid()}, f)
    if os.path.exists(STORE):
        try:
            make_backup("_start")   # 시작할 때마다 한 번 백업
        except Exception as e:
            log(f"시작 백업 실패: {e}")
    threading.Thread(target=watchdog, daemon=True).start()
    log(f"서버 시작: {url.split('?')[0]}")
    if no_browser:
        print(url, flush=True)
    else:
        threading.Timer(0.3, open_app_window, args=(url,)).start()
    try:
        State.server.serve_forever(poll_interval=0.5)
    finally:
        State.server.server_close()
        try:
            os.remove(RUNFILE)
        except OSError:
            pass
        log("=== 종료 ===")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        log(traceback.format_exc())
        message_box("실행 중 오류가 발생했습니다.\n자세한 내용은 data\\logs\\server.log 를 확인하세요.")
        sys.exit(1)
