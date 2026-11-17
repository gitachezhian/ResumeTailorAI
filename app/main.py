"""Resume Tailor - desktop app (pywebview + Edge WebView2)."""
from __future__ import annotations

import base64
import ctypes
import ctypes.wintypes as wt
import json
import os
import re
import sys
import threading
import traceback

import webview

import ai
import engine
import layout

APP_NAME = "ResumeTailorAI"
PAGE_LIMIT = 2.0
SAFE_LIMIT = 1.97  # small safety margin for estimation error


# ------------------------------------------------------------------ paths
def base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def resource(path: str) -> str:
    root = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(root, path)


CONFIG_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")


# ------------------------------------------------------------------ DPAPI (encrypt API key at rest)
class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob(data: bytes):
    buf = ctypes.create_string_buffer(data, len(data))
    return _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf


def _dpapi(data: bytes, encrypt: bool) -> bytes:
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    src, _keep = _blob(data)
    out = _Blob()
    fn = crypt32.CryptProtectData if encrypt else crypt32.CryptUnprotectData
    if not fn(ctypes.byref(src), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("DPAPI failed")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        kernel32.LocalFree(out.pbData)


def protect(s: str) -> str:
    return base64.b64encode(_dpapi(s.encode(), True)).decode() if s else ""


def unprotect(s: str) -> str:
    try:
        return _dpapi(base64.b64decode(s), False).decode() if s else ""
    except Exception:
        return ""


# ------------------------------------------------------------------ config
def load_config() -> dict:
    try:
        with open(CONFIG_PATH, encoding="utf8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg: dict):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf8") as f:
        json.dump(cfg, f, indent=2)
    os.replace(tmp, CONFIG_PATH)


# ------------------------------------------------------------------ clipboard
def copy_to_clipboard(text: str):
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    kernel32.GlobalAlloc.restype = wt.HGLOBAL
    kernel32.GlobalAlloc.argtypes = [wt.UINT, ctypes.c_size_t]
    kernel32.GlobalLock.restype = wt.LPVOID
    kernel32.GlobalLock.argtypes = [wt.HGLOBAL]
    kernel32.GlobalUnlock.argtypes = [wt.HGLOBAL]
    user32.SetClipboardData.argtypes = [wt.UINT, wt.HANDLE]
    data = text.encode("utf-16-le") + b"\x00\x00"
    for _ in range(10):
        if user32.OpenClipboard(None):
            break
        threading.Event().wait(0.05)
    else:
        raise OSError("Clipboard is busy")
    try:
        user32.EmptyClipboard()
        h = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        p = kernel32.GlobalLock(h)
        ctypes.memmove(p, data, len(data))
        kernel32.GlobalUnlock(h)
        user32.SetClipboardData(CF_UNICODETEXT, h)
    finally:
        user32.CloseClipboard()


# ------------------------------------------------------------------ helpers
def safe_name(s: str) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", s or "").strip().strip(".")
    return re.sub(r"\s+", " ", s)[:80]


def unique_path(folder: str, stem: str, ext: str = ".docx") -> str:
    path = os.path.join(folder, stem + ext)
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem} ({n}){ext}")
        n += 1
    return path


def md_to_plain(md: str) -> str:
    out = []
    for line in md.replace("\r", "").split("\n"):
        m = re.match(r"^\s*[-*•]\s+(.*)$", line)
        line = ("• " + m.group(1)) if m else line
        out.append(line.replace("**", ""))
    return "\n".join(out)


# ------------------------------------------------------------------ API exposed to JS
class Api:
    def __init__(self):
        self._cfg = load_config()
        self._window = None
        self._resume = {}  # last tailored resume
        self._cover = {}   # last cover letter
        self._lock = threading.Lock()
        self._job = None   # CancelToken of the running job

    # ---- state
    def _templates_dir(self) -> str:
        d = self._cfg.get("templates_dir")
        if d and os.path.isdir(d):
            return d
        return os.path.join(base_dir(), "Templates")

    def _api_key(self) -> str:
        return unprotect(self._cfg.get("api_key_enc", "")) or os.environ.get("ANTHROPIC_API_KEY", "")

    def get_state(self):
        return {
            "templates": self.list_templates(),
            "templates_dir": self._templates_dir(),
            "resume_dir": self._cfg.get("resume_dir", ""),
            "cover_dir": self._cfg.get("cover_dir", ""),
            "has_key": bool(self._api_key()),
            "last_template": self._cfg.get("last_template", ""),
        }

    def list_templates(self):
        d = self._templates_dir()
        try:
            files = sorted(f for f in os.listdir(d) if f.lower().endswith(".docx") and not f.startswith("~$"))
        except FileNotFoundError:
            files = []
        return [{"file": f, "label": self._label(f)} for f in files]

    @staticmethod
    def _label(f: str) -> str:
        stem = os.path.splitext(f)[0]
        parts = [p.strip() for p in stem.split(" - ")]
        return parts[-1] if len(parts) > 1 else stem

    # ---- folders
    def _pick_folder(self, start: str = ""):
        res = self._window.create_file_dialog(webview.FileDialog.FOLDER, directory=start or os.path.expanduser("~"))
        if not res:
            return None
        return res[0] if isinstance(res, (list, tuple)) else res

    def choose_folder(self, which: str):
        key = {"templates": "templates_dir", "resume": "resume_dir", "cover": "cover_dir"}[which]
        current = self._cfg.get(key, "") or (self._templates_dir() if which == "templates" else "")
        folder = self._pick_folder(current)
        if folder:
            self._cfg[key] = folder
            save_config(self._cfg)
        return self.get_state()

    def open_folder(self, which: str):
        key = {"templates": None, "resume": "resume_dir", "cover": "cover_dir"}[which]
        path = self._templates_dir() if key is None else self._cfg.get(key, "")
        if path and os.path.isdir(path):
            os.startfile(path)
            return True
        if which == "templates":
            os.makedirs(path, exist_ok=True)
            os.startfile(path)
            return True
        return False

    # ---- API key
    def set_api_key(self, key: str):
        key = (key or "").strip()
        if not key:
            self._cfg.pop("api_key_enc", None)
            save_config(self._cfg)
            return {"ok": True, "has_key": bool(self._api_key())}
        try:
            ai.check_key(key)
        except Exception as e:
            return {"ok": False, "error": _friendly(e)}
        self._cfg["api_key_enc"] = protect(key)
        save_config(self._cfg)
        return {"ok": True, "has_key": True}

    # ---- generation
    @staticmethod
    def _stopped() -> bool:
        t = getattr(ai._local, "token", None)
        return bool(t and t.cancelled)

    def _progress(self, pct: int, msg: str):
        if self._window and not self._stopped():
            self._window.evaluate_js(f"window.onProgress && window.onProgress({int(pct)}, {json.dumps(msg)})")

    def _guarded(self, fn, *args):
        """Run a job on a worker thread so Stop can abandon it instantly."""
        with self._lock:
            if self._job is not None:
                return {"ok": False, "error": "Already working on one. Hang tight, or press Stop."}
            token = self._job = ai.CancelToken()
        done = threading.Event()
        box = {}

        def work():
            ai.use_token(token)
            try:
                box["r"] = fn(*args)
            except ai.Cancelled:
                box["r"] = {"ok": False, "cancelled": True}
            except Exception as e:
                traceback.print_exc()
                box["r"] = {"ok": False, "error": _friendly(e)}
            finally:
                done.set()

        threading.Thread(target=work, daemon=True).start()
        while not done.wait(0.15):
            if token.cancelled:
                break
        with self._lock:
            if self._job is token:
                self._job = None
        if token.cancelled:
            return {"ok": False, "cancelled": True}
        return box["r"]

    def cancel(self):
        with self._lock:
            job = self._job
        if job:
            job.cancel()
            return True
        return False

    def generate_resume(self, template_file: str, jd: str):
        return self._guarded(self._generate_resume, template_file, jd)

    def generate_cover(self, template_file: str, jd: str):
        return self._guarded(self._generate_cover, template_file, jd)

    def _check_inputs(self, template_file: str, jd: str):
        key = self._api_key()
        if not key:
            return None, None, {"ok": False, "error": "Add your Anthropic API key first (gear button, top right)."}
        if len((jd or "").strip()) < 80:
            return None, None, {"ok": False, "error": "Paste the full job description first (it looks too short)."}
        path = os.path.join(self._templates_dir(), template_file)
        if not os.path.isfile(path):
            return None, None, {"ok": False, "error": f"Template not found: {template_file}"}
        self._cfg["last_template"] = template_file
        save_config(self._cfg)
        return key, path, None

    def _generate_resume(self, template_file: str, jd: str):
        key, path, err = self._check_inputs(template_file, jd)
        if err:
            return err

        self._progress(5, "Reading your template...")
        tdoc, paras = engine.parse_template(path)
        t_pages = layout.estimate_pages(tdoc, path)
        outline = engine.template_outline(paras)
        edit_chars = sum(len(engine.plain(p.text)) for p in paras if not p.locked)
        budget = int(edit_chars * min(1.06, (SAFE_LIMIT - 0.04) / max(t_pages, 0.5)))

        self._progress(15, "Reading the job ad and the industry...")
        result = ai.tailor_resume(key, jd, outline, budget, os.path.splitext(template_file)[0])
        self._progress(70, "Pouring new words into your template...")
        doc, items, pages, trimmed = self._fit(key, jd, outline, path, result["paragraphs"], budget)

        if self._stopped():
            raise ai.Cancelled()
        sev = {"high": 0, "medium": 1, "low": 2}
        gaps = sorted(result.get("gaps", []), key=lambda g: sev.get(g.get("severity"), 3))
        self._resume = {
            "doc": doc, "template": path, "items": items, "jd": jd,
            "outline": outline, "budget": budget,
            "company": result.get("company", "").strip(),
            "role": result.get("role_title", "").strip(),
        }
        self._progress(100, "Done!")
        return {
            "ok": True,
            "blocks": engine.doc_to_blocks(doc),
            "industry": result.get("industry", ""),
            "industry_notes": result.get("industry_notes", ""),
            "keywords": result.get("keywords", []),
            "gaps": gaps,
            "company": self._resume["company"],
            "role": self._resume["role"],
            "pages": round(pages, 2),
            "auto_trimmed": trimmed,
            "file_name": self._stem(self._resume, "resume") + ".docx",
        }

    def _fit(self, key, jd, outline, path, items, budget):
        """Build the doc and keep it within 2 pages. Returns (doc, items, pages, trimmed)."""
        doc = engine.build_resume(path, items)
        pages = layout.estimate_pages(doc, path)
        tries = 0
        while pages > SAFE_LIMIT and tries < 2:
            tries += 1
            self._progress(75 + tries * 8, f"Trimming to fit 2 pages (pass {tries})...")
            over_lines = max(2, int((pages - SAFE_LIMIT) * 60) + 2)
            budget = int(budget * min(0.95, SAFE_LIMIT / pages))
            items = ai.shrink_resume(key, jd, outline, items, over_lines, budget)
            doc = engine.build_resume(path, items)
            pages = layout.estimate_pages(doc, path)
        trimmed = 0
        while pages > SAFE_LIMIT and engine.drop_last_bullet(doc):
            trimmed += 1
            pages = layout.estimate_pages(doc, path)
        if trimmed:  # rebuild so the item list matches what's on the page
            items = engine.live_items(doc, items)
            doc = engine.build_resume(path, items)
        return doc, items, pages, trimmed

    # ---- comments -> batch fix
    def fix_resume(self, comments: list):
        return self._guarded(self._fix_resume, comments)

    def fix_cover(self, comments: list):
        return self._guarded(self._fix_cover, comments)

    def _fix_resume(self, comments):
        r = self._resume
        if not r.get("doc"):
            return {"ok": False, "error": "Tailor a resume first."}
        if not comments:
            return {"ok": False, "error": "Add a comment first: select some text in the resume."}
        key = self._api_key()
        if not key:
            return {"ok": False, "error": "Add your Anthropic API key first (gear button, top right)."}
        items = [dict(it) for it in r["items"]]
        clean = []
        for c in comments:
            try:
                idx = int(c.get("item"))
            except (TypeError, ValueError):
                continue
            if 0 <= idx < len(items) and str(c.get("comment", "")).strip():
                clean.append({"item": idx, "quote": str(c.get("quote", ""))[:400], "comment": str(c["comment"]).strip()})
        if not clean:
            return {"ok": False, "error": "Those comments are on locked lines. Comment on bullets, skills or the profile."}
        self._progress(15, f"Fixing {len(clean)} issue{'s' if len(clean) != 1 else ''}...")
        edits = ai.fix_resume(key, r["jd"], items, clean)
        for e in edits:
            try:
                i = int(e["item"])
            except (KeyError, TypeError, ValueError):
                continue
            if 0 <= i < len(items):
                items[i]["text"] = str(e.get("text", ""))
        items = [it for it in items if it["text"].strip()]
        self._progress(70, "Pouring the fixes into your template...")
        doc, items, pages, trimmed = self._fit(key, r["jd"], r["outline"], r["template"], items, r["budget"])
        if self._stopped():
            raise ai.Cancelled()
        r.update(doc=doc, items=items)
        self._progress(100, "Done!")
        return {"ok": True, "blocks": engine.doc_to_blocks(doc), "pages": round(pages, 2),
                "auto_trimmed": trimmed, "fixed": len(clean)}

    def _fix_cover(self, comments):
        c = self._cover
        if not c.get("letter"):
            return {"ok": False, "error": "Tailor a cover letter first."}
        clean = [{"quote": str(x.get("quote", ""))[:400], "comment": str(x.get("comment", "")).strip()}
                 for x in comments or [] if str(x.get("comment", "")).strip()]
        if not clean:
            return {"ok": False, "error": "Add a comment first: select some text in the letter."}
        key = self._api_key()
        if not key:
            return {"ok": False, "error": "Add your Anthropic API key first (gear button, top right)."}
        self._progress(20, f"Fixing {len(clean)} issue{'s' if len(clean) != 1 else ''}...")
        letter = ai.fix_cover(key, c["jd"], c["letter"], clean).strip()
        if self._stopped():
            raise ai.Cancelled()
        c["letter"] = letter
        self._progress(100, "Done!")
        return {"ok": True, "letter": letter, "fixed": len(clean)}

    def _generate_cover(self, template_file: str, jd: str):
        key, path, err = self._check_inputs(template_file, jd)
        if err:
            return err
        self._progress(10, "Reading your experience...")
        _, paras = engine.parse_template(path)
        resume_text = engine.resume_plain_text(paras)
        self._progress(25, "Writing your cover letter...")
        result = ai.cover_letter(key, jd, resume_text)
        if self._stopped():
            raise ai.Cancelled()
        self._cover = {
            "letter": result["letter"].strip(), "template": path, "jd": jd,
            "company": result.get("company", "").strip(),
            "role": result.get("role_title", "").strip(),
        }
        self._progress(100, "Done!")
        return {
            "ok": True,
            "letter": self._cover["letter"],
            "company": self._cover["company"],
            "role": self._cover["role"],
            "file_name": self._stem(self._cover, "cover") + ".docx",
        }

    # ---- naming / saving
    @staticmethod
    def _stem(info: dict, kind: str) -> str:
        """Resume:       <template name minus its last part> - <role> - <company>
                         e.g. Gita C - Resume - Senior Product Designer - Programa
           Cover letter: <first part of template name> - Coverletter - <role> - <company>"""
        stem = os.path.splitext(os.path.basename(info.get("template", "Resume")))[0]
        parts = [p.strip() for p in stem.split(" - ") if p.strip()]
        if kind == "resume":
            prefix = parts[:-1] if len(parts) >= 3 else parts
        else:
            prefix = [parts[0], "Coverletter"] if len(parts) >= 2 else ["Coverletter"]
        bits = prefix + [safe_name(info.get("role", "")), safe_name(info.get("company", ""))]
        return safe_name(" - ".join(b for b in bits if b))

    def _ensure_dir(self, which: str):
        key = "resume_dir" if which == "resume" else "cover_dir"
        d = self._cfg.get(key, "")
        if not d or not os.path.isdir(d):
            d = self._pick_folder(self._cfg.get(key, "") or self._cfg.get("resume_dir", ""))
            if not d:
                return None
            self._cfg[key] = d
            save_config(self._cfg)
        return d

    def save_resume(self):
        if not self._resume.get("doc"):
            return {"ok": False, "error": "Tailor a resume first."}
        d = self._ensure_dir("resume")
        if not d:
            return {"ok": False, "cancelled": True}
        path = unique_path(d, self._stem(self._resume, "resume"))
        try:
            self._resume["doc"].save(path)
        except PermissionError:
            return {"ok": False, "error": "Couldn't save - is a file with that name open in Word?"}
        return {"ok": True, "path": path, "state": self.get_state()}

    def save_cover(self):
        if not self._cover.get("letter"):
            return {"ok": False, "error": "Tailor a cover letter first."}
        d = self._ensure_dir("cover")
        if not d:
            return {"ok": False, "cancelled": True}
        path = unique_path(d, self._stem(self._cover, "cover"))
        try:
            engine.build_cover_letter(self._cover["template"], self._cover["letter"]).save(path)
        except PermissionError:
            return {"ok": False, "error": "Couldn't save - is a file with that name open in Word?"}
        return {"ok": True, "path": path, "state": self.get_state()}

    def copy_cover(self):
        text = self._cover.get("letter", "")
        if not text:
            return {"ok": False, "error": "Nothing to copy yet."}
        try:
            copy_to_clipboard(md_to_plain(text))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_file(self, path: str):
        if path and os.path.isfile(path):
            os.startfile(path)
            return True
        return False


def _friendly(e: Exception) -> str:
    if isinstance(e, ai.anthropic.AuthenticationError):
        return "That API key was rejected. Double-check it in Settings."
    if isinstance(e, ai.anthropic.PermissionDeniedError):
        return "Your API key doesn't have access to this model."
    if isinstance(e, ai.anthropic.RateLimitError):
        return "Rate limited by the API. Wait a minute and try again."
    if isinstance(e, ai.anthropic.APIConnectionError):
        return "Can't reach the Anthropic API. Check your internet connection."
    if isinstance(e, ai.anthropic.APIStatusError):
        return f"API error ({e.status_code}): {getattr(e, 'message', e)}"
    return str(e) or e.__class__.__name__


def main():
    api = Api()
    window = webview.create_window(
        "Resume Tailor",
        url=resource(os.path.join("ui", "index.html")),
        js_api=api,
        width=1380, height=920, min_size=(1080, 720),
        background_color="#bfeaff",
        text_select=True,
    )
    api._window = window
    webview.start(gui="edgechromium", debug=bool(os.environ.get("RT_DEBUG")))


if __name__ == "__main__":
    main()
