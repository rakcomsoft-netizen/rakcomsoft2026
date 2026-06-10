# -*- coding: utf-8 -*-
"""
updater.py — RakComSoft 2026 Update Engine
────────────────────────────────────────────────────────────
ระบบอัปเดตโปรแกรม (Offline-First friendly) ใช้ GitHub Releases เป็น Update Server

ออกแบบสำหรับ deploy แบบ Python (.py) ที่ D:\\pos
- เช็กเวอร์ชันล่าสุด (GitHub Releases API)
- เปรียบเทียบเวอร์ชัน (semantic)
- ดาวน์โหลด update.zip (มี progress + กัน partial/disconnect)
- Backup ไฟล์เดิม (ยกเว้น *.db — ฐานข้อมูลห้ามแตะ)
- Apply อัปเดต (replace เฉพาะไฟล์ใน zip, กัน path traversal)
- Rollback คืนไฟล์เดิม
- Log ทุกขั้นตอน

ใช้ stdlib ล้วน (urllib/zipfile/shutil/json) — ไม่ต้องลง package เพิ่ม

⚙️ ตั้งค่าก่อนใช้งานจริง: แก้ GITHUB_REPO ด้านล่างเป็น owner/repo ของคุณ
"""
import os, sys, json, ssl, shutil, zipfile, urllib.request, urllib.error
from datetime import datetime

# ════════════════════════════════════════════════════════════
# CONFIG — แก้ตรงนี้ก่อนใช้งานจริง
# ════════════════════════════════════════════════════════════
APP_VERSION  = "2.2.0"                       # เวอร์ชันปัจจุบันของโปรแกรม
GITHUB_REPO  = "rakcomsoft-netizen/rakcomsoft2026"  # ⚠️ แก้เป็น owner/repo ของคุณ
ASSET_NAME   = "update.zip"                  # ชื่อไฟล์ asset ใน GitHub Release
USER_AGENT   = "RakComSoft-Updater/2026"
DB_EXCLUDE   = (".db", ".db-journal", ".db-wal", ".db-shm")  # ห้าม backup/replace
DIR_EXCLUDE  = ("Backup", "__pycache__", ".git", "logs")

def app_dir() -> str:
    """โฟลเดอร์ติดตั้ง (ที่อยู่ของ updater.py = D:\\pos)"""
    return os.path.dirname(os.path.abspath(__file__))

# ════════════════════════════════════════════════════════════
# VERSION COMPARE
# ════════════════════════════════════════════════════════════
def parse_version(v) -> tuple:
    """'v1.2.3' / '1.2' / '2.0.0-beta' → (1,2,3) เปรียบเทียบได้"""
    s = str(v or "0").strip().lstrip("vV")
    s = s.split("-")[0].split("+")[0]          # ตัด -beta / +build
    parts = []
    for p in s.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3: parts.append(0)
    return tuple(parts[:4])

def is_newer(latest, current) -> bool:
    return parse_version(latest) > parse_version(current)

# ════════════════════════════════════════════════════════════
# CHECK FOR UPDATE (GitHub Releases)
# ════════════════════════════════════════════════════════════
# ════════════════════════════════════════════════════════════
# CHECK FOR UPDATE (version.json → fallback GitHub Releases API)
# ════════════════════════════════════════════════════════════
# version.json = manifest บน raw GitHub (เร็ว ไม่ติด rate limit 60/hr ของ API)
VERSION_JSON_URL = ("https://raw.githubusercontent.com/"
                    + GITHUB_REPO + "/main/version.json")

def _http_json(url, timeout=8):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT,
                                               "Accept": "application/vnd.github+json"})
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
        return json.loads(r.read().decode("utf-8"))

def _blank(current):
    return {"ok": False, "available": False, "latest": current, "notes": "",
            "published": "", "url": "", "error": "", "mandatory": False}

def check_version_json(current_version=APP_VERSION, url=None, channel="stable", timeout=8) -> dict:
    """ตรวจอัปเดตจาก version.json (วิธีหลัก — แนะนำ)"""
    out = _blank(current_version)
    try:
        u = url or VERSION_JSON_URL
        if channel == "beta":
            u = u.replace("version.json", "version-beta.json")
        data = _http_json(u, timeout)
        latest = data.get("version", "0.0.0")
        notes = data.get("notes", "")
        if isinstance(notes, list):
            notes = "\n".join("• " + str(n) for n in notes)
        out.update(latest=latest, notes=notes or "(ไม่มีรายละเอียด)",
                   published=data.get("published", ""), url=data.get("url", ""),
                   mandatory=bool(data.get("mandatory", False)),
                   available=is_newer(latest, current_version), ok=True)
    except urllib.error.HTTPError as e:
        out["error"] = f"version.json: HTTP {e.code}"
    except urllib.error.URLError as e:
        out["error"] = f"เชื่อมต่ออินเทอร์เน็ตไม่ได้: {e.reason}"
    except Exception as e:
        out["error"] = f"version.json: {e}"
    return out

def _check_via_api(current_version, repo, channel, timeout) -> dict:
    """fallback — GitHub Releases API"""
    out = _blank(current_version)
    try:
        rels = _http_json(f"https://api.github.com/repos/{repo}/releases", timeout)
        if isinstance(rels, dict):
            raise IOError(rels.get("message", "ไม่พบ releases"))
        cand = [r for r in rels if not r.get("draft")]
        if channel != "beta":
            cand = [r for r in cand if not r.get("prerelease")]
        if not cand:
            out["ok"] = True; out["error"] = "ยังไม่มี release"; return out
        rel = cand[0]
        latest = rel.get("tag_name") or rel.get("name") or "0.0.0"
        out.update(latest=latest, notes=rel.get("body") or "(ไม่มีรายละเอียด)",
                   published=(rel.get("published_at") or "")[:10])
        for a in rel.get("assets", []):
            if a.get("name", "").lower() == ASSET_NAME.lower():
                out["url"] = a.get("browser_download_url", ""); break
        if not out["url"]:
            for a in rel.get("assets", []):
                if a.get("name", "").lower().endswith(".zip"):
                    out["url"] = a.get("browser_download_url", ""); break
        out["available"] = is_newer(latest, current_version); out["ok"] = True
    except urllib.error.URLError as e:
        out["error"] = f"เชื่อมต่ออินเทอร์เน็ตไม่ได้: {e.reason}"
    except Exception as e:
        out["error"] = f"GitHub API: {e}"
    return out

def check_for_update(current_version=APP_VERSION, repo=GITHUB_REPO,
                     channel="stable", timeout=8) -> dict:
    """
    ตรวจอัปเดต: ลอง version.json ก่อน (เร็ว) → fallback GitHub Releases API
    คืน dict: {ok, available, latest, notes, published, url, mandatory, error}
    """
    rj = check_version_json(current_version, channel=channel, timeout=timeout)
    if rj["ok"]:
        return rj
    ra = _check_via_api(current_version, repo, channel, timeout)
    if ra["ok"]:
        return ra
    ra["error"] = ra["error"] or rj["error"] or "ตรวจสอบอัปเดตไม่สำเร็จ"
    return ra

# ════════════════════════════════════════════════════════════
# DOWNLOAD (มี progress + กัน partial/disconnect)
# ════════════════════════════════════════════════════════════
def download_update(url, dest_path, progress_cb=None, timeout=30) -> str:
    """ดาวน์โหลด url → dest_path. progress_cb(downloaded, total). ตรวจขนาดครบ."""
    tmp = dest_path + ".part"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    ctx = ssl.create_default_context()
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
            total = int(r.headers.get("Content-Length") or 0)
            done = 0
            with open(tmp, "wb") as f:
                while True:
                    chunk = r.read(65536)
                    if not chunk: break
                    f.write(chunk); done += len(chunk)
                    if progress_cb: progress_cb(done, total)
        if total and os.path.getsize(tmp) != total:
            raise IOError("ดาวน์โหลดไม่ครบ (ไฟล์ขนาดไม่ตรง)")
        # ตรวจว่าเป็น zip จริง
        if not zipfile.is_zipfile(tmp):
            raise IOError("ไฟล์ที่ดาวน์โหลดไม่ใช่ไฟล์ zip ที่ถูกต้อง")
        if os.path.exists(dest_path): os.remove(dest_path)
        os.rename(tmp, dest_path)
        return dest_path
    finally:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except Exception: pass

# ════════════════════════════════════════════════════════════
# BACKUP / APPLY / ROLLBACK
# ════════════════════════════════════════════════════════════
def backup_current(target_dir=None, exclude_ext=DB_EXCLUDE) -> str:
    """สำรองไฟล์โปรแกรมปัจจุบัน → /Backup/<วันที่_เวลา>/  (ยกเว้น *.db)"""
    ad = app_dir()
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    bdir = os.path.join(target_dir or os.path.join(ad, "Backup"), stamp)
    os.makedirs(bdir, exist_ok=True)
    for root, dirs, files in os.walk(ad):
        dirs[:] = [d for d in dirs if d not in DIR_EXCLUDE]
        for fn in files:
            if fn.lower().endswith(exclude_ext): continue
            src = os.path.join(root, fn)
            rel = os.path.relpath(src, ad)
            dst = os.path.join(bdir, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            try: shutil.copy2(src, dst)
            except Exception: pass
    return bdir

def _common_top(names) -> str:
    """ถ้าทุกไฟล์ใน zip อยู่ใต้โฟลเดอร์เดียวกัน (เช่น GitHub source zip) → คืนชื่อโฟลเดอร์นั้น"""
    tops = set()
    for n in names:
        n = n.replace("\\", "/")
        if not n or n.startswith("/"): continue
        tops.add(n.split("/", 1)[0] + "/")
    return tops.pop() if len(tops) == 1 else ""

def apply_update(zip_path, target_dir=None, exclude_ext=DB_EXCLUDE) -> int:
    """แตก zip ทับโปรแกรม (replace เฉพาะไฟล์ใน zip). คืนจำนวนไฟล์ที่เขียน.
    - กัน path traversal
    - ไม่แตะ *.db
    - strip โฟลเดอร์ครอบ (GitHub source zip) อัตโนมัติ
    """
    ad = os.path.normpath(target_dir or app_dir())
    written = 0
    with zipfile.ZipFile(zip_path) as z:
        bad = z.testzip()
        if bad: raise IOError(f"ไฟล์อัปเดตเสีย (corrupt): {bad}")
        names = z.namelist()
        strip = _common_top(names)
        for m in names:
            if m.endswith("/"): continue
            rel = m.replace("\\", "/")
            if strip and rel.startswith(strip):
                rel = rel[len(strip):]
            if not rel: continue
            target = os.path.normpath(os.path.join(ad, rel))
            if not (target == ad or target.startswith(ad + os.sep)):
                raise IOError(f"ปฏิเสธ path ผิดปกติ: {m}")     # path traversal
            if target.lower().endswith(exclude_ext):
                continue                                       # ห้ามทับ DB
            base = os.path.basename(target).lower()
            if base in ("rakupdater.pyw", "update_job.json"):
                continue                                       # ห้ามทับตัว updater ที่กำลังรัน
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with z.open(m) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            written += 1
    return written

def rollback(backup_dir, target_dir=None, exclude_ext=DB_EXCLUDE) -> int:
    """คืนไฟล์จาก backup_dir → โปรแกรม (ยกเว้น *.db). คืนจำนวนไฟล์ที่คืน."""
    ad = target_dir or app_dir()
    restored = 0
    for root, _dirs, files in os.walk(backup_dir):
        for fn in files:
            if fn.lower().endswith(exclude_ext): continue
            src = os.path.join(root, fn)
            rel = os.path.relpath(src, backup_dir)
            dst = os.path.join(ad, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            try: shutil.copy2(src, dst); restored += 1
            except Exception: pass
    return restored

# ════════════════════════════════════════════════════════════
# LOGGING
# ════════════════════════════════════════════════════════════
def log(msg, target_dir=None):
    ad = target_dir or app_dir()
    try:
        with open(os.path.join(ad, "update_log.txt"), "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass

# ════════════════════════════════════════════════════════════
# JOB FILE (ส่งงานให้ RakUpdater.pyw)
# ════════════════════════════════════════════════════════════
def write_job(zip_path, new_version, main_script="rakcomsoft.py", launcher="Rakcomsoft.bat"):
    """เขียน update_job.json ให้ RakUpdater อ่านหลังโปรแกรมปิด"""
    ad = app_dir()
    job = {
        "zip": zip_path, "app_dir": ad, "new_version": new_version,
        "from_version": APP_VERSION, "main_script": main_script,
        "launcher": launcher, "pid": os.getpid(),
        "created": datetime.now().isoformat(timespec="seconds"),
    }
    p = os.path.join(ad, "update_job.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(job, f, ensure_ascii=False, indent=2)
    return p


if __name__ == "__main__":
    print("RakComSoft Updater Engine")
    print("เวอร์ชันปัจจุบัน:", APP_VERSION, "| repo:", GITHUB_REPO)
    print("กำลังตรวจสอบอัปเดต...")
    r = check_for_update()
    print(json.dumps(r, ensure_ascii=False, indent=2))
