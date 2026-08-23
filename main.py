"""
我的相簿 App - MVP 後端
(OneDrive 背景掃描 + Google 相簿 Library API 日期區間全自動匯入 + SQLite 快取 + pHash 重複照片偵測與自動刪除
 + 自動分類 + 回憶影片 + 安全快速清理(按月分組+移動至雲端垃圾桶) + Apple 風格手機版 UI)
"""

import asyncio
import io
import json
import os
import re
import shutil
import sqlite3
import subprocess
import uuid
from collections import defaultdict
from datetime import datetime, timedelta
from urllib.parse import urlencode

import httpx
import imagehash
import msal
from dotenv import load_dotenv
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image
from starlette.middleware.sessions import SessionMiddleware

load_dotenv()

# ---------------------------------------------------------------------------
# OneDrive / Microsoft 設定(必填)
# ---------------------------------------------------------------------------
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
TENANT_ID = os.environ.get("TENANT_ID", "common")
REDIRECT_URI = os.environ.get("REDIRECT_URI", "http://localhost:8000/callback")
SESSION_SECRET = os.environ.get("SESSION_SECRET", "dev-secret-change-me")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.ReadWrite.All", "User.Read"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Google 相簿(Library API)設定(選填)
# ---------------------------------------------------------------------------
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET", "")
GOOGLE_REDIRECT_URI = os.environ.get("GOOGLE_REDIRECT_URI", "http://localhost:8000/google/callback")
GOOGLE_ENABLED = bool(GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET)

GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"
GOOGLE_LIBRARY_API = "https://photoslibrary.googleapis.com/v1"
GOOGLE_SCOPE = "https://www.googleapis.com/auth/photoslibrary.readonly"

BASE_DIR = os.path.dirname(__file__)
DB_PATH = os.path.join(BASE_DIR, "photos.db")
RENDER_DIR = os.path.join(BASE_DIR, "renders")
MUSIC_DIR = os.path.join(BASE_DIR, "music")
MEDIA_DIR = os.path.join(BASE_DIR, "media_cache")
GOOGLE_MEDIA_DIR = os.path.join(MEDIA_DIR, "google")
os.makedirs(RENDER_DIR, exist_ok=True)
os.makedirs(MUSIC_DIR, exist_ok=True)
os.makedirs(GOOGLE_MEDIA_DIR, exist_ok=True)

DEFAULT_DUP_THRESHOLD = 5
DEFAULT_EVENT_GAP_HOURS = 48
LOCATION_PRECISION = 2

# ---------------------------------------------------------------------------
# Apple 風格共用 UI
# ---------------------------------------------------------------------------
APPLE_CSS = """
<style>
:root {
    --bg: #f2f2f7; --card-bg: #ffffff; --text: #1c1c1e;
    --secondary: #8a8a8e; --accent: #0a84ff; --danger: #ff3b30;
    --success: #34c759; --warning: #ff9f0a; --border: rgba(60,60,67,0.13);
    --topbar-bg: rgba(249,249,249,0.82); --tabbar-bg: rgba(249,249,249,0.86);
    --radius: 14px; --safe-top: env(safe-area-inset-top, 0px); --safe-bottom: env(safe-area-inset-bottom, 0px);
}
@media (prefers-color-scheme: dark) {
    :root {
        --bg: #000000; --card-bg: #1c1c1e; --text: #f2f2f7;
        --secondary: #98989d; --border: rgba(84,84,88,0.45);
        --topbar-bg: rgba(20,20,22,0.78); --tabbar-bg: rgba(20,20,22,0.86);
    }
}
* { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
html, body {
    margin: 0; padding: 0; background: var(--bg); color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang TC", "Helvetica Neue", "Microsoft JhengHei", sans-serif;
    -webkit-font-smoothing: antialiased;
}
a { color: inherit; text-decoration: none; }
.app-shell { min-height: 100vh; display: flex; flex-direction: column; }
.topbar {
    position: sticky; top: 0; z-index: 50; padding-top: var(--safe-top);
    background: var(--topbar-bg); backdrop-filter: saturate(180%) blur(20px);
    -webkit-backdrop-filter: saturate(180%) blur(20px); border-bottom: 0.5px solid var(--border);
}
.topbar-inner { display: flex; align-items: center; height: 44px; padding: 0 12px; max-width: 640px; margin: 0 auto; position: relative; }
.topbar-back { color: var(--accent); font-size: 17px; padding: 6px 8px 6px 0; display: flex; align-items: center; gap: 2px; white-space: nowrap; }
.topbar-title { position: absolute; left: 50%; transform: translateX(-50%); font-size: 17px; font-weight: 600; max-width: 60%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.content { flex: 1; max-width: 640px; width: 100%; margin: 0 auto; padding: 16px 14px calc(96px + var(--safe-bottom)); }
.content.no-tabbar { padding-bottom: 28px; }
h1.page-h1 { font-size: 30px; font-weight: 700; margin: 6px 2px 16px; letter-spacing: -0.02em; }
h2.section-h2 { font-size: 20px; font-weight: 700; margin: 22px 2px 10px; letter-spacing: -0.01em; }
.section-title { font-size: 13px; color: var(--secondary); text-transform: uppercase; letter-spacing: 0.04em; margin: 20px 6px 8px; font-weight: 600; }
.secondary { color: var(--secondary); font-size: 14px; line-height: 1.5; }
.card { background: var(--card-bg); border-radius: var(--radius); padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
.card.banner-info { background: rgba(10,132,255,0.10); }
.card.banner-success { background: rgba(52,199,89,0.12); }
.card.banner-error { background: rgba(255,59,48,0.12); }
.card.banner-warning { background: rgba(255,159,10,0.14); }
.hero { text-align: center; padding: 48px 12px 24px; }
.hero-icon { font-size: 56px; margin-bottom: 10px; }
.hero h1 { font-size: 28px; font-weight: 700; margin: 0 0 8px; }
.list-group { background: var(--card-bg); border-radius: var(--radius); overflow: hidden; margin-bottom: 12px; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }
.list-row { display: flex; align-items: center; gap: 12px; padding: 13px 16px; border-bottom: 0.5px solid var(--border); font-size: 16px; min-height: 22px; }
.list-row:last-child { border-bottom: none; }
.list-row .row-icon { font-size: 20px; width: 26px; text-align: center; flex-shrink: 0; }
.list-row .row-label { flex: 1; }
.list-row .row-value { color: var(--secondary); font-size: 15px; }
.list-row .chevron { color: var(--secondary); font-size: 15px; }
.list-row.danger { color: var(--danger); }
.list-row.tappable:active { background: rgba(120,120,128,0.12); }
.btn-primary { display: block; width: 100%; text-align: center; background: var(--accent); color: #fff; font-size: 17px; font-weight: 600; padding: 14px 20px; border: none; border-radius: 980px; margin: 6px 0; cursor: pointer; }
.btn-primary:active { opacity: 0.75; }
.btn-secondary { display: block; width: 100%; text-align: center; background: rgba(10,132,255,0.12); color: var(--accent); font-size: 16px; font-weight: 600; padding: 12px 20px; border: none; border-radius: 980px; margin: 6px 0; cursor: pointer; }
.btn-row { display: flex; gap: 8px; }
.btn-row > * { flex: 1; }
select.apple-select, input.apple-input { width: 100%; padding: 10px 12px; border-radius: 10px; border: 0.5px solid var(--border); background: var(--bg); color: var(--text); font-size: 15px; margin: 8px 0; }
.pill-link { display: inline-block; background: rgba(10,132,255,0.12); color: var(--accent); font-size: 13px; font-weight: 600; padding: 6px 14px; border-radius: 980px; margin: 2px 4px 2px 0; }
.stat-num { font-size: 34px; font-weight: 700; letter-spacing: -0.02em; }
.photo-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(96px, 1fr)); gap: 2px; margin: 0 -2px 14px; }
.photo-tile { position: relative; aspect-ratio: 1 / 1; overflow: hidden; border-radius: 3px; background: var(--border); }
.photo-tile img { width: 100%; height: 100%; object-fit: cover; display: block; cursor: zoom-in; }
.photo-badge { position: absolute; left: 4px; bottom: 4px; font-size: 12px; background: rgba(0,0,0,0.55); color: #fff; border-radius: 6px; padding: 1px 5px; line-height: 1.4; }
.group-card { margin-bottom: 14px; }
.group-title { font-weight: 600; font-size: 15px; margin-bottom: 8px; }
.group-sub { color: var(--secondary); font-size: 12px; margin-bottom: 8px; }
.tabbar { position: fixed; left: 0; right: 0; bottom: 0; z-index: 50; display: flex; background: var(--tabbar-bg); backdrop-filter: saturate(180%) blur(20px); -webkit-backdrop-filter: saturate(180%) blur(20px); border-top: 0.5px solid var(--border); padding-bottom: var(--safe-bottom); }
.tabbar-item { flex: 1; text-align: center; padding: 8px 2px 6px; color: var(--secondary); display: flex; flex-direction: column; align-items: center; gap: 2px; }
.tabbar-item .tab-icon { font-size: 22px; line-height: 1; }
.tabbar-item .tab-label { font-size: 10px; font-weight: 500; }
.tabbar-item.active { color: var(--accent); }
form.inline-form { margin: 0; }
</style>
"""

LIGHTBOX_ASSETS = """
<style>
#lightbox-overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.92); z-index: 9999; align-items: center; justify-content: center; flex-direction: column; padding: calc(16px + env(safe-area-inset-top,0px)) 12px calc(16px + env(safe-area-inset-bottom,0px)); }
#lightbox-overlay.active { display: flex; }
#lightbox-img { width: 94vw; height: 76vh; max-width: 1100px; border-radius: 10px; object-fit: contain; }
#lightbox-caption { color: #fff; margin-top: 12px; font-size: 13px; text-align: center; opacity: 0.85; }
.lb-nav { position: fixed; top: 50%; transform: translateY(-50%); background: rgba(120,120,128,0.32); color: #fff; border: none; font-size: 22px; width: 42px; height: 42px; border-radius: 50%; cursor: pointer; }
#lb-prev { left: 12px; } #lb-next { right: 12px; }
#lb-close { position: fixed; top: calc(16px + env(safe-area-inset-top,0px)); right: 16px; color: #fff; font-size: 20px; background: rgba(120,120,128,0.32); border: none; width: 32px; height: 32px; border-radius: 50%; cursor: pointer; }
</style>
<div id="lightbox-overlay">
    <button id="lb-close" onclick="closeLightbox()">&times;</button>
    <button id="lb-prev" class="lb-nav" onclick="navLightbox(-1)">&#8249;</button>
    <img id="lightbox-img" src="" />
    <div id="lightbox-caption"></div>
    <button id="lb-next" class="lb-nav" onclick="navLightbox(1)">&#8250;</button>
</div>
<script>
let lbThumbs = []; let lbIndex = 0;
function initLightbox() { lbThumbs = Array.from(document.querySelectorAll('.lb-thumb')); lbThumbs.forEach((img, i) => { img.addEventListener('click', () => openLightbox(i)); }); }
function openLightbox(i) { lbIndex = i; showLightbox(); document.getElementById('lightbox-overlay').classList.add('active'); }
function showLightbox() { const t = lbThumbs[lbIndex]; document.getElementById('lightbox-img').src = t.dataset.full || t.src; document.getElementById('lightbox-caption').textContent = (t.dataset.name || '') + (t.dataset.taken ? '  ・  ' + t.dataset.taken : ''); }
function navLightbox(delta) { if (lbThumbs.length === 0) return; lbIndex = (lbIndex + delta + lbThumbs.length) % lbThumbs.length; showLightbox(); }
function closeLightbox() { document.getElementById('lightbox-overlay').classList.remove('active'); }
document.addEventListener('keydown', (e) => { if (!document.getElementById('lightbox-overlay').classList.contains('active')) return; if (e.key === 'Escape') closeLightbox(); if (e.key === 'ArrowLeft') navLightbox(-1); if (e.key === 'ArrowRight') navLightbox(1); });
document.addEventListener('DOMContentLoaded', initLightbox);
</script>
"""

TABS = [
    ("home", "/", "🏠", "首頁"),
    ("gallery", "/gallery", "🖼", "圖庫"),
    ("albums", "/albums", "🗂", "分類"),
    ("memories", "/memories", "✨", "回憶"),
    ("more", "/more", "⋯", "更多"),
]

def page_shell(
    title: str, body_html: str, active_tab: str | None = None, show_tabbar: bool = True, show_topbar: bool = True,
    back_href: str | None = None, meta_refresh: int | None = None, include_lightbox: bool = True,
) -> str:
    refresh_tag = f'<meta http-equiv="refresh" content="{meta_refresh}">' if meta_refresh else ""
    
    topbar_html = ""
    if show_topbar:
        back_html = f'<a class="topbar-back" href="{back_href}">‹ 返回</a>' if back_href else ""
        topbar_html = f"""
        <div class="topbar">
            <div class="topbar-inner">
                {back_html}
                <div class="topbar-title">{title}</div>
            </div>
        </div>
        """
        
    tabbar_html = ""
    if show_tabbar:
        items = "".join(f'<a class="tabbar-item{" active" if key == active_tab else ""}" href="{href}"><span class="tab-icon">{icon}</span><span class="tab-label">{label}</span></a>' for key, href, icon, label in TABS)
        tabbar_html = f'<nav class="tabbar">{items}</nav>'
        
    content_cls = "content" if show_tabbar else "content no-tabbar"
    
    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover" />
    <title>{title} - 我的相簿</title>
    {refresh_tag}
    {APPLE_CSS}
</head>
<body>
    <div class="app-shell">
        {topbar_html}
        <div class="{content_cls}">
            {body_html}
        </div>
        {tabbar_html}
    </div>
    {LIGHTBOX_ASSETS if include_lightbox else ""}
</body>
</html>
"""

def lb_img_tag(item: dict, badge: str | None = None) -> str:
    thumb = item.get("thumbnail_url") or item.get("thumbnailUrl") or ""
    full = item.get("thumbnail_large_url") or item.get("thumbnailLargeUrl") or item.get("web_url") or item.get("webUrl") or thumb
    name = (item.get("name") or "").replace('"', "&quot;")
    taken = (item.get("taken_date_time") or item.get("takenDateTime") or "").replace('"', "&quot;")
    if badge is None: badge = "G" if item.get("source") == "google" else ""
    badge_html = f'<div class="photo-badge">{badge}</div>' if badge else ""
    return f"""
    <div class="photo-tile">
        <img class="lb-thumb" src="{thumb}" data-full="{full}" data-name="{name}" data-taken="{taken}" loading="lazy" />
        {badge_html}
    </div>
    """

app = FastAPI(title="我的相簿 App")
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET)
app.mount("/media", StaticFiles(directory=MEDIA_DIR), name="media")

TOKEN_STORE: dict[str, dict] = {}
SCAN_STATUS: dict[str, dict] = {}
SCAN_TASKS: dict[str, asyncio.Task] = {}
MEMORY_JOBS: dict[str, dict] = {}
GOOGLE_IMPORT_STATUS: dict[str, dict] = {}

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS photos (
            sid TEXT NOT NULL, id TEXT NOT NULL, name TEXT, mime_type TEXT, size INTEGER, web_url TEXT, thumbnail_url TEXT, thumbnail_large_url TEXT, taken_date_time TEXT, latitude REAL, longitude REAL, phash TEXT, width INTEGER, height INTEGER,
            PRIMARY KEY (sid, id)
        )
    """)
    for column, col_type in [("phash", "TEXT"), ("width", "INTEGER"), ("height", "INTEGER"), ("thumbnail_large_url", "TEXT"), ("source", "TEXT DEFAULT 'onedrive'")]:
        try:
            conn.execute(f"ALTER TABLE photos ADD COLUMN {column} {col_type}")
        except sqlite3.OperationalError:
            pass
    conn.commit()
    conn.close()

init_db()

def _msal_app() -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(client_id=CLIENT_ID, client_credential=CLIENT_SECRET, authority=AUTHORITY)

def _store_ms_token(sid: str, result: dict):
    TOKEN_STORE.setdefault(sid, {})
    TOKEN_STORE[sid]["access_token"] = result["access_token"]
    if result.get("refresh_token"):
        TOKEN_STORE[sid]["refresh_token"] = result["refresh_token"]
    expires_in = result.get("expires_in", 3600)
    TOKEN_STORE[sid]["expires_at"] = datetime.utcnow() + timedelta(seconds=int(expires_in) - 120)

def _store_google_token(sid: str, data: dict):
    TOKEN_STORE.setdefault(sid, {})
    TOKEN_STORE[sid]["google_access_token"] = data["access_token"]
    if data.get("refresh_token"):
        TOKEN_STORE[sid]["google_refresh_token"] = data["refresh_token"]
    expires_in = data.get("expires_in", 3600)
    TOKEN_STORE[sid]["google_expires_at"] = datetime.utcnow() + timedelta(seconds=int(expires_in) - 120)

async def get_ms_token(sid: str | None) -> str | None:
    """回傳有效的 OneDrive access_token,過期前 2 分鐘自動用 refresh_token 換新。"""
    if not sid or sid not in TOKEN_STORE:
        return None
    entry = TOKEN_STORE[sid]
    if entry.get("access_token") and entry.get("expires_at") and datetime.utcnow() < entry["expires_at"]:
        return entry["access_token"]
    refresh_token = entry.get("refresh_token")
    if not refresh_token:
        return entry.get("access_token")
    result = _msal_app().acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)
    if "access_token" not in result:
        TOKEN_STORE.pop(sid, None)
        return None
    _store_ms_token(sid, result)
    return result["access_token"]

async def get_google_token(sid: str | None) -> str | None:
    """回傳有效的 Google access_token,過期前 2 分鐘自動用 refresh_token 換新。"""
    if not sid or sid not in TOKEN_STORE:
        return None
    entry = TOKEN_STORE[sid]
    if entry.get("google_access_token") and entry.get("google_expires_at") and datetime.utcnow() < entry["google_expires_at"]:
        return entry["google_access_token"]
    refresh_token = entry.get("google_refresh_token")
    if not refresh_token:
        return entry.get("google_access_token")
    async with httpx.AsyncClient(timeout=20) as client:
        resp = await client.post(GOOGLE_TOKEN_URL, data={
            "client_id": GOOGLE_CLIENT_ID, "client_secret": GOOGLE_CLIENT_SECRET,
            "refresh_token": refresh_token, "grant_type": "refresh_token",
        })
    data = resp.json()
    if "access_token" not in data:
        entry.pop("google_access_token", None)
        return None
    _store_google_token(sid, data)
    return data["access_token"]

# ---------------------------------------------------------------------------
# 首頁 / 更多頁
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    sid = request.session.get("sid")
    onedrive_connected = bool(sid and TOKEN_STORE.get(sid, {}).get("access_token"))
    google_connected = bool(sid and TOKEN_STORE.get(sid, {}).get("google_access_token"))

    if not onedrive_connected:
        body = """
        <div class="hero">
            <div class="hero-icon">📸</div>
            <h1>我的相簿</h1>
            <p class="secondary">整理 OneDrive、連結 Google 相簿,一站管理所有照片</p>
        </div>
        <a class="btn-primary" href="/login">用 Microsoft 帳號登入</a>
        """
        return HTMLResponse(page_shell("我的相簿", body, active_tab="home", show_tabbar=False))

    photo_count = len(db_get_photos(sid))
    
    if google_connected:
        google_row = """
        <a class="list-row tappable" href="/google/sync/options">
            <span class="row-icon">🟢</span>
            <span class="row-label">Google 相簿</span>
            <span class="row-value">已連結・同步設定</span>
            <span class="chevron">›</span>
        </a>
        """
    else:
        status_text = "尚未連結" if GOOGLE_ENABLED else "尚未設定"
        google_row = f"""
        <a class="list-row tappable" href="/google/login">
            <span class="row-icon">⚪️</span>
            <span class="row-label">Google 相簿</span>
            <span class="row-value">{status_text}</span>
            <span class="chevron">›</span>
        </a>
        """
    
    body = f"""
    <div class="card" style="text-align:center;">
        <div class="stat-num">{photo_count}</div>
        <div class="secondary">張照片已整理</div>
    </div>
    <div class="section-title">快速功能</div>
    <div class="list-group">
        <a class="list-row tappable" href="/gallery"><span class="row-icon">🖼</span><span class="row-label">我的圖庫</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/albums"><span class="row-icon">🗂</span><span class="row-label">自動分類</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/memories"><span class="row-icon">✨</span><span class="row-label">回憶影片</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/duplicates/view"><span class="row-icon">🧬</span><span class="row-label">疑似重複照片</span><span class="chevron">›</span></a>
    </div>
    <div class="section-title">帳號連結</div>
    <div class="list-group">
        <div class="list-row"><span class="row-icon">🔵</span><span class="row-label">OneDrive</span><span class="row-value">已連結</span></div>
        {google_row}
    </div>
    <div class="list-group">
        <a class="list-row tappable danger" href="/logout"><span class="row-icon">🚪</span><span class="row-label">登出</span></a>
    </div>
    """
    return HTMLResponse(page_shell("我的相簿", body, active_tab="home"))

@app.get("/more", response_class=HTMLResponse)
async def more_page(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login")
    
    google_connected = bool(TOKEN_STORE.get(sid, {}).get("google_access_token"))
    
    if google_connected:
        google_row = """
        <a class="list-row tappable" href="/google/sync/options">
            <span class="row-icon">🟢</span>
            <span class="row-label">從 Google 相簿自動同步</span>
            <span class="chevron">›</span>
        </a>
        """
    else:
        status_text = "" if GOOGLE_ENABLED else "尚未設定"
        google_row = f"""
        <a class="list-row tappable" href="/google/login">
            <span class="row-icon">⚪️</span>
            <span class="row-label">連結 Google 相簿</span>
            <span class="row-value">{status_text}</span>
            <span class="chevron">›</span>
        </a>
        """
        
    google_hint = "" if GOOGLE_ENABLED else """
    <p class="secondary" style="margin:0 6px 14px;">尚未設定 Google API 憑證,請參考程式檔案最上方的說明設定 GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET 後再連結。</p>
    """
    
    body = f"""
    <div class="list-group">
        <a class="list-row tappable" href="/duplicates/view"><span class="row-icon">🧬</span><span class="row-label">疑似重複照片</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/cleanup"><span class="row-icon">🧹</span><span class="row-label">快速清理(安全緩衝模式)</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/scan/start"><span class="row-icon">🔄</span><span class="row-label">重新掃描 OneDrive</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/photos?max_depth=0"><span class="row-icon">🧾</span><span class="row-label">原始照片清單(JSON)</span><span class="chevron">›</span></a>
    </div>
    <div class="section-title">Google 相簿</div>
    {google_hint}
    <div class="list-group">
        {google_row}
    </div>
    <div class="list-group">
        <a class="list-row tappable danger" href="/logout"><span class="row-icon">🚪</span><span class="row-label">登出</span></a>
    </div>
    """
    return HTMLResponse(page_shell("更多", body, active_tab="more", back_href="/"))

@app.get("/login")
async def login(request: Request):
    state = str(uuid.uuid4()); sid = request.session.get("sid") or str(uuid.uuid4())
    request.session["state"] = state; request.session["sid"] = sid
    TOKEN_STORE.setdefault(sid, {})
    auth_url = _msal_app().get_authorization_request_url(scopes=SCOPES, state=state, redirect_uri=REDIRECT_URI)
    return RedirectResponse(auth_url)

@app.get("/callback")
async def callback(request: Request, code: str | None = None, state: str | None = None, error: str | None = None, error_description: str | None = None):
    if error: return JSONResponse({"error": error, "description": error_description}, status_code=400)
    if not code or state != request.session.get("state"): return JSONResponse({"error": "invalid_state_or_missing_code"}, status_code=400)
    result = _msal_app().acquire_token_by_authorization_code(code=code, scopes=SCOPES, redirect_uri=REDIRECT_URI)
    if "access_token" not in result: return JSONResponse({"error": result.get("error"), "description": result.get("error_description")}, status_code=400)
    sid = request.session.get("sid")
    if not sid: return JSONResponse({"error": "missing_session"}, status_code=400)
    _store_ms_token(sid, result)
    await restore_db_from_onedrive(sid, result["access_token"])
    return RedirectResponse("/")

@app.get("/logout")
async def logout(request: Request):
    sid = request.session.get("sid")
    if sid: TOKEN_STORE.pop(sid, None)
    request.session.clear()
    return RedirectResponse("/")

# ---------------------------------------------------------------------------
# Google 相簿與 OneDrive 上傳機制
# ---------------------------------------------------------------------------
async def upload_to_onedrive(client: httpx.AsyncClient, token: str, filename: str, content: bytes) -> str | None:
    url = f"{GRAPH_BASE}/me/drive/root:/GooglePhotosImport/{filename}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }
    try:
        resp = await client.put(url, headers=headers, content=content, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return data.get("id")
    except Exception as e:
        print(f"上傳 {filename} 到 OneDrive 失敗: {e}")
        return None

@app.get("/google/login")
async def google_login(request: Request):
    if not GOOGLE_ENABLED: 
        return HTMLResponse(page_shell("錯誤", "<div class='card banner-error'><p class='secondary'>尚未設定 Google API 憑證，請檢查 .env 檔案中的 GOOGLE_CLIENT_ID 與 GOOGLE_CLIENT_SECRET。</p></div>", active_tab="more", back_href="/more"), status_code=400)
    
    sid = request.session.get("sid") or str(uuid.uuid4())
    request.session["sid"] = sid; state = str(uuid.uuid4()); request.session["google_state"] = state
    params = {"client_id": GOOGLE_CLIENT_ID, "redirect_uri": GOOGLE_REDIRECT_URI, "response_type": "code", "scope": GOOGLE_SCOPE, "access_type": "offline", "prompt": "consent", "state": state}
    return RedirectResponse(f"{GOOGLE_AUTH_URL}?{urlencode(params)}")

@app.get("/google/callback")
async def google_callback(request: Request, code: str | None = None, state: str | None = None, error: str | None = None):
    if error: return JSONResponse({"error": error}, status_code=400)
    if not code or state != request.session.get("google_state"): return JSONResponse({"error": "invalid_state_or_missing_code"}, status_code=400)
    async with httpx.AsyncClient(timeout=20) as client:
        resp = await client.post(GOOGLE_TOKEN_URL, data={"code": code, "client_id": GOOGLE_CLIENT_ID, "client_secret": GOOGLE_CLIENT_SECRET, "redirect_uri": GOOGLE_REDIRECT_URI, "grant_type": "authorization_code"})
    data = resp.json()
    if "access_token" not in data: return JSONResponse({"error": data.get("error"), "description": data.get("error_description")}, status_code=400)
    
    sid = request.session.get("sid")
    if not sid: return JSONResponse({"error": "missing_session"}, status_code=400)
    _store_google_token(sid, data)
    return RedirectResponse("/google/sync/options")

@app.get("/google/sync/options", response_class=HTMLResponse)
async def google_sync_options(request: Request):
    sid = request.session.get("sid")
    token = await get_google_token(sid)
    if not token: return RedirectResponse("/google/login")
    
    current_year = datetime.now().year
    default_start = f"{current_year}-01-01"
    default_end = f"{current_year}-12-31"

    body = f"""
    <div class="card">
        <h2 style="margin-top:0;">自動同步設定</h2>
        <p class="secondary">請設定要從 Google 相簿抓取的日期區間，防止大量資料拖垮系統。留空代表同步所有照片（不建議）。</p>
        <form method="post" action="/google/sync/start">
            <div style="margin-bottom: 12px;">
                <label class="section-title" style="margin-left: 0;">起始日期</label>
                <input class="apple-input" type="date" name="start_date" value="{default_start}">
            </div>
            <div style="margin-bottom: 12px;">
                <label class="section-title" style="margin-left: 0;">結束日期</label>
                <input class="apple-input" type="date" name="end_date" value="{default_end}">
            </div>
            <button class="btn-primary" type="submit" style="margin-top:20px;">開始背景同步</button>
        </form>
    </div>
    """
    return page_shell("設定 Google 同步", body, active_tab="more", back_href="/more")

@app.post("/google/sync/start")
async def google_sync_start(request: Request, start_date: str = Form(""), end_date: str = Form("")):
    sid = request.session.get("sid")
    google_token = await get_google_token(sid)
    onedrive_token = await get_ms_token(sid)
    if not google_token: return RedirectResponse("/google/login")

    GOOGLE_IMPORT_STATUS[sid] = {"status": "importing", "count": 0, "total": "計算中..."}
    asyncio.create_task(run_google_auto_import(sid, google_token, onedrive_token, start_date, end_date))
    return RedirectResponse("/gallery", status_code=303)

def google_media_fs_path(url_path: str) -> str:
    prefix = "/media/"
    if url_path.startswith(prefix): return os.path.join(MEDIA_DIR, url_path[len(prefix):])
    return url_path

async def import_one_google_item(client: httpx.AsyncClient, sid: str, google_token: str, onedrive_token: str, media_item: dict):
    mi_id = media_item.get("id")
    base_url = media_item.get("baseUrl")
    if not mi_id or not base_url: return

    mime = media_item.get("mimeType", "image/jpeg")
    filename = media_item.get("filename", f"{mi_id}.jpg")
    meta = media_item.get("mediaMetadata", {})
    width = meta.get("width")
    height = meta.get("height")
    taken = meta.get("creationTime")

    user_dir = os.path.join(GOOGLE_MEDIA_DIR, sid)
    os.makedirs(user_dir, exist_ok=True)
    thumb_fs_path = os.path.join(user_dir, f"{mi_id}_thumb.jpg")
    large_fs_path = os.path.join(user_dir, f"{mi_id}_large.jpg")
    headers = {"Authorization": f"Bearer {google_token}"}

    thumb_resp = await client.get(f"{base_url}=w480-h480-c", headers=headers, timeout=30)
    if thumb_resp.status_code == 200:
        with open(thumb_fs_path, "wb") as f: f.write(thumb_resp.content)
    try:
        large_resp = await client.get(f"{base_url}=w1600", headers=headers, timeout=30)
        large_resp.raise_for_status()
        with open(large_fs_path, "wb") as f: f.write(large_resp.content)
    except Exception:
        if os.path.exists(thumb_fs_path): shutil.copyfile(thumb_fs_path, large_fs_path)

    phash = None
    try:
        img = Image.open(thumb_fs_path).convert("RGB")
        phash = str(imagehash.phash(img))
    except Exception:
        pass

    download_param = "=dv" if mime.startswith("video/") else "=d"
    file_bytes = None
    try:
        orig_resp = await client.get(f"{base_url}{download_param}", headers=headers, timeout=60, follow_redirects=True)
        orig_resp.raise_for_status()
        file_bytes = orig_resp.content
    except Exception:
        pass

    new_onedrive_id = None
    if onedrive_token and file_bytes:
        new_onedrive_id = await upload_to_onedrive(client, onedrive_token, filename, file_bytes)

    if new_onedrive_id:
        item = {
            "id": new_onedrive_id, "name": filename, "mimeType": mime, "size": len(file_bytes) if file_bytes else None,
            "webUrl": None, "thumbnailUrl": f"/media/google/{sid}/{mi_id}_thumb.jpg",
            "thumbnailLargeUrl": f"/media/google/{sid}/{mi_id}_large.jpg", "takenDateTime": taken,
            "latitude": None, "longitude": None, "width": width, "height": height, "phash": phash, "source": "onedrive", 
        }
    else:
        item = {
            "id": f"google:{mi_id}", "name": filename, "mimeType": mime, "size": None, "webUrl": None,
            "thumbnailUrl": f"/media/google/{sid}/{mi_id}_thumb.jpg", "thumbnailLargeUrl": f"/media/google/{sid}/{mi_id}_large.jpg",
            "takenDateTime": taken, "latitude": None, "longitude": None, "width": width, "height": height, "phash": phash, "source": "google",
        }
    db_upsert_photo(sid, item)

async def run_google_auto_import(sid: str, google_token: str, onedrive_token: str, start_dt: str = "", end_dt: str = ""):
    try:
        media_items: list[dict] = []
        page_token = None
        req_body = {"pageSize": 100}
        if start_dt and end_dt:
            try:
                s_year, s_month, s_day = map(int, start_dt.split("-"))
                e_year, e_month, e_day = map(int, end_dt.split("-"))
                req_body["filters"] = {"dateFilter": {"ranges": [{"startDate": {"year": s_year, "month": s_month, "day": s_day}, "endDate": {"year": e_year, "month": e_month, "day": e_day}}]}}
            except Exception:
                pass
        
        async with httpx.AsyncClient(timeout=30) as client:
            while True:
                if page_token: req_body["pageToken"] = page_token
                resp = await client.post(f"{GOOGLE_LIBRARY_API}/mediaItems:search", headers={"Authorization": f"Bearer {google_token}"}, json=req_body)
                resp.raise_for_status()
                data = resp.json()
                fetched_items = data.get("mediaItems", [])
                if fetched_items: media_items.extend(fetched_items)
                GOOGLE_IMPORT_STATUS[sid] = {"status": "importing", "count": 0, "total": len(media_items)}
                page_token = data.get("nextPageToken")
                if not page_token: break

        if not media_items:
            GOOGLE_IMPORT_STATUS[sid] = {"status": "done", "count": 0, "total": 0}
            return

        count = 0
        async with httpx.AsyncClient(timeout=60) as client:
            for media_item in media_items:
                try: await import_one_google_item(client, sid, google_token, onedrive_token, media_item)
                except Exception: pass
                count += 1
                GOOGLE_IMPORT_STATUS[sid] = {"status": "importing", "count": count, "total": len(media_items)}

        GOOGLE_IMPORT_STATUS[sid] = {"status": "done", "count": count, "total": len(media_items)}
        await backup_db_to_onedrive(sid, onedrive_token)
    except Exception as e:
        import traceback
        traceback.print_exc()
        GOOGLE_IMPORT_STATUS[sid] = {"status": "error", "error": str(e)}

# ---------------------------------------------------------------------------
# OneDrive 掃描
# ---------------------------------------------------------------------------
async def compute_phash(thumbnail_url: str | None) -> str | None:
    if not thumbnail_url: return None
    try:
        async with httpx.AsyncClient(timeout=15) as client:
            resp = await client.get(thumbnail_url)
            resp.raise_for_status()
        img = Image.open(io.BytesIO(resp.content)).convert("RGB")
        return str(imagehash.phash(img))
    except Exception:
        return None

async def fetch_media_iter(token: str, folder_path: str = "root", depth: int = 0, max_depth: int = 15):
    if depth > max_depth: return
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/me/drive/{folder_path}/children?$top=200&$expand=thumbnails"
    
    async with httpx.AsyncClient(timeout=30) as client:
        while url:
            for attempt in range(5):
                resp = await client.get(url, headers=headers)
                if resp.status_code == 429:
                    wait_time = int(resp.headers.get("Retry-After", 3))
                    await asyncio.sleep(wait_time)
                    continue
                resp.raise_for_status()
                break
                
            data = resp.json()
            for item in data.get("value", []):
                if "folder" in item:
                    async for sub_item in fetch_media_iter(token, f"items/{item['id']}", depth + 1, max_depth): 
                        yield sub_item
                    continue
                
                file_info = item.get("file")
                if not file_info: continue
                mime = file_info.get("mimeType", "")
                if not (mime.startswith("image/") or mime.startswith("video/")): continue
                
                photo_meta = item.get("photo", {})
                location = item.get("location", {})
                image_meta = item.get("image", {})
                thumbs = item.get("thumbnails", [{}])[0] if item.get("thumbnails") else {}
                
                yield {
                    "id": item["id"], "name": item["name"], "mimeType": mime, "size": item.get("size"),
                    "webUrl": item.get("webUrl"), "thumbnailUrl": thumbs.get("medium", {}).get("url"),
                    "thumbnailLargeUrl": thumbs.get("large", {}).get("url"), "takenDateTime": photo_meta.get("takenDateTime"),
                    "latitude": location.get("latitude"), "longitude": location.get("longitude"),
                    "width": image_meta.get("width"), "height": image_meta.get("height"), "source": "onedrive",
                }
            url = data.get("@odata.nextLink")

async def fetch_all_media(token: str, folder_path: str = "root", depth: int = 0, max_depth: int = 6) -> list[dict]:
    items = []
    async for item in fetch_media_iter(token, folder_path, depth, max_depth): items.append(item)
    return items

def db_upsert_photo(sid: str, item: dict):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        INSERT INTO photos (sid, id, name, mime_type, size, web_url, thumbnail_url, thumbnail_large_url, taken_date_time, latitude, longitude, phash, width, height, source)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(sid, id) DO UPDATE SET
            name=excluded.name, mime_type=excluded.mime_type, size=excluded.size, web_url=excluded.web_url, thumbnail_url=excluded.thumbnail_url, thumbnail_large_url=excluded.thumbnail_large_url,
            taken_date_time=excluded.taken_date_time, latitude=excluded.latitude, longitude=excluded.longitude, phash=excluded.phash, width=excluded.width, height=excluded.height, source=excluded.source
        """,
        (sid, item["id"], item.get("name"), item.get("mimeType"), item.get("size"), item.get("webUrl"), item.get("thumbnailUrl"), item.get("thumbnailLargeUrl"), item.get("takenDateTime"), item.get("latitude"), item.get("longitude"), item.get("phash"), item.get("width"), item.get("height"), item.get("source", "onedrive")),
    )
    conn.commit()
    conn.close()

def db_get_photos(sid: str) -> list[dict]:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM photos WHERE sid = ? ORDER BY taken_date_time DESC", (sid,)).fetchall()
    conn.close()
    return [dict(row) for row in rows]

def db_restore_sid_rows(sid: str, rows: list[dict]):
    conn = sqlite3.connect(DB_PATH)
    for r in rows:
        conn.execute(
            """
            INSERT INTO photos (sid, id, name, mime_type, size, web_url, thumbnail_url, thumbnail_large_url, taken_date_time, latitude, longitude, phash, width, height, source)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(sid, id) DO UPDATE SET
                name=excluded.name, mime_type=excluded.mime_type, size=excluded.size, web_url=excluded.web_url, thumbnail_url=excluded.thumbnail_url, thumbnail_large_url=excluded.thumbnail_large_url,
                taken_date_time=excluded.taken_date_time, latitude=excluded.latitude, longitude=excluded.longitude, phash=excluded.phash, width=excluded.width, height=excluded.height, source=excluded.source
            """,
            (sid, r.get("id"), r.get("name"), r.get("mime_type"), r.get("size"), r.get("web_url"), r.get("thumbnail_url"), r.get("thumbnail_large_url"), r.get("taken_date_time"), r.get("latitude"), r.get("longitude"), r.get("phash"), r.get("width"), r.get("height"), r.get("source", "onedrive")),
        )
    conn.commit()
    conn.close()

# ---------------------------------------------------------------------------
# photos.db 備份到使用者自己的 OneDrive / 從 OneDrive 還原
# 注意:只會存取「該 sid 自己」的資料列,絕不會把別人的照片資料寫進任何人的 OneDrive
# ---------------------------------------------------------------------------
BACKUP_FOLDER_NAME = "MyAlbumApp_Backup"
BACKUP_FILENAME = "photos_backup.json"

async def backup_db_to_onedrive(sid: str, token: str | None):
    if not token: return
    rows = db_get_photos(sid)
    if not rows: return
    url = f"{GRAPH_BASE}/me/drive/root:/{BACKUP_FOLDER_NAME}/{BACKUP_FILENAME}:/content"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.put(url, headers=headers, content=json.dumps(rows, ensure_ascii=False).encode("utf-8"))
            resp.raise_for_status()
    except Exception as e:
        print(f"備份 photos.db 到 OneDrive 失敗(sid={sid}): {e}")

async def restore_db_from_onedrive(sid: str, token: str | None):
    if not token: return
    if db_get_photos(sid): return  # 本機已經有資料,不用還原,避免蓋掉更新的資料
    url = f"{GRAPH_BASE}/me/drive/root:/{BACKUP_FOLDER_NAME}/{BACKUP_FILENAME}:/content"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.get(url, headers=headers)
            if resp.status_code == 404: return  # 使用者還沒備份過,略過即可
            resp.raise_for_status()
        rows = json.loads(resp.content)
        if rows: db_restore_sid_rows(sid, rows)
    except Exception as e:
        print(f"從 OneDrive 還原 photos.db 失敗(sid={sid}): {e}")

async def run_background_scan(sid: str, token: str):
    SCAN_STATUS[sid] = {"status": "scanning", "count": 0}
    try:
        count = 0
        existing_photos = {p["id"]: p.get("phash") for p in db_get_photos(sid)}
        async for item in fetch_media_iter(token, max_depth=15):
            item_id = item["id"]
            if item_id in existing_photos and existing_photos[item_id]:
                item["phash"] = existing_photos[item_id]
            elif item["mimeType"].startswith("image/"):
                item["phash"] = await compute_phash(item.get("thumbnailUrl"))
            else:
                item["phash"] = None
                
            db_upsert_photo(sid, item)
            count += 1
            SCAN_STATUS[sid]["count"] = count
        SCAN_STATUS[sid] = {"status": "done", "count": count}
        await backup_db_to_onedrive(sid, token)
    except Exception as e:
        SCAN_STATUS[sid] = {"status": "error", "error": str(e)}

def start_scan_if_needed(sid: str, token: str):
    if SCAN_STATUS.get(sid, {}).get("status") == "scanning": return
    SCAN_TASKS[sid] = asyncio.create_task(run_background_scan(sid, token))

@app.get("/scan/start")
async def scan_start(request: Request):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not token: return JSONResponse({"error": "not_logged_in"}, status_code=401)
    start_scan_if_needed(sid, token)
    return RedirectResponse("/gallery")

# ---------------------------------------------------------------------------
# 重複照片偵測(pHash) 與 自動刪除機制 (含 429 防呆暫停)
# ---------------------------------------------------------------------------
def hamming_distance(hash1: str, hash2: str) -> int:
    return imagehash.hex_to_hash(hash1) - imagehash.hex_to_hash(hash2)

def find_duplicate_groups(items: list[dict], threshold: int = DEFAULT_DUP_THRESHOLD) -> list[list[dict]]:
    hashed_items = [item for item in items if item.get("phash")]
    n = len(hashed_items); parent = list(range(n))
    def find(x):
        while parent[x] != x: parent[x] = parent[parent[x]]; x = parent[x]
        return x
    def union(x, y):
        rx, ry = find(x), find(y)
        if rx != ry: parent[rx] = ry
    for i in range(n):
        for j in range(i + 1, n):
            if hamming_distance(hashed_items[i]["phash"], hashed_items[j]["phash"]) <= threshold: union(i, j)
    groups: dict[int, list[dict]] = {}
    for i in range(n): groups.setdefault(find(i), []).append(hashed_items[i])
    return [g for g in groups.values() if len(g) > 1]

@app.get("/duplicates/view", response_class=HTMLResponse)
async def duplicates_view(request: Request, threshold: int = DEFAULT_DUP_THRESHOLD):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login")
    
    groups = find_duplicate_groups(db_get_photos(sid), threshold=threshold)
    if not groups:
        body = """<div class="card"><p class="secondary">目前沒有偵測到疑似重複的照片(或者掃描/匯入還沒完成)。</p></div>"""
    else:
        rows = "".join(f"""
            <div class="card group-card">
                <div class="group-title">重複群組 #{idx}(共 {len(group)} 張)</div>
                <div class="photo-grid">{"".join(lb_img_tag(it) for it in group)}</div>
            </div>
            """ for idx, group in enumerate(groups, start=1))
        body = rows

    threshold_row = f"""
    <div class="card">
        <p class="secondary" style="margin:0 0 10px;">
            相似度門檻:<b>{threshold}</b>。門檻越小越嚴格(只抓幾乎一模一樣的),越大越寬鬆(可能抓到只是「很像」的不同照片)。
        </p>
        <div class="btn-row">
            <a class="pill-link" href="/duplicates/view?threshold=2">更嚴格(2)</a>
            <a class="pill-link" href="/duplicates/view?threshold={DEFAULT_DUP_THRESHOLD}">預設({DEFAULT_DUP_THRESHOLD})</a>
            <a class="pill-link" href="/duplicates/view?threshold=10">更寬鬆(10)</a>
        </div>
        <form action="/duplicates/auto-clean" method="post" onsubmit="return confirm('確定要自動刪除所有重複照片嗎？這個動作會直接刪除 OneDrive 雲端檔案，無法復原！');">
            <input type="hidden" name="threshold" value="{threshold}" />
            <button type="submit" class="btn-primary" style="background: var(--danger); margin-top: 14px;">
                自動清除所有重複照片 (每組保留一份)
            </button>
        </form>
    </div>
    """
    return HTMLResponse(page_shell("疑似重複照片", threshold_row + body, active_tab="more", back_href="/more"))

@app.post("/duplicates/auto-clean")
async def auto_clean_duplicates(request: Request, threshold: int = Form(DEFAULT_DUP_THRESHOLD)):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token: return RedirectResponse("/login")

    groups = find_duplicate_groups(db_get_photos(sid), threshold=threshold)
    headers = {"Authorization": f"Bearer {token}"}
    
    async with httpx.AsyncClient(timeout=30) as client:
        for group in groups:
            for d_item in group[1:]:
                if d_item.get("source") == "onedrive":
                    del_url = f"{GRAPH_BASE}/me/drive/items/{d_item['id']}"
                    for attempt in range(5):
                        try:
                            resp = await client.delete(del_url, headers=headers)
                            if resp.status_code == 429:
                                wait_time = int(resp.headers.get("Retry-After", 3))
                                await asyncio.sleep(wait_time)
                                continue
                            resp.raise_for_status()
                            break
                        except Exception:
                            pass
                
                conn = sqlite3.connect(DB_PATH)
                conn.execute("DELETE FROM photos WHERE sid = ? AND id = ?", (sid, d_item['id']))
                conn.commit()
                conn.close()
                await asyncio.sleep(0.3)

    await backup_db_to_onedrive(sid, token)
    return RedirectResponse(f"/duplicates/view?threshold={threshold}", status_code=303)

# ---------------------------------------------------------------------------
# 快速清理 (安全模式：按月分組 + 移至 OneDrive 待清理資料夾)
# ---------------------------------------------------------------------------
SCREENSHOT_NAME_PATTERN = re.compile(r"(screenshot|screen[\s_-]?shot|截圖|截图)", re.IGNORECASE)
COMMON_SCREEN_RATIOS = [9 / 16, 16 / 9, 3 / 4, 4 / 3, 9 / 19.5, 19.5 / 9, 1.0]

def is_screenshot(item: dict) -> bool:
    if SCREENSHOT_NAME_PATTERN.search(item.get("name") or ""): return True
    w, h = item.get("width"), item.get("height")
    if w and h and not item.get("taken_date_time"):
        if any(abs((w / h) - r) < 0.03 for r in COMMON_SCREEN_RATIOS): return True
    return False

def get_cleanup_items(items: list[dict]) -> dict[str, list[dict]]:
    screenshots = []
    low_quality = []
    for it in items:
        if (it.get("mime_type") or "").startswith("video/"): continue
        if is_screenshot(it):
            screenshots.append(it)
            continue
        size = it.get("size")
        is_small_file = (size is not None and size < 102400)
        w, h = it.get("width"), it.get("height")
        is_low_res = (w is not None and h is not None and w < 800 and h < 800)
        if is_small_file or is_low_res:
            low_quality.append(it)
    return {"screenshots": screenshots, "low_quality": low_quality}

@app.get("/cleanup", response_class=HTMLResponse)
async def cleanup_view(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login")
    
    items = db_get_photos(sid)
    cleanup_data = get_cleanup_items(items)
    
    def group_by_month(item_list):
        grouped = defaultdict(list)
        for it in item_list:
            dt = parse_taken(it)
            month_key = dt.strftime("%Y-%m") if dt else "未知日期"
            grouped[month_key].append(it)
        return grouped

    grouped_screenshots = group_by_month(cleanup_data["screenshots"])
    grouped_low_quality = group_by_month(cleanup_data["low_quality"])
    
    def build_month_sections(groups_dict, icon, title_prefix):
        html = ""
        for month_key in sorted(groups_dict.keys(), reverse=True):
            group_items = groups_dict[month_key]
            thumbs = "".join(lb_img_tag(it) for it in group_items)
            ids_str = ",".join(it["id"] for it in group_items)
            
            html += f"""
            <div class="card group-card">
                <div class="group-title">{icon} {title_prefix} - {month_key} (共 {len(group_items)} 張)</div>
                <div class="group-sub" style="color: var(--warning); font-weight: bold;">向下滑動確認是否全為廢圖 👇</div>
                <div class="photo-grid" style="margin-bottom: 12px;">{thumbs}</div>
                <form action="/cleanup/batch-move" method="post" onsubmit="return confirm('確定要將 {month_key} 的這 {len(group_items)} 張照片，移至 OneDrive 的「待清理資料夾」嗎？');">
                    <input type="hidden" name="ids" value="{ids_str}" />
                    <button type="submit" class="btn-primary" style="background: var(--warning); color: #000; margin-top: 4px;">
                        📦 移至「待清理資料夾」
                    </button>
                </form>
            </div>
            """
        return html

    body = f"""
    <div class="card banner-success">
        <h3 style="margin:0 0 8px 0;">🛡️ 啟用安全清理模式</h3>
        <p class="secondary" style="margin:0;">
            <b>1. 分批預覽：</b>每次只處理單個月份，保證能看完全部照片不怕遺漏。<br>
            <b>2. 安全緩衝區：</b>照片不會被刪除，而是移至 OneDrive 的 <b>「我的相簿_待清理垃圾桶」</b> 資料夾。確認沒問題後，再自行去 OneDrive 清空即可。
        </p>
    </div>
    """
    
    if not cleanup_data["screenshots"] and not cleanup_data["low_quality"]:
        body += """<div class="card"><p class="secondary">太棒了！目前圖庫裡沒有發現截圖或低畫質的垃圾照片。</p></div>"""
    else:
        if cleanup_data["screenshots"]:
            body += f"""<div class="section-title">螢幕截圖分批清理</div>"""
            body += build_month_sections(grouped_screenshots, "📱", "截圖")
        if cleanup_data["low_quality"]:
            body += f"""<div class="section-title">低畫質小檔案分批清理</div>"""
            body += build_month_sections(grouped_low_quality, "🗑️", "小檔案")
            
    return HTMLResponse(page_shell("安全快速清理", body, active_tab="more", back_href="/more"))

@app.post("/cleanup/batch-move")
async def cleanup_batch_move(request: Request, ids: str = Form(...)):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token: return RedirectResponse("/login")

    id_list = [i for i in ids.split(",") if i]
    if not id_list:
        return RedirectResponse("/cleanup", status_code=303)
    
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    async with httpx.AsyncClient(timeout=30) as client:
        folder_name = "我的相簿_待清理垃圾桶"
        folder_id = None
        
        resp = await client.get(f"{GRAPH_BASE}/me/drive/root:/{folder_name}", headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200:
            folder_id = resp.json().get("id")
        else:
            create_payload = {
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename"
            }
            create_resp = await client.post(f"{GRAPH_BASE}/me/drive/root/children", headers=headers, json=create_payload)
            if create_resp.status_code in (200, 201):
                folder_id = create_resp.json().get("id")
                
        if not folder_id:
            return HTMLResponse("建立或取得垃圾桶資料夾失敗，請稍後再試或檢查 API 權限。", status_code=500)
            
        for d_id in id_list:
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            photo = conn.execute("SELECT source FROM photos WHERE sid = ? AND id = ?", (sid, d_id)).fetchone()
            
            if photo and photo["source"] == "onedrive":
                patch_url = f"{GRAPH_BASE}/me/drive/items/{d_id}"
                payload = {
                    "parentReference": {
                        "id": folder_id
                    }
                }
                for attempt in range(5):
                    try:
                        resp = await client.patch(patch_url, headers=headers, json=payload)
                        if resp.status_code == 429:
                            wait_time = int(resp.headers.get("Retry-After", 3))
                            await asyncio.sleep(wait_time)
                            continue
                        break
                    except Exception:
                        pass
                        
            conn.execute("DELETE FROM photos WHERE sid = ? AND id = ?", (sid, d_id))
            conn.commit()
            conn.close()
            await asyncio.sleep(0.3)

    await backup_db_to_onedrive(sid, token)
    return RedirectResponse("/cleanup", status_code=303)

# ---------------------------------------------------------------------------
# 自動分類
# ---------------------------------------------------------------------------
def parse_taken(item: dict) -> datetime | None:
    dt_str = item.get("taken_date_time") or item.get("takenDateTime")
    if not dt_str: return None
    try: return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except ValueError: return None

def cluster_events(items: list[dict], gap_hours: int = DEFAULT_EVENT_GAP_HOURS) -> list[dict]:
    dated = sorted([(dt, it) for it in items if (dt := parse_taken(it))], key=lambda x: x[0])
    events: list[dict] = []; current = None
    for dt, it in dated:
        if current is None or (dt - current["end"]) > timedelta(hours=gap_hours):
            current = {"start": dt, "end": dt, "items": [it]}
            events.append(current)
        else:
            current["items"].append(it); current["end"] = dt
    return sorted(events, key=lambda e: e["start"], reverse=True)

def cluster_by_location(items: list[dict], precision: int = LOCATION_PRECISION) -> dict[tuple, list[dict]]:
    buckets: dict[tuple, list[dict]] = {}
    for it in items:
        lat, lng = it.get("latitude"), it.get("longitude")
        if lat is not None and lng is not None: buckets.setdefault((round(lat, precision), round(lng, precision)), []).append(it)
    return buckets

def render_albums_html(items: list[dict]) -> str:
    screenshots = [it for it in items if is_screenshot(it)]
    videos = [it for it in items if (it.get("mime_type") or "").startswith("video/")]
    exclude_ids = {it["id"] for it in screenshots} | {it["id"] for it in videos}
    photos_only = [it for it in items if it["id"] not in exclude_ids]

    events = cluster_events(photos_only); location_buckets = cluster_by_location(photos_only)
    def thumb_row(group, limit=12): return "".join(lb_img_tag(it) for it in group[:limit] if it.get("thumbnail_url"))

    events_html = "".join(f"""
    <div class="card group-card">
        <div class="group-title">{ev["start"].strftime("%Y-%m-%d")}{"" if ev["start"].date() == ev["end"].date() else " ~ " + ev["end"].strftime("%Y-%m-%d")}</div>
        <div class="group-sub">共 {len(ev["items"])} 張</div>
        <div class="photo-grid">{thumb_row(ev["items"])}</div>
    </div>
    """ for ev in events) or """<div class="card"><p class="secondary">還沒有足夠的拍攝時間資料可以分組。</p></div>"""
    
    location_html = "".join(f"""
    <div class="card group-card">
        <div class="group-title">地點 #{idx}</div>
        <div class="group-sub">共 {len(group)} 張・<a class="pill-link" href="https://www.google.com/maps?q={coords[0]},{coords[1]}" target="_blank">在地圖上看</a></div>
        <div class="photo-grid">{thumb_row(group)}</div>
    </div>
    """ for idx, (coords, group) in enumerate(sorted(location_buckets.items(), key=lambda kv: -len(kv[1])), start=1)) or """<div class="card"><p class="secondary">目前沒有帶 GPS 座標的照片。</p></div>"""

    screenshot_empty = "" if screenshots else """<p class="secondary">目前沒有偵測到螢幕截圖。</p>"""
    video_empty = "" if videos else """<p class="secondary">目前沒有影片。</p>"""

    body = f"""
    <p class="secondary" style="margin:0 6px 14px;">全部用現有的拍攝時間 / GPS / 檔名資料分類,沒有呼叫任何外部 AI 服務。</p>
    <div class="section-title">螢幕截圖({len(screenshots)} 張)</div>
    <div class="card">
        <div class="photo-grid">{thumb_row(screenshots, limit=24) or ""}</div>
        {screenshot_empty}
    </div>
    <div class="section-title">影片({len(videos)} 支)</div>
    <div class="card">
        <div class="photo-grid">{thumb_row(videos, limit=24) or ""}</div>
        {video_empty}
    </div>
    <div class="section-title">依日期分的事件相簿({len(events)} 組)</div>
    {events_html}
    <div class="section-title">依地點分的相簿({len(location_buckets)} 組)</div>
    {location_html}
    """
    return page_shell("自動分類", body, active_tab="albums")

@app.get("/albums", response_class=HTMLResponse)
async def albums(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login")
    return HTMLResponse(render_albums_html(db_get_photos(sid)))

# ---------------------------------------------------------------------------
# 回憶影片
# ---------------------------------------------------------------------------
def list_music_files() -> list[str]:
    try: return sorted(f for f in os.listdir(MUSIC_DIR) if f.lower().endswith((".mp3", ".m4a", ".wav")))
    except FileNotFoundError: return []

def photos_on_this_day(items: list[dict]) -> dict[int, list[dict]]:
    today = datetime.now(); result: dict[int, list[dict]] = {}
    for it in items:
        dt = parse_taken(it)
        if dt and dt.month == today.month and dt.day == today.day and dt.year != today.year: result.setdefault(dt.year, []).append(it)
    return result

def prepare_frame(image_bytes: bytes, size: tuple[int, int]) -> Image.Image:
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    img.thumbnail(size, Image.LANCZOS)
    canvas = Image.new("RGB", size, (0, 0, 0))
    x = (size[0] - img.width) // 2; y = (size[1] - img.height) // 2
    canvas.paste(img, (x, y))
    return canvas

def build_xfade_video(image_paths: list[str], output_path: str, seconds_per_photo: float = 3.0, transition_seconds: float = 1.0, audio_path: str | None = None):
    n = len(image_paths)
    if n < 2: raise ValueError("至少需要兩張照片才能產生回憶影片")
    clip_len = seconds_per_photo + transition_seconds
    cmd = ["ffmpeg", "-y"]
    for p in image_paths: cmd += ["-loop", "1", "-t", f"{clip_len:.3f}", "-i", p]
    filter_parts = []; prev_label = "0:v"
    for i in range(1, n):
        offset = i * seconds_per_photo - (i - 1) * transition_seconds
        out_label = f"v{i}" if i < n - 1 else "vout"
        filter_parts.append(f"[{prev_label}][{i}:v]xfade=transition=fade:duration={transition_seconds:.3f}:offset={offset:.3f}[{out_label}]")
        prev_label = out_label
    cmd += ["-filter_complex", ";".join(filter_parts), "-map", f"[{prev_label}]"]
    if audio_path: cmd += ["-i", audio_path, "-map", f"{n}:a", "-shortest", "-c:a", "aac"]
    cmd += ["-r", "25", "-pix_fmt", "yuv420p", output_path]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0: raise RuntimeError(f"ffmpeg 執行失敗: {result.stderr[-2000:]}")

async def download_full_image(client: httpx.AsyncClient, token: str, item_id: str) -> bytes | None:
    try:
        resp = await client.get(f"{GRAPH_BASE}/me/drive/items/{item_id}/content", headers={"Authorization": f"Bearer {token}"}, timeout=30, follow_redirects=True)
        resp.raise_for_status()
        return resp.content
    except Exception: return None

async def get_full_image_bytes(client: httpx.AsyncClient, item: dict, token: str | None) -> bytes | None:
    if item.get("source") == "google":
        url_path = item.get("thumbnail_large_url") or item.get("thumbnail_url")
        if not url_path: return None
        fs_path = google_media_fs_path(url_path)
        if not os.path.isfile(fs_path): return None
        return await asyncio.to_thread(lambda: open(fs_path, "rb").read())
    if not token: return None
    return await download_full_image(client, token, item["id"])

async def render_memory_video(job_id: str, items: list[dict], token: str | None, music_path: str | None, title: str):
    MEMORY_JOBS[job_id] = {"status": "rendering", "progress": 0, "total": len(items), "title": title}
    tmp_dir = os.path.join(RENDER_DIR, job_id); os.makedirs(tmp_dir, exist_ok=True)
    try:
        frame_paths = []
        async with httpx.AsyncClient() as client:
            for idx, item in enumerate(items):
                if not (raw := await get_full_image_bytes(client, item, token)): continue
                try: frame = prepare_frame(raw, (1280, 720))
                except Exception: continue
                frame_path = os.path.join(tmp_dir, f"{idx:03d}.jpg")
                frame.save(frame_path, "JPEG", quality=90)
                frame_paths.append(frame_path); MEMORY_JOBS[job_id]["progress"] = idx + 1
        if len(frame_paths) < 2: raise ValueError("可下載到的照片不足兩張,無法產生回憶影片")
        output_path = os.path.join(RENDER_DIR, f"{job_id}.mp4")
        await asyncio.to_thread(build_xfade_video, frame_paths, output_path, seconds_per_photo=3.0, transition_seconds=1.0, audio_path=music_path)
        MEMORY_JOBS[job_id] = {"status": "done", "video_path": output_path, "title": title}
    except Exception as e:
        MEMORY_JOBS[job_id] = {"status": "error", "error": str(e), "title": title}
    finally: shutil.rmtree(tmp_dir, ignore_errors=True)

def render_memories_html(items: list[dict]) -> str:
    today_groups = photos_on_this_day(items); events = [e for e in cluster_events(items) if len(e["items"]) >= 4][:6]; music_files = list_music_files()
    
    def music_select(field_id: str) -> str:
        options_html = "".join(f"<option value='{f}'>{f}</option>" for f in music_files)
        return f"""
        <select class="apple-select" name="music" id="{field_id}">
            <option value="">不加配樂</option>
            {options_html}
        </select>
        """
        
    def render_group(group: list[dict], label: str, title: str, field_id: str) -> str:
        ids = ",".join(it["id"] for it in group[:20])
        thumbs = "".join(lb_img_tag(it) for it in group[:8] if it.get("thumbnail_url"))
        return f"""
        <div class="card group-card">
            <div class="group-title">{label}</div>
            <div class="group-sub">共 {len(group)} 張,最多取前 20 張做影片</div>
            <div class="photo-grid">{thumbs}</div>
            <form class="inline-form" method="post" action="/memories/render">
                <input type="hidden" name="ids" value="{ids}" />
                <input type="hidden" name="title" value="{title}" />
                {music_select(field_id)}
                <button class="btn-primary" type="submit">產生回憶影片</button>
            </form>
        </div>
        """

    if today_groups:
        today_section = "".join(render_group(group, f"{year} 年的今天", f"{year} 年的今天", f"music-{year}") for year, group in sorted(today_groups.items(), reverse=True))
    else:
        today_section = """<div class="card"><p class="secondary">目前還沒有找到「當年今日」的舊照片。</p></div>"""
        
    if events:
        events_section = "".join(render_group(ev["items"], f"{ev['start'].strftime('%Y-%m-%d')} 開始的事件", f"{ev['start'].strftime('%Y-%m-%d')} 的回憶", f"music-ev-{idx}") for idx, ev in enumerate(events))
    else:
        events_section = """<div class="card"><p class="secondary">目前還沒有偵測到照片數量夠多的事件(需要拍攝時間資料)。</p></div>"""
        
    music_hint = "" if music_files else """
    <p class="secondary" style="margin:0 6px 14px;">目前 music/ 資料夾裡沒有音檔,只能產生無配樂的影片;要加配樂的話把你自己的 mp3/m4a/wav 放進專案的 music/ 資料夾。</p>
    """
    
    body = f"""
    {music_hint}
    <div class="section-title">當年今日</div>
    {today_section}
    <div class="section-title">自動整理的相片事件</div>
    {events_section}
    """
    
    return page_shell("回憶影片", body, active_tab="memories")

@app.get("/memories", response_class=HTMLResponse)
async def memories(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login")
    return HTMLResponse(render_memories_html(db_get_photos(sid)))

@app.post("/memories/render")
async def memories_render(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login")
    token = await get_ms_token(sid)
    form = await request.form(); ids = [i for i in str(form.get("ids", "")).split(",") if i]
    if len(ids) < 2: return JSONResponse({"error": "photos_not_enough"}, status_code=400)
    all_photos = {p["id"]: p for p in db_get_photos(sid)}; items = [all_photos[i] for i in ids if i in all_photos]
    if len(items) < 2: return JSONResponse({"error": "photos_not_found"}, status_code=400)
    music = form.get("music"); music_path = os.path.join(MUSIC_DIR, str(music)) if music else None
    if music_path and not os.path.isfile(music_path): music_path = None
    job_id = str(uuid.uuid4()); title = str(form.get("title", "回憶影片"))
    MEMORY_JOBS[job_id] = {"status": "pending", "title": title}
    asyncio.create_task(render_memory_video(job_id, items, token, music_path, title))
    return RedirectResponse(f"/memories/status/{job_id}", status_code=303)

@app.get("/memories/status/{job_id}", response_class=HTMLResponse)
async def memories_status_page(job_id: str):
    job = MEMORY_JOBS.get(job_id)
    if not job: 
        error_html = """<div class="card"><p class="secondary">找不到這個影片工作。</p></div>"""
        return HTMLResponse(page_shell("回憶影片", error_html, active_tab="memories"), status_code=404)
        
    status = job.get("status")
    if status == "done":
        title_text = job.get("title", "")
        body = f"""
        <div class="card banner-success">
            <div class="group-title">影片已完成: {title_text}</div>
            <video controls style="width:100%; border-radius:10px; margin-top:8px;">
                <source src="/memories/video/{job_id}.mp4" type="video/mp4">
            </video>
            <a class="btn-primary" href="/memories/video/{job_id}.mp4" download>下載影片</a>
        </div>
        """
        refresh = None
    elif status == "error":
        error_msg = job.get("error")
        body = f"""
        <div class="card banner-error">
            <p class="secondary">產生失敗: {error_msg}</p>
        </div>
        """
        refresh = None
    else:
        progress = job.get("progress", 0); total = job.get("total", 0)
        progress_text = f"{progress} / {total}" if total else str(progress)
        body = f"""
        <div class="card banner-info" style="text-align:center;">
            <div class="stat-num">{progress_text}</div>
            <div class="secondary">影片產生中……(下載原圖 + 轉場運算需要一點時間)</div>
        </div>
        """
        refresh = 3
        
    return HTMLResponse(page_shell("回憶影片產生進度", body, active_tab="memories", back_href="/memories", meta_refresh=refresh))

@app.get("/memories/video/{job_id}.mp4")
async def memories_video(job_id: str):
    job = MEMORY_JOBS.get(job_id)
    if not job or job.get("status") != "done": return JSONResponse({"error": "not_ready"}, status_code=404)
    return FileResponse(job["video_path"], media_type="video/mp4")

# ---------------------------------------------------------------------------
# 圖庫清單
# ---------------------------------------------------------------------------
def render_gallery_html(items: list[dict], status: dict, google_status: dict) -> str:
    cards = "".join(lb_img_tag(item) for item in items if item.get("thumbnail_url"))
    scan_state = status.get("status", "idle"); scanned_count = status.get("count", 0); google_state = google_status.get("status", "idle")
    banners = ""
    
    if scan_state == "scanning": 
        banners += f"""<div class="card banner-warning">正在背景整理你的 OneDrive,目前已掃到 {scanned_count} 張……</div>"""
    elif scan_state == "error": 
        banners += f"""<div class="card banner-error">OneDrive 掃描時發生錯誤:{status.get("error")}</div>"""
    elif scan_state == "idle": 
        banners += """<div class="card banner-info">尚未開始整理 OneDrive。<a class="pill-link" href="/scan/start">開始掃描</a></div>"""

    if google_state == "importing": 
        google_count = google_status.get("count",0)
        google_total = google_status.get("total",0)
        banners += f"""<div class="card banner-warning">正在自動匯入 Google 相簿({google_count}/{google_total})……</div>"""

    refresh = 3 if scan_state == "scanning" or google_state in ("importing", ) else None
    google_href = "/google/sync/options" if GOOGLE_ENABLED else "/more"
    empty_state = "" if items else """
    <div class="card">
        <p class="secondary">目前圖庫是空的,掃描 OneDrive 或連結 Google 相簿後照片就會出現在這裡。</p>
    </div>
    """
    
    body = f"""
    {banners}
    <div class="btn-row" style="margin-bottom:14px;">
        <a class="btn-secondary" href="/scan/start">重新掃描 OneDrive</a>
        <a class="btn-secondary" href="{google_href}">同步 Google 相簿</a>
    </div>
    <div class="photo-grid">
        {cards}
    </div>
    {empty_state}
    """
    return page_shell(f"我的圖庫({len(items)})", body, active_tab="gallery", meta_refresh=refresh)

@app.get("/gallery", response_class=HTMLResponse)
async def gallery(request: Request):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token: return RedirectResponse("/login")
    status = SCAN_STATUS.get(sid, {"status": "idle", "count": 0})
    return HTMLResponse(render_gallery_html(db_get_photos(sid), status, GOOGLE_IMPORT_STATUS.get(sid, {"status": "idle"})))

@app.get("/photos")
async def photos(request: Request, max_depth: int = 6):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not token: return JSONResponse({"error": "not_logged_in", "hint": "先前往 /login"}, status_code=401)
    try: return {"count": len(items := await fetch_all_media(token, max_depth=max_depth)), "items": items}
    except httpx.HTTPStatusError as e: return JSONResponse({"error": "graph_api_error", "detail": e.response.text}, status_code=e.response.status_code)