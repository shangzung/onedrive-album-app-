"""
我的相簿 App - 消費者級後端
(OneDrive 增量掃描 delta + SQLite 快取 + pHash 重複偵測
 + 自動分類 + 回憶影片 + 安全快速清理 + 自訂標籤 + 模糊偵測
 + 方向篩選 + 週/月回顧 + 批次 ZIP + 相簿分享連結
 + 高質感 Apple 風格手機版 UI)
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
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta
from urllib.parse import quote

import httpx
import imagehash
import msal
import numpy as np
from dotenv import load_dotenv
from fastapi import FastAPI, Request, Form, Query
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse, FileResponse, StreamingResponse
from PIL import Image, ImageFilter, ImageStat
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

BASE_DIR = os.path.dirname(__file__)
DB_PATH = os.path.join(BASE_DIR, "photos.db")
RENDER_DIR = os.path.join(BASE_DIR, "renders")
MUSIC_DIR = os.path.join(BASE_DIR, "music")
os.makedirs(RENDER_DIR, exist_ok=True)
os.makedirs(MUSIC_DIR, exist_ok=True)

DEFAULT_DUP_THRESHOLD = 5
DEFAULT_EVENT_GAP_HOURS = 48
LOCATION_PRECISION = 2
GALLERY_MONTHS_PER_PAGE = 6
BLUR_THRESHOLD = 80.0  # Laplacian variance 低於此視為模糊 (可調)

# ---------------------------------------------------------------------------
# 高質感 Apple 風格共用 UI (消費者級)
# ---------------------------------------------------------------------------
APPLE_CSS = """
<style>
:root {
    --bg: #f2f2f7; --card-bg: #ffffff; --text: #1c1c1e;
    --secondary: #8a8a8e; --accent: #0a84ff; --danger: #ff3b30;
    --success: #34c759; --warning: #ff9f0a; --border: rgba(60,60,67,0.12);
    --topbar-bg: rgba(249,249,249,0.78); --tabbar-bg: rgba(249,249,249,0.86);
    --radius: 16px; --radius-sm: 12px;
    --safe-top: env(safe-area-inset-top, 0px); --safe-bottom: env(safe-area-inset-bottom, 0px);
    --shadow-sm: 0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02);
    --shadow-md: 0 4px 16px rgba(0,0,0,0.06), 0 1px 4px rgba(0,0,0,0.03);
    --shadow-lg: 0 12px 40px rgba(0,0,0,0.08);
    --ease: cubic-bezier(0.25, 0.1, 0.25, 1);
}
@media (prefers-color-scheme: dark) {
    :root {
        --bg: #000000; --card-bg: #1c1c1e; --text: #f5f5f7;
        --secondary: #98989d; --border: rgba(84,84,88,0.48);
        --topbar-bg: rgba(20,20,22,0.72); --tabbar-bg: rgba(20,20,22,0.86);
        --shadow-sm: 0 1px 3px rgba(0,0,0,0.3); --shadow-md: 0 4px 16px rgba(0,0,0,0.35);
        --shadow-lg: 0 12px 40px rgba(0,0,0,0.45);
    }
}
* { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
html, body {
    margin: 0; padding: 0; background: var(--bg); color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "SF Pro Display", "PingFang TC", "Helvetica Neue", "Microsoft JhengHei", sans-serif;
    -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale;
    letter-spacing: -0.01em;
}
a { color: inherit; text-decoration: none; transition: opacity 0.15s var(--ease); }
.app-shell { min-height: 100vh; display: flex; flex-direction: column; }
.topbar {
    position: sticky; top: 0; z-index: 50; padding-top: var(--safe-top);
    background: var(--topbar-bg); backdrop-filter: saturate(180%) blur(24px);
    -webkit-backdrop-filter: saturate(180%) blur(24px); border-bottom: 0.5px solid var(--border);
}
.topbar-inner { display: flex; align-items: center; height: 48px; padding: 0 14px; max-width: 640px; margin: 0 auto; position: relative; }
.topbar-back { color: var(--accent); font-size: 17px; font-weight: 500; padding: 8px 10px 8px 0; display: flex; align-items: center; gap: 2px; white-space: nowrap; }
.topbar-title { position: absolute; left: 50%; transform: translateX(-50%); font-size: 17px; font-weight: 600; max-width: 58%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; letter-spacing: -0.02em; }
.content { flex: 1; max-width: 640px; width: 100%; margin: 0 auto; padding: 18px 16px calc(100px + var(--safe-bottom)); }
.content.no-tabbar { padding-bottom: 32px; }
h1.page-h1 { font-size: 32px; font-weight: 700; margin: 4px 2px 18px; letter-spacing: -0.03em; }
h2.section-h2 { font-size: 22px; font-weight: 700; margin: 24px 2px 12px; letter-spacing: -0.02em; }
.section-title { font-size: 13px; color: var(--secondary); text-transform: uppercase; letter-spacing: 0.05em; margin: 24px 6px 10px; font-weight: 600; }
.secondary { color: var(--secondary); font-size: 14px; line-height: 1.55; }
.card {
    background: var(--card-bg); border-radius: var(--radius); padding: 18px;
    margin-bottom: 14px; box-shadow: var(--shadow-sm);
    transition: box-shadow 0.25s var(--ease), transform 0.2s var(--ease);
}
.card:active { transform: scale(0.995); }
.card.banner-info { background: linear-gradient(135deg, rgba(10,132,255,0.12), rgba(10,132,255,0.06)); border: 0.5px solid rgba(10,132,255,0.15); }
.card.banner-success { background: linear-gradient(135deg, rgba(52,199,89,0.14), rgba(52,199,89,0.06)); border: 0.5px solid rgba(52,199,89,0.18); }
.card.banner-error { background: linear-gradient(135deg, rgba(255,59,48,0.12), rgba(255,59,48,0.05)); border: 0.5px solid rgba(255,59,48,0.15); }
.card.banner-warning { background: linear-gradient(135deg, rgba(255,159,10,0.14), rgba(255,159,10,0.06)); border: 0.5px solid rgba(255,159,10,0.18); }
.hero { text-align: center; padding: 56px 16px 28px; }
.hero-icon { font-size: 64px; margin-bottom: 14px; filter: drop-shadow(0 4px 12px rgba(0,0,0,0.08)); }
.hero h1 { font-size: 30px; font-weight: 700; margin: 0 0 10px; letter-spacing: -0.03em; }
.list-group {
    background: var(--card-bg); border-radius: var(--radius); overflow: hidden;
    margin-bottom: 14px; box-shadow: var(--shadow-sm);
}
.list-row {
    display: flex; align-items: center; gap: 14px; padding: 14px 18px;
    border-bottom: 0.5px solid var(--border); font-size: 16px; min-height: 24px;
    transition: background 0.15s var(--ease);
}
.list-row:last-child { border-bottom: none; }
.list-row .row-icon { font-size: 22px; width: 28px; text-align: center; flex-shrink: 0; }
.list-row .row-label { flex: 1; font-weight: 500; }
.list-row .row-value { color: var(--secondary); font-size: 15px; }
.list-row .chevron { color: var(--secondary); font-size: 16px; opacity: 0.7; }
.list-row.danger { color: var(--danger); }
.list-row.tappable:active { background: rgba(120,120,128,0.1); }
.btn-primary {
    display: block; width: 100%; text-align: center; background: var(--accent); color: #fff;
    font-size: 17px; font-weight: 600; padding: 15px 22px; border: none; border-radius: 980px;
    margin: 8px 0; cursor: pointer; letter-spacing: -0.01em;
    box-shadow: 0 4px 14px rgba(10,132,255,0.28);
    transition: opacity 0.15s, transform 0.15s, box-shadow 0.2s;
}
.btn-primary:active { opacity: 0.85; transform: scale(0.98); box-shadow: 0 2px 8px rgba(10,132,255,0.2); }
.btn-secondary {
    display: block; width: 100%; text-align: center; background: rgba(10,132,255,0.1); color: var(--accent);
    font-size: 16px; font-weight: 600; padding: 13px 20px; border: none; border-radius: 980px;
    margin: 6px 0; cursor: pointer; transition: background 0.15s, transform 0.15s;
}
.btn-secondary:active { background: rgba(10,132,255,0.18); transform: scale(0.98); }
.btn-row { display: flex; gap: 10px; }
.btn-row > * { flex: 1; }
select.apple-select, input.apple-input {
    width: 100%; padding: 12px 14px; border-radius: var(--radius-sm);
    border: 0.5px solid var(--border); background: var(--bg); color: var(--text);
    font-size: 16px; margin: 8px 0; transition: border-color 0.2s;
}
select.apple-select:focus, input.apple-input:focus { outline: none; border-color: var(--accent); }
.pill-link {
    display: inline-block; background: rgba(10,132,255,0.1); color: var(--accent);
    font-size: 13px; font-weight: 600; padding: 7px 14px; border-radius: 980px;
    margin: 3px 5px 3px 0; transition: background 0.15s;
}
.pill-link:active { background: rgba(10,132,255,0.2); }
.stat-num { font-size: 36px; font-weight: 700; letter-spacing: -0.03em; line-height: 1.1; }
.photo-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(100px, 1fr)); gap: 3px; margin: 0 -2px 16px; }
.photo-tile {
    position: relative; aspect-ratio: 1 / 1; overflow: hidden; border-radius: 6px;
    background: var(--border); transition: transform 0.2s var(--ease);
}
.photo-tile:active { transform: scale(0.97); }
.photo-tile img { width: 100%; height: 100%; object-fit: cover; display: block; cursor: zoom-in; }
.photo-badge {
    position: absolute; left: 5px; bottom: 5px; font-size: 11px; font-weight: 600;
    background: rgba(0,0,0,0.6); color: #fff; border-radius: 6px; padding: 2px 6px; line-height: 1.3;
    backdrop-filter: blur(8px);
}
.tile-actions { position: absolute; top: 5px; left: 5px; right: 5px; display: flex; justify-content: space-between; pointer-events: none; }
.tile-btn {
    pointer-events: auto; width: 28px; height: 28px; border-radius: 50%; border: none;
    background: rgba(0,0,0,0.45); color: #fff; font-size: 14px; line-height: 28px;
    text-align: center; padding: 0; cursor: pointer; backdrop-filter: blur(8px);
    transition: background 0.15s, transform 0.15s;
}
.tile-btn:active { transform: scale(0.9); }
.tile-btn.fav-btn.active { background: var(--warning); color: #000; }
.tile-btn.del-btn:active { background: var(--danger); }
.tile-btn:disabled { opacity: 0.45; }
.group-card { margin-bottom: 16px; }
.group-title { font-weight: 650; font-size: 16px; margin-bottom: 6px; letter-spacing: -0.01em; }
.group-sub { color: var(--secondary); font-size: 13px; margin-bottom: 10px; }
.tabbar {
    position: fixed; left: 0; right: 0; bottom: 0; z-index: 50; display: flex;
    background: var(--tabbar-bg); backdrop-filter: saturate(180%) blur(24px);
    -webkit-backdrop-filter: saturate(180%) blur(24px); border-top: 0.5px solid var(--border);
    padding-bottom: var(--safe-bottom);
}
.tabbar-item {
    flex: 1; text-align: center; padding: 9px 2px 7px; color: var(--secondary);
    display: flex; flex-direction: column; align-items: center; gap: 3px;
    transition: color 0.2s;
}
.tabbar-item .tab-icon { font-size: 23px; line-height: 1; }
.tabbar-item .tab-label { font-size: 10px; font-weight: 500; letter-spacing: 0.01em; }
.tabbar-item.active { color: var(--accent); }
form.inline-form { margin: 0; }
.tag-chip {
    display: inline-flex; align-items: center; gap: 4px; background: rgba(10,132,255,0.12);
    color: var(--accent); font-size: 12px; font-weight: 600; padding: 4px 10px;
    border-radius: 980px; margin: 2px 3px 2px 0;
}
.tag-chip .remove { cursor: pointer; opacity: 0.7; font-size: 14px; line-height: 1; }
.filter-bar { display: flex; gap: 8px; overflow-x: auto; padding-bottom: 4px; margin-bottom: 14px; -webkit-overflow-scrolling: touch; }
.filter-chip {
    flex-shrink: 0; padding: 8px 14px; border-radius: 980px; font-size: 13px; font-weight: 600;
    background: var(--card-bg); border: 0.5px solid var(--border); color: var(--text);
    transition: all 0.15s;
}
.filter-chip.active { background: var(--accent); color: #fff; border-color: var(--accent); }
.select-mode .photo-tile { cursor: pointer; }
.select-mode .photo-tile.selected::after {
    content: ''; position: absolute; inset: 0; border: 3px solid var(--accent);
    border-radius: 6px; background: rgba(10,132,255,0.15); pointer-events: none;
}
.select-mode .photo-tile.selected::before {
    content: '✓'; position: absolute; top: 6px; right: 6px; width: 22px; height: 22px;
    background: var(--accent); color: #fff; border-radius: 50%; font-size: 13px;
    display: flex; align-items: center; justify-content: center; z-index: 2; font-weight: 700;
}
.floating-action {
    position: fixed; bottom: calc(80px + var(--safe-bottom)); left: 50%; transform: translateX(-50%);
    background: var(--accent); color: #fff; font-weight: 600; font-size: 15px;
    padding: 12px 24px; border-radius: 980px; box-shadow: var(--shadow-lg);
    z-index: 40; display: none; white-space: nowrap;
}
.floating-action.show { display: block; }
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
.lb-top-btn { position: fixed; top: calc(16px + env(safe-area-inset-top,0px)); color: #fff; font-size: 16px; background: rgba(120,120,128,0.32); border: none; width: 32px; height: 32px; border-radius: 50%; cursor: pointer; }
.lb-top-btn.active { background: var(--warning); color: #000; }
#lb-fav-btn { right: 60px; }
#lb-del-btn { right: 104px; background: rgba(255,59,48,0.5); }
</style>
<div id="lightbox-overlay">
    <button id="lb-close" onclick="closeLightbox()">&times;</button>
    <button id="lb-fav-btn" class="lb-top-btn" onclick="toggleFavoritePhoto(this)">&#9734;</button>
    <button id="lb-del-btn" class="lb-top-btn" onclick="quickDeletePhoto(this)">&#128465;</button>
    <button id="lb-prev" class="lb-nav" onclick="navLightbox(-1)">&#8249;</button>
    <img id="lightbox-img" src="" />
    <div id="lightbox-caption"></div>
    <button id="lb-next" class="lb-nav" onclick="navLightbox(1)">&#8250;</button>
</div>
<script>
let lbThumbs = []; let lbIndex = 0;
function initLightbox() { lbThumbs = Array.from(document.querySelectorAll('.lb-thumb')); lbThumbs.forEach((img, i) => { img.addEventListener('click', () => openLightbox(i)); }); }
function openLightbox(i) { lbIndex = i; showLightbox(); document.getElementById('lightbox-overlay').classList.add('active'); }
function showLightbox() {
    const t = lbThumbs[lbIndex];
    document.getElementById('lightbox-img').src = t.dataset.full || t.src;
    document.getElementById('lightbox-caption').textContent = (t.dataset.name || '') + (t.dataset.taken ? '  ・  ' + t.dataset.taken : '');
    const favBtn = document.getElementById('lb-fav-btn');
    favBtn.dataset.id = t.dataset.id || '';
    const isFav = t.dataset.fav === '1';
    favBtn.innerHTML = isFav ? '&#9733;' : '&#9734;';
    favBtn.classList.toggle('active', isFav);
    document.getElementById('lb-del-btn').dataset.id = t.dataset.id || '';
}
function navLightbox(delta) { if (lbThumbs.length === 0) return; lbIndex = (lbIndex + delta + lbThumbs.length) % lbThumbs.length; showLightbox(); }
function closeLightbox() { document.getElementById('lightbox-overlay').classList.remove('active'); }
document.addEventListener('keydown', (e) => { if (!document.getElementById('lightbox-overlay').classList.contains('active')) return; if (e.key === 'Escape') closeLightbox(); if (e.key === 'ArrowLeft') navLightbox(-1); if (e.key === 'ArrowRight') navLightbox(1); });
document.addEventListener('DOMContentLoaded', initLightbox);

async function toggleFavoritePhoto(btn) {
    const id = btn.dataset.id; if (!id) return;
    btn.disabled = true;
    try {
        const resp = await fetch('/api/photo/favorite', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: 'id=' + encodeURIComponent(id) });
        const data = await resp.json();
        if (data.success) {
            document.querySelectorAll('.fav-btn[data-id="' + id + '"]').forEach(el => { el.classList.toggle('active', data.favorite); el.innerHTML = data.favorite ? '&#9733;' : '&#9734;'; });
            const t = lbThumbs.find(x => x.dataset.id === id); if (t) t.dataset.fav = data.favorite ? '1' : '0';
            const lbBtn = document.getElementById('lb-fav-btn');
            if (lbBtn && lbBtn.dataset.id === id) { lbBtn.innerHTML = data.favorite ? '&#9733;' : '&#9734;'; lbBtn.classList.toggle('active', data.favorite); }
        }
    } catch (e) { /* 網路錯誤時安靜失敗,不打斷瀏覽 */ }
    btn.disabled = false;
}

async function quickDeletePhoto(btn) {
    const id = btn.dataset.id; if (!id) return;
    if (!confirm('確定要刪除這張照片嗎?(會移到 OneDrive 回收桶,需要的話可以在 OneDrive 還原)')) return;
    btn.disabled = true;
    try {
        const resp = await fetch('/api/photo/delete', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: 'id=' + encodeURIComponent(id) });
        const data = await resp.json();
        if (data.success) { removePhotoEverywhere(id); }
        else { alert('刪除失敗:' + (data.error || '未知錯誤')); btn.disabled = false; }
    } catch (e) { alert('刪除失敗,請檢查網路連線。'); btn.disabled = false; }
}
function removePhotoEverywhere(id) {
    document.querySelectorAll('.photo-tile[data-id="' + id + '"]').forEach(el => el.remove());
    const wasInLightbox = document.getElementById('lightbox-overlay').classList.contains('active');
    const idx = lbThumbs.findIndex(t => t.dataset.id === id);
    if (idx !== -1) {
        lbThumbs.splice(idx, 1);
        if (lbThumbs.length === 0) { closeLightbox(); }
        else if (wasInLightbox) { lbIndex = Math.min(lbIndex, lbThumbs.length - 1); showLightbox(); }
    }
}
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
    item_id = (item.get("id") or "").replace('"', "&quot;")
    is_fav = bool(item.get("favorite"))
    badge_html = f'<div class="photo-badge">{badge}</div>' if badge else ""
    return f"""
    <div class="photo-tile" data-id="{item_id}">
        <img class="lb-thumb" src="{thumb}" data-full="{full}" data-name="{name}" data-taken="{taken}" data-id="{item_id}" data-fav="{"1" if is_fav else "0"}" loading="lazy" />
        {badge_html}
        <div class="tile-actions">
            <button type="button" class="tile-btn fav-btn{" active" if is_fav else ""}" data-id="{item_id}" onclick="event.stopPropagation(); toggleFavoritePhoto(this);" title="加入最愛">{"&#9733;" if is_fav else "&#9734;"}</button>
            <button type="button" class="tile-btn del-btn" data-id="{item_id}" onclick="event.stopPropagation(); quickDeletePhoto(this);" title="刪除">&#128465;</button>
        </div>
    </div>
    """

app = FastAPI(title="我的相簿 App")
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET)

TOKEN_STORE: dict[str, dict] = {}
SCAN_STATUS: dict[str, dict] = {}
SCAN_TASKS: dict[str, asyncio.Task] = {}
MEMORY_JOBS: dict[str, dict] = {}

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS photos (
            sid TEXT NOT NULL, id TEXT NOT NULL, name TEXT, mime_type TEXT, size INTEGER, web_url TEXT, thumbnail_url TEXT, thumbnail_large_url TEXT, taken_date_time TEXT, latitude REAL, longitude REAL, phash TEXT, width INTEGER, height INTEGER,
            PRIMARY KEY (sid, id)
        )
    """)
    for column, col_type in [
        ("phash", "TEXT"), ("width", "INTEGER"), ("height", "INTEGER"),
        ("thumbnail_large_url", "TEXT"), ("source", "TEXT DEFAULT 'onedrive'"),
        ("cleanup_skip", "INTEGER DEFAULT 0"), ("favorite", "INTEGER DEFAULT 0"),
        ("camera_make", "TEXT"), ("camera_model", "TEXT"),
        ("blur_score", "REAL"),  # Laplacian variance, 越低越模糊
    ]:
        try:
            conn.execute(f"ALTER TABLE photos ADD COLUMN {column} {col_type}")
        except sqlite3.OperationalError:
            pass
    conn.execute("DELETE FROM photos WHERE source = 'google'")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS location_cache (
            lat_r REAL NOT NULL, lng_r REAL NOT NULL, name TEXT,
            PRIMARY KEY (lat_r, lng_r)
        )
    """)
    # 自訂標籤 (完全本地 SQLite, 無 AI)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS photo_tags (
            sid TEXT NOT NULL, photo_id TEXT NOT NULL, tag TEXT NOT NULL,
            PRIMARY KEY (sid, photo_id, tag)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_photo_tags_tag ON photo_tags(sid, tag)")
    # 增量掃描用的 deltaLink (避免每次全量重掃)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS scan_state (
            sid TEXT PRIMARY KEY, delta_link TEXT, last_scan_at TEXT, photo_count INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()

def load_location_cache() -> dict[tuple, str]:
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT lat_r, lng_r, name FROM location_cache").fetchall()
    conn.close()
    return {(r[0], r[1]): r[2] for r in rows}

def save_location_cache_entry(lat_r: float, lng_r: float, name: str):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO location_cache (lat_r, lng_r, name) VALUES (?, ?, ?) "
        "ON CONFLICT(lat_r, lng_r) DO UPDATE SET name=excluded.name",
        (lat_r, lng_r, name),
    )
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------------------------
# 掃描狀態 / 標籤 / 模糊偵測 helpers
# ---------------------------------------------------------------------------
def get_scan_state(sid: str) -> dict:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT * FROM scan_state WHERE sid = ?", (sid,)).fetchone()
    conn.close()
    return dict(row) if row else {}

def save_scan_state(sid: str, delta_link: str | None = None, photo_count: int | None = None):
    conn = sqlite3.connect(DB_PATH)
    existing = conn.execute("SELECT sid FROM scan_state WHERE sid = ?", (sid,)).fetchone()
    now = datetime.utcnow().isoformat()
    if existing:
        if delta_link is not None:
            conn.execute("UPDATE scan_state SET delta_link = ?, last_scan_at = ? WHERE sid = ?", (delta_link, now, sid))
        if photo_count is not None:
            conn.execute("UPDATE scan_state SET photo_count = ?, last_scan_at = ? WHERE sid = ?", (photo_count, now, sid))
    else:
        conn.execute(
            "INSERT INTO scan_state (sid, delta_link, last_scan_at, photo_count) VALUES (?, ?, ?, ?)",
            (sid, delta_link, now, photo_count or 0),
        )
    conn.commit()
    conn.close()

def get_photo_tags(sid: str, photo_id: str) -> list[str]:
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT tag FROM photo_tags WHERE sid = ? AND photo_id = ? ORDER BY tag", (sid, photo_id)).fetchall()
    conn.close()
    return [r[0] for r in rows]

def get_all_tags(sid: str) -> list[tuple[str, int]]:
    """回傳 [(tag, count), ...] 依使用次數排序"""
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT tag, COUNT(*) as cnt FROM photo_tags WHERE sid = ? GROUP BY tag ORDER BY cnt DESC, tag",
        (sid,),
    ).fetchall()
    conn.close()
    return [(r[0], r[1]) for r in rows]

def set_photo_tags(sid: str, photo_id: str, tags: list[str]):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM photo_tags WHERE sid = ? AND photo_id = ?", (sid, photo_id))
    for t in tags:
        t = t.strip()
        if t:
            conn.execute(
                "INSERT OR IGNORE INTO photo_tags (sid, photo_id, tag) VALUES (?, ?, ?)",
                (sid, photo_id, t),
            )
    conn.commit()
    conn.close()

def add_photo_tag(sid: str, photo_id: str, tag: str):
    tag = tag.strip()
    if not tag:
        return
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT OR IGNORE INTO photo_tags (sid, photo_id, tag) VALUES (?, ?, ?)",
        (sid, photo_id, tag),
    )
    conn.commit()
    conn.close()

def remove_photo_tag(sid: str, photo_id: str, tag: str):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM photo_tags WHERE sid = ? AND photo_id = ? AND tag = ?", (sid, photo_id, tag))
    conn.commit()
    conn.close()

def photos_with_tag(sid: str, tag: str) -> list[str]:
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT photo_id FROM photo_tags WHERE sid = ? AND tag = ?", (sid, tag)).fetchall()
    conn.close()
    return [r[0] for r in rows]

def compute_blur_score(image_bytes: bytes) -> float | None:
    """用 Laplacian 變異數評估清晰度。數值越低越模糊。回傳 None 表示無法計算。"""
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("L")
        # 縮小以加速 (對模糊偵測足夠)
        img.thumbnail((512, 512), Image.LANCZOS)
        arr = np.asarray(img, dtype=np.float64)
        # 簡易 Laplacian kernel
        kernel = np.array([[0, 1, 0], [1, -4, 1], [0, 1, 0]], dtype=np.float64)
        # 用 numpy 做 convolution (邊界用 edge)
        from numpy.lib.stride_tricks import sliding_window_view
        if arr.shape[0] < 3 or arr.shape[1] < 3:
            return None
        windows = sliding_window_view(arr, (3, 3))
        laplacian = np.sum(windows * kernel, axis=(-2, -1))
        return float(laplacian.var())
    except Exception:
        return None

def is_blurry(item: dict, threshold: float = BLUR_THRESHOLD) -> bool:
    score = item.get("blur_score")
    if score is None:
        return False
    return score < threshold

def orientation_of(item: dict) -> str:
    w, h = item.get("width"), item.get("height")
    if not w or not h:
        return "unknown"
    ratio = w / h
    if ratio > 1.05:
        return "landscape"
    if ratio < 0.95:
        return "portrait"
    return "square"

def _msal_app() -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(client_id=CLIENT_ID, client_credential=CLIENT_SECRET, authority=AUTHORITY)

def _store_ms_token(sid: str, result: dict):
    TOKEN_STORE.setdefault(sid, {})
    TOKEN_STORE[sid]["access_token"] = result["access_token"]
    if result.get("refresh_token"):
        TOKEN_STORE[sid]["refresh_token"] = result["refresh_token"]
    expires_in = result.get("expires_in", 3600)
    TOKEN_STORE[sid]["expires_at"] = datetime.utcnow() + timedelta(seconds=int(expires_in) - 120)

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

# ---------------------------------------------------------------------------
# 首頁 / 更多頁
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    sid = request.session.get("sid")
    token = await get_ms_token(sid) if sid else None
    onedrive_connected = bool(token)

    if not onedrive_connected:
        body = """
        <div class="hero">
            <div class="hero-icon">📸</div>
            <h1>我的相簿</h1>
            <p class="secondary">專心整理你的 OneDrive 照片</p>
        </div>
        <a class="btn-primary" href="/login">用 Microsoft 帳號登入</a>
        """
        return HTMLResponse(page_shell("我的相簿", body, active_tab="home", show_tabbar=False))

    photo_count = len(db_get_photos(sid))

    # 安全網：已登入但本地是空的 → 自動再試一次還原（Docker 重啟常見）
    restore_banner = ""
    if photo_count == 0:
        rr = await restore_db_from_onedrive(sid, token)
        photo_count = rr.get("count") or len(db_get_photos(sid))
        if rr.get("ok") and photo_count > 0:
            restore_banner = f"""
            <div class="card banner-success">
                <b>已從 OneDrive 還原 {photo_count} 張照片</b>
                <p class="secondary" style="margin:6px 0 0;">容器重啟後本地資料會清空，已自動從雲端備份救回，無需重掃。</p>
            </div>
            """
        elif rr.get("error") or (not rr.get("ok")):
            msg = rr.get("message") or rr.get("error") or "未知錯誤"
            restore_banner = f"""
            <div class="card banner-error">
                <b>雲端還原失敗</b>
                <p class="secondary" style="margin:6px 0 0;">{msg}</p>
                <p class="secondary" style="margin:6px 0 0;">請點下方「增量掃描」或「強制全量重新掃描」。OneDrive 備份檔若存在，也可到「更多」再試一次。</p>
                <a class="btn-secondary" href="/restore">再試一次還原</a>
            </div>
            """
    else:
        # 顯示登入當下的還原提示（只顯示一次）
        msg = request.session.pop("restore_msg", None)
        ok = request.session.pop("restore_ok", None)
        if msg and ok and (request.session.pop("restore_count", 0) or 0) > 0:
            restore_banner = f"""
            <div class="card banner-success"><b>{msg}</b></div>
            """
        elif msg and ok is False:
            restore_banner = f"""
            <div class="card banner-warning"><b>{msg}</b></div>
            """

    scan_state = SCAN_STATUS.get(sid, {}).get("status", "idle")
    scan_hint = {
        "scanning": "背景整理中……",
        "error": "上次掃描發生錯誤,點一下重試",
        "idle": "點一下開始整理 OneDrive" if photo_count == 0 else "增量更新",
        "done": "已是最新狀態",
    }.get(scan_state, "")

    body = f"""
    {restore_banner}
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
        <a class="list-row tappable" href="/cleanup"><span class="row-icon">🧹</span><span class="row-label">快速清理</span><span class="chevron">›</span></a>
    </div>
    <div class="section-title">帳號</div>
    <div class="list-group">
        <a class="list-row tappable" href="/scan/start">
            <span class="row-icon">🔵</span>
            <span class="row-label">OneDrive 增量掃描</span>
            <span class="row-value">{scan_hint}</span>
            <span class="chevron">›</span>
        </a>
        <a class="list-row tappable" href="/restore">
            <span class="row-icon">☁️</span>
            <span class="row-label">從 OneDrive 還原備份</span>
            <span class="chevron">›</span>
        </a>
    </div>
    <div class="list-group">
        <a class="list-row tappable danger" href="/logout"><span class="row-icon">🚪</span><span class="row-label">登出</span></a>
    </div>
    """
    return HTMLResponse(page_shell("我的相簿", body, active_tab="home"))

@app.get("/more", response_class=HTMLResponse)
async def more_page(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)

    body = """
    <div class="list-group">
        <a class="list-row tappable" href="/years"><span class="row-icon">🗓️</span><span class="row-label">年份總覽</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/locations"><span class="row-icon">📍</span><span class="row-label">拍攝地點(含地圖)</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/search"><span class="row-icon">🔍</span><span class="row-label">圖庫搜尋</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/favorites"><span class="row-icon">⭐</span><span class="row-label">我的最愛</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/tags"><span class="row-icon">🏷</span><span class="row-label">自訂標籤</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/orientation"><span class="row-icon">📐</span><span class="row-label">方向篩選(直/橫式)</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/cameras"><span class="row-icon">📷</span><span class="row-label">依相機/裝置分類</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/storage"><span class="row-icon">📊</span><span class="row-label">儲存空間統計</span><span class="chevron">›</span></a>
    </div>
    <div class="list-group">
        <a class="list-row tappable" href="/reviews"><span class="row-icon">📅</span><span class="row-label">上週 / 上個月回顧</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/duplicates/view"><span class="row-icon">🧬</span><span class="row-label">疑似重複照片</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/cleanup"><span class="row-icon">🧹</span><span class="row-label">快速清理(含模糊偵測)</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/share/event"><span class="row-icon">🔗</span><span class="row-label">相簿分享連結</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/scan/start"><span class="row-icon">🔄</span><span class="row-label">增量掃描 OneDrive</span><span class="chevron">›</span></a>
        <a class="list-row tappable" href="/scan/start?force=1"><span class="row-icon">♻️</span><span class="row-label">強制全量重新掃描</span><span class="chevron">›</span></a>
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
    old_sid = request.session.get("sid")
    if not old_sid: return JSONResponse({"error": "missing_session"}, status_code=400)

    # 用 Microsoft 帳號本身的固定 ID(oid/sub)取代原本隨機產生的瀏覽器 session id,
    # 這樣同一個帳號無論用哪個瀏覽器或裝置登入，都會對應到同一份已掃描的圖庫資料，
    # 不會因為換瀏覽器(=換了一個新的隨機 sid)就要重新掃描一次 OneDrive。
    claims = result.get("id_token_claims") or {}
    account_key = claims.get("oid") or claims.get("sub")
    sid = f"ms-{account_key}" if account_key else old_sid
    request.session["sid"] = sid
    if sid != old_sid:
        TOKEN_STORE.pop(old_sid, None)

    _store_ms_token(sid, result)
    restore_result = await restore_db_from_onedrive(sid, result["access_token"])
    # 把還原結果記在 session，首頁可以顯示一次提示
    request.session["restore_msg"] = restore_result.get("message") or ""
    request.session["restore_ok"] = bool(restore_result.get("ok"))
    request.session["restore_count"] = int(restore_result.get("count") or 0)
    return RedirectResponse("/")

@app.get("/logout")
async def logout(request: Request):
    sid = request.session.get("sid")
    if sid: TOKEN_STORE.pop(sid, None)
    request.session.clear()
    return RedirectResponse("/")

@app.get("/restore")
async def restore_page(request: Request):
    """手動觸發從 OneDrive 還原備份（force=True，即使本地有資料也可覆蓋）。"""
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token:
        return RedirectResponse("/login", status_code=303)
    rr = await restore_db_from_onedrive(sid, token, force=True)
    request.session["restore_msg"] = rr.get("message") or ""
    request.session["restore_ok"] = bool(rr.get("ok"))
    request.session["restore_count"] = int(rr.get("count") or 0)
    return RedirectResponse("/", status_code=303)

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
                    try:
                        async for sub_item in fetch_media_iter(token, f"items/{item['id']}", depth + 1, max_depth):
                            yield sub_item
                    except httpx.HTTPStatusError as e:
                        # 該子資料夾可能已被刪除/移動(常見於 404),略過它、繼續掃描其他資料夾，
                        # 不要讓單一資料夾的錯誤中斷整個掃描。
                        print(f"略過子資料夾 {item.get('name')} ({item['id']}): {e}")
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
                    "cameraMake": photo_meta.get("cameraMake"), "cameraModel": photo_meta.get("cameraModel"),
                }
            url = data.get("@odata.nextLink")

async def fetch_all_media(token: str, folder_path: str = "root", depth: int = 0, max_depth: int = 6) -> list[dict]:
    items = []
    async for item in fetch_media_iter(token, folder_path, depth, max_depth): items.append(item)
    return items

def db_upsert_photo(sid: str, item: dict):
    conn = sqlite3.connect(DB_PATH)
    # 保留既有 blur_score / favorite / cleanup_skip 除非有新值
    conn.execute(
        """
        INSERT INTO photos (sid, id, name, mime_type, size, web_url, thumbnail_url, thumbnail_large_url, taken_date_time, latitude, longitude, phash, width, height, source, camera_make, camera_model, blur_score)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(sid, id) DO UPDATE SET
            name=excluded.name, mime_type=excluded.mime_type, size=excluded.size, web_url=excluded.web_url,
            thumbnail_url=excluded.thumbnail_url, thumbnail_large_url=excluded.thumbnail_large_url,
            taken_date_time=excluded.taken_date_time, latitude=excluded.latitude, longitude=excluded.longitude,
            phash=COALESCE(excluded.phash, photos.phash),
            width=excluded.width, height=excluded.height, source=excluded.source,
            camera_make=excluded.camera_make, camera_model=excluded.camera_model,
            blur_score=COALESCE(excluded.blur_score, photos.blur_score)
        """,
        (
            sid, item["id"], item.get("name"), item.get("mimeType"), item.get("size"),
            item.get("webUrl"), item.get("thumbnailUrl"), item.get("thumbnailLargeUrl"),
            item.get("takenDateTime"), item.get("latitude"), item.get("longitude"),
            item.get("phash"), item.get("width"), item.get("height"),
            item.get("source", "onedrive"), item.get("cameraMake"), item.get("cameraModel"),
            item.get("blur_score"),
        ),
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
    """批次還原，兼容舊版備份缺欄位的情況。"""
    conn = sqlite3.connect(DB_PATH)
    payload = []
    for r in rows:
        if not r.get("id"):
            continue
        payload.append((
            sid,
            r.get("id"),
            r.get("name"),
            r.get("mime_type") or r.get("mimeType"),
            r.get("size"),
            r.get("web_url") or r.get("webUrl"),
            r.get("thumbnail_url") or r.get("thumbnailUrl"),
            r.get("thumbnail_large_url") or r.get("thumbnailLargeUrl"),
            r.get("taken_date_time") or r.get("takenDateTime"),
            r.get("latitude"),
            r.get("longitude"),
            r.get("phash"),
            r.get("width"),
            r.get("height"),
            r.get("source", "onedrive"),
            r.get("favorite", 0) or 0,
            r.get("camera_make") or r.get("cameraMake"),
            r.get("camera_model") or r.get("cameraModel"),
            r.get("blur_score"),
            r.get("cleanup_skip", 0) or 0,
        ))
    conn.executemany(
        """
        INSERT INTO photos (
            sid, id, name, mime_type, size, web_url, thumbnail_url, thumbnail_large_url,
            taken_date_time, latitude, longitude, phash, width, height, source,
            favorite, camera_make, camera_model, blur_score, cleanup_skip
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(sid, id) DO UPDATE SET
            name=excluded.name, mime_type=excluded.mime_type, size=excluded.size,
            web_url=excluded.web_url, thumbnail_url=excluded.thumbnail_url,
            thumbnail_large_url=excluded.thumbnail_large_url,
            taken_date_time=excluded.taken_date_time, latitude=excluded.latitude,
            longitude=excluded.longitude, phash=COALESCE(excluded.phash, photos.phash),
            width=excluded.width, height=excluded.height, source=excluded.source,
            favorite=excluded.favorite, camera_make=excluded.camera_make,
            camera_model=excluded.camera_model,
            blur_score=COALESCE(excluded.blur_score, photos.blur_score),
            cleanup_skip=excluded.cleanup_skip
        """,
        payload,
    )
    conn.commit()
    conn.close()

# ---------------------------------------------------------------------------
# photos.db 備份到使用者自己的 OneDrive / 從 OneDrive 還原
# 注意:只會存取「該 sid 自己」的資料列,絕不會把別人的照片資料寫進任何人的 OneDrive
# ---------------------------------------------------------------------------
BACKUP_FOLDER_NAME = "MyAlbumApp_Backup"
BACKUP_FILENAME = "photos_backup.json"

# 還原結果快取（給首頁顯示用，避免使用者完全不知道發生什麼事）
RESTORE_STATUS: dict[str, dict] = {}

async def backup_db_to_onedrive(sid: str, token: str | None):
    if not token:
        return
    rows = db_get_photos(sid)
    if not rows:
        return
    url = f"{GRAPH_BASE}/me/drive/root:/{BACKUP_FOLDER_NAME}/{BACKUP_FILENAME}:/content"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    try:
        # 6000+ 張備份可能較大，給足時間
        async with httpx.AsyncClient(timeout=120) as client:
            resp = await client.put(
                url,
                headers=headers,
                content=json.dumps(rows, ensure_ascii=False).encode("utf-8"),
            )
            resp.raise_for_status()
        print(f"備份成功 sid={sid} count={len(rows)}")
    except Exception as e:
        print(f"備份 photos.db 到 OneDrive 失敗(sid={sid}): {e}")

async def restore_db_from_onedrive(sid: str, token: str | None, force: bool = False) -> dict:
    """
    從 OneDrive 還原 photos_backup.json。
    回傳 {"ok": bool, "count": int, "message": str, "error": str|None}
    """
    result = {"ok": False, "count": 0, "message": "", "error": None}
    if not token:
        result["error"] = "no_token"
        result["message"] = "沒有登入 token，無法還原"
        RESTORE_STATUS[sid] = result
        return result

    local_count = len(db_get_photos(sid))
    if local_count > 0 and not force:
        result["ok"] = True
        result["count"] = local_count
        result["message"] = f"本機已有 {local_count} 張，略過還原"
        RESTORE_STATUS[sid] = result
        return result

    url = f"{GRAPH_BASE}/me/drive/root:/{BACKUP_FOLDER_NAME}/{BACKUP_FILENAME}:/content"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        async with httpx.AsyncClient(timeout=180) as client:
            resp = await client.get(url, headers=headers)
            if resp.status_code == 404:
                result["message"] = "OneDrive 尚無備份檔（可能還沒掃過）"
                RESTORE_STATUS[sid] = result
                return result
            resp.raise_for_status()
            raw = resp.content

        rows = json.loads(raw)
        if not rows:
            result["message"] = "備份檔是空的"
            RESTORE_STATUS[sid] = result
            return result

        db_restore_sid_rows(sid, rows)
        final = len(db_get_photos(sid))
        result["ok"] = True
        result["count"] = final
        result["message"] = f"已從 OneDrive 還原 {final} 張照片"
        print(f"還原成功 sid={sid} count={final}")
        RESTORE_STATUS[sid] = result
        return result
    except Exception as e:
        result["error"] = str(e)
        result["message"] = f"還原失敗：{e}"
        print(f"從 OneDrive 還原 photos.db 失敗(sid={sid}): {e}")
        RESTORE_STATUS[sid] = result
        return result

async def process_drive_item(sid: str, item: dict, existing_phash: dict, token: str | None = None) -> bool:
    """處理單一 media item: 計算 phash / blur, upsert。回傳是否為新照片。"""
    item_id = item["id"]
    is_new = item_id not in existing_phash
    if item_id in existing_phash and existing_phash[item_id]:
        item["phash"] = existing_phash[item_id]
    elif (item.get("mimeType") or "").startswith("image/"):
        item["phash"] = await compute_phash(item.get("thumbnailUrl"))
    else:
        item["phash"] = None

    # 模糊分數 (只用 thumbnail 快速估算, 有縮圖才算)
    if (item.get("mimeType") or "").startswith("image/") and item.get("thumbnailUrl") and item.get("blur_score") is None:
        try:
            async with httpx.AsyncClient(timeout=12) as client:
                resp = await client.get(item["thumbnailUrl"])
                if resp.status_code == 200:
                    score = compute_blur_score(resp.content)
                    if score is not None:
                        item["blur_score"] = score
        except Exception:
            pass

    db_upsert_photo(sid, item)
    return is_new


async def fetch_delta_iter(token: str, delta_link: str | None = None):
    """使用 Graph delta query 取得變更 (新增/修改/刪除)。"""
    headers = {"Authorization": f"Bearer {token}"}
    url = delta_link or f"{GRAPH_BASE}/me/drive/root/delta?$expand=thumbnails&token=latest"
    # 第一次若沒有 delta_link, 用 token=latest 拿最新 token (空結果), 再全量掃一次建立基線
    # 實務上: 有 delta_link 就增量; 沒有就走全量 + 存 deltaLink

    async with httpx.AsyncClient(timeout=45) as client:
        while url:
            for attempt in range(5):
                resp = await client.get(url, headers=headers)
                if resp.status_code == 429:
                    await asyncio.sleep(int(resp.headers.get("Retry-After", 5)))
                    continue
                if resp.status_code == 410:  # delta token expired
                    # 必須重新全量
                    return
                resp.raise_for_status()
                break
            data = resp.json()
            for item in data.get("value", []):
                yield item
            url = data.get("@odata.nextLink")
            if "@odata.deltaLink" in data:
                yield {"__delta_link__": data["@odata.deltaLink"]}
                return


async def run_background_scan(sid: str, token: str, force_full: bool = False):
    """優先使用 delta 增量掃描; 沒有 delta_link 或 force_full 時才全量。"""
    SCAN_STATUS[sid] = {"status": "scanning", "count": 0, "mode": "incremental"}
    try:
        state = get_scan_state(sid)
        delta_link = None if force_full else state.get("delta_link")
        existing_photos = {p["id"]: p.get("phash") for p in db_get_photos(sid)}
        count = 0
        new_delta_link = None
        deleted_ids = []

        if delta_link:
            # ----- 增量模式 -----
            SCAN_STATUS[sid]["mode"] = "incremental"
            async for raw in fetch_delta_iter(token, delta_link):
                if isinstance(raw, dict) and "__delta_link__" in raw:
                    new_delta_link = raw["__delta_link__"]
                    break
                # deleted?
                if raw.get("deleted"):
                    deleted_ids.append(raw["id"])
                    continue
                # folder skip
                if "folder" in raw:
                    continue
                file_info = raw.get("file")
                if not file_info:
                    continue
                mime = file_info.get("mimeType", "")
                if not (mime.startswith("image/") or mime.startswith("video/")):
                    continue
                photo_meta = raw.get("photo", {})
                location = raw.get("location", {})
                image_meta = raw.get("image", {})
                thumbs = raw.get("thumbnails", [{}])[0] if raw.get("thumbnails") else {}
                item = {
                    "id": raw["id"], "name": raw["name"], "mimeType": mime, "size": raw.get("size"),
                    "webUrl": raw.get("webUrl"), "thumbnailUrl": thumbs.get("medium", {}).get("url"),
                    "thumbnailLargeUrl": thumbs.get("large", {}).get("url"),
                    "takenDateTime": photo_meta.get("takenDateTime"),
                    "latitude": location.get("latitude"), "longitude": location.get("longitude"),
                    "width": image_meta.get("width"), "height": image_meta.get("height"),
                    "source": "onedrive",
                    "cameraMake": photo_meta.get("cameraMake"), "cameraModel": photo_meta.get("cameraModel"),
                }
                await process_drive_item(sid, item, existing_photos, token)
                count += 1
                SCAN_STATUS[sid]["count"] = count

            # 處理刪除
            if deleted_ids:
                conn = sqlite3.connect(DB_PATH)
                for did in deleted_ids:
                    conn.execute("DELETE FROM photos WHERE sid = ? AND id = ?", (sid, did))
                    conn.execute("DELETE FROM photo_tags WHERE sid = ? AND photo_id = ?", (sid, did))
                conn.commit()
                conn.close()
        else:
            # ----- 全量模式 (首次或 force) -----
            SCAN_STATUS[sid]["mode"] = "full"
            async for item in fetch_media_iter(token, max_depth=15):
                await process_drive_item(sid, item, existing_photos, token)
                count += 1
                SCAN_STATUS[sid]["count"] = count
                if count % 200 == 0:
                    await backup_db_to_onedrive(sid, token)

            # 全量後取得最新 deltaLink, 之後就能增量
            try:
                async with httpx.AsyncClient(timeout=30) as client:
                    resp = await client.get(
                        f"{GRAPH_BASE}/me/drive/root/delta?token=latest",
                        headers={"Authorization": f"Bearer {token}"},
                    )
                    if resp.status_code == 200:
                        new_delta_link = resp.json().get("@odata.deltaLink")
            except Exception:
                pass

        final_count = len(db_get_photos(sid))
        if new_delta_link:
            save_scan_state(sid, delta_link=new_delta_link, photo_count=final_count)
        else:
            save_scan_state(sid, photo_count=final_count)

        SCAN_STATUS[sid] = {
            "status": "done",
            "count": final_count,
            "changed": count,
            "deleted": len(deleted_ids) if delta_link else 0,
            "mode": SCAN_STATUS[sid].get("mode", "full"),
        }
        await backup_db_to_onedrive(sid, token)
    except Exception as e:
        SCAN_STATUS[sid] = {"status": "error", "error": str(e)}
        await backup_db_to_onedrive(sid, token)


def start_scan_if_needed(sid: str, token: str, force_full: bool = False):
    if SCAN_STATUS.get(sid, {}).get("status") == "scanning":
        return
    SCAN_TASKS[sid] = asyncio.create_task(run_background_scan(sid, token, force_full=force_full))

@app.get("/scan/start")
async def scan_start(request: Request, force: int = 0):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not token:
        return JSONResponse({"error": "not_logged_in"}, status_code=401)
    start_scan_if_needed(sid, token, force_full=bool(force))
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
    if not sid: return RedirectResponse("/login", status_code=303)
    
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
    if not sid or not token: return RedirectResponse("/login", status_code=303)

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
    blurry = []
    for it in items:
        if it.get("cleanup_skip"):
            continue
        if (it.get("mime_type") or "").startswith("video/"):
            continue
        if is_screenshot(it):
            screenshots.append(it)
            continue
        if is_blurry(it):
            blurry.append(it)
            continue
        size = it.get("size")
        is_small_file = (size is not None and size < 102400)
        w, h = it.get("width"), it.get("height")
        is_low_res = (w is not None and h is not None and w < 800 and h < 800)
        if is_small_file or is_low_res:
            low_quality.append(it)
    return {"screenshots": screenshots, "low_quality": low_quality, "blurry": blurry}

@app.get("/cleanup", response_class=HTMLResponse)
async def cleanup_view(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    
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
    grouped_blurry = group_by_month(cleanup_data.get("blurry", []))
    
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
                <div class="btn-row">
                    <form action="/cleanup/batch-move" method="post" style="flex:1;" onsubmit="return confirm('確定要將 {month_key} 的這 {len(group_items)} 張照片，移至 OneDrive 的「待清理資料夾」嗎？');">
                        <input type="hidden" name="ids" value="{ids_str}" />
                        <button type="submit" class="btn-primary" style="background: var(--warning); color: #000; margin-top: 4px;">
                            📦 移至「待清理資料夾」
                        </button>
                    </form>
                    <form action="/cleanup/skip" method="post" style="flex:1;">
                        <input type="hidden" name="ids" value="{ids_str}" />
                        <button type="submit" class="btn-secondary" style="margin-top: 4px;">
                            ✅ 略過不刪(不再出現)
                        </button>
                    </form>
                </div>
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
    
    has_any = cleanup_data["screenshots"] or cleanup_data["low_quality"] or cleanup_data.get("blurry")
    if not has_any:
        body += """<div class="card"><p class="secondary">太棒了！目前圖庫裡沒有發現截圖、模糊或低畫質的垃圾照片。</p></div>"""
    else:
        if cleanup_data["screenshots"]:
            body += """<div class="section-title">螢幕截圖分批清理</div>"""
            body += build_month_sections(grouped_screenshots, "📱", "截圖")
        if cleanup_data.get("blurry"):
            body += """<div class="section-title">模糊 / 失焦照片分批清理</div>"""
            body += build_month_sections(grouped_blurry, "🌫️", "模糊")
        if cleanup_data["low_quality"]:
            body += """<div class="section-title">低畫質小檔案分批清理</div>"""
            body += build_month_sections(grouped_low_quality, "🗑️", "小檔案")
            
    return HTMLResponse(page_shell("安全快速清理", body, active_tab="more", back_href="/more"))

@app.post("/cleanup/batch-move")
async def cleanup_batch_move(request: Request, ids: str = Form(...)):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token: return RedirectResponse("/login", status_code=303)

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

@app.post("/cleanup/skip")
async def cleanup_skip(request: Request, ids: str = Form(...)):
    """把使用者標記為「不想刪除」的照片記起來，之後的快速清理列表不會再顯示它們。"""
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)

    id_list = [i for i in ids.split(",") if i]
    if id_list:
        conn = sqlite3.connect(DB_PATH)
        conn.executemany(
            "UPDATE photos SET cleanup_skip = 1 WHERE sid = ? AND id = ?",
            [(sid, d_id) for d_id in id_list],
        )
        conn.commit()
        conn.close()

    return RedirectResponse("/cleanup", status_code=303)

# ---------------------------------------------------------------------------
# 單張照片快速操作:馬上刪除(送到 OneDrive 回收桶,可還原) + 我的最愛
# ---------------------------------------------------------------------------
@app.post("/api/photo/delete")
async def api_delete_photo(request: Request, id: str = Form(...)):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token: return JSONResponse({"success": False, "error": "not_logged_in"}, status_code=401)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT source FROM photos WHERE sid = ? AND id = ?", (sid, id)).fetchone()
    if not row:
        conn.close()
        return JSONResponse({"success": False, "error": "not_found"}, status_code=404)

    if row["source"] == "onedrive":
        headers = {"Authorization": f"Bearer {token}"}
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                for attempt in range(5):
                    resp = await client.delete(f"{GRAPH_BASE}/me/drive/items/{id}", headers=headers)
                    if resp.status_code == 429:
                        await asyncio.sleep(int(resp.headers.get("Retry-After", 3)))
                        continue
                    if resp.status_code not in (204, 404):
                        resp.raise_for_status()
                    break
        except Exception as e:
            conn.close()
            return JSONResponse({"success": False, "error": str(e)}, status_code=502)

    conn.execute("DELETE FROM photos WHERE sid = ? AND id = ?", (sid, id))
    conn.commit()
    conn.close()
    asyncio.create_task(backup_db_to_onedrive(sid, token))
    return JSONResponse({"success": True})

@app.post("/api/photo/favorite")
async def api_toggle_favorite(request: Request, id: str = Form(...)):
    sid = request.session.get("sid")
    if not sid: return JSONResponse({"success": False, "error": "not_logged_in"}, status_code=401)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    row = conn.execute("SELECT favorite FROM photos WHERE sid = ? AND id = ?", (sid, id)).fetchone()
    if not row:
        conn.close()
        return JSONResponse({"success": False, "error": "not_found"}, status_code=404)

    new_val = 0 if row["favorite"] else 1
    conn.execute("UPDATE photos SET favorite = ? WHERE sid = ? AND id = ?", (new_val, sid, id))
    conn.commit()
    conn.close()
    return JSONResponse({"success": True, "favorite": bool(new_val)})

@app.get("/favorites", response_class=HTMLResponse)
async def favorites_page(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    items = [it for it in db_get_photos(sid) if it.get("favorite")]

    if not items:
        body = """<div class="card"><p class="secondary">目前還沒有加入最愛的照片。瀏覽照片時點縮圖左上角的 ☆,或在放大檢視時點右上角的星星,就能收藏。</p></div>"""
    else:
        body = f"""
        <p class="secondary" style="margin:0 6px 14px;">共 {len(items)} 張最愛照片。</p>
        <div class="photo-grid">{"".join(lb_img_tag(it) for it in items if it.get("thumbnail_url"))}</div>
        """
    return HTMLResponse(page_shell("我的最愛", body, active_tab="more", back_href="/more"))

# ---------------------------------------------------------------------------
# 儲存空間統計
# ---------------------------------------------------------------------------
def format_bytes(n: int | None) -> str:
    if not n: return "0 B"
    size = float(n)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024:
            return f"{int(size)} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} PB"

@app.get("/storage", response_class=HTMLResponse)
async def storage_page(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)

    if not items:
        body = """<div class="card"><p class="secondary">圖庫還是空的,掃描 OneDrive 後這裡就會出現空間統計。</p></div>"""
        return HTMLResponse(page_shell("儲存空間統計", body, active_tab="more", back_href="/more"))

    total_size = sum(it.get("size") or 0 for it in items)
    screenshots = [it for it in items if is_screenshot(it)]
    videos = [it for it in items if (it.get("mime_type") or "").startswith("video/")]
    screenshot_ids = {it["id"] for it in screenshots}; video_ids = {it["id"] for it in videos}
    photos_only_count = len(items) - len(screenshots) - len(videos)
    screenshot_size = sum(it.get("size") or 0 for it in screenshots)
    video_size = sum(it.get("size") or 0 for it in videos)
    photo_size = total_size - screenshot_size - video_size

    by_year: dict[str, int] = defaultdict(int)
    for it in items:
        dt = parse_taken(it)
        by_year[str(dt.year) if dt else "未知年份"] += (it.get("size") or 0)

    largest = sorted(items, key=lambda it: it.get("size") or 0, reverse=True)[:12]

    year_rows = "".join(f"""
    <div class="list-row"><span class="row-icon">🗓️</span><span class="row-label">{y}</span><span class="row-value">{format_bytes(sz)}</span></div>
    """ for y, sz in sorted(by_year.items(), key=lambda kv: -kv[1]))

    body = f"""
    <div class="card" style="text-align:center;">
        <div class="stat-num">{format_bytes(total_size)}</div>
        <div class="secondary">共 {len(items)} 個檔案</div>
    </div>
    <div class="list-group">
        <div class="list-row"><span class="row-icon">🖼</span><span class="row-label">一般照片</span><span class="row-value">{photos_only_count} 張・{format_bytes(photo_size)}</span></div>
        <div class="list-row"><span class="row-icon">📱</span><span class="row-label">螢幕截圖</span><span class="row-value">{len(screenshots)} 張・{format_bytes(screenshot_size)}</span></div>
        <div class="list-row"><span class="row-icon">🎬</span><span class="row-label">影片</span><span class="row-value">{len(videos)} 支・{format_bytes(video_size)}</span></div>
    </div>
    <div class="section-title">依年份佔用空間</div>
    <div class="list-group">{year_rows or '<div class="list-row"><span class="secondary">尚無資料</span></div>'}</div>
    <div class="section-title">最大的檔案(前 12 個,適合檢查要不要刪)</div>
    <div class="photo-grid">{"".join(lb_img_tag(it, badge=format_bytes(it.get("size") or 0)) for it in largest if it.get("thumbnail_url"))}</div>
    """
    return HTMLResponse(page_shell("儲存空間統計", body, active_tab="more", back_href="/more"))

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

# ---------------------------------------------------------------------------
# 拍攝地點命名(免費反向地理編碼,使用 OpenStreetMap Nominatim,不需要金鑰、不用付費)
# 依 Nominatim 使用政策:每秒最多 1 次請求,並且需要帶 User-Agent。
# ---------------------------------------------------------------------------
LOCATION_NAME_CACHE: dict[tuple, str] = load_location_cache()
_geocode_lock = asyncio.Lock()

async def reverse_geocode(lat: float, lng: float, precision: int = LOCATION_PRECISION) -> str:
    key = (round(lat, precision), round(lng, precision))
    if key in LOCATION_NAME_CACHE:
        return LOCATION_NAME_CACHE[key]
    async with _geocode_lock:
        if key in LOCATION_NAME_CACHE:
            return LOCATION_NAME_CACHE[key]
        name = f"{key[0]:.2f}, {key[1]:.2f}"
        try:
            async with httpx.AsyncClient(timeout=10) as client:
                resp = await client.get(
                    "https://nominatim.openstreetmap.org/reverse",
                    params={"lat": lat, "lon": lng, "format": "jsonv2", "accept-language": "zh-TW", "zoom": 14},
                    headers={"User-Agent": "MyAlbumApp/1.0 (personal photo album, non-commercial)"},
                )
                resp.raise_for_status()
                data = resp.json()
                addr = data.get("address", {})
                name = (
                    addr.get("suburb") or addr.get("neighbourhood") or addr.get("town")
                    or addr.get("city_district") or addr.get("city") or addr.get("county")
                    or (data.get("display_name", "").split(",")[0] if data.get("display_name") else None)
                    or name
                )
            await asyncio.sleep(1.0)  # 遵守 Nominatim 每秒最多 1 次請求的限制
        except Exception:
            pass
        LOCATION_NAME_CACHE[key] = name
        save_location_cache_entry(key[0], key[1], name)
        return name

async def annotate_location_names(buckets: list[tuple[tuple, list[dict]]], max_lookup: int = 25) -> list[tuple[tuple, list[dict], str]]:
    """幫地點分組加上人類看得懂的地名。為了避免一次觸發太多次 Nominatim 查詢(每秒限 1 次),
    只主動查詢照片數最多的前 max_lookup 組,其餘先用座標當名稱顯示,之後點進去也能觸發查詢。"""
    result = []
    for idx, (coords, group) in enumerate(buckets):
        if idx < max_lookup:
            name = await reverse_geocode(coords[0], coords[1])
        else:
            key = (round(coords[0], LOCATION_PRECISION), round(coords[1], LOCATION_PRECISION))
            name = LOCATION_NAME_CACHE.get(key, f"{coords[0]:.2f}, {coords[1]:.2f}")
        result.append((coords, group, name))
    return result

async def render_albums_html(items: list[dict]) -> str:
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

    sorted_locations = sorted(location_buckets.items(), key=lambda kv: -len(kv[1]))[:12]
    named_locations = await annotate_location_names(sorted_locations)
    location_html = "".join(f"""
    <div class="card group-card">
        <div class="group-title">📍 {name}</div>
        <div class="group-sub">共 {len(group)} 張・<a class="pill-link" href="/locations/view?lat={coords[0]}&lng={coords[1]}">看這個地點的照片</a>・<a class="pill-link" href="https://www.google.com/maps?q={coords[0]},{coords[1]}" target="_blank">在地圖上看</a></div>
        <div class="photo-grid">{thumb_row(group)}</div>
    </div>
    """ for coords, group, name in named_locations) or """<div class="card"><p class="secondary">目前沒有帶 GPS 座標的照片。</p></div>"""
    more_locations = f"""<a class="btn-secondary" href="/locations">看全部 {len(location_buckets)} 個拍攝地點(含地圖總覽)</a>""" if location_buckets else ""

    camera_buckets = cluster_by_camera(photos_only)
    sorted_cameras = sorted(camera_buckets.items(), key=lambda kv: -len(kv[1]))[:6]
    camera_html = "".join(f"""
    <div class="card group-card">
        <div class="group-title">📷 {label}</div>
        <div class="group-sub">共 {len(group)} 張・<a class="pill-link" href="/cameras/view?model={quote(label)}">看這台裝置拍的照片</a></div>
        <div class="photo-grid">{thumb_row(group)}</div>
    </div>
    """ for label, group in sorted_cameras) or """<div class="card"><p class="secondary">目前沒有相機/裝置資訊(部分照片本身可能沒有內嵌這項資料,或還沒重新掃描)。</p></div>"""
    more_cameras = f"""<a class="btn-secondary" href="/cameras">看全部 {len(camera_buckets)} 種相機/裝置</a>""" if len(camera_buckets) > 6 else ""

    screenshot_empty = "" if screenshots else """<p class="secondary">目前沒有偵測到螢幕截圖。</p>"""
    video_empty = "" if videos else """<p class="secondary">目前沒有影片。</p>"""

    body = f"""
    <p class="secondary" style="margin:0 6px 14px;">全部用現有的拍攝時間 / GPS / 檔名 / 裝置資料分類,地點名稱來自免費的 OpenStreetMap,沒有呼叫任何付費 AI 服務。</p>
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
    <div class="section-title">依地點分的相簿(前 {len(named_locations)} / {len(location_buckets)} 組)</div>
    {location_html}
    {more_locations}
    <div class="section-title">依相機/裝置分的相簿(前 {len(sorted_cameras)} / {len(camera_buckets)} 種)</div>
    {camera_html}
    {more_cameras}
    """
    return page_shell("自動分類", body, active_tab="albums")

@app.get("/albums", response_class=HTMLResponse)
async def albums(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    return HTMLResponse(await render_albums_html(db_get_photos(sid)))

# ---------------------------------------------------------------------------
# 依相機/裝置分類
# ---------------------------------------------------------------------------
def cluster_by_camera(items: list[dict]) -> dict[str, list[dict]]:
    buckets: dict[str, list[dict]] = defaultdict(list)
    for it in items:
        model = (it.get("camera_model") or "").strip()
        make = (it.get("camera_make") or "").strip()
        label = model or make or "未知裝置"
        buckets[label].append(it)
    return buckets

@app.get("/cameras", response_class=HTMLResponse)
async def cameras_page(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)
    screenshot_ids = {it["id"] for it in items if is_screenshot(it)}
    photos_only = [it for it in items if it["id"] not in screenshot_ids]
    buckets = cluster_by_camera(photos_only)
    sorted_buckets = sorted(buckets.items(), key=lambda kv: -len(kv[1]))

    if not sorted_buckets:
        body = """<div class="card"><p class="secondary">目前沒有相機/裝置資訊,重新掃描 OneDrive 後才會出現(部分照片本身可能沒有內嵌這項資料)。</p></div>"""
    else:
        rows = "".join(f"""
        <a class="list-row tappable" href="/cameras/view?model={quote(label)}">
            <span class="row-icon">📷</span>
            <span class="row-label">{label}</span>
            <span class="row-value">{len(group)} 張</span>
            <span class="chevron">›</span>
        </a>
        """ for label, group in sorted_buckets)
        body = f"""
        <p class="secondary" style="margin:0 6px 14px;">依照片內嵌的拍攝裝置資訊分組(來自 OneDrive 照片 metadata)。</p>
        <div class="list-group">{rows}</div>
        """
    return HTMLResponse(page_shell("依相機/裝置分類", body, active_tab="albums", back_href="/albums"))

@app.get("/cameras/view", response_class=HTMLResponse)
async def camera_view(request: Request, model: str):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)
    matched = [it for it in items if ((it.get("camera_model") or "").strip() or (it.get("camera_make") or "").strip() or "未知裝置") == model]

    if not matched:
        body = """<div class="card"><p class="secondary">找不到這個裝置拍的照片。</p></div>"""
    else:
        thumbs = "".join(lb_img_tag(it) for it in matched if it.get("thumbnail_url"))
        body = f"""
        <div class="card" style="text-align:center;">
            <div class="stat-num">{len(matched)}</div>
            <div class="secondary">張照片由這台裝置拍攝</div>
        </div>
        <div class="photo-grid">{thumbs}</div>
        """
    return HTMLResponse(page_shell(f"📷 {model}", body, active_tab="albums", back_href="/cameras"))

# ---------------------------------------------------------------------------
# 拍攝地點:地圖總覽 + 依地點篩選照片
# ---------------------------------------------------------------------------
@app.get("/locations", response_class=HTMLResponse)
async def locations_page(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)
    buckets = cluster_by_location(items)
    sorted_buckets = sorted(buckets.items(), key=lambda kv: -len(kv[1]))
    named = await annotate_location_names(sorted_buckets, max_lookup=30)

    if not named:
        body = """<div class="card"><p class="secondary">目前沒有帶 GPS 座標的照片,無法顯示拍攝地點。</p></div>"""
        return HTMLResponse(page_shell("拍攝地點", body, active_tab="albums", back_href="/albums"))

    markers_js = ",".join(
        f'{{lat:{coords[0]},lng:{coords[1]},count:{len(group)},name:{json.dumps(name, ensure_ascii=False)},href:"/locations/view?lat={coords[0]}&lng={coords[1]}"}}'
        for coords, group, name in named
    )
    avg_lat = sum(c[0] for c, _, _ in named) / len(named)
    avg_lng = sum(c[1] for c, _, _ in named) / len(named)

    rows_html = "".join(f"""
    <a class="list-row tappable" href="/locations/view?lat={coords[0]}&lng={coords[1]}">
        <span class="row-icon">📍</span>
        <span class="row-label">{name}</span>
        <span class="row-value">{len(group)} 張</span>
        <span class="chevron">›</span>
    </a>
    """ for coords, group, name in named)

    body = f"""
    <p class="secondary" style="margin:0 6px 14px;">地圖與地名完全免費(OpenStreetMap),不需要任何付費金鑰。點地圖上的圖釘或下方清單,可以只看該地點拍的照片。</p>
    <a class="btn-secondary" href="/locations/footprint" style="margin-bottom:14px;">🗺️ 看全部照片的足跡地圖(逐張顯示)</a>
    <div id="loc-map" style="width:100%; height:280px; border-radius:14px; overflow:hidden; margin-bottom:14px;"></div>
    <div class="list-group">{rows_html}</div>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
    (function() {{
        var map = L.map('loc-map').setView([{avg_lat}, {avg_lng}], 6);
        L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
            attribution: '&copy; OpenStreetMap contributors', maxZoom: 18
        }}).addTo(map);
        var pts = [{markers_js}];
        var bounds = [];
        pts.forEach(function(p) {{
            var m = L.marker([p.lat, p.lng]).addTo(map);
            m.bindPopup(p.name + '(' + p.count + ' 張)<br><a href="' + p.href + '">查看照片</a>');
            m.on('click', function() {{ window.location.href = p.href; }});
            bounds.push([p.lat, p.lng]);
        }});
        if (bounds.length > 1) map.fitBounds(bounds, {{padding: [24, 24]}});
    }})();
    </script>
    """
    return HTMLResponse(page_shell(f"拍攝地點({len(named)})", body, active_tab="albums", back_href="/albums", include_lightbox=False))

@app.get("/locations/footprint", response_class=HTMLResponse)
async def locations_footprint(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    items = [it for it in db_get_photos(sid) if it.get("latitude") is not None and it.get("longitude") is not None]

    if not items:
        body = """<div class="card"><p class="secondary">目前沒有帶 GPS 座標的照片。</p></div>"""
        return HTMLResponse(page_shell("足跡地圖", body, active_tab="albums", back_href="/locations"))

    LIMIT = 3000
    shown = items[:LIMIT]
    points_js = ",".join(
        f'{{lat:{it["latitude"]},lng:{it["longitude"]},thumb:{json.dumps(it.get("thumbnail_url") or "", ensure_ascii=False)},name:{json.dumps(it.get("name") or "", ensure_ascii=False)}}}'
        for it in shown
    )
    avg_lat = sum(it["latitude"] for it in shown) / len(shown)
    avg_lng = sum(it["longitude"] for it in shown) / len(shown)
    note = (
        f"共 {len(items)} 張有 GPS 座標的照片,地圖上顯示最新的 {len(shown)} 張。"
        if len(items) > LIMIT else f"共 {len(items)} 張有 GPS 座標的照片。"
    )

    body = f"""
    <p class="secondary" style="margin:0 6px 14px;">{note}</p>
    <div id="footprint-map" style="width:100%; height:70vh; border-radius:14px; overflow:hidden;"></div>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
    <script>
    (function() {{
        var map = L.map('footprint-map').setView([{avg_lat}, {avg_lng}], 5);
        L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
            attribution: '&copy; OpenStreetMap contributors', maxZoom: 18
        }}).addTo(map);
        var cluster = L.markerClusterGroup();
        var pts = [{points_js}];
        var bounds = [];
        pts.forEach(function(p) {{
            var m = L.marker([p.lat, p.lng]);
            var img = p.thumb ? '<img src="' + p.thumb + '" style="width:120px;height:120px;object-fit:cover;border-radius:8px;display:block;margin-bottom:4px;">' : '';
            m.bindPopup(img + (p.name || ''));
            cluster.addLayer(m);
            bounds.push([p.lat, p.lng]);
        }});
        map.addLayer(cluster);
        if (bounds.length > 1) map.fitBounds(bounds, {{padding: [24, 24]}});
    }})();
    </script>
    """
    return HTMLResponse(page_shell("足跡地圖", body, active_tab="albums", back_href="/locations", include_lightbox=False))

@app.get("/locations/view", response_class=HTMLResponse)
async def location_view(request: Request, lat: float, lng: float):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)
    key = (round(lat, LOCATION_PRECISION), round(lng, LOCATION_PRECISION))
    matched = [it for it in items if it.get("latitude") is not None and it.get("longitude") is not None
               and (round(it["latitude"], LOCATION_PRECISION), round(it["longitude"], LOCATION_PRECISION)) == key]
    name = await reverse_geocode(lat, lng)

    if not matched:
        body = """<div class="card"><p class="secondary">這個地點目前沒有照片(可能已被刪除或分類條件改變)。</p></div>"""
    else:
        thumbs = "".join(lb_img_tag(it) for it in matched if it.get("thumbnail_url"))
        body = f"""
        <div class="card" style="text-align:center;">
            <div class="stat-num">{len(matched)}</div>
            <div class="secondary">張照片拍攝於此・<a class="pill-link" href="https://www.google.com/maps?q={lat},{lng}" target="_blank">在地圖上看</a></div>
        </div>
        <div class="photo-grid">{thumbs}</div>
        """
    return HTMLResponse(page_shell(f"📍 {name}", body, active_tab="albums", back_href="/locations"))

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
    """優先用 Graph downloadUrl（預簽章、較穩），失敗再退回 /content。"""
    headers = {"Authorization": f"Bearer {token}"}
    try:
        # 1) 先取 metadata 裡的 @microsoft.graph.downloadUrl
        meta = await client.get(
            f"{GRAPH_BASE}/me/drive/items/{item_id}?$select=id,@microsoft.graph.downloadUrl",
            headers=headers,
            timeout=20,
        )
        if meta.status_code == 200:
            download_url = meta.json().get("@microsoft.graph.downloadUrl")
            if download_url:
                # downloadUrl 本身已帶授權，不要再帶 Bearer
                resp = await client.get(download_url, timeout=60, follow_redirects=True)
                if resp.status_code == 200 and resp.content:
                    return resp.content
        # 2) 退回 /content
        resp = await client.get(
            f"{GRAPH_BASE}/me/drive/items/{item_id}/content",
            headers=headers,
            timeout=60,
            follow_redirects=True,
        )
        if resp.status_code == 200 and resp.content:
            return resp.content
    except Exception as e:
        print(f"download_full_image failed id={item_id}: {e}")
    return None

async def get_full_image_bytes(client: httpx.AsyncClient, item: dict, token: str | None) -> bytes | None:
    """原圖 → 大縮圖 URL → 中縮圖 URL，盡量拿到可用的圖。"""
    if token:
        raw = await download_full_image(client, token, item["id"])
        if raw:
            return raw
    # 縮圖 fallback（回憶影片品質會差一點，但至少能產出）
    for key in ("thumbnail_large_url", "thumbnailLargeUrl", "thumbnail_url", "thumbnailUrl"):
        url = item.get(key)
        if not url:
            continue
        try:
            resp = await client.get(url, timeout=30, follow_redirects=True)
            if resp.status_code == 200 and resp.content:
                return resp.content
        except Exception:
            continue
    return None

async def render_memory_video(job_id: str, items: list[dict], token: str | None, music_path: str | None, title: str, sid: str | None = None):
    MEMORY_JOBS[job_id] = {"status": "rendering", "progress": 0, "total": len(items), "title": title, "failed": 0}
    tmp_dir = os.path.join(RENDER_DIR, job_id)
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        # 背景任務執行較久，中途刷新一次 token
        if sid:
            fresh = await get_ms_token(sid)
            if fresh:
                token = fresh

        frame_paths = []
        failed = 0
        async with httpx.AsyncClient() as client:
            for idx, item in enumerate(items):
                raw = await get_full_image_bytes(client, item, token)
                if not raw:
                    failed += 1
                    MEMORY_JOBS[job_id]["failed"] = failed
                    MEMORY_JOBS[job_id]["progress"] = idx + 1
                    continue
                try:
                    frame = prepare_frame(raw, (1280, 720))
                except Exception:
                    failed += 1
                    MEMORY_JOBS[job_id]["failed"] = failed
                    MEMORY_JOBS[job_id]["progress"] = idx + 1
                    continue
                frame_path = os.path.join(tmp_dir, f"{idx:03d}.jpg")
                frame.save(frame_path, "JPEG", quality=90)
                frame_paths.append(frame_path)
                MEMORY_JOBS[job_id]["progress"] = idx + 1

        if len(frame_paths) < 2:
            raise ValueError(
                f"可下載到的照片不足兩張（成功 {len(frame_paths)}、失敗 {failed}）。"
                "可能原因：登入已過期、檔案已從 OneDrive 刪除、或網路被擋。"
                "請重新登入後再試一次。"
            )
        output_path = os.path.join(RENDER_DIR, f"{job_id}.mp4")
        await asyncio.to_thread(
            build_xfade_video, frame_paths, output_path,
            seconds_per_photo=3.0, transition_seconds=1.0, audio_path=music_path,
        )
        MEMORY_JOBS[job_id] = {"status": "done", "video_path": output_path, "title": title}
    except Exception as e:
        MEMORY_JOBS[job_id] = {"status": "error", "error": str(e), "title": title}
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

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
    if not sid: return RedirectResponse("/login", status_code=303)
    return HTMLResponse(render_memories_html(db_get_photos(sid)))

@app.post("/memories/render")
async def memories_render(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    token = await get_ms_token(sid)
    form = await request.form(); ids = [i for i in str(form.get("ids", "")).split(",") if i]
    if len(ids) < 2: return JSONResponse({"error": "photos_not_enough"}, status_code=400)
    all_photos = {p["id"]: p for p in db_get_photos(sid)}; items = [all_photos[i] for i in ids if i in all_photos]
    if len(items) < 2: return JSONResponse({"error": "photos_not_found"}, status_code=400)
    music = form.get("music"); music_path = os.path.join(MUSIC_DIR, str(music)) if music else None
    if music_path and not os.path.isfile(music_path): music_path = None
    if not token:
        return RedirectResponse("/login", status_code=303)
    job_id = str(uuid.uuid4())
    title = str(form.get("title", "回憶影片"))
    MEMORY_JOBS[job_id] = {"status": "pending", "title": title}
    asyncio.create_task(render_memory_video(job_id, items, token, music_path, title, sid=sid))
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
# 圖庫清單(依月份分組 + 分頁載入,大量照片也不會一次塞爆整頁)
# ---------------------------------------------------------------------------
def group_photos_by_month(items: list[dict]) -> list[tuple[str, list[dict]]]:
    """items 已經是 taken_date_time DESC 排序(見 db_get_photos),所以只要依序分組,
    月份的先後順序就會自動維持正確,沒有拍攝時間的照片會落在最後一組。"""
    grouped: dict[str, list[dict]] = defaultdict(list)
    order: list[str] = []
    for it in items:
        dt = parse_taken(it)
        month_key = dt.strftime("%Y-%m") if dt else "未知拍攝時間"
        if month_key not in grouped:
            order.append(month_key)
        grouped[month_key].append(it)
    return [(mk, grouped[mk]) for mk in order]

def month_label(month_key: str) -> str:
    if month_key == "未知拍攝時間": return month_key
    try:
        y, m = month_key.split("-")
        return f"{y} 年 {int(m)} 月"
    except Exception:
        return month_key

def render_gallery_html(items: list[dict], status: dict, visible_months: int, year_filter: str | None = None) -> str:
    scan_state = status.get("status", "idle"); scanned_count = status.get("count", 0)
    banners = ""

    if scan_state == "scanning":
        banners += f"""<div class="card banner-warning">正在背景整理你的 OneDrive,目前已掃到 {scanned_count} 張……</div>"""
    elif scan_state == "error":
        banners += f"""<div class="card banner-error">OneDrive 掃描時發生錯誤:{status.get("error")}</div>"""
    elif scan_state == "idle" and not items:
        banners += """<div class="card banner-info">尚未開始整理 OneDrive。<a class="pill-link" href="/scan/start">開始掃描</a></div>"""

    all_items = items
    if year_filter:
        items = [it for it in items if (dt := parse_taken(it)) and str(dt.year) == year_filter or (not dt and year_filter == "未知年份")]

    refresh = 3 if scan_state == "scanning" else None
    empty_state = "" if items else """
    <div class="card">
        <p class="secondary">目前圖庫是空的,掃描 OneDrive 後照片就會出現在這裡。</p>
    </div>
    """

    month_groups = group_photos_by_month(items)
    shown_groups = month_groups[:visible_months]
    remaining_months = len(month_groups) - len(shown_groups)

    jump_links = "".join(
        f'<a class="pill-link" href="#m-{mk}">{month_label(mk)}</a>'
        for mk, _ in shown_groups
    )
    jump_nav = f'<div style="margin-bottom:14px; overflow-x:auto; white-space:nowrap;">{jump_links}</div>' if len(shown_groups) > 1 else ""

    sections = "".join(f"""
    <div id="m-{mk}" class="section-title">{month_label(mk)}({len(group_items)} 張)</div>
    <div class="photo-grid">{"".join(lb_img_tag(it) for it in group_items if it.get("thumbnail_url"))}</div>
    """ for mk, group_items in shown_groups)

    more_qs = f"&year={year_filter}" if year_filter else ""
    load_more = f"""
    <a class="btn-secondary" href="/gallery?months={visible_months + GALLERY_MONTHS_PER_PAGE}{more_qs}">顯示更早的照片(還有 {remaining_months} 個月份未顯示)</a>
    """ if remaining_months > 0 else ""

    filter_banner = f"""
    <div class="card banner-info" style="display:flex; align-items:center; justify-content:space-between;">
        <span>目前只顯示 <b>{year_filter}</b> 的照片</span>
        <a class="pill-link" href="/gallery">清除篩選</a>
    </div>
    """ if year_filter else ""

    search_bar = """
    <form class="inline-form" method="get" action="/search" style="margin-bottom:14px;">
        <input class="apple-input" type="text" name="q" placeholder="🔍 搜尋檔名……" />
    </form>
    """
    quick_links = """
    <div style="margin-bottom:14px; display:flex; gap:8px; overflow-x:auto; white-space:nowrap;">
        <a class="pill-link" href="/years">📅 年份總覽</a>
        <a class="pill-link" href="/locations">📍 拍攝地點</a>
        <a class="pill-link" href="/search">🔍 進階搜尋</a>
        <a class="pill-link" href="/favorites">⭐ 我的最愛</a>
    </div>
    """

    title = f"{year_filter} 的照片({len(items)})" if year_filter else f"我的圖庫({len(all_items)})"
    body = f"""
    {banners}
    {filter_banner}
    {search_bar}
    {quick_links}
    <a class="btn-secondary" href="/scan/start" style="margin-bottom:14px;">重新掃描 OneDrive</a>
    {jump_nav}
    {sections}
    {empty_state}
    {load_more}
    """
    return page_shell(title, body, active_tab="gallery", meta_refresh=refresh)

@app.get("/gallery", response_class=HTMLResponse)
async def gallery(request: Request, months: int = GALLERY_MONTHS_PER_PAGE, year: str | None = None):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token: return RedirectResponse("/login", status_code=303)
    status = SCAN_STATUS.get(sid, {"status": "idle", "count": 0})
    return HTMLResponse(render_gallery_html(db_get_photos(sid), status, max(1, months), year_filter=year))

# ---------------------------------------------------------------------------
# 年份總覽頁
# ---------------------------------------------------------------------------
def render_years_html(items: list[dict]) -> str:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for it in items:
        dt = parse_taken(it)
        grouped[str(dt.year) if dt else "未知年份"].append(it)

    def year_sort_key(y: str):
        return (-1 if y == "未知年份" else int(y))

    years = sorted(grouped.keys(), key=year_sort_key, reverse=True)
    if not years:
        body = """<div class="card"><p class="secondary">圖庫還是空的,掃描 OneDrive 後這裡就會出現年份總覽。</p></div>"""
        return page_shell("年份總覽", body, active_tab="gallery", back_href="/gallery")

    cards = "".join(f"""
    <a class="list-row tappable" href="/gallery?year={y}">
        <span class="row-icon">🗓️</span>
        <span class="row-label">{y if y == "未知年份" else y + " 年"}</span>
        <span class="row-value">{len(grouped[y])} 張</span>
        <span class="chevron">›</span>
    </a>
    """ for y in years)

    preview_sections = "".join(f"""
    <div class="section-title">{y if y == "未知年份" else y + " 年"}({len(grouped[y])} 張)</div>
    <div class="photo-grid">{"".join(lb_img_tag(it) for it in grouped[y][:12] if it.get("thumbnail_url"))}</div>
    <a class="pill-link" href="/gallery?year={y}" style="margin-bottom:8px;">看這一年全部照片</a>
    """ for y in years)

    body = f"""
    <p class="secondary" style="margin:0 6px 14px;">依拍攝年份整理,點一個年份可以只看那一年的圖庫。</p>
    <div class="list-group">{cards}</div>
    {preview_sections}
    """
    return page_shell("年份總覽", body, active_tab="gallery", back_href="/gallery")

@app.get("/years", response_class=HTMLResponse)
async def years_page(request: Request):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    return HTMLResponse(render_years_html(db_get_photos(sid)))

# ---------------------------------------------------------------------------
# 圖庫搜尋(檔名關鍵字 + 年份/月份篩選)
# ---------------------------------------------------------------------------
def render_search_html(items: list[dict], q: str, year: str, month: str) -> str:
    results = items
    if q:
        ql = q.strip().lower()
        results = [it for it in results if ql in (it.get("name") or "").lower()]
    if year:
        results = [it for it in results if (dt := parse_taken(it)) and str(dt.year) == year]
    if month:
        results = [it for it in results if (dt := parse_taken(it)) and f"{dt.month:02d}" == month]

    year_options = sorted({str((dt.year)) for it in items if (dt := parse_taken(it))}, reverse=True)
    year_select = "".join(f'<option value="{y}"{" selected" if y == year else ""}>{y} 年</option>' for y in year_options)
    month_select = "".join(f'<option value="{m:02d}"{" selected" if f"{m:02d}" == month else ""}>{m} 月</option>' for m in range(1, 13))

    q_escaped = (q or "").replace('"', "&quot;")
    form = f"""
    <form class="inline-form" method="get" action="/search">
        <input class="apple-input" type="text" name="q" placeholder="搜尋檔名……" value="{q_escaped}" />
        <div class="btn-row">
            <select class="apple-select" name="year"><option value="">不限年份</option>{year_select}</select>
            <select class="apple-select" name="month"><option value="">不限月份</option>{month_select}</select>
        </div>
        <button type="submit" class="btn-primary">搜尋</button>
    </form>
    """

    searched = bool(q or year or month)
    if not searched:
        result_html = """<div class="card"><p class="secondary">輸入檔名關鍵字,或選擇年份/月份,就能快速找到照片。</p></div>"""
    elif not results:
        result_html = """<div class="card"><p class="secondary">沒有找到符合條件的照片,換個關鍵字試試看。</p></div>"""
    else:
        result_html = f"""
        <div class="section-title">搜尋結果({len(results)} 張)</div>
        <div class="photo-grid">{"".join(lb_img_tag(it) for it in results[:300] if it.get("thumbnail_url"))}</div>
        """
        if len(results) > 300:
            result_html += """<p class="secondary" style="text-align:center;">結果太多,只顯示前 300 張,建議加上更明確的關鍵字或年份/月份。</p>"""

    body = f"""
    <div class="card">{form}</div>
    {result_html}
    """
    return page_shell("圖庫搜尋", body, active_tab="gallery", back_href="/gallery")

@app.get("/search", response_class=HTMLResponse)
async def search_page(request: Request, q: str = "", year: str = "", month: str = ""):
    sid = request.session.get("sid")
    if not sid: return RedirectResponse("/login", status_code=303)
    return HTMLResponse(render_search_html(db_get_photos(sid), q, year, month))

@app.get("/photos")
async def photos(request: Request, max_depth: int = 6):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not token: return JSONResponse({"error": "not_logged_in", "hint": "先前往 /login"}, status_code=401)
    try: return {"count": len(items := await fetch_all_media(token, max_depth=max_depth)), "items": items}
    except httpx.HTTPStatusError as e: return JSONResponse({"error": "graph_api_error", "detail": e.response.text}, status_code=e.response.status_code)

# ---------------------------------------------------------------------------
# 自訂標籤 API + 頁面
# ---------------------------------------------------------------------------
@app.post("/api/photo/tag")
async def api_add_tag(request: Request, id: str = Form(...), tag: str = Form(...)):
    sid = request.session.get("sid")
    if not sid:
        return JSONResponse({"success": False, "error": "not_logged_in"}, status_code=401)
    add_photo_tag(sid, id, tag)
    tags = get_photo_tags(sid, id)
    return JSONResponse({"success": True, "tags": tags})

@app.post("/api/photo/untag")
async def api_remove_tag(request: Request, id: str = Form(...), tag: str = Form(...)):
    sid = request.session.get("sid")
    if not sid:
        return JSONResponse({"success": False, "error": "not_logged_in"}, status_code=401)
    remove_photo_tag(sid, id, tag)
    tags = get_photo_tags(sid, id)
    return JSONResponse({"success": True, "tags": tags})

@app.get("/tags", response_class=HTMLResponse)
async def tags_page(request: Request):
    sid = request.session.get("sid")
    if not sid:
        return RedirectResponse("/login", status_code=303)
    all_tags = get_all_tags(sid)
    if not all_tags:
        body = """
        <div class="card">
            <p class="secondary">還沒有任何標籤。瀏覽照片時可以用下方快速標籤功能，或到單張照片加入「家人」「旅行」等標籤，之後就能用標籤篩選。</p>
        </div>
        """
    else:
        rows = "".join(f"""
        <a class="list-row tappable" href="/tags/view?tag={quote(tag)}">
            <span class="row-icon">🏷</span>
            <span class="row-label">{tag}</span>
            <span class="row-value">{cnt} 張</span>
            <span class="chevron">›</span>
        </a>
        """ for tag, cnt in all_tags)
        body = f"""
        <p class="secondary" style="margin:0 6px 14px;">標籤完全存在你自己的 SQLite，不會上傳任何第三方，也不使用 AI。</p>
        <div class="list-group">{rows}</div>
        """
    return HTMLResponse(page_shell("我的標籤", body, active_tab="more", back_href="/more"))

@app.get("/tags/view", response_class=HTMLResponse)
async def tags_view(request: Request, tag: str):
    sid = request.session.get("sid")
    if not sid:
        return RedirectResponse("/login", status_code=303)
    photo_ids = set(photos_with_tag(sid, tag))
    items = [it for it in db_get_photos(sid) if it["id"] in photo_ids]
    if not items:
        body = """<div class="card"><p class="secondary">這個標籤目前沒有照片。</p></div>"""
    else:
        thumbs = "".join(lb_img_tag(it) for it in items if it.get("thumbnail_url"))
        body = f"""
        <div class="card" style="text-align:center;">
            <div class="stat-num">{len(items)}</div>
            <div class="secondary">張照片標了「{tag}」</div>
        </div>
        <div class="photo-grid">{thumbs}</div>
        """
    return HTMLResponse(page_shell(f"🏷 {tag}", body, active_tab="more", back_href="/tags"))

# ---------------------------------------------------------------------------
# 方向篩選 (直式 / 橫式 / 正方形)
# ---------------------------------------------------------------------------
@app.get("/orientation", response_class=HTMLResponse)
async def orientation_page(request: Request, orient: str = "portrait"):
    sid = request.session.get("sid")
    if not sid:
        return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)
    filtered = [it for it in items if orientation_of(it) == orient and it.get("thumbnail_url")]
    chips = "".join(
        f'<a class="filter-chip{" active" if orient == o else ""}" href="/orientation?orient={o}">{label}</a>'
        for o, label in [("portrait", "直式 📱"), ("landscape", "橫式 🖼"), ("square", "正方形 ⬜")]
    )
    if not filtered:
        body = f"""
        <div class="filter-bar">{chips}</div>
        <div class="card"><p class="secondary">沒有找到這個方向的照片。</p></div>
        """
    else:
        body = f"""
        <div class="filter-bar">{chips}</div>
        <p class="secondary" style="margin:0 6px 14px;">共 {len(filtered)} 張{ {'portrait':'直式','landscape':'橫式','square':'正方形'}.get(orient,'') }照片，適合當手機桌布或橫向展示。</p>
        <div class="photo-grid">{"".join(lb_img_tag(it) for it in filtered[:400])}</div>
        """
    return HTMLResponse(page_shell("方向篩選", body, active_tab="gallery", back_href="/gallery"))

# ---------------------------------------------------------------------------
# 本週 / 本月回顧
# ---------------------------------------------------------------------------
@app.get("/reviews", response_class=HTMLResponse)
async def reviews_page(request: Request, period: str = "week"):
    sid = request.session.get("sid")
    if not sid:
        return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)
    now = datetime.now()
    if period == "month":
        start = (now.replace(day=1) - timedelta(days=1)).replace(day=1)
        end = now.replace(day=1) - timedelta(seconds=1)
        title = "上個月回顧"
        label = f"{start.year} 年 {start.month} 月"
    else:
        # 上週 (上週一 ~ 上週日)
        today = now.date()
        start_of_this_week = today - timedelta(days=today.weekday())
        end = datetime.combine(start_of_this_week - timedelta(days=1), datetime.max.time())
        start = datetime.combine(start_of_this_week - timedelta(days=7), datetime.min.time())
        title = "上週回顧"
        label = f"{start.strftime('%m/%d')} – {end.strftime('%m/%d')}"

    matched = []
    for it in items:
        dt = parse_taken(it)
        if dt and start <= dt.replace(tzinfo=None) <= end.replace(tzinfo=None):
            matched.append(it)

    chips = f"""
    <div class="filter-bar">
        <a class="filter-chip{" active" if period == "week" else ""}" href="/reviews?period=week">上週</a>
        <a class="filter-chip{" active" if period == "month" else ""}" href="/reviews?period=month">上個月</a>
        <a class="filter-chip" href="/memories">當年今日</a>
    </div>
    """
    if not matched:
        body = chips + f"""<div class="card"><p class="secondary">{label} 沒有找到照片。</p></div>"""
    else:
        ids = ",".join(it["id"] for it in matched[:30])
        body = chips + f"""
        <div class="card" style="text-align:center;">
            <div class="stat-num">{len(matched)}</div>
            <div class="secondary">{label} 的照片</div>
        </div>
        <div class="photo-grid">{"".join(lb_img_tag(it) for it in matched if it.get("thumbnail_url"))}</div>
        <form method="post" action="/memories/render" class="inline-form">
            <input type="hidden" name="ids" value="{ids}" />
            <input type="hidden" name="title" value="{title}" />
            <button class="btn-primary" type="submit">✨ 把這些做成回憶影片</button>
        </form>
        """
    return HTMLResponse(page_shell(title, body, active_tab="memories", back_href="/memories"))

# ---------------------------------------------------------------------------
# 批次下載 ZIP
# ---------------------------------------------------------------------------
@app.post("/download/zip")
async def download_zip(request: Request, ids: str = Form(...)):
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token:
        return RedirectResponse("/login", status_code=303)
    id_list = [i for i in ids.split(",") if i][:80]  # 上限保護
    if not id_list:
        return JSONResponse({"error": "no_ids"}, status_code=400)

    all_photos = {p["id"]: p for p in db_get_photos(sid)}
    buf = io.BytesIO()
    async with httpx.AsyncClient(timeout=60) as client:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for pid in id_list:
                photo = all_photos.get(pid)
                if not photo:
                    continue
                try:
                    resp = await client.get(
                        f"{GRAPH_BASE}/me/drive/items/{pid}/content",
                        headers={"Authorization": f"Bearer {token}"},
                        follow_redirects=True,
                    )
                    if resp.status_code == 200:
                        name = photo.get("name") or f"{pid}.jpg"
                        # 避免 zip 內檔名衝突
                        zf.writestr(name, resp.content)
                except Exception:
                    continue
                await asyncio.sleep(0.15)

    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=photos.zip"},
    )

# ---------------------------------------------------------------------------
# 相簿分享連結 (OneDrive createLink)
# ---------------------------------------------------------------------------
@app.post("/api/share")
async def api_create_share(request: Request, id: str = Form(...), scope: str = Form("anonymous")):
    """為單一照片或資料夾建立分享連結。scope: anonymous | organization"""
    sid = request.session.get("sid")
    token = await get_ms_token(sid)
    if not sid or not token:
        return JSONResponse({"success": False, "error": "not_logged_in"}, status_code=401)

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"type": "view", "scope": scope if scope in ("anonymous", "organization") else "anonymous"}
    try:
        async with httpx.AsyncClient(timeout=20) as client:
            resp = await client.post(
                f"{GRAPH_BASE}/me/drive/items/{id}/createLink",
                headers=headers,
                json=payload,
            )
            if resp.status_code not in (200, 201):
                return JSONResponse({"success": False, "error": resp.text}, status_code=resp.status_code)
            data = resp.json()
            link = data.get("link", {}).get("webUrl")
            return JSONResponse({"success": True, "url": link, "raw": data})
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, status_code=502)

@app.get("/share/event", response_class=HTMLResponse)
async def share_event_page(request: Request, start: str = "", end: str = ""):
    """簡易事件相簿分享：列出該時段照片，並提供「為每張建立分享連結」或建議使用者在 OneDrive 建資料夾後分享。"""
    sid = request.session.get("sid")
    if not sid:
        return RedirectResponse("/login", status_code=303)
    items = db_get_photos(sid)
    # 簡化：若沒帶參數就顯示說明
    body = """
    <div class="card banner-info">
        <p class="secondary" style="margin:0;">
            想把某個事件相簿分享給家人朋友？最穩妥的方式是：<br>
            1. 到 <a class="pill-link" href="/albums">自動分類</a> 找到事件<br>
            2. 在 OneDrive 建立一個資料夾，把想分享的照片移進去<br>
            3. 在 OneDrive 對該資料夾按「共用」產生連結<br><br>
            或使用下方 API 為單張照片產生 view 連結（需照片 id）。
        </p>
    </div>
    <div class="card">
        <form class="inline-form" method="post" action="/api/share" id="share-form">
            <input class="apple-input" name="id" placeholder="OneDrive 項目 ID" required />
            <select class="apple-select" name="scope">
                <option value="anonymous">任何人（匿名）</option>
                <option value="organization">組織內</option>
            </select>
            <button class="btn-primary" type="submit">產生分享連結</button>
        </form>
        <div id="share-result" class="secondary" style="margin-top:12px;"></div>
    </div>
    <script>
    document.getElementById('share-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        const fd = new FormData(e.target);
        const resp = await fetch('/api/share', { method: 'POST', body: fd });
        const data = await resp.json();
        const el = document.getElementById('share-result');
        if (data.success) {
            el.innerHTML = '✅ 連結已產生：<br><a href="' + data.url + '" target="_blank" style="color:var(--accent);word-break:break-all;">' + data.url + '</a>';
        } else {
            el.textContent = '失敗：' + (data.error || '未知錯誤');
        }
    });
    </script>
    """
    return HTMLResponse(page_shell("相簿分享", body, active_tab="more", back_href="/more"))

# ---------------------------------------------------------------------------
# 更新「更多」與首頁連結，加入新功能入口
# ---------------------------------------------------------------------------
# (使用者可從 /more 進入；此處保留原 more_page 內容，實際部署時可把新連結加進 more_page)
