# OneDrive Album App — MVP 後端

用你已經在 Azure Entra ID 註冊好的「OneDrive Album App」,做 OAuth 登入並列出使用者 OneDrive 裡的所有照片/影片。

## 1. 從 Azure 拿設定值

回到 Azure Entra ID > 應用程式註冊 > OneDrive Album App:

- **概觀**頁面複製 `應用程式 (用戶端) 識別碼` → 這是 `CLIENT_ID`
- **憑證及祕密** > 新增用戶端密碼,建立後**立刻複製 Value**(離開頁面就看不到了)→ 這是 `CLIENT_SECRET`
- **驗證 (Authentication)** > 新增平台 > Web,填入 Redirect URI:`http://localhost:8000/callback`

## 2. 安裝與設定

```bash
cd onedrive-album-app
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
# 編輯 .env,填入 CLIENT_ID / CLIENT_SECRET
```

## 3. 啟動

```bash
uvicorn main:app --reload --port 8000
```

瀏覽器打開 http://localhost:8000,點「用 Microsoft 帳號登入」,登入你自己的帳號並同意授權後,
前往 http://localhost:8000/photos 應該會看到類似這樣的 JSON:

```json
{
  "count": 42,
  "items": [
    {
      "id": "...",
      "name": "IMG_0001.jpg",
      "mimeType": "image/jpeg",
      "size": 3456789,
      "webUrl": "https://onedrive.live.com/...",
      "thumbnailUrl": "https://...",
      "takenDateTime": "2025-07-01T10:23:00Z",
      "latitude": 25.0,
      "longitude": 121.5
    }
  ]
}
```

## 已知限制(MVP 階段,之後再補)

- Token 只存在 session cookie,重啟伺服器/token 過期(約 1 小時)就要重新登入。正式版要用 `refresh_token`
  搭配後端資料庫做自動換發。
- 大量照片時遞迴呼叫會比較慢,之後可以改用 Graph 的 `/delta` 端點做增量同步,而不是每次全部重新列出。
- 目前 `TENANT_ID=common`,任何 Microsoft 帳號都能登入;若只想開放給組織內部帳號,改成你的 Tenant ID。

## 下一步建議

1. 把 `/photos` 回傳的縮圖網址畫成一個簡單的網格畫面(前端可以用 React + 這個 API)。
2. 把每張照片下載下來算 pHash,做重複照片偵測。
3. 用 CLIP 對縮圖或原圖算 embedding,存進向量資料庫,做自然語言搜尋的雛形。
