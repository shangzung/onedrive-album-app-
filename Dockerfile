FROM python:3.11-slim

# ffmpeg 是產生回憶影片必需的系統工具,pip 裝不到,要用 apt 裝
RUN apt-get update && apt-get install -y --no-install-recommends \
    ffmpeg \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# renders / music / media_cache 是程式會寫入的資料夾,建議搭配平台的 Persistent Disk 掛載在這幾個路徑
RUN mkdir -p renders music media_cache/google

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
