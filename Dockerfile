FROM python:3.11-slim

# 安裝中文字體（ReportLab PDF 中文支援）
RUN apt-get update && apt-get install -y \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 先複製 requirements 利用 Docker layer cache
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製所有程式碼
COPY . .

# Cloud Run 用 /tmp 存放資料（可寫入，ephemeral）
ENV DATA_FILE=/tmp/targets.xlsx
ENV OUTPUTS_DIR=/tmp/outputs
ENV PORT=8080

# 啟動：從 repo root 執行 gunicorn，使用 web.app:app
CMD exec gunicorn -w 1 -b 0.0.0.0:$PORT --timeout 120 web.app:app
