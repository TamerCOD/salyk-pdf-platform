FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# Кириллический шрифт для PDF + базовые tools
RUN apt-get update \
 && apt-get install -y --no-install-recommends fonts-dejavu fonts-dejavu-core \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

ENV PORT=8080
EXPOSE 8080

# 1 worker (in-memory job store для ИНН-парсера требует одного процесса), 8 threads.
# Timeout 120s — фоновая обработка ИНН живёт в Thread'е и не зависит от request timeout.
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-8080} --workers 1 --threads 8 --timeout 120 app:app"]
