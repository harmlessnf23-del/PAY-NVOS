FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    SVERKA_WORK_DIR=/data

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY app.py core.py ./
COPY templates ./templates
COPY static ./static

RUN mkdir -p /data && useradd -r -u 1000 -d /app sverka && \
    chown -R sverka /app /data
USER sverka

EXPOSE 8021

# 2 воркера хватает на 5–10 одновременных пользователей.
# Таймаут 600 с — сверка крупных файлов (десятки тысяч строк) длится до минуты.
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:8021", \
     "--timeout", "600", "--access-logfile", "-", "app:app"]
