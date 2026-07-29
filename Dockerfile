FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY gmail_candidature_tracker.py config.py run_loop.py ./

ENV DATA_DIR=/app/data
ENV OAUTH_PORT=8080
ENV INTERVAL_HOURS=3
ENV PYTHONUNBUFFERED=1

# Port utilisé uniquement lors de la toute première autorisation Gmail
EXPOSE 8080

CMD ["python", "run_loop.py"]