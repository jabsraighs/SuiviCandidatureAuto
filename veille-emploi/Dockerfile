FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY veille_emploi.py run_loop.py config.py gmail_alertes_emploi.py excel_offres.py ./

ENV DATA_DIR=/app/data
ENV OAUTH_PORT=8080
ENV INTERVAL_HOURS=3
ENV PYTHONUNBUFFERED=1

# Port utilisé uniquement lors de la toute première autorisation Gmail
EXPOSE 8080

CMD ["python", "run_loop.py"]