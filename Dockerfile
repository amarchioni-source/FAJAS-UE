FROM python:3.12-slim

# LibreOffice (solo Calc) para convertir los Excel a PDF con el formato exacto de la plantilla
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc fonts-liberation fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY fonts.conf /etc/fonts/local.conf
RUN fc-cache -f

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PYTHONUNBUFFERED=1 \
    HOME=/tmp \
    JOBS_DIR=/tmp/anexos_jobs

# Un solo worker: la conversión a PDF (LibreOffice) se serializa para cuidar la memoria
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 4 --timeout 300 app:app"]
