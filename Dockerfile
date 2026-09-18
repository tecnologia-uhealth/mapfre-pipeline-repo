FROM python:3.11-slim

WORKDIR /app

# LibreOffice es necesario para convertir la Orden de Trabajo de .xls a .pdf
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instala Chromium + todas las librerías de sistema que necesita Playwright
RUN playwright install --with-deps chromium

COPY . .

ENV PORT=5003
EXPOSE 5003

CMD ["python3", "webhook_mapfre.py"]
