FROM python:3.10-slim

# Libre office installation
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    fonts-liberation \
    default-jre-headless \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt


COPY api.py .
COPY module.py .
COPY template_document ./template_document


CMD ["uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000"]