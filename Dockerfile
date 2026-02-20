FROM python:3.10-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt


COPY api.py .
COPY module.py .
COPY template_document ./template_document


CMD ["uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000"]