FROM python:3.12-slim

# ✅ System-Dependencies für pip builds
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    gcc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN python -m pip install --upgrade pip setuptools wheel
RUN pip install --no-cache-dir -r requirements.txt

COPY src ./src

ENV PORT=8080
EXPOSE 8080

CMD ["python", "src/server.py"]
