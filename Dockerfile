FROM python:3.12-slim

# Arbeitsverzeichnis
WORKDIR /app

# System-Dependencies (minimal)
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Requirements zuerst (Docker Layer Caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App Code
COPY app/ ./app/

# Non-root User für Security
RUN useradd -m -u 1000 mcpuser
USER mcpuser

# Port
EXPOSE 8000

# Healthcheck
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

# Start
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--forwarded-allow-ips", "*"]
