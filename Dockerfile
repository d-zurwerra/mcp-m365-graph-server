FROM python:3.12-slim

# --- System dependencies (needed for some pip packages on slim images) ---
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    gcc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# --- Python deps first (better layer caching) ---
COPY requirements.txt .

# Optional but recommended: ensure modern build tooling
RUN python -m pip install --upgrade pip setuptools wheel

# Install deps (verbose to show the REAL failing package if something breaks)
RUN pip install --no-cache-dir -r requirements.txt -vvv

# --- App code ---
COPY src ./src

ENV PORT=8080
EXPOSE 8080

CMD ["python", "src/server.py"]
