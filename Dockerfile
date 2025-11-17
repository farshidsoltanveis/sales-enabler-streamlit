# ---- Base ----
FROM python:3.11-slim

# Keep image lean
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Workdir
WORKDIR /app

# Install system dependencies for PDF processing
RUN apt-get update && apt-get install -y \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies first (cache-friendly)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY . .

# Create necessary directories
RUN mkdir -p input/web_runs output/parsed output/integrated output/insights

# Render provides $PORT; expose default for local runs
ENV PORT=8080
EXPOSE 8080

# Run Flask app (production mode)
CMD python app.py
