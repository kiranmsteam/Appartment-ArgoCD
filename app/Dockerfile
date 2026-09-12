# Dockerfile – used by Fly.io and any Docker-compatible host
# Build:  docker build -t apartment-mgmt .
# Run:    docker run -p 8080:8080 --env-file .env apartment-mgmt

FROM python:3.11-slim

WORKDIR /app

# Install dependencies first (Docker layer-cache optimisation)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create the data directory that the app uses at runtime
RUN mkdir -p data Apartment/uploads Apartment/announcement_uploads

EXPOSE 8080

# Gunicorn binds to the PORT env-var supplied by the platform (default 8080)
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT:-8080} --workers 2 --timeout 120 app:app"]
