# Build stage for generating icons (optional, only if icons need regeneration)
FROM python:3.12-slim AS builder

# Provide uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /bin/

WORKDIR /build

# Install Pillow dependencies for icon generation
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY generate_icons.py .

# Pillow is only needed to generate icons, not part of the project deps
RUN uv pip install --system --no-cache pillow==11.3.0

# Generate icons
RUN mkdir -p static/icons && python generate_icons.py


# Production stage
FROM python:3.12-slim

# Provide uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /bin/

LABEL org.opencontainers.image.title="Summa"
LABEL org.opencontainers.image.description="Invoice management and expense tracking"
LABEL org.opencontainers.image.source="https://github.com/itacentury/summa"

WORKDIR /app

# Install gosu for proper user switching and create non-root user
RUN apt-get update && apt-get install -y --no-install-recommends \
    gosu \
    && rm -rf /var/lib/apt/lists/* \
    && useradd --create-home --shell /bin/bash appuser

# Install runtime dependencies only (no dev group, no editable install of the app)
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project

# Use the project virtualenv for all subsequent commands
ENV PATH="/app/.venv/bin:$PATH"

# Copy application files
COPY app.py .
COPY templates/ templates/
COPY static/ static/

# Copy generated icons from builder stage
COPY --from=builder /build/static/icons/ static/icons/

# Copy and setup entrypoint
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Create data directory for SQLite database
RUN mkdir -p /data && chown appuser:appuser /data

# Set environment variables
ENV FLASK_APP=app.py
ENV FLASK_ENV=production
ENV PYTHONUNBUFFERED=1
ENV DATABASE_PATH=/data/invoices.db

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/')" || exit 1

# Use entrypoint for permission handling, then run gunicorn
ENTRYPOINT ["/entrypoint.sh"]
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "2", "--threads", "4", "app:app"]
