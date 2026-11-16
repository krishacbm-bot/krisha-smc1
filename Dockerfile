# Playwright image already contains Chromium and its system libraries.
# Keep the tag version equal to the playwright version in requirements.txt.
FROM mcr.microsoft.com/playwright/python:v1.49.0-jammy

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV SMC_SERVER=1 PYTHONUNBUFFERED=1
EXPOSE 10000

# ONE worker on purpose: scan progress is kept in memory, so all requests must hit the same process.
CMD ["sh", "-c", "gunicorn smc_dashboard:app --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 16 --timeout 300"]
