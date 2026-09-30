FROM python:3.12-slim-bookworm

# Prevent interactive prompts during apt installs
ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1
ENV HEADLESS=true
ENV HEADLESS_GOFOOD=true
ENV HEADLESS_GRAB=true
ENV HEADLESS_SHOPEE=true
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

# Install required OS libraries, utilities, and fonts for headless browser automation
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    curl \
    gnupg \
    ca-certificates \
    fonts-liberation \
    fonts-noto-color-emoji \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libc6 \
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libexpat1 \
    libfontconfig1 \
    libgbm1 \
    libgcc1 \
    libglib2.0-0 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libstdc++6 \
    libx11-6 \
    libx11-xcb1 \
    libxcb1 \
    libxcomposite1 \
    libxcursor1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxi6 \
    libxrandr2 \
    libxrender1 \
    libxss1 \
    libxtst6 \
    xdg-utils \
    procps \
    && rm -rf /var/lib/apt/lists/*

# Install Google Chrome Stable for Selenium automation
RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /etc/apt/trusted.gpg.d/google.gpg \
    && echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends google-chrome-stable \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright Chromium browser shared system-wide
RUN mkdir -p /ms-playwright \
    && playwright install chromium \
    && chmod -R 777 /ms-playwright

# Copy project files
COPY . .

# Setup application user (UID 1000 matching standard host user)
RUN useradd -m -u 1000 -s /bin/bash appuser \
    && mkdir -p /app/output_owners \
                /app/SHOPEE/data \
                /app/GRAB/sessions \
                /app/GOFOOD/session \
                /app/cronjobsot/reports \
                /app/cache \
    && chown -R appuser:appuser /app

USER appuser

CMD ["python", "discord_bot.py"]
