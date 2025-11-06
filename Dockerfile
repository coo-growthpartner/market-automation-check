FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    && wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /usr/share/keyrings/googlechrome-linux-keyring.gpg \
    && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/googlechrome-linux-keyring.gpg] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google.list

RUN apt-get update && apt-get install -y \
    google-chrome-stable \
    chromium-driver \
    && rm -rf /var/lib/apt/lists/*

RUN useradd -m -d /home/chrome chrome \
    && chown -R chrome:chrome /home/chrome

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /app/logs && \
    chown -R chrome:chrome /app/logs

RUN mkdir -p /home/chrome/.cache/selenium \
    && mkdir -p /home/chrome/chrome-data \
    && chown -R chrome:chrome /app \
    && chown -R chrome:chrome /home/chrome/.cache \
    && chown -R chrome:chrome /home/chrome/chrome-data

USER chrome

# 실행 및 로깅
CMD ["python", "-u", "main.py"]