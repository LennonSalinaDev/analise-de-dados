FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV HOST=0.0.0.0
ENV HOME=/tmp
ENV SAL_USE_VCLPLUGIN=svp

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        default-jre-headless \
        fontconfig \
        fonts-dejavu \
        fonts-liberation \
        fonts-noto-core \
        libreoffice-calc \
        libreoffice-java-common \
        libreoffice-writer \
    && fc-cache -f \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["python", "app.py"]
