FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libcairo2 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf-xlib-2.0-0 \
    libffi8 \
    libxml2 \
    libxslt1.1 \
    shared-mime-info \
    fonts-noto \
    fonts-noto-cjk \
    gcc \
    g++ \
    python3-dev \
    libffi-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir \
        flask>=3.0.0 \
        flask-cors>=4.0.0 \
        gunicorn>=21.0.0 \
        requests>=2.31.0 \
        jinja2>=3.1.0 \
        Pillow>=10.0.0 \
        qrcode>=7.4.0 \
        reportlab>=4.0.0 \
        pypdf>=4.0.0 \
        openpyxl>=3.1.0 \
        weasyprint>=60.0 \
        fonttools>=4.40.0

COPY . .

ENV PORT=8080
ENV PYTHONUNBUFFERED=1

EXPOSE 8080

CMD exec gunicorn \
    --bind 0.0.0.0:$PORT \
    --workers 2 \
    --threads 2 \
    --timeout 120 \
    --worker-class gthread \
    --log-level info \
    app:app
