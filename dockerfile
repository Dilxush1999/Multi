# Python 3.12 base image (Render'ning default'iga mos)
FROM python:3.12-slim

# Tizim paketlarini o'rnatish (apt-get bu yerda ishlaydi, read-only emas)
RUN apt-get update && apt-get install -y \
    libreoffice \
    tesseract-ocr \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# Ishchi papkani o'rnatish
WORKDIR /app

# Requirements.txt ni ko'chirish va kutubxonalarni o'rnatish
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Barcha kodni ko'chirish
COPY . .

# Tesseract yo'lini sozlash (kodda allaqachon bor, lekin ta'minlash uchun)
ENV PYTESSERACT_TESSERACT_CMD=/usr/bin/tesseract

# Start command (Flask/Gunicorn uchun)
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "app:flask_app"]
