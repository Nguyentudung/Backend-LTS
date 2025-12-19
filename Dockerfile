FROM python:3.13.11-slim

# ===============================
# System dependencies
# ===============================
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        imagemagick \
        ghostscript \
        libreoffice-core \
        libreoffice-common \
        libreoffice-draw \
        libreoffice-writer \
        fonts-dejavu \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# ===============================
# ImageMagick 7 policy
# ===============================
RUN sed -i 's/rights="none"/rights="read|write"/g' /etc/ImageMagick-7/policy.xml

# ===============================
# LibreOffice runtime fix
# ===============================
ENV HOME=/tmp
RUN mkdir -p /tmp && chmod 777 /tmp

WORKDIR /app

# ===============================
# Python deps
# ===============================
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PYTHONUNBUFFERED=1

CMD ["uvicorn", "src.server:app", "--host", "0.0.0.0", "--port", "8080"]
