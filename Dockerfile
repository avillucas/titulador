FROM python:3.11-slim

# Install LibreOffice, CUPS printer client, and fonts for A5 PDF conversion
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    libreoffice-java-common \
    cups-client \
    fonts-dejavu \
    fonts-liberation \
    fonts-freefont-ttf \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source and templates
COPY . .

# Default command
ENTRYPOINT ["python", "main.py"]
CMD ["batch"]
