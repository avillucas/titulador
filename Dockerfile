FROM python:3.11-slim

# Install LibreOffice, CUPS printer client, fonts, and Tkinter for GUI support
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    libreoffice-java-common \
    cups-client \
    fonts-dejavu \
    fonts-liberation \
    fonts-freefont-ttf \
    python3-tk \
    tk \
    tcl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source and templates
COPY . .

# Default command
CMD ["python", "main.py", "batch"]
