FROM python:3.12-slim

WORKDIR /app

# catdoc reads the old binary .doc format
RUN apt-get update \
    && apt-get install -y --no-install-recommends catdoc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

EXPOSE 8080

CMD ["streamlit", "run", "app.py", \
     "--server.port=8080", \
     "--server.address=0.0.0.0", \
     "--server.headless=true"]
