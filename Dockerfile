FROM python:3.13-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1
ENV PRICING_SERVICE_ONLY=1

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY preview_server.py uprinting_scraper.py ./

EXPOSE 8877

CMD ["python", "preview_server.py"]
