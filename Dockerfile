FROM python:3.13-slim

# Rebuild pricing when Waterproof Menus quantity stays an option id (41858).
WORKDIR /app

ENV PYTHONUNBUFFERED=1
ENV PRICING_SERVICE_ONLY=1

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

ARG CACHEBUST=wp-41858
RUN echo "cachebust ${CACHEBUST}"
COPY preview_server.py uprinting_scraper.py ./

EXPOSE 8877

CMD ["python", "preview_server.py"]
