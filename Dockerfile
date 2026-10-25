FROM python:3.12-slim

WORKDIR /app
ENV PYTHONPATH=/app/src \
    PYTHONUNBUFFERED=1

RUN pip install --no-cache-dir "mistralai-workflows[mistralai]>=3.10" "pydantic>=2"

COPY src/ ./src/

CMD ["python", "src/entrypoints/worker.py"]
