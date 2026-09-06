# Use official Python image
FROM python:3.11-slim

# Prevent Python from writing .pyc files and buffering logs
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set working directory
WORKDIR /app

# Install system packages for psycopg & Pillow
RUN apt-get update && apt-get install -y \
    build-essential \
    libpq-dev \
    gcc \
    && apt-get clean

# Copy requirements and install dependencies
COPY requirements.txt /app/
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Copy project files
COPY . /app/

# List files for debugging
RUN ls -la /app/

# Collect static files
RUN python manage.py collectstatic --noinput || true

# Expose port
EXPOSE 8000

# Run Django using Gunicorn with verbose logging
CMD ["gunicorn", "certificate_generator.wsgi:application", "--bind", "0.0.0.0:8000", "--workers=2", "--log-level=debug", "--access-logfile=-", "--error-logfile=-"]