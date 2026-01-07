# IT Operations Dashboard - Docker Container
# Uses Python 3.10 slim image for smaller size
# Works identically on Windows and Mac with Docker Desktop

FROM python:3.10-slim

# Set the working directory inside the container
WORKDIR /app

# Copy only requirements first (leverages Docker layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create .env directory if it doesn't exist (for users.json)
RUN mkdir -p .env

# Environment variables for Flask
ENV FLASK_APP=app.py
ENV FLASK_RUN_HOST=0.0.0.0
ENV PYTHONUNBUFFERED=1

# Expose the port the app runs on
EXPOSE 5000

# Start the application using Gunicorn (Production-grade server)
# Workers: 2 for small deployments, increase for more traffic
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--access-logfile", "-", "app:app"]
