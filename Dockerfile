
# Use lightweight Python image
FROM python:3.11-slim

# Disable Streamlit analytics
ENV STREAMLIT_TELEMETRY_DISABLED=true

WORKDIR /app
COPY . .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose Render’s port
EXPOSE 8080

# Start Streamlit
CMD ["streamlit", "run", "streamlit_app.py", "--server.port=8080", "--server.address=0.0.0.0"]
