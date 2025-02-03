# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Install required system packages for LibreOffice (for Linux PDF conversion)
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Copy the entire application
COPY . .

# Expose the Streamlit default port
EXPOSE 8080

# Run the Streamlit app
CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0"]
