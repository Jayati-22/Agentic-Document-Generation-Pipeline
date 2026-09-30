
# FROM python:3.11-slim

# WORKDIR /app

# COPY src/ /app

# # Install the Nasiko A2A framework + Document/Pexels tools
# RUN pip install --no-cache-dir \
#     "a2a-sdk[http-server]>=0.3.0" \
#     click>=8.1.8 \
#     httpx>=0.28.1 \
#     openai>=1.57.0 \
#     pydantic>=2.11.4 \
#     python-dotenv>=1.1.0 \
#     uvicorn>=0.34.2 \
#     python-pptx \
#     openpyxl \
#     requests

# ENV PYTHONUNBUFFERED=1

# CMD ["python", "__main__.py", "--host", "0.0.0.0", "--port", "5000"]

FROM python:3.11-slim

WORKDIR /app

COPY src/ /app

# Install the Nasiko A2A framework + Document/Pexels tools
RUN pip install --no-cache-dir \
    "a2a-sdk[http-server]>=0.3.0,<1.0" \
    "click>=8.1.8" \
    "httpx>=0.28.1" \
    "openai>=1.57.0" \
    "pydantic>=2.11.4" \
    "python-dotenv>=1.1.0" \
    "uvicorn>=0.34.2" \
    python-pptx \
    openpyxl \
    requests

ENV PYTHONUNBUFFERED=1

CMD ["python", "__main__.py", "--host", "0.0.0.0", "--port", "5000"]