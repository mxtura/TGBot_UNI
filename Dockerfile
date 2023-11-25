FROM python:3.10
WORKDIR /app
COPY . .
RUN pip install --upgrade pip
RUN pip install google-api-python-client
RUN pip install google-auth
RUN pip install -r requirements.txt
CMD ["python", "bot.py"]