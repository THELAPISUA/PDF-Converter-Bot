FROM python:3.12-trixie

COPY main.py app/
COPY requirements.txt /

RUN pip install -r requirements.txt

CMD ["python3", "app/main.py"]