FROM python:3.10

ENV PYTHONUNBUFFERED 1

WORKDIR /server

ADD requirements.txt /server
RUN pip3 --default-timeout=300 install -r requirements.txt

ADD . /server/

EXPOSE 8000

CMD gunicorn --config wsgi/gunicorn_config.py wsgi.wsgi:app
