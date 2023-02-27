FROM rust:1.66.1-slim

ENV PYTHONUNBUFFERED 1

ARG GIT_COMMIT
ENV GIT_COMMIT $GIT_COMMIT

WORKDIR /server

RUN apt-get update \
  && apt-get -y upgrade \
  && apt-get install -y --no-install-recommends python3 pip

# compile python SDK
RUN mkdir /calc
COPY ./calc /calc
RUN mkdir /python_sdk
COPY ./sdk/python /python_sdk
RUN pip3 install /python_sdk
RUN rm -rf /calc /python_sdk

ADD ./sheet_ai/server /server/
RUN pip3 --default-timeout=300 install -r requirements.txt

EXPOSE 8000

CMD gunicorn --config wsgi/gunicorn_config.py wsgi.wsgi:app
