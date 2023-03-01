FROM rust:1.66.1-slim

ENV PYTHONUNBUFFERED 1

ARG GIT_COMMIT
ENV GIT_COMMIT $GIT_COMMIT

WORKDIR /server

COPY . /packages

# install system dependencies
RUN apt-get update && apt-get -y upgrade
RUN apt-get install -y --no-install-recommends python3 pip curl pkg-config openssl libssl-dev
RUN curl -fsSL https://deb.nodesource.com/setup_18.x | bash -
RUN apt-get install -y nodejs
RUN npm install -g npm@9.2.0
RUN /packages/sdk/typescript/util/install-wasm-pack.sh

# build the client
RUN cd /packages/sheet_ai/client && npm run build-full
RUN cp -r /packages/sheet_ai/client/build /static
RUN find /static -type f -regex ".*.map" | xargs rm

# compile the python SDK
RUN ln -s /packages/calc /calc  # that's where python SDK looks for its dependencies
RUN pip3 install /packages/sdk/python

# cleanup
RUN rm -rf /packages && rm /calc

ADD ./sheet_ai/server /server/

# install server dependencies
RUN pip3 --default-timeout=300 install -r requirements.txt

EXPOSE 8000

CMD gunicorn --config wsgi/gunicorn_config.py wsgi.wsgi:app
