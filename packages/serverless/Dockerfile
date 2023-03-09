FROM python:3.11

ENV PYTHONUNBUFFERED 1

WORKDIR /server

COPY . /packages

# install system dependencies
RUN apt-get update && apt-get -y upgrade
RUN apt-get install -y --no-install-recommends curl
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
ENV PATH="/root/.cargo/bin:${PATH}"
RUN curl -fsSL https://deb.nodesource.com/setup_18.x | bash -
RUN apt-get install -y nodejs
RUN npm install -g npm@9.2.0

# build the client
RUN cd /packages/serverless/client && npm install && npm run build
RUN mkdir /static
RUN cp -r /packages/serverless/client/build /static/client
RUN find /static/client -type f -regex ".*.map" | xargs rm

# compile the python SDK
RUN ln -s /packages/calc /calc  # that's where python SDK looks for its dependencies
RUN pip3 install /packages/sdk/python

# cleanup
RUN rm -rf /packages && rm /calc

ADD ./serverless/server /server/

# install server dependencies
RUN pip3 --default-timeout=300 install -r requirements.txt

EXPOSE 5000

RUN chmod +x /server/start_server.sh
CMD /server/start_server.sh
