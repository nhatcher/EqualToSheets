FROM python:3.11

ENV PYTHONUNBUFFERED 1

WORKDIR /server

COPY . /packages

# install system dependencies
RUN apt-get update && apt-get -y upgrade
RUN apt-get install -y --no-install-recommends curl
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
ENV PATH="/root/.cargo/bin:${PATH}"

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
