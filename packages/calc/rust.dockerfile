# Rust base image with clippy, rustfmt and wasm-pack

# Must be build using this command from the cal directory - (note the '.' build context):
# docker buildx build --platform linux/arm64,linux/amd64  -f rust.dockerfile . -t equaltodev/rust:{rust-version}-{image-version} --push

FROM rust:1.65.0-slim

ARG DEBIAN_FRONTEND=noninteractive

ADD . /calc
WORKDIR /calc

# Note that we're on Python 3.10, but the `python3-dev` pulls in the dev package for Python 3.9
# (which is in debian-stable, which the rust image is based off). This works for now, but something
# to look at if we start having issues.

RUN apt update \
  && apt -y --no-install-recommends upgrade \
  && apt install --no-install-recommends -y curl make python3-dev git binaryen pkg-config libssl-dev \
  && rm -rf /var/cache/apt/archives/*.deb \
  && rustup component add clippy rustfmt \
  && rustup target add wasm32-unknown-unknown \
  && ./install-wasm  || ( \
      cargo install wasm-pack --version 0.10.3 --locked \
    ) \
  && cargo install wasm-bindgen-cli --version 0.2.83 --locked \
  && make fetch \
  && chmod -R a+rwX /usr/local/cargo/registry


# We run the rust builds/tests as the current host user (via `EQUALTO_DOCKER_USER` - see
# `docker-compose.yml` and elsewhere). However, that user would likely not have a homedir, which
# a few things (like wasm-pack) need. So, we set `HOME` to something writable.

ENV HOME=/tmp
