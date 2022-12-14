FROM equaltodev/rust:1.65.0-2
# FROM equaltodev/rust-arm:1.65.0-2 - to run on ARM

# To build/test cross arch, add a `--platform` arg to the FROM line above.
# For example, to build/test arm64 (M1 macs), the above should say:
#
#     FROM --platform=linux/arm64 equaltodev/rust:1.65.0-1
#
# and then do the following:
#
# 1. docker rmi equalto_rust:latest
# 2. docker rmi equaltodev/rust:1.65.-0  (matching the version above)
# 3. docker-compose build rust
#
# And then test with:
#
#     docker-compose run --rm rust uname -m
#
# which should output `aarch64` for the M1/arm64 case.


ENV HOME=/tmp
