#!/bin/bash

docker-compose run --rm server bash -c "
  mypy ./serverless &&
  flake8 ./serverless &&
  black --check ./serverless &&
  isort --check ./serverless
"
