#!/bin/bash

docker-compose run --rm server bash -c "black ./serverless && isort ./serverless"
