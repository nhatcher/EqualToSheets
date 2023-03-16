#!/bin/bash

docker-compose run --rm server bash -c "black ./serverless ./server && isort ./serverless ./server"
