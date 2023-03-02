#!/bin/bash

# Creates requirements.txt based on requirements.in file
docker-compose run --rm server bash -c "
  python3 -m pip install pip-tools &&
  pip-compile -o requirements.txt requirements.in --resolver=backtracking
"
