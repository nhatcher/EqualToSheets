#!/bin/bash

docker-compose run --rm -e DJANGO_TESTS=True server python3 manage.py test
