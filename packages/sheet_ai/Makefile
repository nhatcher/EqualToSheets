docker:
	@docker-compose down --rmi all
	@GIT_COMMIT=$(shell git rev-parse HEAD) docker-compose up -d --build

container-push:
	cd server && heroku container:push web --app sheet-ai --context-path .. --arg GIT_COMMIT=$(git rev-parse HEAD)
	@echo "\n\nif you are sure: heroku container:release web --app sheet-ai"

.PHONY: docker container-push

