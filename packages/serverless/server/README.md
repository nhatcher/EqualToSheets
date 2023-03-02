# Serverless Sheets

## Deployment to Heroku

Push branch `deployment/serverless-sheets`.

Alternatively, you can follow the manual steps:

1. Log in to Heroku CLI and Heroku container repository.

```bash
heroku login
heroku container:login
```

2. Build and push the container (run in `packages/serverless/server`).

```bash
heroku container:push web --app serverless-sheets --context-path ../..
```

3. Release the new version.

```bash
heroku container:release web --app serverless-sheets
```
