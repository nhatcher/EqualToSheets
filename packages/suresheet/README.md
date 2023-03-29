# SureSheet

Built with:

- [EqualTo Sheets](https://sheets.equalto.com)
- [Next.js](https://nextjs.org/)
- [prisma](https://www.prisma.io/)

## Getting Started (internal use only)

### Running development server

1. First, deploy serverless locally. `.env` file exposes port 5000 by default.
2. Look at `.env` file and supply necessary keys.
3. Run `docker compose up -d` to start up the MongoDB replica set (port 27018).
4. `npm install`
5. `npx prisma generate`
6. `npx prisma db push`
7. Start Next.js development server, in `packages/suresheet` run `npm run dev`.
8. Visit http://localhost:3000/
