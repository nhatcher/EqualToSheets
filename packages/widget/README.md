# EqualTo Spreadsheet [WIP]

This repo contains spreadsheet widget.

## Testing

1. Run `npm run build-full`. It will: rebuild and reinstall TypeScript SDK, build a widget package, reinstall it in provided example.
2. Run `cd example && npm run start` to open dev server with the spreadsheet widget.

You can also build the example and test it using serve (in `./example`):

1. `npm run build`
2. `npx serve -s build`

## Prerequisites for package using widget

This package requires React v18.

Install `@equalto-software/spreadsheet` from the local filesystem by adding this line to `package.json`:
`"@equalto-software/spreadsheet": "file:<PATH_TO_EQUALTO_SPREADSHEET_PACKAGE>",`
