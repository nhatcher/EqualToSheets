// @ts-ignore
const scriptURL = new URL(import.meta.url);
const REQUIRED_SCRIPT_PATH = '/static/v1/equalto.js';

function getHostname(): string {
  // For local development
  if (scriptURL.protocol === 'file:') {
    return 'http://127.0.0.1:5000';
  }
  if (!scriptURL.href.endsWith(REQUIRED_SCRIPT_PATH)) {
    throw new Error(
      `Could not initialize EqualTo Sheets.\nequalto.js can only be hosted in one of two paths: 1) ${REQUIRED_SCRIPT_PATH}, 2) subfolder/${REQUIRED_SCRIPT_PATH}.\nNote that in either case, it must finish "${REQUIRED_SCRIPT_PATH}".`,
    );
  }
  return scriptURL.href.slice(0, scriptURL.href.length - REQUIRED_SCRIPT_PATH.length);
}

export function getGraphqlEndpoint(): string {
  return `${getHostname()}/graphql`;
}

export function getSyncEndpoint(): string {
  return `${getHostname()}/get-updated-workbook`;
}
