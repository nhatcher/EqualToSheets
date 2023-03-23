import { getSyncEndpoint } from './domain';

const ERROR_CHILL_OUT_TIME = 1000;

/**
 * Wraps `fetch` for long polls backend endpoint that returns updates.
 * No retry logic.
 *
 * Use `revision = 0` to download latest revision available at the moment.
 */
const fetchUpdatedWorkbook = async (
  workbookId: string,
  licenseKey: string | null,
  revision: number,
): Promise<Response> => {
  return await fetch(`${getSyncEndpoint()}/${workbookId}/${revision}`, {
    method: 'GET',
    headers: new Headers({
      Authorization: `Bearer ${licenseKey}`,
      'Content-Type': 'application/json',
    }),
    credentials: 'include',
  });
};

/**
 * Continuously tries to long poll backend endpoint to get revision update.
 *
 * Use `revision = 0` to download latest revision available at the moment.
 *
 * Promise will be fulfilled once response is received. Errors will be reported
 * via `onError`, but they will not reject the promise returned by the function.
 */
export const fetchUpdatedWorkbookWithRetries = async ({
  workbookId,
  licenseKey,
  revision,
  onError,
}: {
  workbookId: string;
  licenseKey: string | null;
  revision: number;
  onError: (error: string) => void;
}) => {
  while (true) {
    try {
      const response = await fetchUpdatedWorkbook(workbookId, licenseKey, revision);
      if (response.status === 502) {
        // Connection timeout, let's sync again
      } else if (!response.ok) {
        // Error, let's wait a second
        await new Promise((resolve) => setTimeout(resolve, ERROR_CHILL_OUT_TIME));
        onError(response.statusText);
      } else {
        const responseJson = await response.json();
        return responseJson;
      }
    } catch (e) {
      await new Promise((resolve) => setTimeout(resolve, ERROR_CHILL_OUT_TIME));
      onError(`${e}`);
    }
  }
};
