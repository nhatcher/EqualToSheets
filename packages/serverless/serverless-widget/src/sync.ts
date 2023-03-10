import { getSyncEndpoint } from './domain';

let globalRevision = 0;
const ERROR_CHILL_OUT_TIME = 1000;

/**
 * Long polls the backend for the updates to the workbook
 */
export const sync = async (
  { workbookId, licenseKey, onResponse, onError }: {
    workbookId: string,
    licenseKey: string | null,
    onResponse: (response: any) => void,
    onError: (error: string) => void,
}) => {
  try {
    const response = await fetch(
      `${getSyncEndpoint()}/${workbookId}/${globalRevision}`,
      { 
        method: 'GET',
        headers: new Headers({
          Authorization: `Bearer ${licenseKey}`,
          'Content-Type': 'application/json',
        }),
        credentials: 'include',
      },
    );
    if (response.status === 502) {
      // Connection timeout, let's sync again
      await sync({ workbookId, licenseKey, onResponse, onError });
    } else if (!response.ok) {
      // Error, let's wait a second
      await new Promise(resolve => setTimeout(resolve, ERROR_CHILL_OUT_TIME));
      onError(response.statusText);
      await sync({ workbookId, licenseKey, onResponse, onError });
    } else {
      const responseJson = await response.json();
      globalRevision = responseJson['revision'];
      onResponse(responseJson);
      await sync({ workbookId, licenseKey, onResponse, onError });
    }
  } catch (e) {
    await new Promise(resolve => setTimeout(resolve, ERROR_CHILL_OUT_TIME));
    onError(`${e}`);
    await sync({ workbookId, licenseKey, onResponse, onError });
  }
};
