export function getSheetsApiHost(): string {
  const serverlessHost = process.env.SERVERLESS_HOST;
  if (!serverlessHost) {
    throw new Error('Serverless host is not set.');
  }
  return serverlessHost;
}

export function getSheetsApiLicenseId(): string {
  const licenseId = process.env.SERVERLESS_LICENSE_ID;
  if (!licenseId) {
    throw new Error('Serverless license ID is not set.');
  }
  return licenseId;
}

export async function createWorkbook(options?: { json: string }): Promise<{
  id: string;
}> {
  let body;
  if (options) {
    body = JSON.stringify({
      version: '1',
      workbook_json: options.json,
    });
  }

  let response = await fetch(`${getSheetsApiHost()}api/v1/workbooks`, {
    method: 'POST',
    body,
    headers: new Headers({
      Authorization: `Bearer ${getSheetsApiLicenseId()}`,
      'Content-Type': 'application/json',
    }),
    credentials: 'include',
  });

  if (!response.ok) {
    throw new Error('Request failed. Code=' + response.status + ' Text=' + response.statusText);
  }

  // TODO: Zod on response json?
  return (await response.json()) as { id: string };
}
