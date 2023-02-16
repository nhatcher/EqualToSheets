import { ConversationEntry } from './types';

const API_CONVERSE_URL = `/converse`;

export const sendMessage = async (
  history: ConversationEntry[],
): Promise<{
  data: string[][];
}> => {
  // TODO: Send back workbooks too.
  const formData = new FormData();
  formData.append(
    'prompt',
    JSON.stringify(history.filter((entry) => entry.source === 'user').map((entry) => entry.text)),
  );

  const response = await fetch(API_CONVERSE_URL, {
    method: 'POST',
    headers: {
      Accept: 'application/json',
    },
    body: formData,
  });

  const json = await response.json();
  return { data: json };
};
