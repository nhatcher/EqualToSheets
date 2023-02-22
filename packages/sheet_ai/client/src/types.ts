export type ServerResponse = {
  input: string | number | boolean;
  style?: { bold?: boolean };
}[][];

export type ConversationEntry =
  | {
      source: 'user';
      text: string;
      hasFailed: boolean;
    }
  | {
      source: 'server';
      data: ServerResponse;
      text?: string | null;
      id: string;
    }
  | {
      source: 'server-error';
      text: string;
    };
