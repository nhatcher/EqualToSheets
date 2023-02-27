export type ServerResponse = {
  input: string | number | boolean;
  style?: { bold?: boolean; num_fmt?: string };
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
