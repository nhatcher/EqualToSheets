export type ConversationEntry =
  | {
      source: 'user';
      text: string;
      hasFailed: boolean;
    }
  | {
      source: 'server';
      data: { input: string | number | boolean }[][];
      text?: string | null;
    }
  | {
      source: 'server-error';
      text: string;
    };
