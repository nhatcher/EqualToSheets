export type ConversationEntry =
  | {
      source: 'user';
      text: string;
    }
  | {
      source: 'server';
      data: string[][];
      text?: string | null;
    };
