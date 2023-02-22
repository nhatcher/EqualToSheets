import * as Workbook from '@equalto-software/spreadsheet';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { saveAs } from 'file-saver';
import { Download } from 'lucide-react';
import { useCallback, useRef, useState } from 'react';
import styled from 'styled-components/macro';
import { CircularSpinner } from './components/circularSpinner';
import {
  ErrorMessageBubble,
  SystemMessageBubble,
  SystemMessageThread,
  UserMessageBubble,
} from './components/messageBubbles';
import { PromptEditor } from './components/promptEditor';
import { ConversationEntry } from './types';
import { useSessionCookie } from './useSessionCookie';

const queryClient = new QueryClient();

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <ChatRoot />
    </QueryClientProvider>
  );
}

function ChatRoot() {
  const sessionCookie = useSessionCookie();

  if (sessionCookie === 'loading') {
    return null;
  } else if (sessionCookie === 'set') {
    return <AppContent />;
  } else if (sessionCookie === 'rate-limited') {
    // TODO: Nicer layout
    return (
      <div>{'Rate limit for your IP has been reached. Please try again later.'}</div>
    );
  } else if (sessionCookie === 'error') {
    // TODO: Nicer layout
    return <div>{'Could not connect to the chat service. Please try again later.'}</div>;
  } else {
    const impossibleResult: never = sessionCookie;
    throw new Error(`Session cookie has unknown state: ${impossibleResult}`);
  }
}

function AppContent() {
  const [conversation, setConversation] = useState<ConversationEntry[]>([]);
  const [pendingMessage, setPendingMessage] = useState<string | null>(null);

  const discussionViewRef = useRef<HTMLDivElement | null>(null);
  const scrollToBottom = useCallback(
    () =>
      setTimeout(() => {
        if (discussionViewRef.current) {
          discussionViewRef.current.scrollTop = discussionViewRef.current.scrollHeight;
        }
      }, 0),
    [],
  );

  const handleMessageSend = useCallback(
    async (prompt: string) => {
      const newMessage = {
        source: 'user' as const,
        text: prompt,
        hasFailed: false,
      };

      const newMessages: ConversationEntry[] = [...conversation, newMessage];
      setPendingMessage(prompt);
      scrollToBottom();

      try {
        const formData = new FormData();
        formData.append(
          'prompt',
          JSON.stringify(
            newMessages
              .filter((entry) => entry.source === 'user' && !entry.hasFailed)
              .map((entry) => entry.text),
          ),
        );

        const fetchResponse = await fetch('/converse', {
          method: 'POST',
          headers: {
            Accept: 'application/json',
          },
          body: formData,
        });

        if (!fetchResponse.ok) {
          let error: Extract<ConversationEntry, { source: 'server-error' }>;
          if (fetchResponse.status === 400) {
            error = {
              source: 'server-error',
              text: 'Your request could not be processed. Please try again.',
            };
          } else if (fetchResponse.status === 404) {
            error = {
              source: 'server-error',
              text: 'Could not generate workbook for given prompt. Please try again.',
            };
          } else if (fetchResponse.status === 429) {
            error = {
              source: 'server-error',
              text: 'Request quota has been exceeded. Thank you for trying us out!',
            };
          } else {
            // 401 included here on purpose.
            error = {
              source: 'server-error',
              text: 'An error has occurred. Please try again.',
            };
          }

          setPendingMessage(null);
          newMessage.hasFailed = true;
          setConversation((currentMessages) => [...currentMessages, newMessage, error]);
          return false;
        }

        const json = await fetchResponse.json();
        const data = json as { input: string | number | boolean }[][];

        newMessages.push({ source: 'server', data });

        setPendingMessage(null);
        setConversation(newMessages);
        scrollToBottom();

        return true;
      } catch {
        setPendingMessage(null);

        newMessage.hasFailed = true;
        setConversation((currentMessages) => [
          ...currentMessages,
          newMessage,
          {
            source: 'server-error',
            text: 'Could not communicate with the server, please try again later.',
          },
        ]);

        return false;
      }
    },
    [conversation, scrollToBottom],
  );

  return (
    <div>
      <TopContainer>
        <ChatWidget>
          <DiscussionView ref={discussionViewRef}>
            <Discussion>
              {conversation.map((entry, index) => (
                <ConversationMessageBlock key={index} entry={entry} />
              ))}
              {pendingMessage !== null && (
                <>
                  <UserMessageBubble $pending>{pendingMessage}</UserMessageBubble>
                  <SpinnerContainer>
                    <PositionedSpinner $color="#70D379" />
                  </SpinnerContainer>
                </>
              )}
            </Discussion>
          </DiscussionView>
          <DiscussionFooter>
            <PromptEditor onSubmit={handleMessageSend} />
          </DiscussionFooter>
        </ChatWidget>
      </TopContainer>
    </div>
  );
}
const WorkbookRoot = styled(Workbook.Root)`
  border: 1px solid #c6cae3;
  filter: drop-shadow(0px 2px 2px rgba(33, 36, 58, 0.15));
  border-radius: 10px;
`;
const Worksheet = styled(Workbook.Worksheet)`
  border-bottom-right-radius: 10px;
  border-bottom-left-radius: 10px;
`;

const ChatWidget = styled.div`
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
`;

const Discussion = styled.div`
  display: grid;
  gap: 10px;
`;

const DiscussionView = styled.div`
  flex: 1 1 auto;
  padding: 20px 20px 10px 20px;
  overflow: auto;
`;

const DiscussionFooter = styled.div`
  padding: 0 20px 20px 20px;
`;

const TopContainer = styled.div`
  margin: 20px 20px;
  height: 600px;
  border: 1px solid #d0d0d0;
`;

const SpinnerContainer = styled.div`
  position: relative;
  margin-top: 20px;
  height: 85px;
`;

const PositionedSpinner = styled(CircularSpinner)`
  position: absolute;
  top: 0;
  left: calc(50% - 40px);
`;

const ConversationMessageBlock = (properties: { entry: ConversationEntry }) => {
  const { entry } = properties;

  if (entry.source === 'user') {
    return <UserMessageBubble $hasFailed={entry.hasFailed}>{entry.text}</UserMessageBubble>;
  }

  if (entry.source === 'server') {
    return <ServerMessageBlock entry={entry} />;
  }

  if (entry.source === 'server-error') {
    return (
      <SystemMessageThread>
        <ErrorMessageBubble>{entry.text}</ErrorMessageBubble>
      </SystemMessageThread>
    );
  }

  return null;
};

const ServerMessageBlock = (properties: {
  entry: Extract<ConversationEntry, { source: 'server' }>;
}) => {
  const { entry } = properties;

  const COLUMNS = 10;
  const ROWS = 10;

  const [model, setModel] = useState<Workbook.Model | null>(null);
  const onModelCreate = useCallback(
    (model: Workbook.Model) => {
      for (let rowIndex = 0; rowIndex < entry.data.length; ++rowIndex) {
        let row = entry.data[rowIndex];
        for (let columnIndex = 0; columnIndex < row.length; ++columnIndex) {
          const cell = row[columnIndex];
          const input = cell['input'];
          model.setCellValue(0, rowIndex + 1, columnIndex + 1, removeFormatting(`${input}`));
        }
      }
      setModel(model);
    },
    [entry],
  );

  const downloadXlsx = useCallback(() => {
    if (!model) {
      return;
    }

    const buffer = model.saveToXlsx();
    const blob = new Blob([buffer], { type: 'application/vnd.ms-excel' });
    saveAs(blob, 'Spreadsheet.xlsx');
  }, [model]);

  return (
    <SystemMessageThread>
      <div style={{ height: ROWS * 24 + 74, overflow: 'visible' }}>
        <WorkbookRoot lastRow={COLUMNS + 10} lastColumn={ROWS + 10} onModelCreate={onModelCreate}>
          <Workbook.FormulaBar />
          <Worksheet />
        </WorkbookRoot>
      </div>
      <DownloadButton disabled={model === null} type="button" onClick={downloadXlsx}>
        <Download size={18} />
      </DownloadButton>
      {entry.text && <SystemMessageBubble>{entry.text}</SystemMessageBubble>}
    </SystemMessageThread>
  );
};

function removeFormatting(text: string): string {
  if (text.includes('$')) {
    return text.replace(/[^0-9.-]+/g, '');
  }
  if (text.includes('%')) {
    return text.replace(/[^0-9.-]+/g, '');
  }
  return text;
}

const DownloadButton = styled.button`
  justify-self: right;
  transition: background-color 0.2s ease-in-out;
  cursor: pointer;
  background: #f8f8f8;
  &:hover {
    background: #d0d2f8;
  }
  width: 30px;
  height: 30px;
  padding: 6px;
  border-radius: 15px;
  border: none;
  display: flex;
  align-items: center;
  justify-content: center;
`;

export default App;
