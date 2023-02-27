import * as Workbook from '@equalto-software/spreadsheet';
import { createTheme, CssBaseline, ThemeProvider } from '@mui/material';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { saveAs } from 'file-saver';
import uniqueId from 'lodash/uniqueId';
import { AlertOctagon, Download } from 'lucide-react';
import { useCallback, useRef, useState } from 'react';
import styled from 'styled-components/macro';
import { ChatWelcomeView } from './components/chatWelcomeView';
import { CircularSpinner } from './components/circularSpinner';
import { EmailCollectionBar } from './components/emailCollectionBar';
import {
  ErrorMessageBubble,
  SystemMessageBubble,
  SystemMessageThread,
  UserMessageBubble,
} from './components/messageBubbles';
import { PromptEditor } from './components/promptEditor';
import { TopBar } from './components/topBar';
import { ConversationEntry, ServerResponse } from './types';
import { useSessionCookie } from './useSessionCookie';

const queryClient = new QueryClient();

export const fonts = {
  regular: 'Inter, sans-serif',
  mono: '"Fira Mono", serif',
};

const theme = createTheme({
  spacing: 5,
  palette: {
    primary: {
      light: '#E8F8E9',
      main: '#72ED79',
      dark: '#70D379',
      contrastText: '#FFF',
    },
    secondary: {
      light: '#6A6E88',
      main: '#292C42',
    },
    error: {
      main: '#E06276',
    },
    warning: {
      main: '#F5BB49',
    },
    info: {
      main: '#B4B7D1',
    },
    success: {
      main: '#70D379',
    },
    text: {
      primary: '#21243A',
      secondary: '#7D8EC2',
      disabled: '#ADAFBE',
    },
    background: {
      default: '#FFFFFF',
    },
  },
  typography: {
    fontFamily: fonts.regular,
    fontSize: 14,
    h1: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 24,
    },
    h2: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 20,
    },
    h3: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 16,
    },
    subtitle1: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 14,
    },
    subtitle2: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 14,
    },
    body1: {
      fontFamily: fonts.regular,
      fontSize: 14,
    },
    body2: {
      fontFamily: fonts.regular,
      fontSize: 14,
    },
    button: {
      fontFamily: fonts.regular,
      fontSize: 14,
      fontWeight: 500,
      textTransform: 'none',
    },
  },
});

function App() {
  return (
    <>
      <CssBaseline />
      <ThemeProvider theme={theme}>
        <QueryClientProvider client={queryClient}>
          <RootLayout>
            <TopBar />
            <EmailCollectionBar />
            <ChatRoot />
          </RootLayout>
        </QueryClientProvider>
      </ThemeProvider>
    </>
  );
}

const RootLayout = styled.div`
  display: grid;
  grid-template-columns: 100%;
  grid-template-rows: auto auto 1fr;
  height: 100%;
`;

function ChatRoot() {
  const sessionCookie = useSessionCookie();

  if (sessionCookie === 'loading') {
    return null;
  } else if (sessionCookie === 'set') {
    return <AppContent />;
  } else if (sessionCookie === 'rate-limited') {
    // TODO: Nicer layout
    return (
      <ErrorContainer>
        <ErrorBox>
          <AlertOctagon size={18} />
          <div>{'Rate limit for your IP has been reached. Please try again later.'}</div>
        </ErrorBox>
      </ErrorContainer>
    );
  } else if (sessionCookie === 'error') {
    return (
      <ErrorContainer>
        <ErrorBox>
          <AlertOctagon size={18} />
          <div>{'Could not connect to the chat service. Please try again later.'}</div>
        </ErrorBox>
      </ErrorContainer>
    );
  } else {
    const impossibleResult: never = sessionCookie;
    throw new Error(`Session cookie has unknown state: ${impossibleResult}`);
  }
}

const ErrorContainer = styled.div`
  padding: 0 20px;
`;

const ErrorBox = styled.div`
  padding: 20px;
  margin: 40px auto 0 auto;
  max-width: 600px;
  border-radius: 10px;
  text-align: center;
  background-color: #fceff1;
  border: 1px solid #fddbe0;
  color: ${({ theme }) => theme.palette.error.main};
`;

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
    async (newPrompt: string) => {
      const newMessage = {
        source: 'user' as const,
        text: newPrompt,
        hasFailed: false,
      };

      const newMessages: ConversationEntry[] = [...conversation, newMessage];
      setPendingMessage(newPrompt);
      scrollToBottom();

      try {
        const fetchResponse = await fetch('./converse', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json',
          },
          body: buildRequestData(newMessages, models.current),
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
        const data = json as ServerResponse;

        newMessages.push({ source: 'server', id: uniqueId(), data });

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

  const models = useRef<Record<string, Workbook.Model>>({});
  const onModelCreate = useCallback((id: string, model: Workbook.Model) => {
    models.current[id] = model;
  }, []);

  return (
    <ChatWidget>
      {conversation.length === 0 && pendingMessage === null ? (
        <WelcomeContainer>
          <ChatWelcomeView />
        </WelcomeContainer>
      ) : (
        <DiscussionView ref={discussionViewRef}>
          <Discussion>
            {conversation.map((entry, index) => (
              <ConversationMessageBlock key={index} entry={entry} onModelCreate={onModelCreate} />
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
      )}
      <DiscussionFooter>
        <PromptEditor onSubmit={handleMessageSend} />
        <LinksFooter>
          {/* eslint-disable-next-line react/jsx-no-target-blank */}
          <EqualToFooterLink href="https://www.equalto.com/" target="_blank">
            equalto.com
          </EqualToFooterLink>
        </LinksFooter>
      </DiscussionFooter>
    </ChatWidget>
  );
}

const EqualToFooterLink = styled.a`
  font-weight: 400;
  font-size: 9px;
  line-height: 11px;
  &:link,
  &:visited,
  &:hover,
  &:active {
    text-decoration: none;
    color: #8b8fad;
  }
`;

const WorkbookRoot = styled(Workbook.Root)`
  border: 1px solid #c6cae3;
  box-shadow: 0px 2px 2px rgba(33, 36, 58, 0.15);
  border-radius: 5px;
`;
const Worksheet = styled(Workbook.Worksheet)`
  border-bottom-right-radius: 5px;
  border-bottom-left-radius: 5px;
`;

function buildRequestData(
  conversation: ConversationEntry[],
  models: Record<string, Workbook.Model>,
): string {
  const rawData = conversation.map((entry) => {
    if (entry.source === 'user' && entry.hasFailed === false) {
      return entry.text;
    }

    if (entry.source === 'server') {
      if (models.hasOwnProperty(entry.id)) {
        const model = models[entry.id];
        const { maxColumn, maxRow } = model.getSheetDimensions(0);
        const workbookData: string[][] = [];
        for (let row = 1; row <= maxRow; ++row) {
          const rowData: (typeof workbookData)[number] = [];
          for (let column = 1; column <= maxColumn; ++column) {
            const cell = model.getUICell(0, row, column);
            const textRepresentation = cell.formula || cell.formattedValue || '';
            rowData.push(textRepresentation);
          }
          workbookData.push(rowData);
        }
        return workbookData.map((row) => '|' + row.join('|') + '|').join('\n');
      }

      // No data for workbook. Indicates a bug.
      return null;
    }

    return null;
  });

  console.log('Extracted data: ', rawData);
  return JSON.stringify({ prompt: rawData.filter((entry) => typeof entry === 'string') });
}

const WelcomeContainer = styled.div`
  display: flex;
  flex: 1;
  align-items: center;
  justify-content: center;
  min-height: 0;
`;

const ChatWidget = styled.div`
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
  width: 100%;
`;

const DiscussionView = styled.div`
  flex: 1 1 auto;
  overflow: auto;
  width: 100%;
  padding: 20px;
`;

const Discussion = styled.div`
  display: grid;
  gap: 10px;
  max-width: 600px;
  width: 100%;
  margin: 0 auto;
`;

const DiscussionFooter = styled.div`
  max-width: 600px;
  width: 100%;
  padding: 5px 20px 20px 20px;
  margin: 0 auto;
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

const LinksFooter = styled.div`
  display: flex;
  justify-content: center;
  font-size: 9px;
  margin-top: 10px;
`;

const ConversationMessageBlock = (properties: {
  entry: ConversationEntry;
  onModelCreate?: (id: string, model: Workbook.Model) => void;
}) => {
  const { entry, onModelCreate } = properties;

  if (entry.source === 'user') {
    return <UserMessageBubble $hasFailed={entry.hasFailed}>{entry.text}</UserMessageBubble>;
  }

  if (entry.source === 'server') {
    return <ServerMessageBlock entry={entry} onModelCreate={onModelCreate} />;
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
  onModelCreate?: (id: string, model: Workbook.Model) => void;
}) => {
  const { entry, onModelCreate: onModelCreateFromProperties } = properties;

  const sizeRef = useRef({
    columns: entry.data.length > 0 ? entry.data[0].length : 0,
    rows: entry.data.length,
  });
  const { columns, rows } = sizeRef.current;

  const [model, setModel] = useState<Workbook.Model | null>(null);
  const onModelCreate = useCallback(
    (model: Workbook.Model) => {
      for (let rowIndex = 0; rowIndex < entry.data.length; ++rowIndex) {
        let row = entry.data[rowIndex];
        for (let columnIndex = 0; columnIndex < row.length; ++columnIndex) {
          const cell = row[columnIndex];

          // TODO: Widget doesn't support bold rendering I think
          // HACK: Inversed order to apply style and then cause redraw
          if (cell.style?.bold) {
            const uiCell = model.getUICell(0, rowIndex + 1, columnIndex + 1);
            uiCell.style.font.bold = true;
          }

          const input = cell['input'];
          model.setCellValue(0, rowIndex + 1, columnIndex + 1, `${input}`);

          if (cell.style?.num_fmt) {
            model.setNumberFormat(
              0,
              {
                rowStart: rowIndex + 1,
                rowEnd: rowIndex + 1,
                columnStart: columnIndex + 1,
                columnEnd: columnIndex + 1,
              },
              cell.style.num_fmt,
            );
          }
        }
      }
      setModel(model);
      onModelCreateFromProperties?.(entry.id, model);
    },
    [entry, onModelCreateFromProperties],
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
      <div style={{ height: Math.min(rows, 10) * 24 + 74, overflow: 'visible' }}>
        <WorkbookRoot lastRow={rows + 10} lastColumn={columns + 10} onModelCreate={onModelCreate}>
          <Workbook.FormulaBar />
          <Worksheet />
        </WorkbookRoot>
      </div>
      <DownloadButton disabled={model === null} type="button" onClick={downloadXlsx}>
        <DownloadIconWrapper>
          <Download size={10} />
        </DownloadIconWrapper>
        {'Download (.xlsx)'}
      </DownloadButton>
      {entry.text && <SystemMessageBubble>{entry.text}</SystemMessageBubble>}
    </SystemMessageThread>
  );
};

const DownloadButton = styled.button`
  cursor: pointer;

  color: #8b8fad;
  background: #f1f2f8;

  transition: background-color 0.2s ease-in-out;
  &:hover {
    background: #d0d2f8;
  }

  padding: 5px 10px;
  border-radius: 15px;
  border: none;
  font-size: 9px;
  line-height: 12px;

  justify-self: start;
`;

const DownloadIconWrapper = styled.div`
  display: inline-block;
  margin-right: 5px;
  position: relative;
  top: 1px;
`;

export default App;
