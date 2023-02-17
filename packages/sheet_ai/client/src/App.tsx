import {
  UserMessageBubble,
  SystemMessageBubble,
  SystemMessageThread,
} from './components/messageBubbles';
import styled from 'styled-components/macro';
import { PromptEditor } from './components/promptEditor';
import { useCallback, useRef, useState } from 'react';
import { sendMessage } from './serverApi';
import { ConversationEntry } from './types';
import { CircularSpinner } from './components/circularSpinner';
import * as Workbook from '@equalto-software/spreadsheet';
import { saveAs } from 'file-saver';
import { Download } from 'lucide-react';

function App() {
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
      try {
        const newMessages: ConversationEntry[] = [
          ...conversation,
          {
            source: 'user',
            text: prompt,
          },
        ];

        setPendingMessage(prompt);
        scrollToBottom();

        const response = await sendMessage(newMessages);
        newMessages.push({
          source: 'server',
          data: response.data,
        });

        setPendingMessage(null);
        setConversation(newMessages);
        scrollToBottom();

        return true;
      } catch {
        // TODO: Error handling
        setPendingMessage(null);
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
                  <CenteredCircularSpinner $color="#70D379" />
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

const CenteredCircularSpinner = styled(CircularSpinner)`
  justify-self: center;
  margin-top: 20px;
`;

const ConversationMessageBlock = (properties: { entry: ConversationEntry }) => {
  const { entry } = properties;

  if (entry.source === 'user') {
    return <UserMessageBubble>{entry.text}</UserMessageBubble>;
  }

  if (entry.source === 'server') {
    return <ServerMessageBlock entry={entry} />;
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
          // @ts-ignore
          const input = cell['input'];
          model.setCellValue(0, rowIndex + 1, columnIndex + 1, removeFormatting(input));
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
