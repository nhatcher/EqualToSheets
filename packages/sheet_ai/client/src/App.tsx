import {
  UserMessageBubble,
  SystemMessageBubble,
  SystemMessageThread,
} from './components/messageBubbles';
import { Spreadsheet } from './components/spreadsheet';
import styled from 'styled-components/macro';
import { PromptEditor } from './components/promptEditor';
import { useCallback, useRef, useState } from 'react';
import { sendMessage } from './serverApi';
import { ConversationEntry } from './types';
import { CircularSpinner } from './components/circularSpinner';

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
    return (
      <SystemMessageThread>
        <Spreadsheet />
        <SimpleTable>
          {entry.data.map((row, index) => (
            <tr key={index}>
              {row.map((cell, index) => {
                // @ts-ignore
                const input = cell['input'];
                return <td key={index}>{input}</td>;
              })}
            </tr>
          ))}
        </SimpleTable>
        {entry.text && <SystemMessageBubble>{entry.text}</SystemMessageBubble>}
      </SystemMessageThread>
    );
  }

  return null;
};

const SimpleTable = styled.table`
  border-collapse: collapse;
  td {
    border: 1px solid #d0d0d0;
  }
`;

export default App;
