import {
  UserMessageBubble,
  SystemMessageBubble,
  SystemMessageThread,
} from './components/messageBubbles';
import { Spreadsheet } from './components/spreadsheet';
import styled from 'styled-components/macro';
import { PromptEditor } from './components/promptEditor';
import { useRef, useState } from 'react';

function App() {
  const [myMessages, setMyMessages] = useState<string[]>(['Thanks']);
  const discussionViewRef = useRef<HTMLDivElement | null>(null);

  return (
    <div>
      <TopContainer>
        <ChatWidget>
          <DiscussionView ref={discussionViewRef}>
            <Discussion>
              <UserMessageBubble>
                {'Please make me a greatest spreadsheet in the world.'}
              </UserMessageBubble>
              <SystemMessageThread>
                <Spreadsheet />
                <SystemMessageBubble ownComment={false}>{'Here it is 👆.'}</SystemMessageBubble>
              </SystemMessageThread>
              {myMessages.map((message, index) => (
                <UserMessageBubble key={index}>{message}</UserMessageBubble>
              ))}
            </Discussion>
          </DiscussionView>
          <DiscussionFooter>
            <PromptEditor
              onSubmit={(prompt) => {
                return new Promise((resolve, reject) => {
                  setTimeout(() => {
                    setMyMessages((currentMessages) => [...currentMessages, prompt]);
                    setTimeout(() => {
                      if (discussionViewRef.current) {
                        discussionViewRef.current.scrollTop =
                          discussionViewRef.current.scrollHeight;
                      }
                    }, 0);
                    resolve(true);
                  }, 1000);
                });
              }}
            />
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

export default App;
