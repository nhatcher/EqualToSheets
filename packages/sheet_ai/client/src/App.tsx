import {
  UserMessageBubble,
  SystemMessageBubble,
  SystemMessageThread,
} from './components/messageBubbles';
import { Spreadsheet } from './components/spreadsheet';
import styled from 'styled-components/macro';
import { PromptEditor } from './components/promptEditor';

function App() {
  return (
    <div>
      <TopContainer>
        <ChatWidget>
          <DiscussionView>
            <Discussion>
              <UserMessageBubble>
                {'Please make me a greatest spreadsheet in the world.'}
              </UserMessageBubble>
              <SystemMessageThread>
                <Spreadsheet />
                <SystemMessageBubble ownComment={false}>{'Here it is 👆.'}</SystemMessageBubble>
              </SystemMessageThread>
              <UserMessageBubble>{'Thanks.'}</UserMessageBubble>
            </Discussion>
          </DiscussionView>
          <DiscussionFooter>
            <PromptEditor />
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
  max-width: 300px;
  margin: 20px auto;
  height: 400px;
  border: 1px solid #d0d0d0;
`;

export default App;
