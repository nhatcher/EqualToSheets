import styled from 'styled-components/macro';
import { ReactComponent as EqualToMessageLogo } from './equaltoMessageLogo.svg';
import { ReactComponent as OpenAILogo } from './openAiLogo.svg';

export function ChatWelcomeView() {
  return (
    <Container>
      <LogosBox>
        <EqualToMessageLogo width={30} height={30} />
        <div>+</div>
        <OpenAILogo width={30} height={30} />
      </LogosBox>
      <Heading>EqualTo Chat</Heading>
      <BodyText>Ask for a spreadsheet and we'll try to create it for you</BodyText>
      <ExampleText>
        "Create a detailed project plan for the restoration of a 19th century home."
      </ExampleText>
      <ExampleText>"Create a monthly budget for a student starting in university."</ExampleText>
      <BodyText>Use follow-up questions to refine your spreadsheet</BodyText>
      <ExampleText>"Change the currency to euro"</ExampleText>
      <ExampleText>"Add a row with the totals"</ExampleText>
    </Container>
  );
}

const Container = styled.div`
  max-width: 480px;
  padding: 60px 80px;
  background: linear-gradient(180deg, rgba(241, 242, 248, 0.5) 0%, rgba(241, 242, 248, 0) 100%);
  border-radius: 20px;
  text-align: center;
`;

const LogosBox = styled.div`
  display: flex;
  flex-direction: row;
  align-items: center;
  justify-content: center;
  gap: 10px;
  color: #8b8fad;
  > div {
    font-size: 16px;
  }
`;

const Heading = styled.div`
  margin: 20px 0;
  color: #21243a;
  font-weight: 600;
  font-size: 16px;
  line-height: 19px;
`;

const BodyText = styled.div`
  font-weight: 400;
  font-size: 13px;
  line-height: 16px;
  text-align: center;
  color: #5f6989;
  margin-bottom: 10px;
`;

const ExampleText = styled.div`
  font-style: italic;
  font-weight: 400;
  font-size: 13px;
  line-height: 16px;
  text-align: center;
  color: #8b8fad;
  & + & {
    margin-top: 5px;
  }
  & + ${BodyText} {
    margin-top: 20px;
  }
`;
