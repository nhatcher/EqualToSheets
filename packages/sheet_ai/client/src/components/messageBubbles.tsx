import { PropsWithChildren } from 'react';
import styled from 'styled-components/macro';
import { ReactComponent as EqualToMessageLogo } from './equaltoMessageLogo.svg';

export const UserMessageBubble = styled.div`
  padding: 10px;
  white-space: pre-wrap;
  word-break: break-word;
  border-radius: 10px 0 10px 10px;
  background: #5879f0;
  color: #ffffff;
  max-width: 500px;
  justify-self: end;
`;

export const SystemMessageBubble = styled.div.attrs(() => ({}))<{ ownComment: boolean }>`
  padding: 10px;
  white-space: pre-wrap;
  word-break: break-word;
  border-radius: 0 10px 10px 10px;
  background: #dee0ee;
  color: #292c42;
  max-width: 500px;
`;

export const SystemMessageThread = (properties: PropsWithChildren<{}>) => {
  return (
    <SystemMessageThreadLayout>
      <EqualToMessageLogo />
      <SystemMessagesContainer>{properties.children}</SystemMessagesContainer>
    </SystemMessageThreadLayout>
  );
};

const SystemMessageThreadLayout = styled.div`
  display: grid;
  grid-template-columns: 36px 1fr;
  column-gap: 4px;
`;

const SystemMessagesContainer = styled.div`
  display: grid;
  gap: 10px;
`;
