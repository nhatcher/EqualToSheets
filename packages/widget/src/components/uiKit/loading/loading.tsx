import React, { FunctionComponent } from 'react';
import styled from 'styled-components';

const Container = styled.div<{ marginTop?: string; paddingTop?: string }>`
  margin-top: ${(properties): string => properties.marginTop || '0px'};
  padding-top: ${(properties): string => properties.paddingTop || '0px'};
  display: flex;
  height: 100%;
  justify-content: center;
`;

type LoadingProps = {
  loading?: boolean;
  marginTop?: string;
  paddingTop?: string;
  className?: string;
};

const Loading: FunctionComponent<LoadingProps> = ({
  marginTop,
  paddingTop,
  className,
}: LoadingProps) => (
  <Container marginTop={marginTop} paddingTop={paddingTop} className={className}>
    {'Loading...'}
  </Container>
);

export default Loading;
