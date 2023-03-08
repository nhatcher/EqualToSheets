import styled from 'styled-components/macro';

export const Box = styled.div<{ $maxWidth: number }>`
  padding: 0 20px 20px 20px;

  width: 100%;
  max-width: ${({ $maxWidth }) => $maxWidth}px;
  align-self: center;

  background: rgba(255, 255, 255, 0.05);
  border: 1px solid rgba(255, 255, 255, 0.1);
  box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1);
  border-radius: 16px;
`;

export const ExternalLink = styled.a`
  :link,
  :visited,
  :hover,
  :active {
    color: #5879f0;
    text-decoration: none;
  }
`;
