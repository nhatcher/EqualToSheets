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

export const DualBox = styled.div`
  display: grid;
  grid-template-columns: 6fr 4fr;

  align-self: center;
  width: 100%;
  max-width: 960px;

  border: 1px solid #46495e;
  filter: drop-shadow(0px 2px 4px rgba(0, 0, 0, 0.1));
  border-radius: 16px;

  @media (max-width: 768px) {
    display: flex;
    flex-direction: column-reverse;
    border: none;
  }
`;

export const LeftSide = styled.div`
  padding: 50px;
  background: rgba(255, 255, 255, 0.03);
  box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1);
`;

export const RightSide = styled.div`
  padding: 20px 50px;
  filter: drop-shadow(0px 2px 4px rgba(0, 0, 0, 0.1));
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  align-items: center;
  min-height: 300px;
`;

export const HeadingText = styled.h1`
  font-style: normal;
  font-weight: 600;
  font-size: 28px;
  line-height: 34px;
  color: #ffffff;
  margin: 0;
  em {
    font-style: normal;
    color: #72ed79;
  }
`;

export const Subtitle = styled.p`
  margin: 10px 0 0 0;
  font-weight: 400;
  font-size: 16px;
  line-height: 19px;
  display: flex;
  color: #b4b7d1;
`;

export const VideoPlaceholder = styled.div`
  margin-top: 30px;
  background: #f2f2f2;
  border-radius: 10px;
  width: 100%;
  height: 250px;
  text-align: center;
  padding: 40px;
`;
