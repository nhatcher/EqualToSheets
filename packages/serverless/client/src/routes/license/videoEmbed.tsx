import styled from 'styled-components/macro';

export const VideoEmbed = () => {
  return (
    <FrameContainer>
      <div />
      <iframe
        src="https://www.youtube.com/embed/HobLlnD7Im0"
        title="YouTube video player"
        frameBorder="0"
        allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
        allowFullScreen
      ></iframe>
    </FrameContainer>
  );
};

const FrameContainer = styled.div`
  margin-top: 30px;
  position: relative;
  div:first-child {
    padding-right: 100%;
    padding-bottom: 56.25%; // 315/560
  }
  iframe {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
  }
`;
