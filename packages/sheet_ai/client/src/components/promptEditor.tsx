import { FunctionComponent, useRef } from 'react';
import styled from 'styled-components/macro';
import { Send } from 'lucide-react';

const MIN_TEXT_AREA_HEIGHT = 25;
const MAX_TEXT_AREA_HEIGHT = 100;

export const PromptEditor: FunctionComponent<{}> = (properties) => {
  const textAreaRef = useRef<HTMLTextAreaElement | null>(null);

  return (
    <Form
      onSubmit={(event) => {
        if (!textAreaRef.current) {
          return;
        }
        const textArea = textAreaRef.current;
        console.log('Submitted prompt: ', textArea.value);
        event.preventDefault();
      }}
    >
      <TextArea
        name="prompt"
        ref={textAreaRef}
        onKeyUp={() => {
          if (textAreaRef.current) {
            const textArea = textAreaRef.current;
            textArea.style.height = 'auto';
            const height = Math.max(
              MIN_TEXT_AREA_HEIGHT,
              Math.min(MAX_TEXT_AREA_HEIGHT, textArea.scrollHeight),
            );
            textArea.style.height = `${height}px`;
          }
        }}
      />
      <Button type="submit" aria-label="Send prompt">
        <Send size={18} />
      </Button>
    </Form>
  );
};

const Form = styled.form`
  position: relative;
  display: grid;
  grid-template-columns: 1fr 30px;
  gap: 5px;
`;

const TextArea = styled.textarea`
  resize: none;
  max-height: ${MAX_TEXT_AREA_HEIGHT}px;
  padding: 5px;
  border: 1px solid #aeaeae;
  border-radius: 5px;
`;

const Button = styled.button`
  display: inline-flex;
  border: none;
  background: none;
  width: 24px;
  height: 24px;
  line-height: 18px;
  align-items: center;
  justify-content: center;
  &:hover {
    background: #f2f2f2;
  }

  align-self: end;
`;
