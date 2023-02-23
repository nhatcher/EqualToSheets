import { FunctionComponent, useCallback, useRef, useState } from 'react';
import styled from 'styled-components/macro';
import { TextareaAutosize } from '@mui/material';

export const PromptEditor: FunctionComponent<{
  onSubmit: (prompt: string) => Promise<boolean>;
}> = (properties) => {
  const { onSubmit } = properties;
  const textAreaRef = useRef<HTMLTextAreaElement | null>(null);

  const [duringRequest, setDuringRequest] = useState(false);
  const handleSubmit = useCallback(() => {
    if (!textAreaRef.current) {
      return;
    }

    if (duringRequest) {
      return;
    }
    setDuringRequest(true);

    const textArea = textAreaRef.current;
    const prompt = textArea.value;

    onSubmit(prompt)
      .then((shouldClear) => {
        if (shouldClear) {
          textArea.value = '';
        }
        textArea.focus();
      })
      .finally(() => {
        setDuringRequest(false);
      });
  }, [duringRequest, onSubmit]);

  return (
    <Form
      onSubmit={(event) => {
        handleSubmit();
        event.preventDefault();
      }}
    >
      <Textarea
        name="prompt"
        minRows={1}
        maxRows={4}
        ref={textAreaRef}
        disabled={duringRequest}
        placeholder="Type in your prompt here"
        onKeyDown={(event) => {
          if (!event.shiftKey && event.key === 'Enter') {
            event.preventDefault();
            handleSubmit();
            return;
          }
        }}
      />
      <HRule />
      <Button type="submit" disabled={duringRequest}>
        Send
      </Button>
    </Form>
  );
};

const Form = styled.form`
  position: relative;
  display: grid;
  background: #f1f2f8;
  border-radius: 10px;
`;

const Textarea = styled(TextareaAutosize)`
  resize: none;
  padding: 10px 10px 10px 10px;
  border: none;
  background: transparent;
  border-radius: 10px 10px 0 0;
  z-index: 1;
`;

const HRule = styled.div`
  margin: 0 10px;
  height: 1px;
  border-bottom: 1px solid #e2e2e2;
`;

const Button = styled.button`
  display: inline-flex;
  border: none;
  background: none;
  align-items: center;
  justify-content: center;

  font-weight: 700;
  color: #587af0;

  &:disabled {
    color: #d0d0d0;
  }

  cursor: pointer;

  align-self: end;
  justify-self: end;
  padding: 10px;
`;
