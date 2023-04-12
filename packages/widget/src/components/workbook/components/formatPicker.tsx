import styled from 'styled-components';
import React, { FunctionComponent } from 'react';
import Dialog from 'src/components/uiKit/dialog';
import Button from 'src/components/uiKit/button';
import TextField from 'src/components/uiKit/textField';

type FormatPickerProps = {
  className?: string;
  open: boolean;
  onClose: () => void;
  onExited?: () => void;
  numFmt: string;
  onChange: (numberFmt: string) => void;
};

const FormatPicker: FunctionComponent<FormatPickerProps> = (properties) => {
  const onSubmit: React.FormEventHandler<HTMLFormElement> = (event) => {
    event.preventDefault();
    const formData = new FormData(event.target as HTMLFormElement);
    properties.onChange((formData.get('formatCodeInput') as string) ?? '');
    properties.onClose();
  };
  return (
    <Dialog
      $width="380px"
      title="Custom number format"
      open={properties.open}
      onClose={properties.onClose}
      onExited={properties.onExited}
      closeTitle="Cancel"
    >
      <Body onSubmit={onSubmit}>
        <FormatInput
          defaultValue={properties.numFmt}
          label="Number format"
          name="formatCodeInput"
        />
        <Footer>
          <Button variant="primary" type="submit" width="100%">
            {'Save'}
          </Button>
        </Footer>
      </Body>
    </Dialog>
  );
};

const FormatInput = styled(TextField)`
  border-width: 0px;
`;

const Body = styled.form`
  display: flex;
  flex-direction: column;
`;

const Footer = styled.div`
  display: flex;
  flex-direction: row;
  justify-content: flex-end;
  padding-top: 20px;
`;

export default FormatPicker;
