import React, { ReactNode, useCallback, useState } from 'react';
import styled from 'styled-components';
import { BodyText } from 'src/components/uiKit/typography';
import ReserveSpaceForFontWeight from '../reserveSpaceForFontWeight';

const Container = styled.div`
  margin: 20px 20px 0px;
`;

const ToggleableFontWeight: React.FC<{ children?: ReactNode }> = (properties) => {
  const [fontWeightType, setFontWeightType] = useState('primary');
  const toggleFontWeight = useCallback(() => {
    setFontWeightType((old) => (old === 'primary' ? 'secondary' : 'primary'));
  }, []);

  return (
    <button
      type="button"
      onClick={toggleFontWeight}
      style={{ fontWeight: fontWeightType === 'primary' ? 300 : 700 }}
    >
      {properties.children}
    </button>
  );
};

export const Index: React.FC = () => <StableInlineChain />;

export const StableInlineChain: React.FC = () => (
  <>
    <BodyText>
      {'Sometimes we want to change font weight on elements which size affect layout. ' +
        'That can create undesirable shifts, this component is meant to mitigate the problem.'}
    </BodyText>
    <div>
      <ToggleableFontWeight>
        {'I should have UNSTABLE width on toggling (click me).'}
      </ToggleableFontWeight>
      <ToggleableFontWeight>
        <ReserveSpaceForFontWeight maxFontWeight="700">
          {'I should have stable width on toggling (click me).'}
        </ReserveSpaceForFontWeight>
      </ToggleableFontWeight>
      {'I am just a normal text.'}
    </div>
  </>
);

export default {
  title: 'UI Kit/Typography/Util',
  component: ReserveSpaceForFontWeight,
  decorators: [
    (Story: React.FunctionComponent): React.ReactNode => (
      <Container>
        <Story />
      </Container>
    ),
  ],
};
