import styled from 'styled-components';

const StylelessButton = styled.button`
  border: 0px;
  margin: 0px;
  padding: 0px;
  line-height: 0px;
  background-color: inherit;
  cursor: ${({ disabled }): string => (!disabled ? 'pointer' : 'default')};
  &:focus {
    outline: 0px;
  }
`;

export default StylelessButton;
