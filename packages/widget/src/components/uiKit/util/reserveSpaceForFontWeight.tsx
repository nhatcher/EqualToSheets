import React from 'react';
import styled from 'styled-components';

type SpaceReservationElementType = {
  maxFontWeight: string;
};

const Aligner = styled.span`
  display: inline-block;
  text-align: center;
  position: relative;
`;

const TextDisplay = styled.span`
  position: absolute;
  left: 0;
  right: 0;
  text-align: center;
`;

const SpaceReservationElement = styled.span<SpaceReservationElementType>`
  height: 0;
  font-weight: ${(properties): string => properties.maxFontWeight};
  visibility: hidden;
  display: inline-block;
`;

type TextReserveForBoldType = {
  children?: React.ReactNode;
  maxFontWeight?: string;
};

/**
 * Be careful with this component since it will draw children twice. Once with default settings, and once with
 * bold set for space reservation purposes.
 */
const ReserveSpaceForFontWeight: React.FC<TextReserveForBoldType> = ({
  children,
  maxFontWeight,
}: TextReserveForBoldType) => (
  <Aligner>
    <TextDisplay>{children}</TextDisplay>
    <SpaceReservationElement aria-hidden="true" maxFontWeight={maxFontWeight || 'bold'}>
      {children}
    </SpaceReservationElement>
  </Aligner>
);

export default ReserveSpaceForFontWeight;
