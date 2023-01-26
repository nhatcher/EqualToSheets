type Color = {
  RGB: string;
};

export type TabsInput = {
  index: number;
  sheet_id: number;
  state: string;
  // TODO [MVP]: Color should not be optional, but currently our mock plans have workbooks without colors
  // Model might create some default colors for tabs in the future
  color?: Color;
  name: string;
};
