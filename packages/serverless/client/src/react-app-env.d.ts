/// <reference types="react-scripts" />

declare module '*.svg' {
  const content: string;
  export const ReactComponent: React.FC;
  export default content;
}
