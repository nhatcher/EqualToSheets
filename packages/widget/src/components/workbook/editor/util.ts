interface CaretPosition {
  start: number;
  end: number;
}

export interface EditorSelection {
  start: number;
  end: number;
  text: string;
}

export const editorClass = 'equalto-cell-editor';

/**
 * Returns full text and the start and end of the selection.
 * start:  position of the first character of the selection
 * end: position of the last character of the selection
 * NOTE: start < end
 * Assumes that the structure is <div>[<span>text</span>]+</span>
 * like <div><span>=</span>A1<span>+</span><span>1</span></div>
 *
 * See: https://developer.mozilla.org/en-US/docs/Web/API/Selection
 * In particular:
 *
 * anchorNode:
 *   * Node in which the selection begins
 * anchorOffset:
 *   * If anchorNode is a text node: number of characters within anchorNode preceding the anchor
 *   * If anchorNode is an element: number of child nodes of the anchorNode preceding the anchor.
 * For us anchorNode is always a text node
 *
 * In the same vein. (focusNode, focusOffset) is the node and offset where the selection ends
 *
 */
export function getSelectedRangeInEditor(): EditorSelection | null {
  const sel = window.getSelection();
  const anchor = sel?.anchorNode;
  const focus = sel?.focusNode;
  let parent = anchor?.parentElement;
  if (parent && parent.tagName !== 'DIV') {
    // This is the normal situation. anchor is a TEXT_NODE, it's parent is a span
    // and its parent parent the div.
    // If this branch does not happen if the input is just text: <div>EqualTo</div>
    parent = parent?.parentElement;
  }

  // Silence the linter
  if (!anchor || !focus || !parent) {
    return null;
  }

  // If the parent is not the editor we try to get the contents of the editor. Otherwise we bail
  if (!parent.classList.contains(editorClass)) {
    const cellEditors = document.querySelectorAll(`.${editorClass}[contenteditable="true"]`);
    if (cellEditors.length === 1 && cellEditors[0].textContent) {
      const text = cellEditors[0].textContent;
      const { length } = text;
      return { start: length, end: length, text };
    }
    return null;
  }

  const text = parent.textContent || '';
  const rangeStart = document.createRange();
  rangeStart.setStart(parent, 0);
  rangeStart.setEnd(anchor, sel.anchorOffset || 0);
  const start = rangeStart.toString().length;

  const rangeEnd = document.createRange();
  rangeEnd.setStart(parent, 0);
  rangeEnd.setEnd(focus, sel.focusOffset || 0);
  const end = rangeEnd.toString().length;

  return {
    start: start < end ? start : end,
    end: end > start ? end : start,
    text,
  };
}

export default function setCaretPosition(editor: HTMLDivElement, caret: CaretPosition): void {
  let { start, end } = caret;
  let startSet = false;
  let endSet = false;
  let anchorNode = null;
  let anchorOffset = -1;
  let focusNode = null;
  let focusOffset = -1;
  const text = editor.textContent || '';
  if (start < 0 || end > text.length || end < start) {
    return;
  }
  Array.prototype.forEach.call(editor.children, (child: HTMLDivElement) => {
    const l = (child?.textContent || '').length;
    if (!startSet) {
      if (start <= l) {
        startSet = true;
        anchorNode = child.firstChild;
        anchorOffset = start;
      } else {
        start -= l;
      }
    }
    if (!endSet) {
      if (end <= l) {
        endSet = true;
        focusNode = child.firstChild;
        focusOffset = end;
      } else {
        end -= l;
      }
    }
  });
  if (anchorNode === null || focusNode === null) {
    return;
  }
  const selection = window.getSelection();
  selection?.setBaseAndExtent(anchorNode, anchorOffset, focusNode, focusOffset);
}
