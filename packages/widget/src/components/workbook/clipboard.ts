import Model from './model';
import { Area, Cell } from './util';

const CLIPBOARD_ID_SESSION_STORAGE_KEY = 'equalTo_clipboardId';
const getNewClipboardId = () => new Date().toISOString();

export const onPaste =
  (model: Model | null, selectedSheet: number, selectedCell: Cell, selectedArea: Area) =>
  (event: React.ClipboardEvent) => {
    if (!model) {
      return;
    }
    const { items } = event.clipboardData;
    if (!items) {
      return;
    }
    const mimeTypes = ['application/json', 'text/plain', 'text/csv', 'text/html'];
    let mimeType;
    let value;
    const l = mimeTypes.length;
    for (let index = 0; index < l; index += 1) {
      mimeType = mimeTypes[index];
      value = event.clipboardData.getData(mimeType);
      if (value) {
        break;
      }
    }
    if (!mimeType || !value) {
      // No clipboard data to paste
      return;
    }
    if (mimeType === 'application/json') {
      // We are copying from within the application
      const targetArea = {
        sheet: selectedSheet,
        ...selectedArea,
      };

      try {
        const source = JSON.parse(value);
        const clipboardId = sessionStorage.getItem(CLIPBOARD_ID_SESSION_STORAGE_KEY);
        let sourceType = source.type;
        if (clipboardId !== source.clipboardId) {
          sourceType = 'copy';
        }
        model.paste(source.area, targetArea, source.sheetData, sourceType);
        if (sourceType === 'cut') {
          event.clipboardData.clearData();
        }
      } catch {
        // Trying to paste incorrect JSON
        // FIXME: We should validate the JSON and not try/catch
        // If JSON is invalid We should try to paste 'text/plain' content if it exist
      }
    } else if (mimeType === 'text/plain') {
      const targetArea = {
        sheet: selectedSheet,
        ...selectedArea,
      };
      model.pasteText(selectedSheet, targetArea, value);
    } else {
      // NOT IMPLEMENTED
    }
    event.preventDefault();
    event.stopPropagation();
  };

export const onCut =
  (model: Model | null, selectedSheet: number, selectedArea: Area) =>
  (event: React.ClipboardEvent<Element>) => {
    if (!model) {
      return;
    }
    const { tsv, area, sheetData } = model.copy({ sheet: selectedSheet, ...selectedArea });
    let clipboardId = sessionStorage.getItem(CLIPBOARD_ID_SESSION_STORAGE_KEY);
    if (!clipboardId) {
      clipboardId = getNewClipboardId();
      sessionStorage.setItem(CLIPBOARD_ID_SESSION_STORAGE_KEY, clipboardId);
    }
    event.clipboardData.setData('text/plain', tsv);
    event.clipboardData.setData(
      'application/json',
      JSON.stringify({ type: 'cut', area, sheetData, clipboardId }),
    );
    event.preventDefault();
    event.stopPropagation();
  };

// Other
/**
 * The clipboard allows us to attach different values to different mime types.
 * When copying we return two things: A TSV string (tab separated values).
 * And a an string representing the area we are copying.
 * We attach the tsv string to "text/plain" useful to paste to a different application
 * We attach the area to 'application/json' useful to paste within the application.
 *
 * FIXME: This second part is cheesy and will produce unexpected results:
 *      1. User copies an area (call it's contents area1)
 *      2. User modifies the copied area (call it's contents area2)
 *      3. User paste content to a different place within the application
 *
 * To the user surprise area2 will be pasted. The fix for this ius ro actually return a json with the actual content.
 */
// FIXME: Copy only works for areas [<top,left>, <bottom,right>]
export const onCopy =
  (model: Model | null, selectedSheet: number, selectedArea: Area) =>
  (event: React.ClipboardEvent<Element>) => {
    if (!model) {
      return;
    }
    const { tsv, area, sheetData } = model.copy({ sheet: selectedSheet, ...selectedArea });
    let clipboardId = sessionStorage.getItem(CLIPBOARD_ID_SESSION_STORAGE_KEY);
    if (!clipboardId) {
      clipboardId = getNewClipboardId();
      sessionStorage.setItem(CLIPBOARD_ID_SESSION_STORAGE_KEY, clipboardId);
    }
    event.clipboardData.setData('text/plain', tsv);
    event.clipboardData.setData(
      'application/json',
      JSON.stringify({ type: 'copy', area, sheetData, clipboardId }),
    );
    event.preventDefault();
    event.stopPropagation();
  };
