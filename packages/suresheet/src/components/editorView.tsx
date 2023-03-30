import { postShareWorkbook } from '@/lib/client/postShare';
import {
  Button,
  CircularProgress,
  Divider,
  Popover,
  Stack,
  TextField,
  Typography,
} from '@mui/material';
import { Copy, Download } from 'lucide-react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { ShareButton } from './shareButton';
import { ToolbarTool } from './toolbar';
import styles from './editorView.module.css';
import getConfig from 'next/config';

const { publicRuntimeConfig } = getConfig();

export type EditorViewProperties = {
  /**
   * Workbook ID of workbook to load from serverless. It's not mutated, and this ID shouldn't
   * be changed without component reload (use key if needed to force remount).
   */
  workbookId: string;
  onNew?: () => void;
};

export default function EditorView(properties: EditorViewProperties) {
  const { workbookId, onNew } = properties;

  const [latestJson, setLatestJson] = useState<string | null>(null);
  const [sharedWorkbookId, setSharedWorkbookId] = useState<string | null>(null);
  const [isShareOpen, setShareOpen] = useState(false);
  const shareButtonReference = useRef<HTMLButtonElement | null>(null);

  const onChange = useCallback((json: string) => {
    setSharedWorkbookId(null);
    setLatestJson(json);
  }, []);

  const share = () => {
    setShareOpen(true);
    if (latestJson !== null && sharedWorkbookId === null) {
      postShareWorkbook({ workbookJson: latestJson }).then(({ id }) => {
        setSharedWorkbookId(id);
      });
    }
  };

  return (
    <>
      <ToolbarTool>
        <Stack direction="row" justifyContent="space-between">
          <div>
            <Button onClick={onNew} sx={{ marginLeft: 2 }}>
              New
            </Button>
          </div>
          <ShareButton disabled={latestJson === null} onClick={share} ref={shareButtonReference} />
        </Stack>
        <Popover
          className={styles.popover}
          open={isShareOpen}
          onClose={() => setShareOpen(false)}
          anchorEl={shareButtonReference.current}
          anchorOrigin={{
            vertical: 'bottom',
            horizontal: 'right',
          }}
          transformOrigin={{
            vertical: 'top',
            horizontal: 'right',
          }}
          marginThreshold={5}
        >
          <Stack direction="column" spacing={1}>
            <Typography>Anyone with the link can access the spreadsheet.</Typography>
            <TextField
              name="share-link"
              disabled
              value={sharedWorkbookId ? getWorkbookViewLink(sharedWorkbookId) : ''}
              InputProps={{
                startAdornment:
                  sharedWorkbookId === null ? (
                    <Stack direction="column" justifyContent="center">
                      <CircularProgress size={18} />
                    </Stack>
                  ) : undefined,
                endAdornment:
                  sharedWorkbookId !== null ? (
                    <Button
                      onClick={() => {
                        navigator.clipboard.writeText(getWorkbookViewLink(sharedWorkbookId));
                      }}
                    >
                      <Copy />
                      Copy
                    </Button>
                  ) : undefined,
              }}
            />
            <Divider />
            <Button
              disabled={sharedWorkbookId === null}
              href={`${publicRuntimeConfig.basePath}/api/sheets-proxy/api/v1/workbooks/${sharedWorkbookId}/xlsx`}
              type="button"
              variant="contained"
              color="secondary"
              startIcon={<Download size={18} />}
            >
              Download .xlsx file
            </Button>
          </Stack>
        </Popover>
      </ToolbarTool>
      <Workbook workbookId={workbookId} onChange={onChange} />
    </>
  );
}

/**
 * Important note: Properties updates are ignored at the moment. Attach key to force remount.
 *
 * Workbook ID is not updated on serverless! Save requests are stopped.
 */
function Workbook(properties: { workbookId: string; onChange?: (json: string) => void }) {
  const workbookIdRef = useRef(properties.workbookId);
  const onChangeRef = useRef(properties.onChange);

  useEffect(() => {
    const workbookSlot = document.getElementById('workbook-slot');
    if (workbookSlot) {
      const { unmount } = EqualToSheets.load(workbookIdRef.current, workbookSlot, {
        syncChanges: false,
        onJSONChange: (_workbookId, workbookJson) => {
          onChangeRef.current?.(workbookJson);
        },
      });

      return () => {
        unmount();
      };
    }
  }, []);

  return <div id="workbook-slot" style={{ position: 'absolute', inset: 10 }} />;
}

function getWorkbookViewLink(workbookId: string) {
  return `${window.location.origin}${publicRuntimeConfig.basePath}/view/${encodeURIComponent(
    workbookId,
  )}`;
}
