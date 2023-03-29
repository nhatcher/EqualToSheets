import EditorView from '@/components/editorView';
import { Button, Stack, Typography } from '@mui/material';
import clsx from 'clsx';
import { File, Upload } from 'lucide-react';
import { GetServerSideProps } from 'next';
import Head from 'next/head';
import { useState } from 'react';
import { useDropzone } from 'react-dropzone';
import styles from './index.module.css';

export const getServerSideProps: GetServerSideProps = async ({ params }) => {
  return {
    props: {},
  };
};

export default function Home(properties: {}) {
  const [workbookId, setWorkbookId] = useState<string | null>(null);

  return (
    <>
      <Head>
        <title>EqualTo SureSheet</title>
      </Head>
      {workbookId === null ? (
        <NewWorkbookChoice setWorkbookId={setWorkbookId} />
      ) : (
        <EditorView workbookId={workbookId} />
      )}
    </>
  );
}

function NewWorkbookChoice(properties: { setWorkbookId: (workbookId: string) => void }) {
  const { setWorkbookId } = properties;

  const onDrop = (acceptedFiles: File[]) => {
    if (acceptedFiles.length === 0) {
      throw new Error('There are no accepted files.');
    }
    const firstFile = acceptedFiles[0];
    uploadXlsxWorkbook(firstFile).then(({ workbookId }) => {
      setWorkbookId(workbookId);
    });
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
    },
    onDrop,
  });

  const onNew = () => {
    createEmptyWorkbook().then(({ workbookId }) => setWorkbookId(workbookId));
  };

  return (
    <div className={styles.newWorkbookContainer}>
      <Stack direction="column" spacing={2}>
        <Button type="button" onClick={onNew} startIcon={<File />}>
          Start with a blank workbook
        </Button>
        <Stack direction="row" alignItems="center" spacing={1}>
          <div className={styles.divider} />
          <Typography>OR</Typography>
          <div className={styles.divider} />
        </Stack>
        <div {...getRootProps({ className: clsx(styles.dropzone, isDragActive && styles.active) })}>
          <input {...getInputProps()} />
          <Stack direction="column" spacing={1} alignItems="center">
            <Upload />
            <Typography>Upload a workbook (.xlsx)</Typography>
          </Stack>
        </div>
      </Stack>
    </div>
  );
}

async function createEmptyWorkbook() {
  const response = await fetch('/api/sheets-proxy/api/v1/workbooks', {
    method: 'POST',
  });
  const json = await response.json();
  return { workbookId: json['id'] as string };
}

async function uploadXlsxWorkbook(xlsxFile: File) {
  const body = new FormData();
  body.append('xlsx-file', xlsxFile);

  const response = await fetch('/api/sheets-proxy/create-workbook-from-xlsx', {
    method: 'POST',
    body,
  });

  // TODO: Endpoint should return JSON to avoid this parsing.
  const text = await response.text();
  const regex = /^Workbook Id: (.*)$/gm;
  const matches = regex.exec(text);
  if (matches && matches.length == 2) {
    const workbookId = matches[1];
    return { workbookId };
  }

  throw new Error('Could not parse the response.');
}
