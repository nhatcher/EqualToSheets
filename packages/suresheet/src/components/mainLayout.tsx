import styles from './mainLayout.module.css';
import Image from 'next/image';
import { ArrowUpRight } from 'lucide-react';
import { Stack } from '@mui/system';
import { PropsWithChildren, useMemo, useState } from 'react';
import { ToolbarContext, ToolbarContextValue } from './toolbar';
import getConfig from 'next/config';
import Link from 'next/link';

const { publicRuntimeConfig } = getConfig();

export default function MainLayout(properties: PropsWithChildren<{}>) {
  const { children } = properties;

  // State is used instead of ref on purpose. This way context refreshes appropriately.
  const [toolbarNode, setToolbarNode] = useState<HTMLDivElement | null>(null);
  const toolbarContext = useMemo((): ToolbarContextValue => ({ toolbarNode }), [toolbarNode]);

  return (
    <div className={styles.layout}>
      <div className={styles.infoBar}>
        {
          'EqualTo SureSheet is an open source tech demo, showing how you can easily build software '
        }
        {'using our product EqualTo Sheets - "Spreadsheets as a service" for developers. '}
        <a href="https://www.github.com/EqualTo-Software/SureSheet" target="_blank">
          See project on GitHub <ArrowUpRight size={9} />
        </a>
      </div>
      <div className={styles.menuBar}>
        <Stack direction="row" justifyContent="space-between">
          <Stack direction="row" alignItems="center" spacing={2}>
            <a href="https://www.equalto.com/" target="_blank">
              <Image
                priority
                src={`${publicRuntimeConfig.basePath}/images/equalto.svg`}
                width={32}
                height={32}
                alt="EqualTo Logo"
              />
            </a>
            <Stack direction="row" alignItems="center" spacing={0}>
              <code className={styles.siteName}>SureSheet</code>
              <span className={styles.betaTag}>Beta</span>
            </Stack>
          </Stack>
          <div style={{ flex: 1 }} ref={(ref) => setToolbarNode(ref)} />
        </Stack>
      </div>
      <div className={styles.content}>
        <ToolbarContext.Provider value={toolbarContext}>{children}</ToolbarContext.Provider>
      </div>
    </div>
  );
}
