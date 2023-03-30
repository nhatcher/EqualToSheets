import { Html, Head, Main, NextScript } from 'next/document';
import Script from 'next/script';
import getConfig from 'next/config';

const { publicRuntimeConfig } = getConfig();

export default function Document() {
  return (
    <Html>
      <Head>
        <link rel="icon" href={`${publicRuntimeConfig.basePath}/favicon.ico`} />
      </Head>
      <body>
        <Main />
        <NextScript />
        <Script
          src={`${publicRuntimeConfig.basePath}/api/sheets-proxy/static/v1/equalto.js`}
          strategy="beforeInteractive"
        />
      </body>
    </Html>
  );
}
