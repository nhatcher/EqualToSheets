import MainLayout from '@/components/mainLayout';
import '@/styles/globals.css';
import '@fontsource/inter/variable.css';
import '@fontsource/jetbrains-mono/variable.css';
import { createTheme, ThemeProvider } from '@mui/material/styles';
import type { AppProps } from 'next/app';
import Head from 'next/head';

const theme = createTheme({
  spacing: 5,
  palette: {
    secondary: { main: '#21243A' },
  },
  typography: {
    fontFamily: "'InterVariable'",
    button: {
      textTransform: 'none',
    },
  },
});

export default function App({ Component, pageProps }: AppProps) {
  return (
    <>
      <Head>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
      </Head>
      <ThemeProvider theme={theme}>
        <MainLayout>
          <Component {...pageProps} />
        </MainLayout>
      </ThemeProvider>
    </>
  );
}
