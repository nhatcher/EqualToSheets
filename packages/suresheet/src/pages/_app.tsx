import '@fontsource/inter';
import '@fontsource/jetbrains-mono';
import '@/styles/globals.css';
import type { AppProps } from 'next/app';
import { ThemeProvider, createTheme } from '@mui/material/styles';
import MainLayout from '@/components/mainLayout';

const theme = createTheme({
  palette: {
    secondary: { main: '#21243A' },
  },
  typography: {
    fontFamily: "'Inter'",
    button: {
      textTransform: 'none',
    },
  },
});

export default function App({ Component, pageProps }: AppProps) {
  return (
    <ThemeProvider theme={theme}>
      <MainLayout>
        <Component {...pageProps} />
      </MainLayout>
    </ThemeProvider>
  );
}
