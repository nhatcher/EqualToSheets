import { RouterProvider } from 'react-router-dom';
import router from './routes/router';
import { createTheme, CssBaseline, ThemeProvider } from '@mui/material';

export const fonts = {
  regular: 'Inter, sans-serif',
  mono: '"Fira Mono", serif',
};

const theme = createTheme({
  spacing: 5,
  palette: {
    primary: {
      light: '#E8F8E9',
      main: '#72ED79',
      dark: '#70D379',
      contrastText: '#FFF',
    },
    secondary: {
      light: '#6A6E88',
      main: '#292C42',
    },
    error: {
      main: '#E06276',
    },
    warning: {
      main: '#F5BB49',
    },
    info: {
      main: '#B4B7D1',
    },
    success: {
      main: '#70D379',
    },
    text: {
      primary: '#21243A',
      secondary: '#7D8EC2',
      disabled: '#ADAFBE',
    },
    background: {
      default: '#FFFFFF',
    },
  },
  typography: {
    fontFamily: fonts.regular,
    fontSize: 14,
    h1: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 24,
    },
    h2: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 20,
    },
    h3: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 16,
    },
    subtitle1: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 14,
    },
    subtitle2: {
      fontFamily: fonts.regular,
      fontWeight: 600,
      fontSize: 14,
    },
    body1: {
      fontFamily: fonts.regular,
      fontSize: 14,
    },
    body2: {
      fontFamily: fonts.regular,
      fontSize: 14,
    },
    button: {
      fontFamily: fonts.regular,
      fontSize: 14,
      fontWeight: 500,
      textTransform: 'none',
    },
  },
});

function App() {
  return (
    <>
      <CssBaseline />
      <ThemeProvider theme={theme}>
        <RouterProvider router={router} />
      </ThemeProvider>
    </>
  );
}

export default App;
