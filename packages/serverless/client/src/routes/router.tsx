import { createHashRouter, Navigate } from 'react-router-dom';
import { LicenseActivationPage } from './license/licenseActivation';
import { LicenseLayout } from './license/licenseLayout';
import { LicenseRequestPage } from './license/licenseRequest';
import { LicenseSentPage } from './license/licenseSent';

const RedirectToRequest = () => {
  return <Navigate to="/license/request" />;
};

export default createHashRouter([
  {
    path: '/license/',
    element: <LicenseLayout />,
    children: [
      {
        path: 'request',
        element: <LicenseRequestPage />,
      },
      {
        path: 'sent',
        element: <LicenseSentPage />,
      },
      {
        path: 'activate/:licenseId',
        element: <LicenseActivationPage />,
      },
    ],
  },
  {
    path: '*',
    element: <RedirectToRequest />,
  },
]);
