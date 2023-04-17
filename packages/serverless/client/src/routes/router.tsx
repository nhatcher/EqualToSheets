import { createHashRouter, Navigate } from 'react-router-dom';
import { LicenseActivationPage } from './license/licenseActivation';
import { LicenseLayout } from './license/licenseLayout';
import { LicenseRequestPage } from './license/licenseRequest';
import { LicenseSentPage } from './license/licenseSent';

export default createHashRouter([
  {
    path: '/license',
    element: <LicenseLayout />,
    children: [
      {
        path: 'request',
        element: <Navigate to="/license/access" replace />,
      },
      {
        path: 'access',
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
    path: 'access',
    element: <Navigate to="/license/access" replace />,
  },
  {
    path: '*',
    element: <Navigate to="/license/access" />,
  },
]);
