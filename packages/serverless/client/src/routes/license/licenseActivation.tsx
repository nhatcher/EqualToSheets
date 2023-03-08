import { CircularProgress } from '@mui/material';
import { Stack } from '@mui/system';
import { useEffect, useState } from 'react';
import { useParams } from 'react-router-dom';
import styled from 'styled-components/macro';
import { Brand } from '../../utils';
import { Box, ExternalLink } from './common';

type LicenseId = Brand<string, 'LicenseId'>;
type LicenseKey = Brand<string, 'LicenseKey'>;

function isLicenseIdParam(id: string | null | undefined): id is LicenseId {
  if (!id) {
    return false;
  }
  return true;
}

export const LicenseActivationPage = () => {
  const params = useParams();

  const licenseId = params.hasOwnProperty('licenseId') ? params.licenseId : null;
  if (!isLicenseIdParam(licenseId)) {
    return <div>{'Invalid activation URL. License ID is missing or invalid.'}</div>;
  }

  return <LicenseActivation licenseId={licenseId} />;
};

const getSnippet = (
  licenseKey: LicenseKey,
  workbookId: string,
) => `<script language="text/javascript" src="https://sheets.equalto.com/static/v1/equalto.js"/>
<script language="text/javascript">
  // Sets the Authorization: Bearer <key> http header
  // For production use, use a proxy instead of exposing this
  // in your client code
  equalto.set_license_key({
    licenseKey: "${licenseKey}"
  });
  // Insert spreadsheet widget into the DOM
  equalto.load({
    workbookId: "${workbookId}",
    elementId: "workbook-div"
  });
</script>
`;

const LicenseActivation = (parameters: { licenseId: LicenseId }) => {
  const { licenseId } = parameters;
  const [activationState, setActivationState] = useState<
    | { type: 'loading' }
    | {
        type: 'success';
        licenseKey: LicenseKey;
        workbookId: string;
      }
    | { type: 'error'; error: string }
  >({ type: 'loading' });

  useEffect(() => {
    activateLicense();
    async function activateLicense() {
      setActivationState({ type: 'loading' });

      let response;
      try {
        response = await fetch(`./activate-license-key/${encodeURIComponent(licenseId)}`);
      } catch {
        setActivationState({
          type: 'error',
          error: 'Could not connect to a server. Please try again.',
        });
        return;
      }

      if (!response.ok) {
        setActivationState({ type: 'error', error: 'Could not fetch details about your license.' });
      }

      try {
        const json = (await response.json()) as {
          license_key: LicenseKey;
          workbook_id: string;
        };
        setActivationState({
          type: 'success',
          licenseKey: json.license_key,
          workbookId: json.workbook_id,
        });
      } catch {
        setActivationState({
          type: 'error',
          error: 'Unexpected error happened. Please contact our support.',
        });
      }
    }
  }, [licenseId]);

  if (activationState.type === 'loading') {
    return (
      <LoadingBox>
        <CircularProgress />
      </LoadingBox>
    );
  }

  if (activationState.type === 'error') {
    return (
      <StyledBox $maxWidth={660}>
        <ErrorText>{activationState.error}</ErrorText>
      </StyledBox>
    );
  }

  return (
    <StyledBox $maxWidth={660}>
      <Stack gap={2} alignItems="center">
        <LicenseStatusText>
          Your license is <em>active</em>
        </LicenseStatusText>
        <DescriptionText>
          Copy and paste the code snippet into the appropriate location within your website or app.
        </DescriptionText>
      </Stack>
      <CodeArea>
        <code>{getSnippet(activationState.licenseKey, activationState.workbookId)}</code>
      </CodeArea>
      <FooterText>
        {'Need help? Contact us at '}
        <ExternalLink href="mailto:support@equalto.com">support@equalto.com</ExternalLink>.
      </FooterText>
    </StyledBox>
  );
};

const LoadingBox = styled.div`
  text-align: center;
`;

const StyledBox = styled(Box)`
  padding: 60px 35px 45px 35px;
`;

const LicenseStatusText = styled.p`
  font-weight: 600;
  font-size: 24px;
  line-height: 29px;
  text-align: center;
  color: #ffffff;
  em {
    font-style: normal;
    color: #70d379;
  }
  margin: 0;
`;

const DescriptionText = styled.p`
  font-weight: 400;
  font-size: 16px;
  line-height: 19px;
  text-align: center;
  color: #b4b7d1;
  max-width: 400px;
  margin: 0 0 30px 0;
`;

const CodeArea = styled.pre`
  position: relative;
  background: rgba(255, 255, 255, 0.05);
  border: 1px solid rgba(180, 183, 209, 0.2);
  border-radius: 5px;
  padding: 10px;
  overflow: auto;

  color: #8b8fad;
  font-weight: 300;
  font-size: 14px;
  line-height: 18px;
`;

const FooterText = styled.div`
  margin-top: 32px;
  font-weight: 400;
  font-size: 14px;
  line-height: 17px;
  text-align: center;
  color: #8b8fad;
`;

const ErrorText = styled.div`
  color: #e06276;
  font-weight: 400;
  font-size: 16px;
  line-height: 19px;
  text-align: center;
`;
