import { CircularProgress } from '@mui/material';
import { Stack } from '@mui/system';
import Tabs from '@mui/material/Tabs';
import Tab from '@mui/material/Tab';
import MaterialBox from '@mui/material/Box';
import { ReactNode, useEffect, useState } from 'react';
import { useParams } from 'react-router-dom';
import Snackbar from '@mui/material/Snackbar';
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
) => `<div id="workbook-slot" style="height:100%;min-height:300px;"></div>
<script src="${window.location.origin}/static/v1/equalto.js">
</script>
<script>
  // WARNING: you should not expose your license key in client,
  //          instead you should proxy calls to EqualTo.
  EqualToSheets.setLicenseKey(
    "${licenseKey}"
  );
  // Insert spreadsheet widget into the DOM
  EqualToSheets.load(
    "${workbookId}",
    document.getElementById("workbook-slot")
  );
</script>
`;

const getCurlSnippet = (licenseKey: LicenseKey) =>
  `curl -F xlsx-file=@/path/to/file.xlsx -H "Authorization: Bearer ${licenseKey}" ${window.location.origin}/create-workbook-from-xlsx`;

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

  const [openAlert, setAlertOpen] = useState(false);
  const [selectedTab, setTab] = useState(0);

  const handleTabChange = (event: React.SyntheticEvent, newValue: number) => {
    setTab(newValue);
  };

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

  const copyToClipboard = (text: string) => () => {
    navigator.clipboard.writeText(text).then(() => {
      setAlertOpen(true);
    });
  };

  const { licenseKey, workbookId } = activationState;

  return (
    <StyledBox $maxWidth={720}>
      <Stack gap={2} alignItems="center" marginBottom="50px">
        <LicenseStatusText>
          Your license is <em>active</em>
        </LicenseStatusText>
      </Stack>
      <Stack gap={1} marginBottom="20px">
        <InputLabel>License key</InputLabel>
        <Input value={licenseKey} readOnly onClick={copyToClipboard(licenseKey)}></Input>
        <InputDescription>
          Protect your license key, it grants full access to the data you store in EqualTo Sheets.
        </InputDescription>
      </Stack>
      <Panel>
        <TabsBox>
          <Tabs value={selectedTab} onChange={handleTabChange} aria-label="Tabs">
            <Tab label="Get started now" {...a11yProps(0)} />
            <Tab label="Upload XLSX" {...a11yProps(1)} />
            <Tab label="GraphQL" {...a11yProps(2)} />
            <Tab label="Rest API" {...a11yProps(3)} />
            <Tab label="Simulation API" {...a11yProps(4)} />
          </Tabs>
        </TabsBox>
        <TabPanel value={selectedTab} index={0}>
          <GetStartedPanel
            licenseKey={licenseKey}
            workbookId={workbookId}
            copyToClipboard={copyToClipboard}
          />
        </TabPanel>
        <TabPanel value={selectedTab} index={1}>
          <UploadPanel
            licenseKey={licenseKey}
            workbookId={workbookId}
            copyToClipboard={copyToClipboard}
          />
        </TabPanel>
        <TabPanel value={selectedTab} index={2}>
          <GraphQLPanel licenseKey={licenseKey} workbookId={workbookId} />
        </TabPanel>
        <TabPanel value={selectedTab} index={3}>
          <RestAPIPanel licenseKey={licenseKey} workbookId={workbookId} />
        </TabPanel>
        <TabPanel value={selectedTab} index={4}>
          <SimulationAPIPanel licenseKey={licenseKey} workbookId={workbookId} />
        </TabPanel>
      </Panel>
      <Stack direction="row" alignItems="baseline" justifyContent="space-between" marginTop="24px">
        <FooterText>
          {'Need help? Take a look at the '}
          <DocumentationLink>documentation</DocumentationLink>.
        </FooterText>
        <BookmarkText>We recommend bookmarking this page</BookmarkText>
      </Stack>
      <Snackbar
        open={openAlert}
        autoHideDuration={3000}
        onClose={() => setAlertOpen(false)}
        message="Copied to the clipboard!"
      />
    </StyledBox>
  );
};

const DocumentationLink = ({ children }: { children: ReactNode }) => {
  return (
    <ExternalLink
      href="https://docs.google.com/document/d/19cDrllxVKid0HbSG8uXjDSs7iRDlitKM-9OLgkKiQV4/edit?usp=sharing"
      target="_blank"
    >
      {children}
    </ExternalLink>
  );
};

type Workbook = {
  id: string;
  name: string | null;
};
const WorkbookList = ({ licenseKey }: { licenseKey: string }) => {
  const [workbooks, setWorkbooks] = useState<Workbook[]>([]);
  useEffect(() => {
    getWorkbookList();
    async function getWorkbookList() {
      try {
        const response = await fetch('./api/v1/workbooks', {
          method: 'GET',
          headers: new Headers({
            Authorization: `Bearer ${licenseKey}`,
            'Content-Type': 'application/json',
          }),
          credentials: 'include',
        });
        if (response.ok) {
          const responseWorkbooks = await response.json();
          setWorkbooks(responseWorkbooks['workbooks']);
        }
      } catch (e) {
        console.error(e);
      }
    }
  }, [licenseKey]);

  return (
    <ul>
      {workbooks.map((workbook) => (
        <li key={workbook.id}>
          <ExternalGraphQLLink
            href={`./unsafe-just-for-beta/edit-workbook/${licenseKey}/${workbook.id}`}
            target="_blank"
          >
            {workbook.name ?? workbook.id}
          </ExternalGraphQLLink>
        </li>
      ))}
    </ul>
  );
};

const GetStartedPanel = ({
  licenseKey,
  workbookId,
  copyToClipboard,
}: {
  licenseKey: LicenseKey;
  workbookId: string;
  copyToClipboard: (text: string) => () => void;
}) => {
  return (
    <>
      <TabTextSection gap={2} direction="row" alignItems="center" $border>
        <div>
          We've created a sample workbook for you. Paste code snippet below inside the &lt;body&gt;
          tag and use it immediately.
        </div>
        <a
          className="link-button"
          href={`./unsafe-just-for-beta/edit-workbook/${licenseKey}/${workbookId}`}
          target="_blank"
          rel="noreferrer"
        >
          Preview workbook
        </a>
      </TabTextSection>
      <CodeArea onClick={copyToClipboard(getSnippet(licenseKey, workbookId))}>
        <code>{getSnippet(licenseKey, workbookId)}</code>
      </CodeArea>
    </>
  );
};
const UploadPanel = ({
  licenseKey,
  workbookId,
  copyToClipboard,
}: {
  licenseKey: LicenseKey;
  workbookId: string;
  copyToClipboard: (text: string) => () => void;
}) => {
  return (
    <>
      <TabTextSection direction="column" alignItems="flex-start" $border>
        <div>
          You can use this <em>curl</em> command to upload an XLSX file:
        </div>
        <CodeArea onClick={copyToClipboard(getCurlSnippet(licenseKey))}>
          <code style={{ whiteSpace: 'break-spaces' }}>{getCurlSnippet(licenseKey)}</code>
        </CodeArea>
      </TabTextSection>
      <TabTextSection>
        <div>
          List of all your workbooks:
          <WorkbookList licenseKey={licenseKey} />
        </div>
      </TabTextSection>
    </>
  );
};
const GraphQLPanel = ({ licenseKey, workbookId }: { licenseKey: string; workbookId: string }) => {
  return (
    <TabTextSection>
      You can use GraphQL to explore and manipulate your data:
      <ul>
        <li>
          <ExternalGraphQLLink
            target="_blank"
            href={`./graphql?license=${licenseKey}#query=%23%20WARNING%3A%20you%20should%20not%20share%20the%20above%20URL.%20It%20contains%0A%23%20%20%20%20%20%20%20%20%20%20your%20license%20key%2C%20which%20grants%20full%20access%20to%20all%0A%23%20%20%20%20%20%20%20%20%20%20your%20EqualTo%20Sheets%20data.%0A%0Aquery%20%7B%0A%20%20workbooks%20%7B%0A%20%20%20%20id%0A%20%20%20%20name%0A%20%20%7D%0A%7D%0A%0A%0A`}
          >
            List all workbooks
          </ExternalGraphQLLink>
        </li>
        <li>
          <ExternalGraphQLLink
            target="_blank"
            href={`./graphql?license=${licenseKey}#query=%23%20WARNING%3A%20you%20should%20not%20share%20the%20above%20URL.%20It%20contains%0A%23%20%20%20%20%20%20%20%20%20%20your%20license%20key%2C%20which%20grants%20full%20access%20to%20all%0A%23%20%20%20%20%20%20%20%20%20%20your%20EqualTo%20Sheets%20data.%0A%0Amutation%20%7B%0A%09%20%20createWorkbook%20%7B%0A%20%20%20%20%09workbook%20%7B%0A%20%20%20%20%20%20%09id%0A%20%20%20%20%20%20%20%20name%0A%20%20%20%20%7D%0A%20%20%7D%0A%7D%0A%0A%0A%0A`}
          >
            Create a new workbook
          </ExternalGraphQLLink>
        </li>
        <li>
          <ExternalGraphQLLink
            target="_blank"
            href={`./graphql?license=${licenseKey}#query=%23%20WARNING%3A%20you%20should%20not%20share%20the%20above%20URL.%20It%20contains%0A%23%20%20%20%20%20%20%20%20%20%20your%20license%20key%2C%20which%20grants%20full%20access%20to%20all%0A%23%20%20%20%20%20%20%20%20%20%20your%20EqualTo%20Sheets%20data.%0A%0Aquery%20%7B%0A%20%20workbook(workbookId%3A%20"${workbookId}")%20%7B%0A%20%20%20%20id%0A%20%20%20%20name%0A%20%20%20%20sheets%20%7B%0A%20%20%20%20%20%20name%0A%20%20%20%20%7D%0A%20%20%7D%0A%7D%0A%0A%0A`}
          >
            List all sheets in your sample workbook
          </ExternalGraphQLLink>
        </li>
        <li>
          <ExternalGraphQLLink
            target="_blank"
            href={`./graphql?license=${licenseKey}#query=%23%20WARNING%3A%20you%20should%20not%20share%20the%20above%20URL.%20It%20contains%0A%23%20%20%20%20%20%20%20%20%20%20your%20license%20key%2C%20which%20grants%20full%20access%20to%20all%0A%23%20%20%20%20%20%20%20%20%20%20your%20EqualTo%20Sheets%20data.%0A%0Aquery%20%7B%0A%20%20workbook(workbookId%3A%20"${workbookId}")%20%7B%0A%20%20%20%20id%0A%20%20%20%20name%0A%20%20%20%20sheet(sheetId%3A%201)%20%7B%0A%20%20%20%20%20%20cell(ref%3A%20"A1")%20%7B%0A%20%20%20%20%20%20%20%20value%20%7B%0A%20%20%20%20%20%20%20%20%20%20boolean%0A%20%20%20%20%20%20%20%20%20%20text%0A%20%20%20%20%20%20%20%20%20%20number%0A%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20formattedValue%0A%20%20%20%20%20%20%7D%0A%20%20%20%20%7D%0A%20%20%7D%0A%7D%0A%0A%0A`}
          >
            View the value in cell Sheet1!A1 of your sample workbook
          </ExternalGraphQLLink>
        </li>
        <li>
          <ExternalGraphQLLink
            target="_blank"
            href={`./graphql?license=${licenseKey}#query=%23%20WARNING%3A%20you%20should%20not%20share%20the%20above%20URL.%20It%20contains%0A%23%20%20%20%20%20%20%20%20%20%20your%20license%20key%2C%20which%20grants%20full%20access%20to%20all%0A%23%20%20%20%20%20%20%20%20%20%20your%20EqualTo%20Sheets%20data.%0A%0Amutation%20%7B%0A%20%20setCellInput(workbookId%3A"${workbookId}"%2C%20sheetId%3A%201%2C%20row%3A%201%2C%20col%3A%201%2C%20input%3A%20"300%24")%20%7B%0A%20%20%20%20__typename%0A%20%20%7D%20%0A%7D%0A%0A%0A`}
          >
            Change the value of cell Sheet1!A1 in your sample workbook
          </ExternalGraphQLLink>
        </li>
      </ul>
    </TabTextSection>
  );
};
const RestAPIPanel = ({ licenseKey, workbookId }: { licenseKey: string; workbookId: string }) => {
  const [response, setResponse] = useState('');

  const setInitialInvestment = async () => {
    const response = await fetch(`./api/v1/workbooks/${workbookId}/sheets/1/cells/6/3`, {
      method: 'PUT',
      headers: new Headers({
        Authorization: `Bearer ${licenseKey}`,
        'Content-Type': 'application/json',
      }),
      credentials: 'include',
      body: JSON.stringify({ value: 25000 }),
    });
    const responseJson = await response.json();
    setResponse(JSON.stringify(responseJson, null, 2));
  };

  const getInterestEarned = async () => {
    const response = await fetch(`./api/v1/workbooks/${workbookId}/sheets/1/cells/10/6`, {
      method: 'GET',
      headers: new Headers({
        Authorization: `Bearer ${licenseKey}`,
        'Content-Type': 'application/json',
      }),
      credentials: 'include',
    });
    const responseJson = await response.json();
    setResponse(JSON.stringify(responseJson, null, 2));
  };

  const getAllWorkbooks = async () => {
    const response = await fetch('./api/v1/workbooks', {
      method: 'GET',
      headers: new Headers({
        Authorization: `Bearer ${licenseKey}`,
        'Content-Type': 'application/json',
      }),
      credentials: 'include',
    });
    const responseJson = await response.json();
    setResponse(JSON.stringify(responseJson, null, 2));
  };

  const createNewWorkbook = async () => {
    const response = await fetch('./api/v1/workbooks', {
      method: 'POST',
      headers: new Headers({
        Authorization: `Bearer ${licenseKey}`,
        'Content-Type': 'application/json',
      }),
      credentials: 'include',
    });
    const responseJson = await response.json();
    setResponse(JSON.stringify(responseJson, null, 2));
  };

  return (
    <TabTextSection>
      Create, read and update your workbook using our REST API.
      <br /> <br />
      For the sample workbook in your account:
      <ul>
        <ClickableLi onClick={setInitialInvestment}>
          Set the initial investment (Sheet!C6) to $25,000
        </ClickableLi>
        <ClickableLi onClick={getInterestEarned}>Get the interest earned (Sheet!F10)</ClickableLi>
      </ul>
      Other useful REST APIs:
      <ul>
        <ClickableLi onClick={getAllWorkbooks}>Get a list of all my workbooks</ClickableLi>
        <ClickableLi onClick={createNewWorkbook}>Create new workbook</ClickableLi>
      </ul>
      <div style={{ fontSize: 'inherit' }}>
        Read <DocumentationLink>our documentation</DocumentationLink> for the full list of REST
        APIs.
      </div>
      <br />
      {response !== '' ? (
        <>
          Response:
          <CodeArea>
            <code>{response}</code>
          </CodeArea>
        </>
      ) : null}
    </TabTextSection>
  );
};

const SimulationAPIPanel = ({
  licenseKey,
  workbookId,
}: {
  licenseKey: string;
  workbookId: string;
}) => {
  const [response, setResponse] = useState('');
  const simulateInterest = async () => {
    const url = new URL(`${window.location.origin}/api/v1/workbooks/${workbookId}/simulate`);
    url.searchParams.append('inputs', JSON.stringify({ Sheet1: { C6: 50000 } }));
    url.searchParams.append('outputs', JSON.stringify({ Sheet1: ['F10'] }));
    const response = await fetch(url.href, {
      method: 'GET',
      headers: new Headers({
        Authorization: `Bearer ${licenseKey}`,
        'Content-Type': 'application/json',
      }),
    });
    const responseJson = await response.json();
    setResponse(JSON.stringify(responseJson, null, 2));
  };

  const simulateCapital = async () => {
    const url = new URL(`${window.location.origin}/api/v1/workbooks/${workbookId}/simulate`);
    url.searchParams.append('inputs', JSON.stringify({ Sheet1: { F4: 15 } }));
    url.searchParams.append('outputs', JSON.stringify({ Sheet1: ['F6'] }));
    const response = await fetch(url.href, {
      method: 'GET',
      headers: new Headers({
        Authorization: `Bearer ${licenseKey}`,
        'Content-Type': 'application/json',
      }),
    });
    const responseJson = await response.json();
    setResponse(JSON.stringify(responseJson, null, 2));
  };
  return (
    <TabTextSection>
      Run "what if" scenarios for your workbook.
      <br /> <br />
      For the sample workbook in your account:
      <ul>
        <ClickableLi onClick={simulateInterest}>
          The interest earned, if the initial investment is $50,000
        </ClickableLi>
        <ClickableLi onClick={simulateCapital}>
          Capital at the end of a term of 15 years
        </ClickableLi>
      </ul>
      <div style={{ fontSize: 'inherit' }}>
        Read <DocumentationLink>our documentation</DocumentationLink> for the full list of REST
        APIs.
      </div>
      <br />
      {response !== '' ? (
        <>
          Response:
          <CodeArea>
            <code>{response}</code>
          </CodeArea>
        </>
      ) : null}
    </TabTextSection>
  );
};

interface TabPanelProps {
  children?: React.ReactNode;
  index: number;
  value: number;
}

function a11yProps(index: number) {
  return {
    id: `simple-tab-${index}`,
    'aria-controls': `simple-tabpanel-${index}`,
  };
}

function TabPanel(props: TabPanelProps) {
  const { children, value, index, ...other } = props;

  return (
    <div
      role="tabpanel"
      hidden={value !== index}
      id={`simple-tabpanel-${index}`}
      aria-labelledby={`simple-tab-${index}`}
      {...other}
    >
      {value === index && <>{children}</>}
    </div>
  );
}

const ClickableLi = styled.li`
  cursor: pointer;
`;

const TabsBox = styled(MaterialBox)`
  background: rgb(73, 74, 99);
  border-bottom: 1px solid #545971;
  position: sticky;
  top: 0;
`;

const Input = styled.input`
  unset: all;
  cursor: pointer;
  display: flex;
  flex-direction: row;
  align-items: center;
  padding: 5px 4px 5px 10px;
  gap: 10px;
  color: #fff;
  background: #3e415a;
  border: 1px solid rgba(180, 183, 209, 0.2);
  border-radius: 5px;
  font-family: inherit;
  user-select: all;
`;
const InputLabel = styled.label`
  font-weight: 500;
  font-size: 13px;
  line-height: 16px;
  color: #ffffff;
`;
const InputDescription = styled.div`
  font-size: 13px;
  line-height: 16px;
  color: #8b8fad;
`;

const Panel = styled.div`
  background: #3e415a;
  border: 1px solid rgba(180, 183, 209, 0.2);
  border-radius: 5px;
  flex-grow: 1;
  overflow: scroll;
`;

const BookmarkText = styled.div`
  font-size: 11px;
  line-height: 13px;
  color: #f5bb49;
`;

const TabTextSection = styled(Stack)<{ $border?: boolean }>`
  color: #fff;
  font-style: normal;
  font-weight: 400;
  font-size: 14px;
  line-height: 17px;
  padding: 11px 13px;
  ${({ $border }) => ($border ? 'border-bottom: 1px solid #545971;' : '')}

  em {
    color: #8b8fad;
  }
  a.link-button {
    unset: all;
    text-decoration: none;
    display: flex;
    flex-direction: row;
    justify-content: center;
    align-items: center;
    gap: 5px;
    color: #fff;
    font-style: normal;
    font-weight: 400;
    font-size: 14px;
    line-height: 17px;
    padding: 10px 15px;
    border-radius: 5px;
    background: #565972;
    width: 180px;
    white-space: nowrap;
  }
  li {
    margin-bottom: 10px;
    cursor: pointer;
  }
`;

const ExternalGraphQLLink = styled(ExternalLink)`
  :link,
  :visited,
  :hover,
  :active {
    color: #fff;
  }
`;

const LoadingBox = styled.div`
  text-align: center;
`;

const StyledBox = styled(Box)`
  padding: 60px 35px 24px 35px;
  height: 750px;
  display: flex;
  flex-direction: column;
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

const CodeArea = styled.pre`
  cursor: pointer;
  padding: 10px;
  overflow: auto;

  color: #8b8fad;
  font-weight: 300;
  font-size: 14px;
  line-height: 18px;
  margin: 0px;
  flex-grow: 1;
`;

const FooterText = styled.div`
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
