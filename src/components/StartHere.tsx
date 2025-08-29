import { useContext } from "react";
import { TeamsFxContext } from "./Context";
import React from "react";
import { applyTheme, PeoplePicker } from "@microsoft/mgt-react";
import { Button } from "@fluentui/react-components";

import { Providers, ProviderState } from "@microsoft/mgt-react";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import {
  TeamsUserCredential,
  TeamsUserCredentialAuthConfig,
} from "@microsoft/teamsfx";
import { Dashboard } from "./custom/Dashboard";
import config from "./custom/lib/config";

const authConfig: TeamsUserCredentialAuthConfig = {
  clientId: config.clientId!,
  initiateLoginEndpoint: config.initiateLoginEndpoint!,
};


// const scopes = ["User.Read"];
const scopes = [
  "User.Read",
  "People.Read",
  "Contacts.Read",
  "Contacts.Read.Shared",
  "User.ReadBasic.All",
];
const credential = new TeamsUserCredential(authConfig);
const provider = new TeamsFxProvider(credential, scopes);
Providers.globalProvider = provider;

export default function StartHere() {
  const { themeString } = useContext(TeamsFxContext);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [consentNeeded, setConsentNeeded] = React.useState<boolean>(false);

  React.useEffect(() => {
    const init = async () => {
      try {
        await credential.getToken(scopes);
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } catch (error) {
        setConsentNeeded(true);
      }
    };

    init();
  }, []);

  const consent = async () => {
    setLoading(true);
    await credential.login(scopes);
    Providers.globalProvider.setState(ProviderState.SignedIn);
    setLoading(false);
    setConsentNeeded(false);
  };

  React.useEffect(() => {
    applyTheme(themeString === "default" ? "light" : "dark");
  }, [themeString]);

  return (
    <div>
      {consentNeeded && (
        <>
          <p>Bitte authorisieren Sie sich über den folgenden Button.</p>
          <Button appearance="primary" disabled={loading} onClick={consent}>
            Authorize
          </Button>
        </>
      )}

      {!consentNeeded && (
        <>
          <Dashboard></Dashboard>
        </>
      )}
    </div>
  );
}