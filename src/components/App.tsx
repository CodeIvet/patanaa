// https://fluentsite.z22.web.core.windows.net/quick-start
import {
  FluentProvider,
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
  Spinner,
  tokens,
} from "@fluentui/react-components";
import { HashRouter as Router, Navigate, Route, Routes } from "react-router-dom";
import { useTeamsUserCredential } from "@microsoft/teamsfx-react";
import StartHere from "./StartHere";
import { TeamsFxContext } from "./Context";
import config from "./custom/lib/config";
import { createDarkTheme, createLightTheme } from "@fluentui/react-components";
import type { BrandVariants, Theme } from "@fluentui/react-components";
import About from "./About";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */

const as01: BrandVariants = {
  10: "#070109",
  20: "#240837",
  30: "#37016B",
  40: "#3E0097",
  50: "#3800CA",
  60: "#0B10FF",
  70: "#2B34FF",
  80: "#3E4BFF",
  90: "#4F5FFF",
  100: "#6071FF",
  110: "#7083FF",
  120: "#8094FF",
  130: "#90A5FF",
  140: "#A2B5FF",
  150: "#B4C6FF",
  160: "#C7D6FF",
};

export const lightTheme: Theme = {
  ...createLightTheme(as01),
};

export const darkTheme: Theme = {
  ...createDarkTheme(as01),
};

darkTheme.colorBrandForeground1 = as01[110];
darkTheme.colorBrandForeground2 = as01[120];

export default function App() {
  const { loading, theme, themeString, teamsUserCredential } = useTeamsUserCredential({
    initiateLoginEndpoint: config.initiateLoginEndpoint!,
    clientId: config.clientId!,
  });

  return (
    <TeamsFxContext.Provider value={{ theme, themeString, teamsUserCredential }}>
      <FluentProvider
        theme={
          themeString === "dark"
            ? darkTheme
            : themeString === "contrast"
            ? teamsHighContrastTheme
            : {
                ...lightTheme,
                colorNeutralBackground3: "#eeeeee",
              }
        }
        style={{ background: tokens.colorNeutralBackground3 }}
      >
        <Router>
          {loading ? (
            <Spinner style={{ margin: 100 }} />
          ) : (
            <Routes>
              <Route path="/starthere" element={<StartHere />} />
              <Route path="/about" element={<About />} />
              <Route path="*" element={<Navigate to={"/starthere"} />}></Route>
            </Routes>
          )}
        </Router>
      </FluentProvider>
    </TeamsFxContext.Provider>
  );
}