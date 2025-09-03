import * as appInsights from "applicationinsights";
import config from "./config";

let telemetryClient: appInsights.TelemetryClient;

export const getTelemetryClient = (): appInsights.TelemetryClient => {
  if (!telemetryClient) {
    const aiConnectionString = config.applicationInsightsConnectionString;
    appInsights
      .setup(aiConnectionString)
      .setAutoCollectConsole(false, false)
      .start();
    telemetryClient = appInsights.defaultClient;
  }

  return telemetryClient;
};