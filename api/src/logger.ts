import { InvocationContext } from "@azure/functions";
import { TelemetryClient } from "applicationinsights";
import { getTelemetryClient } from "./appInsights";

export class Logger {
  private telemetryClient: TelemetryClient;
  private context: InvocationContext;

  constructor(context: InvocationContext) {
    this.context = context;
    this.telemetryClient = getTelemetryClient();
  }

  logTrace(message: string, properties?: { [key: string]: any }) {
    // Write to Azure Log Stream
    this.context.log(`[TRACE] ${message}`);

    // Send to Application Insights
    this.telemetryClient.trackTrace({
      message,
      properties: {
        functionName: this.context.functionName,
        isCustom: true,
        ...properties,
      },
    });
  }

  logError(error: Error, properties?: { [key: string]: any }) {
    // Write error to Azure Log Stream
    this.context.error(`[ERROR] ${error.message}`);

    // Send error to Application Insights
    this.telemetryClient.trackException({
      exception: error,
      properties: {
        functionName: this.context.functionName,
        isCustom: true,
        ...properties,
      },
    });
  }

  logEvent(name: string, properties?: { [key: string]: any }) {
    // Write event to Azure Log Stream
    this.context.log(`[EVENT] ${name}`);

    // Send event to Application Insights
    this.telemetryClient.trackEvent({
      name,
      properties: {
        functionName: this.context.functionName,
        iscustom: true,
        ...properties,
      },
    });
  }
}