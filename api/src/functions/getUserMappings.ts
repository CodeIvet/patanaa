import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { getTelemetryClient } from "../appInsights";
import { Logger } from "../logger";
import { randomUUID } from 'crypto';

export async function getUserMappings(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const telemetryClient = getTelemetryClient();
  context.log("Processing getUserMappings function.");
  const logger = new Logger(context);

  // Prepare access token.
  const accessToken: string | undefined = request.headers
    .get("Authorization")
    ?.replace("Bearer ", "")
    .trim();
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  try {
    // Define your SQL query and parameters
    const query = "SELECT * FROM UserMappings ORDER BY DisplayName ASC";

    // Execute the query using the helper class
    const userMappings = await DatabaseHelper.executeQuery(query);

    // Transform the userMappings to match the UserMapping type
    const transformedUserMappings = userMappings.map((user: any) => ({
      upn: user.Upn || "",
      displayName: user.DisplayName || "",
      id: randomUUID(),
    }));

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: transformedUserMappings,
    };
  } catch (err) {
    logger.logError(
      err instanceof Error ? err : new Error(typeof err === "string" ? err : JSON.stringify(err)),
      { additionalInfo: "dummy" }
    );
    context.error("Error:", err);
    return {
      status: 500,
      body: "Internal server error",
    };
  }
}

app.http("getUserMappings", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: getUserMappings,
});