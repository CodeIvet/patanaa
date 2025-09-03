import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext,
} from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { getTelemetryClient } from "../appInsights";
import { Logger } from "../logger";

export async function getBoardMeetings(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const telemetryClient = getTelemetryClient();
  context.log("Processing getBoardMeetings function.");
  const logger = new Logger(context);

  // Prepare access token.
  // const accessToken = request.headers
  //   .get("Authorization")
  //   ?.replace("Bearer ", "")
  //   .trim();
  // if (!accessToken) {
  //   return {
  //     status: 400,
  //     body: JSON.stringify({
  //       error: "No access token was found in request header.",
  //     }),
  //   };
  // }

  try {
    // Define your SQL query and parameters
    const query = "SELECT * FROM BoardMeetings ORDER BY StartTime ASC";

    // Execute the query using the helper class
    const boardMeetings = await DatabaseHelper.executeQuery(query);

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: boardMeetings,
    };
  } catch (err) {
    const errorObj = err instanceof Error ? err : new Error(String(err));
    logger.logError(errorObj, { additionalInfo: "dummy" });
    context.error("Error:", err);
    return {
      status: 500,
      body: "Internal server error",
    };
  }
}

app.http("getBoardMeetings", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: getBoardMeetings,
});