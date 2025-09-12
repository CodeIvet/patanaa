import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { getTelemetryClient } from "../appInsights";
import { Logger } from "../logger";
import { Client } from "@microsoft/microsoft-graph-client";
import { createGraphClient } from "../helper";
import config from "../config";

export async function deleteAgendaItem(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const telemetryClient = getTelemetryClient();
  context.log("Processing deleteAgendaItem function.");
  const logger = new Logger(context);

  // Initialize response.
  const res: HttpResponseInit = {
    status: 200,
  };
  const body = Object();

  // Put an echo into response body.
  body.receivedHTTPRequestBody = (await request.text()) || "";
  let requestBody: any;
  try {
    requestBody = body.receivedHTTPRequestBody
      ? JSON.parse(body.receivedHTTPRequestBody)
      : {};
  } catch (parseError) {
    if (parseError instanceof Error) {
      context.log(`Error parsing request body: ${parseError.message}`);
    } else {
      context.log("Error parsing request body: Unknown error");
    }
    throw new Error("Invalid JSON in request body");
  }

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
    const agendaItemId = requestBody.itemId;
    const agendaItemEventId = requestBody.eventId;
    const agendaItemFileLocationId = requestBody.fileLocationId;
    //const mailbox: string = config.eventMailbox; // implement when mailbox is implemented
    const params = { Id: agendaItemId };
    const graphClient: Client = createGraphClient("App");

    // 1. Cancel agendaItem event
    if (agendaItemEventId) {
      //try {
        //await graphClient
          //.api(`/users/${mailbox}/calendar/events/${agendaItemEventId}/cancel`)
          //.post({});
      //} catch (err) {
        //console.log(err);
      //}
    }

    // 2. Delete file structure of agenda item
    if (agendaItemFileLocationId) {
      try {
        await graphClient
          .api(
            `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${agendaItemFileLocationId}`
          )
          .delete();
      } catch (err) {
        console.log(err);
      }
    }

    // 3. Delete Agenda Item from database
    const query = "DELETE FROM AgendaItems WHERE ID = @Id";

    // Execute the query using the helper class
    const agendaItemPurgeResult = await DatabaseHelper.executeQuery(query, params);

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(agendaItemPurgeResult),
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

app.http("deleteAgendaItem", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: deleteAgendaItem,
});