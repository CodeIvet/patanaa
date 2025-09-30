import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { getTelemetryClient } from "../appInsights";
import { Logger } from "../logger";
import { Client } from "@microsoft/microsoft-graph-client";
import { createGraphClient, getSafeString } from "../helper";
import config from "../config";

export async function deleteBoardMeeting(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const telemetryClient = getTelemetryClient();
  context.log("Processing deleteBoardMeeting function.");
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
  const accessTokenHeader = request.headers.get("Authorization");
  const accessToken: string | undefined = accessTokenHeader
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
    const boardMeetingId = requestBody.meetingId;
    const boardMeetingEventId = requestBody.eventId;
    const boardMeetingFileLocationId = requestBody.fileLocationId;
    const mailbox: string = config.eventMailbox;
    const params = { Id: boardMeetingId };
    const graphClient: Client = createGraphClient("App");

    // 0. Get agendaItems
    const getAgendaItemsQuery = `
      SELECT * FROM AgendaItems 
      WHERE BoardMeeting = @Id
      ORDER BY OrderIndex ASC;
      `;

    // Execute the query using the helper class
    const agendaItems = await DatabaseHelper.executeQuery(getAgendaItemsQuery, params);

    // 1. Cancel agendaItem events and boardmeeting event
    // Cancel each agenda item event if eventId is not null, undefined, or empty string
    for (const item of agendaItems) {
      if (item.EventId) {
        try {
          await graphClient
            .api(`/users/${mailbox}/calendar/events/${item.EventId}/cancel`)
            .post({});
        } catch (err) {
          console.log(err);
        }
      }
    }

    // Cancel boardMeeting event if eventId is not null, undefined, or empty string
    if (boardMeetingEventId) {
      try {
        await graphClient
          .api(`/users/${mailbox}/calendar/events/${boardMeetingEventId}/cancel`)
          .post({});
      } catch (err) {
        // ok, when ItemNotFound (deleted otherwise by user)
        if (typeof err === "object" && err !== null && "code" in err && (err as any).code !== "ErrorItemNotFound") {
          throw err;
        }
      }
    }

    // 2. Unassign agendaItems from boardMeeting and set EventId to null
    const unassignQuery =
      "UPDATE AgendaItems SET EventId = NULL, BoardMeeting = NULL WHERE BoardMeeting = @Id";
    const unassignResult = await DatabaseHelper.executeQuery(unassignQuery, params);

    // 3. Move agenda item file structure to topic storage
    for (const item of agendaItems) {
      let url = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${item.FileLocationId}`;
      await graphClient.api(url).patch({
        name: getSafeString(item.Title),
        parentReference: { id: config.sharePointUnassignedTopsFolderId },
        "@microsoft.graph.conflictBehavior": "rename",
      });
    }

    // 4. Delete file structure of boardMeeting
    if (boardMeetingFileLocationId) {
      try {
        await graphClient
          .api(
            `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${boardMeetingFileLocationId}`
          )
          .delete();
      } catch (err) {
        console.log(err);
      }
    }

    // 5. Delete boardMeeting
    // Define your SQL query and parameters
    const query = "DELETE FROM BoardMeetings WHERE ID = @Id";

    // Execute the query using the helper class
    const boardMeetings = await DatabaseHelper.executeQuery(query, params);

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(boardMeetings),
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

app.http("deleteBoardMeeting", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: deleteBoardMeeting,
});
