import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
// import { getTelemetryClient } from "../appInsights";
// import { Logger } from "../logger";
// import { Client } from "@microsoft/microsoft-graph-client";
// import { createGraphClient, getSafeString } from "../helper";
// import config from "../config";

export async function deleteBoardMeeting(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing deleteBoardMeeting function.");

  // Initialize response
  const res: HttpResponseInit = {
    status: 200,
  };
  const body: any = {};

  // Parse request body
  try {
    body.receivedHTTPRequestBody = (await request.text()) || "";
    var requestBody = body.receivedHTTPRequestBody
      ? JSON.parse(body.receivedHTTPRequestBody)
      : {};
  } catch (parseError: unknown) {
  const err = parseError as any;
  context.log(`Error parsing request body: ${err.message}`);
}

  try {
    const boardMeetingId = requestBody.meetingId;

    if (!boardMeetingId) {
      return { status: 400, body: "Missing meetingId in request body" };
    }

    const params = { Id: boardMeetingId };

    // 1. Delete agenda items for the board meeting
    const deleteAgendaItemsQuery = "DELETE FROM AgendaItems WHERE BoardMeeting = @Id";
    await DatabaseHelper.executeQuery(deleteAgendaItemsQuery, params);

    // 2. Delete the board meeting
    const deleteBoardMeetingQuery = "DELETE FROM BoardMeetings WHERE ID = @Id";
    const deletedMeeting = await DatabaseHelper.executeQuery(deleteBoardMeetingQuery, params);

    // Return success
    return {
      status: 200,
      body: JSON.stringify({ message: "Board meeting deleted successfully", deletedMeeting }),
    };
  } catch (err: any) {
    context.log("Error deleting board meeting:", err);
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
