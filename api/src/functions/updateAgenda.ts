import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { AgendaItem, createGraphClient, ensureFileStructure } from "../helper";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../config";

export async function updateAgenda(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing updateAgenda function.");
  let statusMessage = "";

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
  const accessToken: string =
    (request.headers.get("Authorization")?.replace("Bearer ", "").trim()) || "";
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  try {
    // Extract `agendaItems` and `unassignedAgendaItems` from request body
    const agendaItems: AgendaItem[] = requestBody.agendaItems || [];
    const unassignedAgendaItems: AgendaItem[] = requestBody.unassignedAgendaItems || [];
    const boardMeetingId: number = requestBody.boardMeetingId;
    //const mailbox: string = config.eventMailbox;
    const graphClient: Client = createGraphClient("App");

    // A Issue database inserts / updates for current Agenda Items
    // Insert new AgendaItems
    const insertQuery = `
      INSERT INTO [AgendaItems] (
        [DurationInMinutes],
        [Title],
        [AdditionalParticipants],
        [FileLocationId],
        [ProtocolLocationId],
        [OrderIndex],
        [IsMisc],
        [NeedsDecision],
        [BoardMeeting],
        [Remarks]

      ) VALUES (
        @durationInMinutes,
        @title,
        @additionalParticipants,
        @fileLocationId,
        @protocolLocationId,
        @orderIndex,
        @isMisc,
        @needsDecision,
        @boardMeeting,
        @remarks
      )
    `;

    // Update assigned AgendaItems
    const updateQuery = `
        UPDATE [AgendaItems]
        SET
            [DurationInMinutes] = @durationInMinutes,
            [Title] = @title,
            [AdditionalParticipants] = @additionalParticipants,
            [FileLocationId] = @fileLocationId,
            [ProtocolLocationId] = @protocolLocationId,
            [OrderIndex] = @orderIndex,
            [IsMisc] = @isMisc,
            [NeedsDecision] = @needsDecision,
            [BoardMeeting] = @boardMeeting,
            [Remarks] = @remarks
        WHERE [ID] = @id;
    `;

    try {
      for (let i = 0; i < agendaItems.length; i++) {
        const item = agendaItems[i];
        const params = {
          id: item.id,
          durationInMinutes: item.durationInMinutes,
          title: item.title,
          additionalParticipants: item.additionalParticipants,
          fileLocationId: item.fileLocationId,
          protocolLocationId: item.protocolLocationId,
          orderIndex: i,
          isMisc: item.isMisc,
          needsDecision: item.needsDecision,
          boardMeeting: boardMeetingId,
          remarks: item.remarks,
        };

        // Execute the query using the helper class
        if (item.isNew) {
          await DatabaseHelper.executeQuery(insertQuery, params);
        } else {
          await DatabaseHelper.executeQuery(updateQuery, params);
        }
      }
      statusMessage = "All agenda items assigned and updated successfully.";
      console.log(statusMessage);
    } catch (error) {
      statusMessage = "Error assigning and updating agenda items.";
      console.error(statusMessage);
      throw error;
    }

    // B. Process UNASSIGNED agenda items
    // B1. Cancel each UNASSIGNED agenda item event if eventId is not null, undefined, or empty string
    for (const item of unassignedAgendaItems) {
      if (item.eventId) {
        try {
          await graphClient
            //.api(`/users/${mailbox}/calendar/events/${item.eventId}/cancel`)
            //.post({});
        } catch (err) {
          console.log(err);
        }
      }
    }

    // B2. Unassign agendaItems from boardMeeting and set EventId to null
    const unassignQuery = `
        UPDATE [AgendaItems]
        SET
            [BoardMeeting] = NULL,
            [EventId] = NULL
        WHERE [ID] = @id;
    `;

    try {
      for (let i = 0; i < unassignedAgendaItems.length; i++) {
        const item = unassignedAgendaItems[i];
        const params = {
          id: item.id,
        };

        // 2. Execute the query using the helper class
        await DatabaseHelper.executeQuery(unassignQuery, params);
      }
      statusMessage = "Agenda items unassigned successfully.";
      console.log(statusMessage);
    } catch (error) {
      statusMessage = "Error unassigning agenda items.";
      console.error(statusMessage);
      throw error;
    }

    // C. Ensure file structure
    try {
      const fileResult = await ensureFileStructure(boardMeetingId);
      // D. Update AgendaItems with FileLocationId
      if (fileResult.agendaItems && fileResult.agendaItems.length > 0) {
        for (const item of fileResult.agendaItems) {
          const updateQuery = `
            UPDATE AgendaItems
            SET FileLocationId = @fileLocationId
            WHERE Id = @id;
          `;
          const params = {
            id: item.id,
            fileLocationId: item.fileLocationId,
          };

          await DatabaseHelper.executeQuery(updateQuery, params);
        }
      }
    } catch (error) {
      statusMessage = "Error ensuring file structure.";
      console.error(error);
      throw error;
    }

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(statusMessage),
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Error in updateAgenda",
    };
  }
}

app.http("updateAgenda", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: updateAgenda,
});