import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { BoardMeeting, ensureFileStructure } from "../helper";

export async function updateAgendaItem(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing updateAgendaItem function.");

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
      context.log("Error parsing request body:", parseError);
    }
    throw new Error("Invalid JSON in request body");
  }

  // Prepare access token.
  const accessToken: string =
    request.headers.get("Authorization")?.replace("Bearer ", "").trim() || "";
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  try {
    const agendaItemId = requestBody.agendaItemId;
    const eventId = requestBody.eventId;

    const updateQuery = `
      UPDATE [AgendaItems]
      SET [EventId] = @eventId      
      WHERE [Id] = @id;
  `;

    const params = {
      id: agendaItemId,
      eventId: eventId,
    };

    // Execute the query using the helper class
    const updatedAgendaItem = await DatabaseHelper.executeQuery(updateQuery, params);

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(updatedAgendaItem),
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Error in updateAgendaItem",
    };
  }
}

app.http("updateAgendaItem", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: updateAgendaItem,
});