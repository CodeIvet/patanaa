import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { AgendaItem } from "../helper";


export async function getAgendaItems(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing getAgendaItems function.");

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


  // Initialize response.
  const meetingId = request.query.get("boardmeeting");

  // Set the MeetingId to `null` if it's empty, `null`, or an empty string
  const sanitizedMeetingId = meetingId ? meetingId : null;

  try {
    // Define your SQL query and parameters
    const query = `
      SELECT * FROM AgendaItems 
      WHERE (@MeetingId IS NULL AND BoardMeeting IS NULL) 
      OR (BoardMeeting = @MeetingId)
      ORDER BY OrderIndex ASC;
      `;
    const params = { MeetingId: sanitizedMeetingId };

    // Execute the query using the helper class
    const agendaItems = await DatabaseHelper.executeQuery(query, params);

    // Transform the JSON data to an array of BoardMeeting objects
    const transformedData: AgendaItem[] = agendaItems.map((item: any) => ({
      id: item.ID.toString(), // Convert ID to string
      durationInMinutes: item.DurationInMinutes,
      title: item.Title,
      additionalParticipants: item.AdditionalParticipants,
      fileLocationId: item.FileLocationId,
      ProtocolLocationId: item.ProtocolLocationId,
      orderIndex: item.OrderIndex,
      isMisc: item.IsMisc,
      needsDecision: item.NeedsDecision,
      boardMeeting: item.BoardMeeting,
      eventId: item.EventId,
      remarks: item.Remarks,
    }));

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(transformedData),
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Internal server error",
    };
  }
}

app.http("getAgendaItems", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: getAgendaItems,
});