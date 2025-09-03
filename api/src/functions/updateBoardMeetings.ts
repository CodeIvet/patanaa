import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { BoardMeeting, ensureFileStructure } from "../helper";
import { DateTime } from "luxon";

export async function updateBoardMeeting(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing updateBoardMeeting function.");

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
    const errorMessage = (parseError instanceof Error) ? parseError.message : String(parseError);
    context.log(`Error parsing request body: ${errorMessage}`);
    throw new Error("Invalid JSON in request body");
  }

  // Prepare access token.
  const accessToken: string = request.headers
    .get("Authorization")
    ?.replace("Bearer ", "")
    .trim() ?? "";
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  try {
    const boardMeeting: BoardMeeting = requestBody.boardmeeting;
    const shouldEnsureFileStructure = requestBody.ensureFileStructure;

    const updateQuery = `
    UPDATE [BoardMeetings]
    SET 
        [StartTime] = @startTime,
        [Title] = @title,
        [FixedParticipants] = @fixedParticipants,
        [Remarks] = @remarks,
        [Location] = @location,
        [EventId] = @eventId,
        [TimeZone] = @timeZone,
        [MeetingLink] = @meetingLink,
        [Room] = @room
    WHERE [Id] = @id;
`;

    const params = {
      id: boardMeeting.id,
      startTime: boardMeeting.startTime,
      title: boardMeeting.title,
      fixedParticipants: boardMeeting.fixedParticipants,
      remarks: boardMeeting.remarks,
      location: boardMeeting.location,
      eventId: boardMeeting.eventId,
      timeZone: boardMeeting.timeZone,
      meetingLink: boardMeeting.meetingLink || "",
      room: boardMeeting.room,
    };

    // Execute the query using the helper class
    const newMeeting = await DatabaseHelper.executeQuery(updateQuery, params);

    if (!boardMeeting.id) {
  throw new Error("BoardMeeting id is missing");
}
await ensureFileStructure(boardMeeting.id);


    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(newMeeting),
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Error in updateBoardMeeting",
    };
  }
}

app.http("updateBoardMeeting", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: updateBoardMeeting,
});