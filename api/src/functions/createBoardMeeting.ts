import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { BoardMeeting, ensureFileStructure } from "../helper";
import config from "../config"; // ✅ import centralized config

export async function createBoardMeeting(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing createBoardMeeting function.");

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
    const boardMeeting: BoardMeeting = requestBody;

    // 1. Create database entry
    const insertQuery = `
      INSERT INTO [BoardMeetings] 
      ([StartTime], [Title], [FixedParticipants], [Remarks], [Location], [TimeZone], [Room])
      OUTPUT INSERTED.*
      VALUES (@startTime, @title, @fixedParticipants, @remarks, @location, @timeZone, @room);
  `;
    const params = {
      startTime: boardMeeting.startTime,
      title: boardMeeting.title,
      fixedParticipants: boardMeeting.fixedParticipants,
      remarks: boardMeeting.remarks,
      location: boardMeeting.location,
      timeZone: boardMeeting.timeZone,
      room: boardMeeting.room,
    };
    const [newMeeting] = await DatabaseHelper.executeQuery(insertQuery, params);

    // 2.Insert Sharepoint Filestructure

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(newMeeting),
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Error in createBoardMeeting",
    };
  }
}

app.http("createBoardMeeting", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: createBoardMeeting,
});