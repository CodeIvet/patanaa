import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { createGraphClient, getMailBody } from "../helper";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../config";
import { DateTime } from "luxon";

export async function createUpdateAgendaItemCalendarItem(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("createUpdateAgendaItemCalendarItem called at", new Date().toISOString());
  context.log("Processing createUpdateAgendaItemCalendarItem function.");

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
  const rawAuthHeader = request.headers.get("Authorization");
  const accessToken: string =
    rawAuthHeader?.replace("Bearer ", "").trim() ?? "";
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  const graphClient: Client = createGraphClient("App");
  const mailbox: string = config.eventMailbox;

  try {
    const startDateTime = DateTime.fromISO(requestBody.startTime, {
      setZone: true,
    }).toISO({
      includeOffset: false,
    });
    const endDateTime = DateTime.fromISO(requestBody.endTime, {
      setZone: true,
    }).toISO({
      includeOffset: false,
    });

    // Convert semicolon-separated string to an array and map to attendee objects
    const combinedParticipants =
      `${requestBody.mainMeeting.fixedParticipants};${requestBody.participants}`.replace(
        /;;/g,
        ";"
      );
    const attendees = combinedParticipants
  .split(";")
  .map((email) => email.trim())
  .filter((email) => email.length > 0)
  .map((email) => ({
    emailAddress: { address: email, name: "" },
    type: "required",
  }));

    const mailBody = await getMailBody(requestBody.mainMeeting.eventId, graphClient);

    // Define Calendar Event
    const event = {
      subject: requestBody.title,
      start: { dateTime: startDateTime, timeZone: requestBody.timeZone },
      end: {
        dateTime: endDateTime,
        timeZone: requestBody.timeZone,
      },
      location: { displayName: requestBody.mainMeeting.room },
      attendees: attendees.length > 0 ? attendees : [],
      body: {
        contentType: "HTML",
        content: mailBody,
      },
      isDraft: !requestBody.isAlreadySent, // Marking it as a draft when not sent
      reminderMinutesBeforeStart: 0, // No reminder
      isReminderOn: false, // Disable reminder
    };

    let response;
    if (requestBody.isCreateAsNew) {
      // Create Event in Calendar
      response = await graphClient.api(`/users/${mailbox}/calendar/events`).post(event);
    } else {
      // Update Event in Calendar
      response = await graphClient
        .api(`/users/${mailbox}/calendar/events/${requestBody.eventId}`)
        .patch(event);
    }

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: response.id,
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: `Error in createUpdateAgendaItemCalendarItem: ${
        err instanceof Error ? err.message : String(err)
      }`,
    };
  }
}

app.http("createUpdateAgendaItemCalendarItem", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: createUpdateAgendaItemCalendarItem,
});