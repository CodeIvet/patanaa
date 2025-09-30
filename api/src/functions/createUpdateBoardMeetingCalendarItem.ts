import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { calculateEndTime, createGraphClient } from "../helper";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../config";
import { DateTime } from "luxon";

export async function createUpdateBoardMeetingCalendarItem(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing createUpdateBoardMeetingCalendarItem function.");

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
      context.log(`Error parsing request body: ${String(parseError)}`);
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

  const graphClient: Client = createGraphClient("App");
  const mailbox: string = config.eventMailbox;
  context.log("EVENT_MAILBOX from config:", mailbox);
  let endTime: DateTime;

  try {
    endTime = calculateEndTime(
      DateTime.fromISO(requestBody.startTime, { setZone: requestBody.timeZone }),
      requestBody.agendaItems
    ).endTime;

    const startDateTime = DateTime.fromISO(requestBody.startTime, {
      setZone: true,
    }).toISO({
      includeOffset: false,
    });
    const endDateTime = endTime.toISO({ includeOffset: false });

    // Convert semicolon-separated string to an array and map to attendee objects
    const attendees = config.onlineMeetingHosts
      .split(";")
      .map((email) => email.trim()) // Trim spaces
      .filter((email) => email.length > 0) // Remove empty values
      .map((email) => ({
        emailAddress: { address: email, name: "" }, // Name can be left empty or fetched separately
        type: "required",
      }));

    // Define Calendar Event
    let formattedFirstName = "";
    try {
      // Step 1: Split by semicolon and trim each part
      const firstEmail = config.onlineMeetingHosts.split(";")[0].trim();

      // Step 2: Extract the part before the first dot
      const localPart = firstEmail.split("@")[0];
      const firstName = localPart.split(".")[0];

      // Step 3: Capitalize the first letter
      formattedFirstName =
        firstName.charAt(0).toUpperCase() + firstName.slice(1).toLowerCase();
    } catch (error) {
      if (error instanceof Error) {
        context.error(`Error formatting first name: ${error.message}`);
      } else {
        context.error(`Error formatting first name: ${String(error)}`);
      }
    }

    const event = {
      subject: requestBody.title,
      start: { dateTime: startDateTime, timeZone: requestBody.timeZone },
      end: {
        dateTime: endDateTime,
        timeZone: requestBody.timeZone,
      },
      location: { displayName: requestBody.room },
      attendees,
      body: {
        contentType: "HTML",
        content: `Dear participants,<br>
        Should you wish to present a deck or one-pager, please send it to <a href="mailto:patanaa@axelspringer.com">patanaa@axelspringer.com</a> 
        <u><strong>48 hours prior to the meeting.</strong></u>
        It will then be forwarded to the board.
        <br>
        Best,<br>
        ${formattedFirstName}`,
      },
      isOnlineMeeting: true,
      // isDraft: false, // this will always be sent directly because we have a fixed attendee and want to get the onlineMeetingOptions
      onlineMeetingProvider: "teamsForBusiness",
      reminderMinutesBeforeStart: 0, // No reminder
      isReminderOn: false, // Disable reminder
    };

    const patchEvent = {
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location,
    };

    let response;
    if (requestBody.isCreateAsNew) {
      // Create Event in Calendar
      context.log(
        `createUpdateBoardMeetingCalendarItem: Creating event in calendar of ${mailbox}`,
        event
      );
      response = await graphClient.api(`/users/${mailbox}/calendar/events`).post(event);

      // Get id of calendar mailbox
      context.log(
        `createUpdateBoardMeetingCalendarItem: Getting user profile of ${config.eventMailbox}`
      );
      const userProfileUrl = `/users/${config.eventMailbox}`;
      const userProfileData = await graphClient.api(userProfileUrl).get();

      // Set meeting options
      const joinWebUrl = response.onlineMeeting.joinUrl;
      const meetingOptionsUrl = `https://graph.microsoft.com/v1.0/users/${userProfileData.id}/OnlineMeetings?$filter=JoinWebUrl+eq+'${joinWebUrl}'`;

      context.log(
        `createUpdateBoardMeetingCalendarItem: Getting meeting options for joinWebUrl '${joinWebUrl}' from ${userProfileData.id}`
      );
      const oldMeetingOptions = await graphClient.api(meetingOptionsUrl).get();
      const meetingOptionsPatchUrl = `https://graph.microsoft.com/beta/users/${userProfileData.id}/OnlineMeetings/${oldMeetingOptions.value[0].id}`;

      const coOrganizers = config.onlineMeetingHosts
        .split(";")
        .map((email) => ({ upn: email.trim(), role: "coorganizer" }));

      const newMeetingOptions = {
        lobbyBypassSettings: {
          isDialInBypassEnabled: false,
          scope: "organizer",
        },
        allowedLobbyAdmitters: "organizerAndCoOrganizers",
        participants: {
          attendees: coOrganizers,
        },
      };

      context.log(
        `createUpdateBoardMeetingCalendarItem: Updating meeting options via '${meetingOptionsPatchUrl}' with ${JSON.stringify(
          newMeetingOptions
        )}`
      );

      const patchResult = await graphClient
        .api(meetingOptionsPatchUrl)
        .header("Prefer", "include-unknown-enum-members")
        .patch(newMeetingOptions);
    } else {
      // Update Event in Calendar
      context.log(
        `createUpdateBoardMeetingCalendarItem: Updating event '${requestBody.eventId}' in calendar of ${mailbox}`,
        patchEvent
      );
      response = await graphClient
        .api(`/users/${mailbox}/calendar/events/${requestBody.eventId}`)
        .patch(patchEvent);
    }

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: {
        eventId: response.id,
        joinUrl: response.onlineMeeting.joinUrl,
      },
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: `Error in createUpdateBoardMeetingCalendarItem: ${
        err instanceof Error ? err.message : String(err)
      }`,
    };
  }
}

app.http("createUpdateBoardMeetingCalendarItem", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: createUpdateBoardMeetingCalendarItem,
});