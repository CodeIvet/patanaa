import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import {
  AgendaItem,
  BoardMeeting,
  createGraphClient,
  createJsonForProtocolTemplate,
  streamToBuffer,
  calculateTimestamps,
  convertPascalToCamel,
  convertToPdf,
} from "../helper";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import config from "../config";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

export async function getCalendarItem(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing getCalendarItem function.");

  // Initialize response.
  const res: HttpResponseInit = {
    status: 200,
  };
  const body = Object();

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

  // Put an echo into response body.
  body.receivedHTTPRequestBody = (await request.text()) || "";
  // Initialize response.
  const eventId = request.query.get("eventId");
  const mailbox = config.eventMailbox;

  const graphClient: Client = createGraphClient("App");

  try {
    // Fetch the input file content from SharePoint
    try {
      const eventUrl = `https://graph.microsoft.com/v1.0/users/${mailbox}/calendar/events/${eventId}`;
      const event = await graphClient.api(eventUrl).get();
      // Return the data in the HTTP response
      return {
        status: 200,
        jsonBody: event,
      };
    } catch (err) {
      if (typeof err === "object" && err !== null && "code" in err && (err as any).code === "ErrorItemNotFound") {
        context.log("Item not found:", err);
        return {
          status: 200,
          body: "false",
        };
      } else {
        throw err;
      }
    }
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Error in getCalendarItem",
    };
  }
}

app.http("getCalendarItem", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: getCalendarItem,
});