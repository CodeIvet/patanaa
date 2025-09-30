import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../config";

export async function getDefaultRooms(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing getDefaultRooms function.");

  // Prepare access token.
  const rawAuthHeader = request.headers.get("Authorization");
  if (!rawAuthHeader) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }
  const accessToken: string = rawAuthHeader.replace("Bearer ", "").trim();

  // Initialize response.
  const defaultParticipantGroups = config.defaultRooms;

  return {
    status: 200,
    jsonBody: defaultParticipantGroups,
  };
}

app.http("getDefaultRooms", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: getDefaultRooms,
});