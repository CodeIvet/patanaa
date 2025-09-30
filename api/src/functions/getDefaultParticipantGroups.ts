import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../config";

export async function getDefaultParticipantGroups(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing getDefaultParticipantGroups function.");

  // Prepare access token.
  const authHeader = request.headers.get("Authorization");
  const accessToken: string | undefined = authHeader
    ? authHeader.replace("Bearer ", "").trim()
    : undefined;
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  // Initialize response.
  const defaultParticipantGroups = config.defaultParticipantGroups;

  return {
    status: 200,
    jsonBody: defaultParticipantGroups,
  };
}

app.http("getDefaultParticipantGroups", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: getDefaultParticipantGroups,
});