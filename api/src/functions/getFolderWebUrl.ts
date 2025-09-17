import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { createGraphClient, getFolderLink } from "../helper";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../config";

export async function getFolderWebUrl(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing getFolderWebUrl function.");

  // Prepare access token
  const rawAuthHeader = request.headers.get("Authorization");
  const accessToken: string | undefined = rawAuthHeader
    ? rawAuthHeader.replace("Bearer ", "").trim()
    : undefined;
  if (!accessToken) {
    context.error("No access token found in request header.");
    return {
      status: 400,
      body: "No access token found in request header.",
    };
  }

  // Validate query parameters
  const driveName = request.query.get("driveName");
  const fileLocationId = request.query.get("fileLocationId");
  if (!driveName) {
    context.error("Missing required query parameter: driveName");
    return {
      status: 400,
      body: "Missing required query parameter: driveName",
    };
  }
  if (!fileLocationId) {
    context.error("Missing required query parameter: fileLocationId");
    return {
      status: 400,
      body: "Missing required query parameter: fileLocationId",
    };
  }

  // Select driveId based on driveName
  let driveId: string;
  if (driveName === "Meetings") {
    driveId = config.sharePointMeetingsDriveId;
  } else if (driveName === "Assets") {
    driveId = config.assetsDriveId;
  } else {
    context.error(`Unknown driveName: ${driveName}`);
    return {
      status: 400,
      body: `Unknown driveName: ${driveName}`,
    };
  }

  try {
    const graphClient: Client = createGraphClient("App", accessToken);
    const webUrl = await getFolderLink(driveId, fileLocationId, graphClient);
    if (!webUrl || typeof webUrl !== "string") {
      context.error("Could not retrieve folder webUrl.");
      return {
        status: 500,
        body: "Could not retrieve folder webUrl.",
      };
    }
    // Return the plain URL in the HTTP response
    return {
      status: 200,
      body: webUrl,
    };
  } catch (err) {
    context.error("Error in getFolderWebUrl:", err);
    return {
      status: 500,
      body: err instanceof Error ? err.message : String(err),
    };
  }
}

app.http("getFolderWebUrl", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: getFolderWebUrl,
});