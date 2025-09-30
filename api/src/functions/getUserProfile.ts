/* This code sample provides a starter kit to implement server side logic for your Teams App in TypeScript,
 * refer to https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference for complete Azure Functions
 * developer guide.
 */

// Import polyfills for fetch required by msgraph-sdk-javascript.
import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { createGraphClient } from "../helper";

/**
 * This function handles requests from teamsapp client.
 * The HTTP request should contain an SSO token queried from Teams in the header.
 *
 * This function initializes the teamsapp SDK with the configuration and calls these APIs:
 * - new OnBehalfOfUserCredential(accessToken, oboAuthConfig) - Construct OnBehalfOfUserCredential instance with the received SSO token and initialized configuration.
 * - getUserInfo() - Get the user's information from the received SSO token.
 *
 * The response contains multiple message blocks constructed into a JSON object, including:
 * - An echo of the request body.
 * - The display name encoded in the SSO token.
 * - Current user's Microsoft 365 profile if the user has consented.
 *
 * @param {InvocationContext} context - The Azure Functions context object.
 * @param {HttpRequest} req - The HTTP request.
 */
export async function getUserProfiles(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("HTTP trigger function processed a request.");

  // Initialize response.
  const res: HttpResponseInit = {
    status: 200,
  };
  const body = Object();

  // Put an echo into response body.
  body.receivedHTTPRequestBody = (await req.text()) || "";
  let requestBody: any;
  try {
    requestBody = body.receivedHTTPRequestBody
      ? JSON.parse(body.receivedHTTPRequestBody)
      : {};
  } catch (parseError) {
    const message = typeof parseError === "object" && parseError !== null && "message" in parseError ? (parseError as any).message : String(parseError);
    context.log(`Error parsing request body: ${message}`);
    throw new Error("Invalid JSON in request body");
  }

  // Prepare access token.
  const accessTokenHeader = req.headers.get("Authorization");
  const accessToken: string = accessTokenHeader ? accessTokenHeader.replace("Bearer ", "").trim() : "";
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({
        error: "No access token was found in request header.",
      }),
    };
  }

  const graphClient: Client = createGraphClient("App");

  const batchRequests = (requestBody.upns as string[]).map((upn: string, index: number) => ({
    id: `${index}`,
    method: "GET",
    url: `/users/${upn}?$select=displayName,mail`,
  }));

  // Send batch request for user display names
  const batchResponse = await graphClient
    .api("/$batch")
    .post({ requests: batchRequests });

  // Extract user profiles with display names
  const userProfiles = batchResponse.responses.map((response: any) => {
    const upn = requestBody.upns[parseInt(response.id)];

    return response.status === 200
      ? {
          displayName: response.body.displayName,
          mail: response.body.mail,
          upn,
        }
      : {
          displayName: upn, // Use UPN for external users
          mail: upn,
          upn,
        };
  });

  return {
    ...res,
    body: JSON.stringify(userProfiles),
  };
}

app.http("getUserProfiles", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: getUserProfiles,
});
