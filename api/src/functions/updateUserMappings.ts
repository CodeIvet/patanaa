import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import { BoardMeeting, ensureFileStructure } from "../helper";
import { DateTime } from "luxon";
import sql from "mssql";

export async function updateUserMappings(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing updateUserMappings function.");

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
  const accessToken = request.headers
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
    const userMappings = requestBody;

    const columns = [
      { name: "Upn", type: sql.NVarChar(255), nullable: false },
      { name: "DisplayName", type: sql.NVarChar(255), nullable: false },
    ];

    const transformedUserMappings = userMappings.map((user: any) => ({
      Upn: user.upn || "",
      DisplayName: user.displayName || "",
    }));

    // Delete all existing entries
    const query = "DELETE FROM UserMappings";
    await DatabaseHelper.executeQuery(query);

    const insertResult = await DatabaseHelper.bulkInsert(
      "UserMappings",
      columns,
      transformedUserMappings
    );

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify(insertResult),
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Error in updateUserMappings",
    };
  }
}

app.http("updateUserMappings", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: updateUserMappings,
});