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
  parseDbRecordToBoardMeeting,
  convertCamelToPascal,
} from "../helper";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import config from "../config";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { DateTime } from "luxon";

export async function processProtocolTemplate(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing processProtocolTemplate function.");

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

  try {
    const boardMeeting: BoardMeeting = parseDbRecordToBoardMeeting(
      convertCamelToPascal([requestBody.boardMeeting])[0]
    );

    const graphClient: Client = createGraphClient("App");

    const getAgendaItemsQuery =
      "SELECT * FROM AgendaItems WHERE BoardMeeting = @Id ORDER BY OrderIndex ASC";
    const getAgendaItemsParams = { Id: boardMeeting.id };
    const rawAgendaItems = await DatabaseHelper.executeQuery(
      getAgendaItemsQuery,
      getAgendaItemsParams
    );
    const agendaItems: AgendaItem[] = rawAgendaItems
      ? calculateTimestamps(
          boardMeeting.startTime ? boardMeeting.startTime : DateTime.now(),
          convertPascalToCamel(rawAgendaItems) as AgendaItem[]
        )
      : [];

    const templateFiles = [
      { fileId: config.protocolTemplateFileIdDe, language: "DE" },
      { fileId: config.protocolTemplateFileIdEn, language: "EN" },
    ];

    try {
      // Fetch the input file content from SharePoint
      for (const templateFile of templateFiles) {
        const json = await createJsonForProtocolTemplate(
          boardMeeting,
          agendaItems,
          templateFile.language,
          graphClient
        );

        const fileUrl = `/sites/${config.sharePointWebsite}/drives/${config.assetsDriveId}/items/${templateFile.fileId}/content`;
        const responseStream = await graphClient.api(fileUrl).getStream();
        const contentBuffer = await streamToBuffer(responseStream);
        const zip = new PizZip(contentBuffer);
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
        });

        doc.render(json);

        const updatedDocument = doc.getZip().generate({
          type: "nodebuffer",
          compression: "DEFLATE",
        });

        const uploadUrl = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${boardMeeting.fileLocationId}:/Protocol DRAFT ${templateFile.language}.docx:/content`;
        try {
          await graphClient.api(uploadUrl).put(updatedDocument);
        } catch (err) {
          throw err;
        }
      }
    } catch (err) {
      context.error("Error:", err);
      return {
        status: 500,
        body: err instanceof Error ? err.message : String(err),
      };
    }

    // Return the data in the HTTP response
    return {
      status: 200,
      jsonBody: JSON.stringify("newMeeting"),
    };
  } catch (err) {
    context.error("Error:", err);
    return {
      status: 500,
      body: "Error in processProtocolTemplate",
    };
  }
}

app.http("processProtocolTemplate", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: processProtocolTemplate,
});