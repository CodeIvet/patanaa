import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import DatabaseHelper from "../DatabaseHelper";
import {
  AgendaItem,
  BoardMeeting,
  createGraphClient,
  streamToBuffer,
  calculateTimestamps,
  convertPascalToCamel,
  convertToPdf,
  createJsonForAgendaTemplate,
  parseDbRecordToBoardMeeting,
  convertCamelToPascal,
  ensureFileStructure,
} from "../helper";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../config";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { DateTime } from "luxon";


export async function createAgendaPdf(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  context.log("Processing createAgendaPdf function.");

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

  // Prepare access token from request header
const accessTokenHeader = request.headers.get("Authorization");
const accessToken = accessTokenHeader?.replace("Bearer ", "").trim();

if (!accessToken) {
  return {
    status: 400,
    body: JSON.stringify({ error: "No access token found in request header" }),
  };
}
context.log("Access Token:", accessToken);


// Create the Graph client with the token
const graphClient: Client = createGraphClient("App", accessToken);


  try {
    const boardMeeting: BoardMeeting = parseDbRecordToBoardMeeting(
      convertCamelToPascal([requestBody.boardMeeting])[0]
    );

  // ensure folder is created for board meeting
if (typeof boardMeeting.id !== "number") {
  throw new Error("BoardMeeting id is undefined. Cannot ensure file structure.");
}
const fileStructure = await ensureFileStructure(boardMeeting.id);
boardMeeting.fileLocationId = fileStructure.boardMeetingFileLocationId;


    //create Graph client with access token
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
      // { fileId: config.agendaPdfTemplateFileIdDe, language: "DE", includeRemarks: true },
      // { fileId: config.agendaPdfTemplateFileIdDe, language: "DE", includeRemarks: false },
      { fileId: config.agendaPdfTemplateFileIdEn, language: "EN", includeRemarks: true },
      { fileId: config.agendaPdfTemplateFileIdEn, language: "EN", includeRemarks: false },
    ];

    try {
      for (const templateFile of templateFiles) {
        const json = await createJsonForAgendaTemplate(
          boardMeeting,
          agendaItems,
          templateFile.language,
          templateFile.includeRemarks,
          graphClient
        );

        const templateFileId = templateFile.fileId;

        // Fetch the input file content from SharePoint
        const fileUrl = `/sites/${config.sharePointWebsite}/drives/${config.assetsDriveId}/items/${templateFileId}/content`;
        const responseStream = await graphClient.api(fileUrl).getStream();
        // Convert the ReadableStream to a Buffer
        const contentBuffer = await streamToBuffer(responseStream);
        // Load the DOCX content using PizZip
        const zip = new PizZip(contentBuffer);
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
        });

        doc.render(json);

        // Generate the updated document as a Node.js Buffer
        const updatedDocument = doc.getZip().generate({
          type: "nodebuffer",
          compression: "DEFLATE",
        });

        // Upload the updated document back to SharePoint
  if (!boardMeeting.fileLocationId) {
    throw new Error("BoardMeeting has no fileLocationId, cannot upload/convert.");
  }

  const uploadUrl = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${boardMeeting.fileLocationId}:/Agenda_temp_${templateFile.language}.docx:/content`;

  try {
    const newDocument = await graphClient.api(uploadUrl).put(updatedDocument);
    const withRemarksString = templateFile.includeRemarks ? "" : " clean";

    // Convert the document to PDF
    await convertToPdf(
      config.sharePointMeetingsDriveId,
      newDocument.id,
      config.sharePointMeetingsDriveId,
      boardMeeting.fileLocationId,
      `Agenda-${boardMeeting.title}${withRemarksString}.pdf`,
      graphClient
    );

    // Remove the temporary file
    const deleteUrl = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${newDocument.id}`;
    await graphClient.api(deleteUrl).delete();
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
      body: "Error in createAgendaPdf",
    };
  }
}

app.http("createAgendaPdf", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: createAgendaPdf,
});