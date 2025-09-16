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

  let requestBody: any;
  try {
    const rawBody = (await request.text()) || "{}";
    requestBody = JSON.parse(rawBody);
  } catch (err) {
    context.error("Failed to parse request body:", err);
    return { status: 400, body: "Invalid JSON in request body" };
  }

  // Extract access token
  const accessTokenHeader = request.headers.get("Authorization");
  const accessToken = accessTokenHeader?.replace("Bearer ", "").trim();
  if (!accessToken) {
    return {
      status: 400,
      body: JSON.stringify({ error: "No access token found in request header" }),
    };
  }

  // Create Graph client once with token
  const graphClient: Client = createGraphClient("App", accessToken);

  try {
    // Parse boardMeeting
    const boardMeeting: BoardMeeting = parseDbRecordToBoardMeeting(
      convertCamelToPascal([requestBody.boardMeeting])[0]
    );

   
    // Ensure file structure exists
    context.log("Parsed boardMeeting:", boardMeeting);
const boardMeetingId = Number(boardMeeting.id);
if (!boardMeetingId || isNaN(boardMeetingId)) {
  throw new Error("BoardMeeting id is undefined or not a valid number. Cannot ensure file structure.");
}

const fileStructure = await ensureFileStructure(boardMeetingId);
boardMeeting.fileLocationId = fileStructure.boardMeetingFileLocationId;

    // Fetch agenda items
    const getAgendaItemsQuery =
      "SELECT * FROM AgendaItems WHERE BoardMeeting = @Id ORDER BY OrderIndex ASC";
    const rawAgendaItems = await DatabaseHelper.executeQuery(getAgendaItemsQuery, {
      Id: boardMeeting.id,
    });
    const agendaItems: AgendaItem[] = rawAgendaItems
      ? calculateTimestamps(
          boardMeeting.startTime ?? DateTime.now(),
          convertPascalToCamel(rawAgendaItems) as AgendaItem[]
        )
      : [];

    // Define template files
    const templateFiles = [
      { fileId: config.agendaPdfTemplateFileIdEn, language: "EN", includeRemarks: true },
      { fileId: config.agendaPdfTemplateFileIdEn, language: "EN", includeRemarks: false },
    ];

    for (const templateFile of templateFiles) {
      // Generate JSON for template
      const json = await createJsonForAgendaTemplate(
        boardMeeting,
        agendaItems,
        templateFile.language,
        templateFile.includeRemarks,
        graphClient
      );

      // Fetch DOCX template from SharePoint
      const fileUrl = `/sites/${config.sharePointWebsite}/drives/${config.assetsDriveId}/items/${templateFile.fileId}/content`;
      const responseStream = await graphClient.api(fileUrl).getStream();
      const contentBuffer = await streamToBuffer(responseStream);

      // Render DOCX with Docxtemplater
      const zip = new PizZip(contentBuffer);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
      doc.render(json);
      const updatedDocument = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

      // Ensure fileLocationId exists
      if (!boardMeeting.fileLocationId) {
        throw new Error("BoardMeeting has no fileLocationId. Cannot upload or convert file.");
      }

      // Upload DOCX
      const uploadUrl = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${boardMeeting.fileLocationId}:/Agenda_temp_${templateFile.language}.docx:/content`;
      const newDocument = await graphClient.api(uploadUrl).put(updatedDocument);

      // Convert to PDF
      const pdfName = `Agenda-${boardMeeting.title}${templateFile.includeRemarks ? "" : " clean"}.pdf`;
      await convertToPdf(
        config.sharePointMeetingsDriveId,
        newDocument.id,
        config.sharePointMeetingsDriveId,
        boardMeeting.fileLocationId,
        pdfName,
        graphClient
      );

      // Delete temporary DOCX
      const deleteUrl = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${newDocument.id}`;
      await graphClient.api(deleteUrl).delete();
    }

    return { status: 200, jsonBody: JSON.stringify("newMeeting") };
  } catch (err) {
    context.error("Error in createAgendaPdf:", err);
    return {
      status: 500,
      body: err instanceof Error ? err.message : String(err),
    };
  }
}

app.http("createAgendaPdf", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: createAgendaPdf,
});