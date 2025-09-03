import { TokenCredential } from "@azure/identity";
import { DateTime } from "luxon";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import {
  AppCredential,
  AppCredentialAuthConfig,
  OnBehalfOfCredentialAuthConfig,
  OnBehalfOfUserCredential,
} from "@microsoft/teamsfx";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import config from "./config";
import DatabaseHelper from "./DatabaseHelper";
import { format } from "path";

export function createGraphClient(type: "App" | "User", accessToken?: string): Client {
  const credential = createTeamsFxCredential(type, accessToken);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });

  // Initialize the Graph client
  const graphClient = Client.initWithMiddleware({
    authProvider: authProvider,
  });
  return graphClient;
}

// Function to get the webUrl of a folder
export async function getFolderLink(
  driveId: string,
  folderId: string,
  graphClient: Client
): Promise<string> {
  try {
    // Fetch folder details
    const url = `/sites/${config.sharePointWebsite}/drives/${driveId}/items/${folderId}`;
    const folder = await graphClient.api(url).get();

    // Extract the webUrl
    const folderLink = folder.webUrl;

    console.log("Folder link generated:", folderLink);
    return folderLink;
  } catch (error) {
    console.error("Error fetching folder link:", error);
    throw error;
  }
}

// Helper function to convert ReadableStream to Buffer
export async function streamToBuffer(
  readableStream: NodeJS.ReadableStream
): Promise<Buffer> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of readableStream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

export async function createLinkFile(
  hostDriveId: string,
  hostFolderId: string,
  httpLinkUrl: string,
  linkTitle: string,
  graphClient: Client
) {
  try {
    // Define the file name and content
    const linkFileName = `${linkTitle}.url`;
    const fileContent = `[InternetShortcut]\nURL=${httpLinkUrl}`;
    // const folderPath = await getFolderPathFromId(hostFolderId, hostDriveId);

    // Define the URL to upload the file

    // https://graph.microsoft.com/v1.0/sites/moveoffice.sharepoint.com,3f4c5443-1d29-439e-af22-7a2cc2749fa0,c807d726-1282-4311-b8bf-df332cdf4279/drives/b!Q1RMPykdnkOvInoswnSfoCbXB8iCEhFDuL_fMyzfQnl4hfi-BolbT607GNpdjVbN/items/01ZMORK2JZOS7ICSDQ6VA3FH2EDJEAIGXU:/dummy.txt:/content

    const url = `/sites/${config.sharePointWebsite}/drives/${hostDriveId}/items/${hostFolderId}:/${linkFileName}:/content`;

    // Upload the link file
    const linkFile = await graphClient
      .api(url)
      .headers({ "Content-Type": "text/plain" }) // Specify the content type
      .query({ "@microsoft.graph.conflictBehavior": "replace" }) // Enforce overwrite
      .put(fileContent);

    console.log("Link file created successfully:", linkFile);
    return linkFile;
  } catch (error) {
    console.error("Error creating link file:", error);
    throw error;
  }
}

function formatDateForFolder(date: DateTime, includeTime: boolean = false) {
  const datePart = date.toFormat("yyyy-MM-dd"); // Formats as YYYY-MM-DD

  if (includeTime) {
    const timePart = date.toFormat("HHmm"); // Formats as HHmm (24-hour format, no colon)
    return `${datePart}_${timePart}`;
  } else {
    return datePart;
  }
}

export const calculateTimestamps = (
  startTime: DateTime,
  items: AgendaItem[]
): AgendaItem[] => {
  let cumulativeDuration = 0;

  return items.map((item: AgendaItem): AgendaItem => {
    // Calculate the timestamp for the current item
    const itemTimestamp: DateTime = startTime.plus({ minutes: cumulativeDuration }); // Keeps the timezone
    // Update cumulative duration for the next item
    cumulativeDuration += item.durationInMinutes;

    // Return the item with the added timestamp property
    return {
      ...item,
      startTime: itemTimestamp,
    };
  });
};

export function getSafeString(unsafeString: string) {
  return unsafeString
    .substring(0, 40)
    .replace(/[^\w0-9\s\-\_äöüÄÖÜß.]/g, "_")
    .trim();
}

export async function ensureFileStructure(boardMeetingId: number) {
  // init reusables
  const graphClient: Client = createGraphClient("App");

  try {
    // 0a. Get board meeting
    const getMeetingQuery = `
    SELECT *
    FROM BoardMeetings
    WHERE Id = @Id;
  `;
    const getMeetingParams = { Id: boardMeetingId };
    const rawBoardMeetingArray = await DatabaseHelper.executeQuery(
      getMeetingQuery,
      getMeetingParams
    );

    const boardMeeting = rawBoardMeetingArray?.[0];
if (!boardMeeting) {
  throw new Error("No board meeting found");
}


    // 0b. Get agenda items for board meeting
    const getAgendaItemsQuery =
      "SELECT * FROM AgendaItems WHERE BoardMeeting = @Id ORDER BY OrderIndex ASC";
    const getAgendaItemsParams = { Id: boardMeetingId };
    const rawAgendaItems = await DatabaseHelper.executeQuery(
      getAgendaItemsQuery,
      getAgendaItemsParams
    );

    const agendaItems: AgendaItem[] = rawAgendaItems
      ? calculateTimestamps(
          boardMeeting.startTime ?? DateTime.now(),
          convertPascalToCamel(rawAgendaItems) as AgendaItem[]
        )
      : [];

    // 0c. Get orphaned agenda items
    const getOrphanedAgendaItemsQuery =
      "SELECT * FROM AgendaItems WHERE BoardMeeting IS NULL OR BoardMeeting = ''";
    const rawOrphanedAgendaItems = await DatabaseHelper.executeQuery(
      getOrphanedAgendaItemsQuery
    );
    const orphanedAgendaItems: AgendaItem[] = rawOrphanedAgendaItems
      ? (convertPascalToCamel(rawOrphanedAgendaItems) as AgendaItem[])
      : [];

    console.log(
      "Board meeting and agenda items fetched:",
      boardMeeting,
      agendaItems.length,
      orphanedAgendaItems.length
    );

    try {
      // 1. Ensure meeting folder
      const meetingFolderName =
        formatDateForFolder(boardMeeting.startTime, false) +
        " - " +
        getSafeString(boardMeeting.title);
      if (!boardMeeting?.fileLocationId) {
        // create as new meeting folder
        let url = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${config.sharePointMeetingFolderId}/children`;
        console.log(
          "Creating new meeting folder with url:",
          `"${url}"`,
          "and name:",
          `"${meetingFolderName}"`
        );
        const meetingFolder = await graphClient.api(url).post({
          name: meetingFolderName,
          folder: {},
          "@microsoft.graph.conflictBehavior": "rename",
        });
        boardMeeting.fileLocationId = meetingFolder.id;
        console.log("Successfully created meeting folder:", meetingFolderName);
      } else {
        // update existing meeting folder
        const url = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${boardMeeting.fileLocationId}`;
        console.log(
          "Updating existing meeting folder with url:",
          `"${url}"`,
          "and name:",
          `"${meetingFolderName}"`
        );
        await graphClient.api(url).patch({
          name: meetingFolderName,
          "@microsoft.graph.conflictBehavior": "rename",
        });
        console.log("Successfully updated meeting folder:", meetingFolderName);
      }
    } catch (error) {
      console.error("Error ensuring meeting folder:", error);
      throw error;
    }

    try {
      // 2. Ensure agenda folders
      for (const item of agendaItems) {
        const folderName =
          (item.orderIndex + 1).toString().padStart(2, "0") +
          " - " +
          getSafeString(item.title);
        if (!item?.fileLocationId) {
          // create as new agendaitem folder
          let url = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${boardMeeting.fileLocationId}/children`;
          console.log(
            "Creating new agenda item folder with url:",
            `"${url}"`,
            "and name:",
            `"${folderName}"`
          );
          const folderItem = await graphClient.api(url).post({
            name: folderName,
            folder: {},
            "@microsoft.graph.conflictBehavior": "rename",
          });

          item.fileLocationId = folderItem.id;
          console.log("Successfully created agenda item folder:", folderName);
        } else {
          // update existing agendaitem folder
          const url = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${item.fileLocationId}`;
          console.log(
            "Updating existing agenda item folder with url:",
            `"${url}"`,
            "and name:",
            `"${folderName}"`
          );
          await graphClient.api(url).patch({
            name: folderName,
            parentReference: { id: boardMeeting.fileLocationId },
            "@microsoft.graph.conflictBehavior": "rename",
          });
          console.log("Successfully updated agenda item folder:", folderName);
        }
      }
    } catch (error) {
      console.error("Error ensuring agenda item folders:", error);
      throw error;
    }

    try {
      // 4. Move orphaned agenda folders to topic storage
      for (const item of orphanedAgendaItems) {
        console.log("Moving orphaned agenda item folder:", item.title);
        console.log(JSON.stringify(item, null, 2));
        console.log("safe string:", getSafeString(item.title));
        let url = `/sites/${config.sharePointWebsite}/drives/${config.sharePointMeetingsDriveId}/items/${item.fileLocationId}`;
        try {
          await graphClient.api(url).patch({
            name: getSafeString(item.title),
            parentReference: { id: config.sharePointUnassignedTopsFolderId },
            "@microsoft.graph.conflictBehavior": "rename",
          });

          console.log("successfully moved orphaned agenda item folder:", item.title);
        } catch (error: any) {
          if (error?.statusCode === 404) {
            console.error(
              `Orphaned agenda item folder not found: ${item.title} (${item.fileLocationId})`
            );
          } else {
            console.error(
              `Failed to move orphaned agenda item folder: ${item.title} (${item.fileLocationId})`
            );
          }
        }
      }
    } catch (error) {
      console.error("Error moving orphaned agenda item folders:", error);
      throw error;
    }

    console.log("File structure created successfully.");

    // Return JSON object with fileLocationId of meeting and agenda items
    return {
      boardMeetingFileLocationId: boardMeeting.fileLocationId,
      agendaItems: agendaItems.map((item) => ({
        id: item.id,
        title: item.title,
        fileLocationId: item.fileLocationId,
      })),
    };
  } catch (error) {
    console.error("Error fetching board meeting and agenda items:", error);
    throw error;
  }
}

export function convertPascalToCamel(
  items: Record<string, any>[]
): Record<string, any>[] {
  return items.map((item) => {
    const camelCasedItem: Record<string, any> = {};
    for (const key in item) {
      if (key === "ID") {
        camelCasedItem["id"] = item[key];
      } else {
        const camelKey = key.charAt(0).toLowerCase() + key.slice(1);
        camelCasedItem[camelKey] = item[key];
      }
    }
    return camelCasedItem;
  });
}

export function convertCamelToPascal(
  items: Record<string, any>[]
): Record<string, any>[] {
  return items.map((item) => {
    const pascalCasedItem: Record<string, any> = {};
    for (const key in item) {
      if (key === "ID") {
        pascalCasedItem["id"] = item[key];
      } else {
        const camelKey = key.charAt(0).toUpperCase() + key.slice(1);
        pascalCasedItem[camelKey] = item[key];
      }
    }
    return pascalCasedItem;
  });
}

export function parseDbRecordToBoardMeeting(row: any): BoardMeeting {
  const boardMeeting: BoardMeeting = {
    id: row.Id,
    startTime: DateTime.isDateTime(row.StartTime)
      ? row.StartTime
      : row.StartTime instanceof Date
      ? DateTime.fromJSDate(row.StartTime, { zone: row.TimeZone })
      : DateTime.fromISO(row.StartTime, { setZone: row.TimeZone }),
    title: row.Title,
    fixedParticipants: row.FixedParticipants,
    remarks: row.Remarks,
    location: row.Location,
    meetingLink: row.MeetingLink ?? undefined,
    fileLocationId: row.FileLocationId ?? undefined,
    eventId: row.EventId ?? undefined,
    timeZone: row.TimeZone,
    room: row.Room ?? undefined,
  };

  return boardMeeting;
}

async function deleteLinkFiles(
  driveId: string,
  fileLocationId: string,
  graphClient: Client
) {
  const folderUrl = `/sites/${config.sharePointWebsite}/drives/${driveId}/items/${fileLocationId}/children`;
  const folderItems = await graphClient.api(folderUrl).get();
  const filesToDelete = folderItems.value.filter((item: any) => {
    return item.file && item.name.endsWith(".url");
  });
  for (const file of filesToDelete) {
    const fileUrl = `/sites/${config.sharePointWebsite}/drives/${driveId}/items/${file.id}`;
    await graphClient.api(fileUrl).delete();
  }
  console.log(
    `Deleted ${filesToDelete.length} .url file(s) from /drives/${driveId}/items/${fileLocationId} .`
  );
}

function createTeamsFxCredential(
  type: "App" | "User",
  accessToken?: string
): TokenCredential {
  const authConfig = {
    authorityHost: config.authorityHost,
    clientId: config.clientId,
    tenantId: config.tenantId,
    clientSecret: config.clientSecret,
  };

  if (type === "User") {
    if (!accessToken) {
      throw new Error("Access token is required for User credential type.");
    }
    return new OnBehalfOfUserCredential(
      accessToken,
      authConfig as OnBehalfOfCredentialAuthConfig
    );
  } else {
    return new AppCredential(authConfig as AppCredentialAuthConfig);
  }
}

export async function createJsonForProtocolTemplate(
  boardMeeting: BoardMeeting,
  agendaItems: AgendaItem[],
  languageCode: string,
  graphClient: Client
) {
  const formattedDate: string = boardMeeting.startTime
    .setLocale(languageCode.toLowerCase())
    .toLocaleString(DateTime.DATE_SHORT);

  const rawFixedParticipants = boardMeeting.fixedParticipants.split(";");
  const rawAllParticipants = new Set(rawFixedParticipants);
  agendaItems.forEach((item) => {
    const additionalParticipants = item.additionalParticipants.split(";");
    additionalParticipants.forEach((participant) => rawAllParticipants.add(participant));
  });
  const uniqueParticipants = Array.from(rawAllParticipants);
  const participantsInfo = await getDisplayNamesBatch(graphClient, uniqueParticipants);

  const lookupDisplayName = (upn: string): string => {
    const [lastName, firstName] = participantsInfo[upn]?.split(", ") || [];
    if (firstName?.trim() && lastName?.trim()) {
      return `${firstName} ${lastName}`;
    } else if (firstName?.trim()) {
      return firstName;
    } else if (lastName?.trim()) {
      return lastName;
    } else {
      return "Unknown Participant";
    }
  };

  // Populate fixedParticipants
  const fixedParticipants = rawFixedParticipants.map((participant) => ({
    fixedPerson: lookupDisplayName(participant),
    totalTops: agendaItems.length,
  }));

  // Populate topsDetails
  const topsDetails = agendaItems.map((item, index) => ({
    agendaTitle: item.title,
    i: (index + 1).toString(), // 1-based index as string
    additionalParticipants: item.additionalParticipants
      .split(";")
      .map((participant) => lookupDisplayName(participant))
      .join(", "), // Convert to display names and join with commas
    isMisc: item.isMisc,
    hasBody: !item.isMisc && !!item.remarks,
    isDecision: item.needsDecision,
    hasAdditionalParticipants: !!item.additionalParticipants, // True if additionalParticipants is not empty
    remarks: item.remarks,
  }));

  // Construct the JSON object
  const resultJson = {
    meetingTitle: boardMeeting.title,
    meetingDate: formattedDate,
    meetingLocation: boardMeeting.location,
    topsCount: agendaItems.length,
    fixedParticipants: fixedParticipants,
    topsDetails: topsDetails,
  };

  return resultJson;
}

export async function createJsonForAgendaTemplate(
  boardMeeting: BoardMeeting,
  agendaItems: AgendaItem[],
  languageCode: string,
  includeRemarks: boolean,
  graphClient: Client
) {
  const formattedStartDate = boardMeeting.startTime
    .setLocale(languageCode.toLowerCase())
    .toLocaleString(DateTime.DATE_HUGE);
  const formattedStartTime = boardMeeting.startTime.toFormat("HH:mm");

  const currentFormattedDate =
    DateTime.now()
      .setZone("Europe/Berlin")
      .setLocale(languageCode.toLowerCase())
      .toLocaleString(DateTime.DATE_SHORT) +
    " " +
    DateTime.now().setZone("Europe/Berlin").toFormat("HH:mm"); // Add time in 24-hour format

  const formattedEndTime = calculateEndTime(
    boardMeeting.startTime,
    agendaItems
  ).formattedEndTime;

  const rawFixedParticipants = boardMeeting.fixedParticipants.split(";");
  const rawAllParticipants = new Set(rawFixedParticipants);
  agendaItems.forEach((item) => {
    const additionalParticipants = item.additionalParticipants.split(";");
    additionalParticipants.forEach((participant) => rawAllParticipants.add(participant));
  });
  const uniqueParticipants = Array.from(rawAllParticipants);
  const participantsInfo = await getDisplayNamesBatch(graphClient, uniqueParticipants);

  const lookupDisplayName = (upn: string): string => {
    const [lastName, firstName] = participantsInfo[upn]?.split(", ") || [];
    if (firstName?.trim() && lastName?.trim()) {
      return `${firstName} ${lastName}`;
    } else if (firstName?.trim()) {
      return firstName;
    } else if (lastName?.trim()) {
      return lastName;
    } else {
      return "Unknown Participant";
    }
  };

  // Populate fixedParticipants
  const fixedParticipants = rawFixedParticipants.map((participant) => ({
    fixedPerson: lookupDisplayName(participant),
    totalTops: agendaItems.length,
  }));

  // Populate topsDetails
  const topsDetails = agendaItems.map((item, index) => ({
    agendaTitle: item.title,
    i: (index + 1).toString(), // 1-based index as string
    additionalParticipants: item.additionalParticipants
      .split(";")
      .map((participant) => lookupDisplayName(participant))
      .join(", "), // Convert to display names and join with commas
    isMisc: item.isMisc,
    hasBody: !item.isMisc,
    isDecision: item.needsDecision,
    hasAdditionalParticipants: !!item.additionalParticipants, // True if additionalParticipants is not empty
    durationInMinutes: item.durationInMinutes,
    startTime: item.startTime?.toFormat("HH:mm") ?? "",
    includeRemarks: includeRemarks,
    ...(item.remarks ? { remarks: item.remarks } : {}), // Conditionally add remarks
  }));

  // Construct the JSON object
  const resultJson = {
    meetingTitle: boardMeeting.title,
    meetingDate: formattedStartDate,
    meetingTime: formattedStartTime,
    meetingLocation: boardMeeting.location,
    fixedParticipants: fixedParticipants,
    topsDetails: topsDetails,
    creationDate: currentFormattedDate,
    meetingEndTime: formattedEndTime,
  };

  return resultJson;
}

export function calculateEndTime(startTime: DateTime, agendaItems: AgendaItem[]) {
  // Sum up the total duration in minutes
  const totalDurationInMinutes = agendaItems.reduce(
    (total, item) => total + item.durationInMinutes,
    0
  );
  // Create a new Date instance for the end time
  const endTime = startTime.plus({ minutes: totalDurationInMinutes });
  const formattedEndTime = endTime
    .setLocale("de-DE")
    .toLocaleString(DateTime.TIME_24_SIMPLE);

  return { endTime, formattedEndTime };
}

export async function convertToPdf(
  sourceFolderDriveId: string,
  sourceFileId: string,
  targetFolderDriveId: string,
  targetFolderId: string,
  targetFileName: string,
  graphClient: Client
) {
  try {
    const response = await graphClient
      .api(
        `/sites/${config.sharePointWebsite}/drives/${sourceFolderDriveId}/items/${sourceFileId}/content?format=pdf`
      )
      .responseType(ResponseType.BLOB)
      .get();
    const pdfBlob = response;
    const uploadUrl = `/sites/${config.sharePointWebsite}/drives/${targetFolderDriveId}/items/${targetFolderId}:/${targetFileName}:/content`;
    const newFile = await graphClient.api(uploadUrl).put(pdfBlob);
    return newFile;
  } catch (error) {
    throw error;
  }
}

export type BoardMeeting = {
  id?: number;
  startTime: DateTime;
  title: string;
  fixedParticipants: string;
  remarks: string;
  location: string;
  meetingLink?: string;
  fileLocationId?: string;
  eventId?: string;
  timeZone: string;
  room?: string;
};

export type AgendaItem = {
  id: number;
  durationInMinutes: number;
  title: string;
  additionalParticipants: string;
  fileLocationId: string;
  protocolLocationId: string;
  orderIndex: number;
  isMisc: boolean;
  needsDecision: boolean;
  boardMeeting: number;
  startTime?: DateTime;
  isNew?: boolean;
  eventId?: string;
  remarks?: string;
};

// Helper to split an array into chunks
function chunkArray<T>(array: T[], chunkSize: number): T[][] {
  const chunks: T[][] = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    chunks.push(array.slice(i, i + chunkSize));
  }
  return chunks;
}

async function getDisplayNamesBatch(
  graphClient: Client,
  upns: string[]
): Promise<Record<string, string>> {
  const displayNames: Record<string, string> = {};
  const batches = chunkArray(upns, 20); // Split UPNs into chunks of 20

  for (const batch of batches) {
    const batchRequestBody: { requests: Array<{ id: string; method: string; url: string }> } = { requests: [] };

    // Create a batch request for this chunk
    batch.forEach((upn, index) => {
      batchRequestBody.requests.push({
        id: `${index}`,
        method: "GET",
        url: `/users/${upn}?$select=displayName`,
      });
    });

    try {
      // Send the batch request
      const batchResponse = await graphClient.api("/$batch").post(batchRequestBody);

      // Process responses
      for (const response of batchResponse.responses) {
        if (response.status === 200) {
          const { displayName } = response.body;
          const upn = batch[parseInt(response.id)]; // Map response back to UPN
          displayNames[upn] = displayName;
        }
        if (response.status === 404) {
          // external users
          const upn = batch[parseInt(response.id)]; // Map response back to UPN
          displayNames[upn] = upn;
        }
      }
    } catch (error) {
      console.error("Error fetching display names for a batch:", error);
    }
  }

  // Process custom userMappings
  // Define your SQL query and parameters
  const query = "SELECT * FROM UserMappings ORDER BY DisplayName ASC";
  const userMappings = await DatabaseHelper.executeQuery(query);
  userMappings.forEach((mapping: { Upn: string; DisplayName: string }) => {
    if (displayNames[mapping.Upn]) {
      displayNames[mapping.Upn] = mapping.DisplayName;
    }
  });

  return displayNames;
}

export function toBerlinTimeISOString(date: DateTime): string {
  // Ensure the DateTime object is in the "Europe/Berlin" time zone
  const berlinTime = date.setZone("Europe/Berlin");

  // Format it as an ISO 8601 string
  return berlinTime.toISO({ suppressMilliseconds: true }) ?? ""; // "2025-01-28T07:00:00+01:00"
}

export async function getMailBody(eventId: string, graphClient: Client): Promise<string> {
  const url = `https://graph.microsoft.com/v1.0/users/${config.eventMailbox}/events/${eventId}?$select=body`;

  try {
    const response = await graphClient.api(url).get();
    let htmlContent = response.body.content;

    // Remove \r\n and unnecessary spaces
    htmlContent = htmlContent.replace(/\r\n/g, " ").replace(/\s+/g, " ").trim();
    return htmlContent;
  } catch (error) {
    console.error("Error fetching meeting footer:", error);
    throw error;
  }
}