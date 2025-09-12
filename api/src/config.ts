import dotenv from "dotenv";
dotenv.config({ path: ".env.local" });

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing environment variable: ${name}`);
  }
  return value;
}

const config = { 
authorityHost: requireEnv("M365_AUTHORITY_HOST"),
tenantId: process.env.M365_TENANT_ID!, 
clientId: process.env.M365_CLIENT_ID!, 
clientSecret: process.env.M365_CLIENT_SECRET!, 
dbUser: process.env.DB_USER!, 
dbPassword: process.env.DB_PASSWORD!, 
dbServer: process.env.DB_SERVER!, 
dbName: process.env.DB_NAME!,



 
 

// Add stubs for now: 

  sharePointWebsite: process.env.SharePointWebsite ?? "",
  sharePointMeetingsDriveId: process.env.SharePointMeetingsDriveId ?? "",
  sharePointMeetingFolderId: process.env.SharePointMeetingFolderId ?? "",
  sharePointUnassignedTopsFolderId: process.env.SharePointUnassignedTopsFolderId ?? "",
  applicationInsightsConnectionString: process.env.APPLICATIONINSIGHTS_CONNECTION_STRING ?? "",
  protocolTemplateFileIdDe: process.env.PROTOCOL_TEMPLATE_FILE_ID_DE ?? "",
  agendaPdfTemplateFileIdDe: process.env.AGENDA_PDF_TEMPLATE_FILE_ID_DE ?? "",
  protocolTemplateFileIdEn: process.env.PROTOCOL_TEMPLATE_FILE_ID_EN ?? "",
  agendaPdfTemplateFileIdEn: process.env.AGENDA_PDF_TEMPLATE_FILE_ID_EN ?? "",
  assetsDriveId: process.env.AssetsDriveId ?? "",
  eventMailbox: process.env.EVENT_MAILBOX ?? "",

}; 

export default config; 

 
 

 