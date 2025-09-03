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

sharePointWebsite: process.env.SP_WEBSITE! || "dummy", 
sharePointMeetingsDriveId: process.env.SP_MEETINGS_DRIVE_ID! || "dummy", 
sharePointMeetingFolderId: process.env.SP_MEETING_FOLDER_ID! || "dummy", 
sharePointUnassignedTopsFolderId: process.env.SP_UNASSIGNED_FOLDER_ID! || "dummy", 
eventMailbox: process.env.EVENT_MAILBOX! || "dummy@domain.com", 
applicationInsightsConnectionString: process.env.APPLICATION_INSIGHTS_CONNECTION_STRING || "",

}; 

export default config; 

 
 

 