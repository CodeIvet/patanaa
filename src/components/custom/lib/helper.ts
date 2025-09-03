import { TeamsUserCredential } from "@microsoft/teamsfx";
import * as axios from "axios";
import { Method } from "axios";
import { defaultDatePickerStrings } from "@fluentui/react-datepicker-compat";
import type { CalendarStrings } from "@fluentui/react-datepicker-compat";
import { DateTime } from "luxon";
import config from "./config"; // adjust path if needed

export type BoardMeeting = {
  id: number;
  startTime: DateTime;
  title: string;
  fixedParticipants: string;
  remarks: string;
  location: string;
  fileLocationId: string;
  meetingLink?: string;
  eventId?: string;
  timeZone: string;
  room?: string;
};

export type AgendaItem = {
  id: number;
  durationInMinutes: number;
  title: string;
  additionalParticipants: string;
  fileLocationId?: number;
  protocolLocationId?: number;
  orderIndex?: number; // no need to update manually
  isMisc: boolean;
  needsDecision: boolean;
  boardMeeting?: number; // no need to update manually
  startTime?: DateTime;
  isNew?: boolean;
  eventId?: string;
  remarks?: string;
};

export const timeZoneOptions = [
  { text: "Berlin", value: "Europe/Berlin" },
  { text: "New York", value: "America/New_York" },
  { text: "Tel Aviv", value: "Asia/Tel_Aviv" },
];

export const truncateText = (text: string) => {
  return text.length > 70 ? `${text.substring(0, 70)}...` : text;
};

export function dateCustomFormatting(date: Date): string {
  const padStart = (value: number): string => value.toString().padStart(2, "0");
  return `${padStart(date.getDate())}.${padStart(
    date.getMonth() + 1
  )}.${date.getFullYear()} 
          ${padStart(date.getHours())}:${padStart(date.getMinutes())}`;
}

export async function callBackend(
  functionName: string,
  method: Method,
  teamsUserCredential?: TeamsUserCredential,
  body?: any,
  queryParams?: string[]
): Promise<any> {
  try {
    let query: string = "";
    if (!!queryParams) {
      query = "?" + queryParams.join("&");
    }

    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }

    // const cred = teamsfx.getCredential();
    const token = await teamsUserCredential.getToken(""); // Get SSO token for the user
    const apiEndpoint = config.apiEndpoint; // ✅ comes from import.meta.env
    const response = await axios.default.request({
      url: apiEndpoint + "/api/" + functionName + query,
      method: method,
      data: body,
      headers: {
        authorization: "Bearer " + token?.token || "",
      },
    });
    return response.data;
  } catch (err: unknown) {
    throw err;
  }
}

export type User = {
  displayName: string;
  emailAddress: string;
  avatarUrl: string;
};

export const localizedCalendarStrings: CalendarStrings = {
  ...defaultDatePickerStrings,
  days: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"],
  shortDays: ["S", "M", "D", "M", "D", "F", "S"],
  months: [
    "Januar",
    "Februar",
    "März",
    "April",
    "Mai",
    "Juni",
    "Juli",
    "August",
    "September",
    "Oktober",
    "November",
    "Dezember",
  ],
  shortMonths: [
    "Jan",
    "Feb",
    "Mär",
    "Apr",
    "Mai",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Okt",
    "Nov",
    "Dez",
  ],
  goToToday: "Heute",
};

export const onFormatDate = (date?: Date) => {
  return !date
    ? ""
    : `${date.getDate()}. ${
        localizedCalendarStrings.months[date.getMonth()]
      } ${date.getFullYear()}`;
};

// Function to add the timestamp property to each item
export const calculateTimestamps = (
  startTime: DateTime,
  items: AgendaItem[]
): AgendaItem[] => {
  let cumulativeDuration = 0;

  return items.map((item: AgendaItem): AgendaItem => {
    // Calculate the timestamp for the current item
    const itemTimestamp = startTime.plus({ minutes: cumulativeDuration });
    // Update cumulative duration for the next item
    cumulativeDuration += item.durationInMinutes;

    // Return the item with the added timestamp property
    return {
      ...item,
      startTime: itemTimestamp,
    };
  });
};

// Function to split an array into chunks of size `size`
// Function to split an array into chunks of size `size`
export const chunkArray = <T>(arr: T[], size: number): T[][] => {
  return arr.reduce<T[][]>((acc, _, i) => {
    if (i % size === 0) acc.push(arr.slice(i, i + size));
    return acc;
  }, []);
};

export async function isAttendeesMatch(
  calendarData: any,
  dbparticipantList: string,
  teamsUserCredential: TeamsUserCredential
): Promise<boolean> {
  try {
    // Extract email addresses from the attendees list
    const calendarAttendeeEmails =
      calendarData.attendees?.map((attendee: any) => attendee.emailAddress.address) || [];

    // Get primary email address for each upn in participants list
    const dbUpnArray = dbparticipantList
      .split(";")
      .map((participant) => participant.trim())
      .filter((participant) => participant.length > 0);

    if (dbUpnArray.length === 0) {
      if (calendarAttendeeEmails.length === 0) {
        return true;
      } else {
        return false;
      }
    } else {
      const participantsData = await callBackend(
        "getUserProfiles",
        "POST",
        teamsUserCredential,
        {
          upns: dbUpnArray, // Sending the list of UPNs
        }
      );

      // Remove duplicates from the list of participants
      const dbAttendees = Array.from(
        new Set(participantsData.map((user: { mail: string }) => user.mail))
      );
      // Sort both arrays for comparison (optional but helps ensure order doesn't matter)
      calendarAttendeeEmails.sort();
      dbAttendees.sort();
      // Compare arrays (ensuring every element matches)
      return JSON.stringify(calendarAttendeeEmails) === JSON.stringify(dbAttendees);
    }
  } catch (error) {
    return false;
  }
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

export function areDatesEqual(
  graphCalendarDate: { dateTime: string; timeZone: string },
  tsDate: DateTime
): boolean {
  // Convert the event's start time to a Luxon DateTime object with the correct time zone
  const eventStartDateTime: DateTime = DateTime.fromISO(graphCalendarDate.dateTime, {
    zone: graphCalendarDate.timeZone,
  });
  return eventStartDateTime.toUTC().equals(tsDate.toUTC());
}