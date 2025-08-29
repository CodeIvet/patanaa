import { defaultDatePickerStrings } from "@fluentui/react-datepicker-compat";
import type { CalendarStrings } from "@fluentui/react-datepicker-compat";
import { DateTime } from "luxon";

// Types
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
  orderIndex?: number;
  isMisc: boolean;
  needsDecision: boolean;
  boardMeeting?: number;
  startTime?: DateTime;
  isNew?: boolean;
  eventId?: string;
  remarks?: string;
};

export type User = {
  displayName: string;
  emailAddress: string;
  avatarUrl: string;
};

// Time zones
export const timeZoneOptions = [
  { text: "Berlin", value: "Europe/Berlin" },
  { text: "New York", value: "America/New_York" },
  { text: "Tel Aviv", value: "Asia/Tel_Aviv" },
];

// Text utils
export const truncateText = (text: string) => {
  return text.length > 70 ? `${text.substring(0, 70)}...` : text;
};

// Date formatting
export function dateCustomFormatting(date: Date): string {
  const padStart = (value: number): string => value.toString().padStart(2, "0");
  return `${padStart(date.getDate())}.${padStart(
    date.getMonth() + 1
  )}.${date.getFullYear()} ${padStart(date.getHours())}:${padStart(date.getMinutes())}`;
}

// Calendar strings (German)
export const localizedCalendarStrings: CalendarStrings = {
  ...defaultDatePickerStrings,
  days: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"],
  shortDays: ["S", "M", "D", "M", "D", "F", "S"],
  months: [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
  ],
  shortMonths: [
    "Jan", "Feb", "Mär", "Apr", "Mai", "Jun",
    "Jul", "Aug", "Sep", "Okt", "Nov", "Dez",
  ],
  goToToday: "Heute",
};

export const onFormatDate = (date?: Date) =>
  !date ? "" : `${date.getDate()}. ${localizedCalendarStrings.months[date.getMonth()]} ${date.getFullYear()}`;

// Add calculated startTime to each agenda item
export const calculateTimestamps = (
  startTime: DateTime,
  items: AgendaItem[]
): AgendaItem[] => {
  let cumulativeDuration = 0;
  return items.map((item) => {
    const itemTimestamp = startTime.plus({ minutes: cumulativeDuration });
    cumulativeDuration += item.durationInMinutes;
    return { ...item, startTime: itemTimestamp };
  });
};

// Array chunking
export const chunkArray = <T>(arr: T[], size: number): T[][] => {
  return arr.reduce<T[][]>((acc, _, i) => {
    if (i % size === 0) acc.push(arr.slice(i, i + size));
    return acc;
  }, []);
};

// Compare attendee lists using email strings
export function isAttendeesMatchLocal(
  calendarAttendeeEmails: string[],
  dbParticipantList: string
): boolean {
  const dbUpnArray = dbParticipantList
    .split(";")
    .map((p) => p.trim())
    .filter(Boolean);

  const cleanEmails = (arr: string[]) =>
    Array.from(new Set(arr.map((e) => e.toLowerCase()))).sort();

  return JSON.stringify(cleanEmails(calendarAttendeeEmails)) ===
         JSON.stringify(cleanEmails(dbUpnArray));
}

// End time calculation
export function calculateEndTime(startTime: DateTime, agendaItems: AgendaItem[]) {
  const totalDurationInMinutes = agendaItems.reduce(
    (total, item) => total + item.durationInMinutes,
    0
  );
  const endTime = startTime.plus({ minutes: totalDurationInMinutes });
  const formattedEndTime = endTime.setLocale("de-DE").toLocaleString(DateTime.TIME_24_SIMPLE);
  return { endTime, formattedEndTime };
}

// Compare two dates from Graph and Luxon
export function areDatesEqual(
  graphCalendarDate: { dateTime: string; timeZone: string },
  tsDate: DateTime
): boolean {
  const eventStartDateTime = DateTime.fromISO(graphCalendarDate.dateTime, {
    zone: graphCalendarDate.timeZone,
  });
  return eventStartDateTime.toUTC().equals(tsDate.toUTC());
}
