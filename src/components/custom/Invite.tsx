import {
  Button,
  Field,
  makeStyles,
  ProgressBar,
  Spinner,
  Text,
  Table,
  TableHeader,
  TableRow,
  TableCell,
  TableBody,
  TableHeaderCell,
  DialogActions,
  tokens,
} from "@fluentui/react-components";
import { useContext, useState, useEffect, useRef } from "react";
import {
  AgendaItem,
  areDatesEqual,
  BoardMeeting,
  calculateEndTime,
  calculateTimestamps,
  callBackend,
  isAttendeesMatch,
} from "./lib/helper";
import { TeamsFxContext } from "../Context";
import { DateTime } from "luxon";

const useStyles = makeStyles({
  columnStyle: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  tableContainer: {
    marginTop: "20px",
    width: "100%",
    fontSize: "14px", // Force table to use 14px
  },
  tableCell: {
    fontSize: "14px", // Ensure all cells have 14px font size
  },
  tableHeaderCell: {
    fontSize: "14px",
    fontWeight: "bold",
  },
  buttonDanger: {
    backgroundColor: tokens.colorStatusDangerBackground3,
    "&:hover": {
      backgroundColor: tokens.colorStatusDangerBackground3Pressed,
    },
  },
});

type InviteStatusType = (typeof InviteStatus)[keyof typeof InviteStatus];

interface InviteItem {
  id: number;
  type: string;
  title: string;
  status: InviteStatusType;
  eventId: string;
  participants: string;
  startTime?: DateTime;
  endTime?: DateTime;
  location?: string;
  webLink?: string;
  room?: string;
}

const InviteStatus = {
  1: {
    message: "Einladung erstellt",
    actionLabel: "Keine Aktion notwendig",
    isActionEnabled: false,
  },
  2: {
    message: "Einladung fehlt",
    actionLabel: "Einladung erstellen",
    isActionEnabled: true,
  },
  3: {
    message: "Einladung erstellt, noch nicht versendet",
    actionLabel: "Öffnen zum Senden",
    isActionEnabled: true,
  },
  4: {
    message: "Einladung erstellt und versendet",
    actionLabel: "Keine Aktion notwendig",
    isActionEnabled: false,
  },
  5: {
    message: "Einladung veraltet, noch nicht versendet",
    actionLabel: "Einladung aktualisieren",
    isActionEnabled: true,
  },
  6: {
    message: "Einladung veraltet aber schon versendet",
    actionLabel: "Einladung aktualisieren und senden",
    isActionEnabled: true,
  },
  7: {
    message: "Einladungstatus unbekannt. Bitte neu laden.",
    actionLabel: "Bitte App neu laden",
    isActionEnabled: true,
  },
} as const;

export function Invite(props: {
  isSmallScreen: boolean;
  currentMeetingItem?: BoardMeeting;
  setOpenDialogInvites: (open: boolean) => void;
  setPreventDialogClose: (open: boolean) => void;
  setShouldReloadBoardMeetings: (open: boolean) => void;
}) {
  const [loading, setLoading] = useState(false);
  const [savingProgress, setSavingProgress] = useState("");
  const savingProgressPercent = useRef(0.0);
  const [error, setError] = useState<string>("");
  const [successMessage, setSuccessMessage] = useState<string>("");
  const [inviteData, setInviteData] = useState<InviteItem[]>([]);
  const [agendaItems, setAgendaItems] = useState<AgendaItem[]>([]);
  const [isOnlineMeetingLinkAvailable, setIsOnlineMeetingLinkAvailable] = useState(false);
  const [showAutomationConfirmation, setShowAutomationConfirmation] = useState(false);
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const styles = useStyles();

  const fetchAgendaItems = async (omitLoadingSpinner: boolean = false) => {
    let returnInviteData: InviteItem[] = [];
    if (!omitLoadingSpinner) {
      setLoading(true);
    }
    try {
      const agendaItems = await callBackend(
        "getAgendaItems",
        "GET",
        teamsUserCredential,
        undefined,
        ["boardmeeting=" + props.currentMeetingItem?.id]
      );
      const agendaItemsWithStartDates = calculateTimestamps(
        props.currentMeetingItem?.startTime ?? DateTime.now(),
        JSON.parse(agendaItems) as AgendaItem[]
      );

      setAgendaItems(agendaItemsWithStartDates);

      // Process invite data after agenda items are set
      if (props.currentMeetingItem && agendaItemsWithStartDates.length > 0) {
        returnInviteData = await processInviteData(
          props.currentMeetingItem,
          agendaItemsWithStartDates
        );
      }
    } catch (error) {
      setError(JSON.stringify(error));
    } finally {
      if (!omitLoadingSpinner) {
        setLoading(false);
      }
      return returnInviteData;
    }
  };

  // Function to process invite data and fetch calendar items
  const processInviteData = async (
    meetingItem: BoardMeeting,
    agendaItems: AgendaItem[]
  ) => {
    setIsOnlineMeetingLinkAvailable(meetingItem.eventId ? true : false);

    const mainMeetingInvite: InviteItem = {
      id: meetingItem.id ?? 0,
      type: "Gesamtsitzung",
      title: meetingItem.title ?? "",
      status: meetingItem.eventId ? InviteStatus[1] : InviteStatus[2],
      eventId: meetingItem.eventId ?? "",
      participants: meetingItem.fixedParticipants,
      startTime: meetingItem.startTime,
      endTime: calculateEndTime(meetingItem.startTime, agendaItems).endTime,
      location: meetingItem.location,
      room: meetingItem.room,
    };

    const newInviteData: InviteItem[] = [
      mainMeetingInvite,
      ...agendaItems.map((item) => ({
        id: item.id,
        type: "TOP",
        title: item.title,
        status: item.eventId ? InviteStatus[1] : InviteStatus[2],
        eventId: item.eventId ?? "undefined",
        participants: item.additionalParticipants,
        startTime: item.startTime,
        endTime: item.startTime!.plus({ minutes: item.durationInMinutes }),
        location: mainMeetingInvite.location,
        room: mainMeetingInvite.room,
      })),
    ];

    // Fetch calendar items in parallel
    const updatedInviteData = await Promise.all(
      newInviteData.map(async (element: InviteItem) => {
        if (element.status && element.status != InviteStatus[2]) {
          try {
            const calendarItem = await callBackend(
              "getCalendarItem",
              "GET",
              teamsUserCredential,
              undefined,
              ["eventId=" + element.eventId]
            );

            let newStatus: InviteStatusType;

            if (calendarItem) {
              const dbCombinedParticipants =
                `${props.currentMeetingItem?.fixedParticipants};${element.participants}`.replace(
                  /;;/g,
                  ";"
                );
              if (
                (calendarItem.subject === element.title ||
                  calendarItem.subject ===
                    props.currentMeetingItem?.title + " - " + element.title) &&
                (element.type === "Gesamtsitzung" ||
                  (await isAttendeesMatch(
                    calendarItem,
                    dbCombinedParticipants,
                    teamsUserCredential!
                  ))) &&
                areDatesEqual(calendarItem.start, element.startTime!) &&
                areDatesEqual(calendarItem.end, element.endTime!) &&
                (calendarItem.location.displayName === element.room ||
                  element.type === "Gesamtsitzung")
              ) {
                if (calendarItem.isDraft) {
                  newStatus = InviteStatus[3];
                } else {
                  newStatus = InviteStatus[4];
                }
              } else {
                if (calendarItem.isDraft) {
                  newStatus = InviteStatus[5];
                } else {
                  newStatus = InviteStatus[6];
                }
              }
              // Save calendar item deeplink
              element.webLink = calendarItem.webLink;
            } else {
              newStatus = InviteStatus[2];
            }

            return { ...element, status: newStatus };
          } catch (error) {
            console.error("Error fetching calendar item:", error);
          }
        }
        return element;
      })
    );

    setInviteData(updatedInviteData);
    return updatedInviteData;
  };

  // Fetch agenda items on mount
  useEffect(() => {
    fetchAgendaItems();
  }, []);

  const handleCreate = async (
    invite: InviteItem,
    omitProgressInfo: boolean = false
  ): Promise<InviteItem[]> => {
    props.setPreventDialogClose(true);
    setError("");
    let updatedInviteData: InviteItem[] = [];

    let progressInfo = "";
    const progressSteps = 2;
    const progressIncrement = 1 / progressSteps;

    if (!omitProgressInfo) {
      setLoading(true);
      progressInfo = "Fenster bitte nicht schließen! Aktualisiere Einladung...";
      savingProgressPercent.current = progressIncrement;
      setSavingProgress(progressInfo);
    }

    try {
      const isGesamtsitzung = invite.type === "Gesamtsitzung";
      const isCreateNew = invite.status === InviteStatus[2];
      const isReadyToBeSent = invite.status === InviteStatus[3];
      const isAlreadySent = invite.status === InviteStatus[6];
      const isUpdate =
        invite.status === InviteStatus[5] || invite.status === InviteStatus[6];

      // case 1: if Gesamtsitzung and Create/Update Invite
      if (isGesamtsitzung && (isCreateNew || isUpdate)) {
        const postCalendarBody = {
          ...props.currentMeetingItem,
          isCreateAsNew: isCreateNew,
          isAlreadySent: isAlreadySent,
          agendaItems: agendaItems.map((item) => ({
            durationInMinutes: item.durationInMinutes,
          })),
        };

        const calendarItem = await callBackend(
          "createUpdateBoardMeetingCalendarItem",
          "POST",
          teamsUserCredential,
          JSON.stringify(postCalendarBody),
          undefined
        );
        props.currentMeetingItem!.eventId = calendarItem.eventId;
        props.currentMeetingItem!.meetingLink = calendarItem.joinUrl;

        if (isCreateNew) {
          // Update eventid in database if this a new calendar item
          const postDatabaseBody = {
            boardmeeting: {
              ...props.currentMeetingItem,
              eventId: calendarItem.eventId,
              meetingLink: calendarItem.joinUrl,
            },
            ensureFileStructure: false,
          };

          const dbResult = await callBackend(
            "updateBoardMeeting",
            "POST",
            teamsUserCredential,
            JSON.stringify(postDatabaseBody),
            undefined
          );
        }
      } else if (!isGesamtsitzung && (isCreateNew || isUpdate)) {
        // TOPS....
        const postCalendarBody = {
          mainMeeting: { ...props.currentMeetingItem },
          isCreateAsNew: isCreateNew,
          isAlreadySent: isAlreadySent,
          timeZone: invite.startTime!.zone.name,
          ...invite,
        };

        const calendarItemId = await callBackend(
          "createUpdateAgendaItemCalendarItem",
          "POST",
          teamsUserCredential,
          JSON.stringify(postCalendarBody),
          undefined
        );

        if (isCreateNew || isUpdate) {
          // Update eventid in database if this a new calendar item
          const postDatabaseBody = {
            agendaItemId: invite.id,
            eventId: calendarItemId,
          };

          const dbResult = await callBackend(
            "updateAgendaItem",
            "POST",
            teamsUserCredential,
            JSON.stringify(postDatabaseBody),
            undefined
          );
        }
      } else if (isReadyToBeSent && invite.webLink) {
        // Open the link in a new tab
        window.open(invite.webLink, "_blank", "noopener,noreferrer");
        props.setShouldReloadBoardMeetings(true);
        setTimeout(() => {
          props.setOpenDialogInvites(false);
          setLoading(false);
          props.setPreventDialogClose(false);
        }, 2000);
      }

      // reload invite data
      updatedInviteData = await fetchAgendaItems(omitProgressInfo);

      if (!omitProgressInfo) {
        progressInfo = "Fertig!";
        savingProgressPercent.current += progressIncrement;
        setSavingProgress(progressInfo);
      }
    } catch (error) {
      setError(
        "Fehler beim Verwalten des Kalendereintrags. Bitte wenden Sie sich an den Administrator. " +
          error
      );
    } finally {
      if (!omitProgressInfo) {
        setLoading(false);
      }
      props.setPreventDialogClose(false);
      return updatedInviteData;
    }
  };

  const handleAutomationConfirm = async () => {
    setLoading(true);
    setError("");
    props.setPreventDialogClose(true);
    let progressInfo = "";
    const progressSteps = inviteData.length + 1;
    const progressIncrement = 1 / progressSteps;

    progressInfo = "Fenster bitte nicht schließen! Bearbeite Gesamtsitzung... ";
    savingProgressPercent.current = progressIncrement;
    setSavingProgress(progressInfo);

    try {
      let newInviteData: InviteItem[] = inviteData;
      // process Gesamtsitzung
      let boardMeetingInvite = newInviteData.find(
        (invite) => invite.type === "Gesamtsitzung"
      );
      while (boardMeetingInvite && boardMeetingInvite.status !== InviteStatus[4]) {
        newInviteData = await handleCreate(boardMeetingInvite, true);
        boardMeetingInvite = newInviteData.find(
          (invite) => invite.type === "Gesamtsitzung"
        );
      }

      // process TOPs
      const tops = newInviteData
        .filter((invite) => invite.type === "TOP")
        .sort((a, b) => a.id - b.id);

      let count = 1;

      for (const topInvite of tops) {
        progressInfo += "Bearbeite TOP " + count++ + "... ";
        setSavingProgress(progressInfo);
        let newTopInvite = topInvite;
        while (newTopInvite.status !== InviteStatus[4]) {
          // if calendar item already exists, we need to ensure that it gets actually sent. this can be done in status 6
          if (
            newTopInvite.status === InviteStatus[3] ||
            newTopInvite.status === InviteStatus[5]
          ) {
            newTopInvite.status = InviteStatus[6];
          }
          newInviteData = await handleCreate(newTopInvite, true);
          newTopInvite = newInviteData.find((invite) => invite.id === topInvite.id)!;
        }
      }

      progressInfo = "Fertig!";
      savingProgressPercent.current += progressIncrement;
      setSavingProgress(progressInfo);
    } catch (error) {
      setError(
        "Fehler beim Verwalten des Kalendereintrags. Bitte wenden Sie sich an den Administrator. " +
          error
      );
    } finally {
      setLoading(false);
      setShowAutomationConfirmation(false);
      props.setPreventDialogClose(false);
    }
  };

  return (
    <div className={styles.columnStyle}>
      <div>
        <Field
          validationMessage={savingProgress}
          validationState={savingProgress !== "" ? "error" : "none"}
        >
          <ProgressBar value={savingProgressPercent.current} color="warning" />
        </Field>
      </div>

      {loading && <Spinner />}
      {showAutomationConfirmation ? (
        <>
          <p>
            Alle Einladungen werden jetzt <strong>automatisch</strong> erstellt /
            aktualisiert und versendet.
            <br />
            Sind Sie sich wirklich sicher?
          </p>
          <DialogActions>
            <Button
              appearance="primary"
              onClick={() => setShowAutomationConfirmation(false)}
              disabled={loading}
            >
              Abbrechen
            </Button>
            <Button
              appearance="primary"
              className={styles.buttonDanger}
              onClick={() => handleAutomationConfirm()}
              disabled={loading}
            >
              Ja, fortfahren
            </Button>
          </DialogActions>
        </>
      ) : (
        <>
          <div className={styles.columnStyle}>
            <Text>
              Hier kannst du die Einladungen für die Sitzung und die einzelnen TOPs
              verwalten. Es werden nur TOPs mit zusätzlichen Teilnehmern (Gästen)
              angezeigt.
            </Text>

            {!loading && successMessage && (
              <div style={{ color: "green" }}>{successMessage}</div>
            )}
            {!loading && error && <div style={{ color: "red" }}>{error}</div>}

            {/* Table Section */}
            <div className={styles.tableContainer}>
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHeaderCell
                      className={styles.tableHeaderCell}
                      style={{ width: "120px" }}
                    >
                      Typ
                    </TableHeaderCell>
                    <TableHeaderCell
                      className={styles.tableHeaderCell}
                      style={{ flexGrow: 1 }}
                    >
                      Titel
                    </TableHeaderCell>
                    <TableHeaderCell
                      className={styles.tableHeaderCell}
                      style={{ width: "150px" }}
                    >
                      Status
                    </TableHeaderCell>
                    <TableHeaderCell
                      className={styles.tableHeaderCell}
                      style={{ flexGrow: 1 }}
                    >
                      Aktion
                    </TableHeaderCell>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {inviteData.map((invite, index) => (
                    <TableRow key={index}>
                      <TableCell className={styles.tableCell} style={{ width: "120px" }}>
                        {invite.type}
                      </TableCell>
                      <TableCell className={styles.tableCell} style={{ flexGrow: 1 }}>
                        {invite.title}
                      </TableCell>
                      <TableCell className={styles.tableCell} style={{ width: "150px" }}>
                        {invite.status.message}
                      </TableCell>
                      <TableCell className={styles.tableCell} style={{ flexGrow: 1 }}>
                        <Button
                          appearance="primary"
                          onClick={() => handleCreate(invite)}
                          disabled={
                            loading ||
                            !invite.status.isActionEnabled ||
                            (invite.type === "TOP" && !isOnlineMeetingLinkAvailable)
                          }
                        >
                          {invite.status.actionLabel}
                        </Button>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <Button
                className={styles.buttonDanger}
                appearance="primary"
                onClick={() => setShowAutomationConfirmation(true)}
                disabled={
                  loading || inviteData.every((item) => item.status === InviteStatus[4])
                }
              >
                Alle aktualisieren und senden
              </Button>
            </div>
          </div>
        </>
      )}
    </div>
  );
}