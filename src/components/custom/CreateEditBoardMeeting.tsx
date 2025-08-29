import {
  Button,
  Dropdown,
  Field,
  Input,
  makeStyles,
  ProgressBar,
  Spinner,
  Textarea,
  tokens,
  Option,
  DialogActions,
} from "@fluentui/react-components";
import React, { useContext } from "react";
import { DatePicker, DatePickerProps } from "@fluentui/react-datepicker-compat";
import { TimePicker, TimePickerProps } from "@fluentui/react-timepicker-compat";
import { PeoplePicker } from "@microsoft/mgt-react";
import { useState } from "react";
import { TeamsFxContext } from "../Context";
import {
  BoardMeeting,
  callBackend,
  localizedCalendarStrings,
  onFormatDate,
  timeZoneOptions,
} from "./lib/helper";
import { DateTime } from "luxon";

const useStyles = makeStyles({
  root: {
    display: "grid",
    columnGap: "20px",
    gridTemplateColumns: "repeat(3, 1fr)",
  },
  columnStyle: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  flexEnd: {
    display: "flex",
    justifyContent: "flex-end",
    gap: "10px",
    paddingTop: "20px",
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "20px", // Optional: Add spacing from the content above
  },
  peoplePicker: {
    "--people-picker-input-background": tokens.colorNeutralBackground1,
    width: "100%", // Ensures the PeoplePicker expands within its container
  },
  dropdown: {
    width: "100%", // Ensures dropdown does not overflow
    minWidth: "0px !important", // Forces Fluent UI to ignore default min-width
  },
  rowContainer: {
    display: "grid",
    gridTemplateColumns: "175px 1fr", // First column fixed, second column takes remaining space
    columnGap: "10px",
    width: "100%", // Ensure it takes full width inside columnStyle
    alignItems: "center", // Align vertically
  },
  buttonDanger: {
    backgroundColor: tokens.colorStatusDangerBackground3,
    "&:hover": {
      backgroundColor: tokens.colorStatusDangerBackground3Pressed,
    },
  },
});

export function CreateEditBoardMeeting(props: {
  isSmallScreen: boolean;
  isEditScreen: boolean;
  currentMeetingItem?: BoardMeeting;
  setOpenDialogCreateEditBoardMeeting: (open: boolean) => void;
  setShouldReloadBoardMeetings: (open: boolean) => void;
  setPreventDialogClose: (open: boolean) => void;
}) {
  const [loading, setLoading] = useState(false);
  const [savingProgress, setSavingProgress] = useState("");
  const savingProgressPercent = React.useRef(0.0);
  const [title, setTitle] = useState(props.currentMeetingItem?.title ?? "");
  const [location, setLocation] = useState(props.currentMeetingItem?.location ?? "");
  const [remarks, setRemarks] = useState(props.currentMeetingItem?.remarks ?? "");
  const [titleMissing, setTitleMissing] = useState(false);
  const [dateMissing, setDateMissing] = useState(false);
  const [error, setError] = useState<string>("");
  const [showDeleteConfirmation, setShowDeleteConfirmation] = useState(false);
  const [selectedDate, setSelectedDate] = React.useState<DateTime | null | undefined>(
    props.currentMeetingItem?.startTime ? props.currentMeetingItem.startTime : null
  );
  const [selectedParticipants, setSelectedParticipants] = useState<string>(
    props.currentMeetingItem?.fixedParticipants ?? ""
  );
  const [selectedTime, setSelectedTime] = React.useState<DateTime | null>(
    props.currentMeetingItem?.startTime ? props.currentMeetingItem.startTime : null
  );
  const [timePickerValue, setTimePickerValue] = React.useState<string>(
    props.currentMeetingItem?.startTime
      ? props.currentMeetingItem.startTime.toFormat("HH:mm")
      : ""
  );
  const [timeZone, setTimeZone] = React.useState(
    props.currentMeetingItem?.timeZone ?? "Europe/Berlin"
  );
  const [defaultParticipantGroups, setDefaultParticipantGroups]: any = React.useState([]);
  const [defaultRooms, setDefaultRooms]: any = React.useState([]);
  const [selectedGroupParticipantOptions, setSelectedGroupParticipantOptions] =
    React.useState<string[]>([]);
  const [selectedRoom, setSelectedRoom] = React.useState<string>(
    props.currentMeetingItem?.room ?? ""
  );

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const styles = useStyles();

  const onSelectDate: DatePickerProps["onSelectDate"] = (date) => {
    if (!date) {
      setSelectedDate(null);
      return;
    }

    const dtSelected = DateTime.fromJSDate(date);
    setSelectedDate(dtSelected);

    setSelectedTime(
      dtSelected.set({
        hour: selectedTime?.hour ?? 0,
        minute: selectedTime?.minute ?? 0,
        second: selectedTime?.second ?? 0,
        millisecond: selectedTime?.millisecond ?? 0,
      })
    );
  };

  const onTimeChange: TimePickerProps["onTimeChange"] = (_ev, data) => {
    if (!data.selectedTime) return;

    // Store time as-is, without conversion
    setSelectedTime(
      DateTime.fromObject({
        year: data.selectedTime.getFullYear(),
        month: data.selectedTime.getMonth() + 1,
        day: data.selectedTime.getDate(),
        hour: data.selectedTime.getHours(),
        minute: data.selectedTime.getMinutes(),
        second: data.selectedTime.getSeconds(),
        millisecond: data.selectedTime.getMilliseconds(),
      })
    );
    setTimePickerValue(data.selectedTimeText ?? "");
  };

  const onTimePickerInput = (ev: React.ChangeEvent<HTMLInputElement>) => {
    setTimePickerValue(ev.target.value);
  };

  React.useEffect(() => {
    async function fetchData() {
      try {
        const [groupsResponse, roomsResponse] = await Promise.all([
          callBackend("getDefaultParticipantGroups", "GET", teamsUserCredential),
          callBackend("getDefaultRooms", "GET", teamsUserCredential),
        ]);

        const groupsData = JSON.parse(groupsResponse);
        const roomsData = roomsResponse.split(";");

        setDefaultParticipantGroups(groupsData.groups);
        setDefaultRooms(roomsData);
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    }

    if (teamsUserCredential) {
      fetchData();
    }
  }, [teamsUserCredential]);

  function validateTitle(text?: string): boolean {
    let valid = true;
    const regex = /^[\w0-9\s\-\_äöüÄÖÜ.]*$/gm;
    let textToValidate = text != null ? text : title;
    if (
      textToValidate === "" ||
      textToValidate.trim() === "" ||
      regex.exec(textToValidate) == null ||
      textToValidate.length > 100
    ) {
      valid = false;
    }

    if (!valid) {
      setTitleMissing(true);
      return false;
    }
    setTitleMissing(false);
    return true;
  }

  const handlePeoplePickerChanged = async (e: any) => {
    let tempPeopleList: string[] = [];
    await Promise.all(
      e.target.selectedPeople.map((currentPerson: any) => {
        if (currentPerson.userPrincipalName) {
          tempPeopleList.push(currentPerson.userPrincipalName);
        } else if (
          currentPerson.scoredEmailAddresses &&
          currentPerson.scoredEmailAddresses.length > 0
        ) {
          tempPeopleList.push(currentPerson.scoredEmailAddresses[0].address);
        } else if (currentPerson.mail && currentPerson.mail.length > 0) {
          tempPeopleList.push(currentPerson.mail);
        }
      })
    );
    const participantsString = tempPeopleList.join(";");
    setSelectedParticipants(participantsString);
  };

  const handleSubmit = async () => {
    const isDateMissing: boolean = selectedDate === undefined || selectedDate === null;
    const isTimeMissing: boolean = selectedTime === undefined || selectedTime === null;
    setDateMissing(isDateMissing || isTimeMissing);
    const isTitleMissing = !validateTitle(title);
    if (isDateMissing || isTimeMissing || isTitleMissing) {
      return;
    }

    setLoading(true);
    props.setPreventDialogClose(true);
    let progressInfo = "";
    const progressSteps = 2;
    const progressIncrement = 1 / progressSteps;

    // Create Board Meeting in Database
    progressInfo = "Fenster bitte nicht schließen! Speichere Sitzung...";
    savingProgressPercent.current = progressIncrement;
    setSavingProgress(progressInfo);

    try {
      const startTimeInTimeZone = selectedTime!.setZone(timeZone, {
        keepLocalTime: true,
      });

      const newBoardMeeting: BoardMeeting = {
        id: props.currentMeetingItem?.id ?? 123456789,
        startTime: startTimeInTimeZone ? startTimeInTimeZone : DateTime.now(),
        title: title,
        fixedParticipants: selectedParticipants,
        remarks: remarks,
        location: location,
        fileLocationId: "",
        timeZone: timeZone,
        room: selectedRoom,
      };

      let backendFunction = "createBoardMeeting";
      let body: any = newBoardMeeting;

      if (props.isEditScreen) {
        backendFunction = "updateBoardMeeting";
        body = {
          boardmeeting: {
            ...newBoardMeeting,
            eventId: props.currentMeetingItem?.eventId,
            meetingLink: props.currentMeetingItem?.meetingLink,
          },
          ensureFileStructure: true,
        };
      }

      const boardmeetings = await callBackend(
        backendFunction,
        "POST",
        teamsUserCredential,
        body,
        undefined
      );

      progressInfo = "Fertig! Das Fenster wird nun automatisch geschlossen...";
      savingProgressPercent.current += progressIncrement;
      setSavingProgress(progressInfo);

      props.setShouldReloadBoardMeetings(true);
      setTimeout(() => {
        props.setOpenDialogCreateEditBoardMeeting(false);
        setLoading(false);
        props.setPreventDialogClose(false);
      }, 2000);
    } catch (error) {
      setError(
        "Fehler beim Erstellen der Sitzung. Bitte wenden Sie sich an den Administrator. " +
          error
      );
      setLoading(false);
      props.setPreventDialogClose(false);
    }
  };

  const handleDelete = async () => {
    setLoading(true);
    props.setPreventDialogClose(true);
    let progressInfo = "";
    const progressSteps = 2;
    const progressIncrement = 1 / progressSteps;

    // DELETE Board Meeting in Database
    progressInfo = "Fenster bitte nicht schließen! Lösche Sitzung...";
    savingProgressPercent.current = progressIncrement;
    setSavingProgress(progressInfo);

    try {
      const body = {
        meetingId: props.currentMeetingItem?.id,
        eventId: props.currentMeetingItem?.eventId,
        fileLocationId: props.currentMeetingItem?.fileLocationId,
      };

      const boardmeetings = await callBackend(
        "deleteBoardMeeting",
        "POST",
        teamsUserCredential,
        JSON.stringify(body),
        undefined
      );

      progressInfo = "Fertig! Das Fenster wird nun automatisch geschlossen...";
      savingProgressPercent.current += progressIncrement;
      setSavingProgress(progressInfo);

      props.setShouldReloadBoardMeetings(true);
      setTimeout(() => {
        props.setOpenDialogCreateEditBoardMeeting(false);
        setLoading(false);
        props.setPreventDialogClose(false);
      }, 2000);
    } catch (error) {
      setError(
        "Fehler beim Erstellen der Sitzung. Bitte wenden Sie sich an den Administrator. " +
          error
      );
      setLoading(false);
      props.setPreventDialogClose(false);
    }
  };

  return (
    <>
      <div className={styles.columnStyle}>
        <div>
          <Field
            validationMessage={savingProgress}
            validationState={savingProgress != "" ? "error" : "none"}
          >
            <ProgressBar value={savingProgressPercent.current} color="warning" />
          </Field>
        </div>
        {loading && <Spinner />}
        {showDeleteConfirmation ? (
          <>
            <p>
              <ul>
                <li>Die Einladung wird im Kalender abgesagt und gelöscht. </li>
                <li>Alle Dokumente in der Dateiablage der Sitzung werden gelöscht.</li>
                <li>
                  Alle TOPs der Sitzung werden dem Pool der verfügbaren TOPs hinzugefügt.{" "}
                </li>
                <li>Dokumente in der Dateiablage der TOPs werden nicht gelöscht.</li>
              </ul>
              <br />
              Sind Sie sich wirklich sicher?
            </p>
            <DialogActions>
              <Button
                appearance="primary"
                onClick={() => setShowDeleteConfirmation(false)}
                disabled={loading}
              >
                Abbrechen
              </Button>
              <Button
                appearance="primary"
                className={styles.buttonDanger}
                onClick={() => handleDelete()}
                disabled={loading}
              >
                Ja, fortfahren
              </Button>
            </DialogActions>
          </>
        ) : (
          <>
            <div className={styles.columnStyle}>
              <div className={styles.root}>
                <Field
                  label="Datum wählen"
                  required
                  validationMessage={
                    dateMissing ? "Bitte Datum und Uhrzeit auswählen" : ""
                  }
                >
                  <DatePicker
                    placeholder="Datum wählen..."
                    value={
                      selectedDate
                        ? new Date(
                            selectedDate.year,
                            selectedDate.month - 1, // JavaScript months are 0-based
                            selectedDate.day,
                            selectedDate.hour,
                            selectedDate.minute,
                            selectedDate.second,
                            selectedDate.millisecond
                          )
                        : null
                    }
                    onSelectDate={onSelectDate}
                    firstDayOfWeek={1}
                    strings={localizedCalendarStrings}
                    formatDate={onFormatDate}
                  />
                </Field>
                <Field
                  label="Zeit wählen oder eingeben (z.B. 13:22)"
                  required
                  validationMessage={dateMissing ? " " : ""}
                >
                  <TimePicker
                    placeholder="Zeit wählen oder eingeben..."
                    freeform
                    dateAnchor={
                      selectedDate
                        ? new Date(
                            selectedDate.year,
                            selectedDate.month - 1, // JavaScript months are 0-based
                            selectedDate.day,
                            selectedDate.hour,
                            selectedDate.minute,
                            selectedDate.second,
                            selectedDate.millisecond
                          )
                        : undefined
                    }
                    selectedTime={
                      selectedTime
                        ? new Date(
                            selectedTime.year,
                            selectedTime.month - 1, // JavaScript months are 0-based
                            selectedTime.day,
                            selectedTime.hour,
                            selectedTime.minute,
                            selectedTime.second,
                            selectedTime.millisecond
                          )
                        : null
                    }
                    onTimeChange={onTimeChange}
                    value={timePickerValue}
                    onInput={onTimePickerInput}
                    hourCycle={"h23"}
                  />
                </Field>
                <Field label="Zeitzone wählen" required>
                  <Dropdown
                    disabled={loading}
                    placeholder="Zeitzone auswählen..."
                    value={timeZoneOptions.find((tz) => tz.value === timeZone)?.text}
                    onOptionSelect={(event, item) => {
                      const selected = timeZoneOptions.find(
                        (tz) => tz.text === item.optionText
                      );
                      if (selected) {
                        setTimeZone(selected.value);
                      }
                    }}
                  >
                    {timeZoneOptions.map((tz) => (
                      <Option key={tz.value} value={tz.value}>
                        {tz.text}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
              </div>

              <Field
                required
                label="Titel"
                validationMessage={
                  titleMissing
                    ? "Titel darf nicht leer und maximal 100 Zeichen lang sein. Nur Buchstaben, Zahlen, '-', '_' und '.' sind erlaubt!"
                    : ""
                }
              >
                <Input
                  disabled={loading}
                  placeholder="Titel der Sitzung..."
                  defaultValue={title}
                  id="txtSummaryField"
                  onChange={(_, data) => {
                    setTitle(data.value.trim());
                    validateTitle(data.value.trim());
                  }}
                />
              </Field>

              <Field label="Ort">
                <Input
                  disabled={loading}
                  placeholder="Ort / Stadt der Sitzung..."
                  defaultValue={location}
                  id="txtLocationField"
                  onChange={(_, data) => {
                    setLocation(data.value);
                  }}
                />
              </Field>

              <div className={styles.rowContainer}>
                <Field label="Raum">
                  <Dropdown
                    className={styles.dropdown}
                    disabled={loading}
                    placeholder="Raum hinzufügen"
                    selectedOptions={[]}
                    onOptionSelect={(_, item) => {
                      if (item.optionValue) {
                        setSelectedRoom(item.optionValue);
                      }
                    }}
                  >
                    {defaultRooms.map((room: string) => (
                      <Option key={room} value={room}>
                        {room}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="&nbsp;">
                  <Input
                    placeholder="oder Raumname eingeben"
                    onChange={(_, data) => setSelectedRoom(data.value)}
                    value={selectedRoom}
                  />
                </Field>
              </div>

              <div className={styles.rowContainer}>
                <Field label="Fixe Teilnehmer">
                  <Dropdown
                    className={styles.dropdown}
                    disabled={loading}
                    placeholder="Gruppe hinzufügen"
                    selectedOptions={selectedGroupParticipantOptions} // Controls the selection
                    onOptionSelect={(event, item) => {
                      const selected = defaultParticipantGroups.find(
                        (group: any) => group.displayName === item.optionText
                      );
                      if (selected) {
                        setSelectedParticipants(selected.mailAddresses);
                        setSelectedGroupParticipantOptions([]);
                      }
                    }}
                  >
                    {defaultParticipantGroups.map((group: any) => (
                      <Option key={group.id} value={group.id}>
                        {group.displayName}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="&nbsp;">
                  <PeoplePicker
                    className={styles.peoplePicker}
                    selectionMode="multiple"
                    type="person"
                    placeholder="Weitere Personen via Texteingabe hinzufügen"
                    allowAnyEmail
                    selectionChanged={handlePeoplePickerChanged}
                    selectedPeople={
                      selectedParticipants.trim() !== ""
                        ? selectedParticipants.split(";").map((email) => ({
                            displayName: email.split("@")[0],
                            email: email.trim(),
                            userPrincipalName: email.trim(),
                          }))
                        : []
                    }
                  />
                </Field>
              </div>
              <Field label="Anmerkungen">
                <Textarea
                  disabled={loading}
                  style={{ height: "200px" }}
                  value={remarks}
                  onChange={(_, data) => {
                    setRemarks(data.value);
                  }}
                />
              </Field>
            </div>
            <div className={styles.buttonRow}>
              <Button
                appearance="primary"
                onClick={() => setShowDeleteConfirmation(true)}
                disabled={loading}
                style={{
                  backgroundColor: "red",
                  color: "white",
                  display: props.isEditScreen ? "block" : "none",
                }}
              >
                {props.isEditScreen ? "Sitzung löschen" : "Sitzung löschen"}
              </Button>
              <Button appearance="primary" onClick={handleSubmit} disabled={loading}>
                {props.isEditScreen ? "Sitzung speichern" : "Sitzung erstellen"}
              </Button>
            </div>
          </>
        )}
        {!loading && error && <div style={{ bottom: "10px" }}>{error}</div>}
      </div>
    </>
  );
}