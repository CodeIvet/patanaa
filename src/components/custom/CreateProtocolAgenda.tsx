import {
  Button,
  Field,
  makeStyles,
  ProgressBar,
  Spinner,
  tokens,
  Text,
  Image,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
} from "@fluentui/react-components";
import React, { useContext } from "react";
import { useState } from "react";
import { BoardMeeting, callBackend } from "./lib/helper";

//import renameFileImage from "./lib/rename_file.png"; // for screenshot
import { TeamsFxContext } from "../Context";

const useStyles = makeStyles({
  columnStyle: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  buttonRow: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "20px", // Optional: Add spacing from the content above
  },
});

export function CreateProtocolAgenda(props: {
  currentMeetingItem?: BoardMeeting;
  setOpenDialogCreateProtocolAgenda: (open: boolean) => void;
  setPreventDialogClose: (open: boolean) => void;
}) {
  const [loading, setLoading] = useState(false);
  const [folderWebUrl, setFolderWebUrl] = useState("");
  const [savingProgress, setSavingProgress] = useState("");
  const savingProgressPercent = React.useRef(0.0);
  const [error, setError] = useState<string>("");
  const [successMessage, setSuccessMessage] = useState<string>("");

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const styles = useStyles();

  const handleSubmit = async (isProtocol: boolean) => {
    setSuccessMessage("");
    setLoading(true);
    props.setPreventDialogClose(true);
    let progressInfo = "";
    const progressSteps = 2;
    const progressIncrement = 1 / progressSteps;

    // Create Board Meeting in Database
    progressInfo = "Erstellung läuft...";
    savingProgressPercent.current = progressIncrement;
    setSavingProgress(progressInfo);

    try {
      if (isProtocol) {
        const result = await callBackend(
          "processProtocolTemplate",
          "POST",
          teamsUserCredential,
          {
            boardMeeting: props.currentMeetingItem,
          }
        );
      } else {
        // Create PDF of agenda items
        const result = await callBackend("createAgendaPdf", "POST", teamsUserCredential, {
          boardMeeting: props.currentMeetingItem,
        });
      }
      progressInfo = " Datei erfolgreich erstellt!";
      savingProgressPercent.current += progressIncrement;
      setSavingProgress(progressInfo);

      setSuccessMessage(
        "Datei erfolgreich erstellt! Du kannst dieses Fenster nun schließen und in der Übersicht über das Büroklammer-Symbol auf die Datei zugreifen."
      );
    } catch (error: any) {
      let errMessage = "";
      if (error.response) {
        // Axios-style error
        console.error("Backend Error:", error.response.data);
        errMessage = error.response.data;
      } else {
        // Generic error
        console.error("Unknown Error:", error.message || error);
        errMessage = "Unknown error";
      }
      setSavingProgress(
        "Fehler beim Erstellen der Datei. Stelle bitte sicher, dass du die Datei nicht geöffnet hast."
      );
      setError(
        "Fehler beim Erstellen der Datei. Stelle bitte sicher, dass du die Datei nicht geöffnet hast: " +
          errMessage
      );
    } finally {
      setLoading(false);
      props.setPreventDialogClose(false);
    }
  };

  React.useEffect(() => {
    const fetchFolderLink = async () => {
      try {
        setLoading(true);
        const folderLink = await callBackend(
          "getFolderWebUrl",
          "GET",
          teamsUserCredential,
          undefined,
          [
            "fileLocationId=" + props.currentMeetingItem?.fileLocationId,
            "driveName=Meetings",
          ]
        );

        setFolderWebUrl(folderLink);
      } catch (error) {
        console.error("Error fetching folder link:", error);
      } finally {
        setLoading(false);
      }
    };

    fetchFolderLink();
  }, [teamsUserCredential]);

  return (
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
      <div className={styles.columnStyle}>
        <div>
          <Text>
            Hier kannst du eine druckbare Agendaübersicht oder eine bearbeitbare
            Protokollvorlage für die aktuelle Sitzung erstellen. Es werden immer englische
            und deutsche Varianten der Datei angelegt. Die Dateien werden in dem{" "}
            <a
              style={{ color: tokens.colorBrandBackground }}
              target="_blank"
              href={folderWebUrl}
            >
              Ordner der Sitzung
            </a>{" "}
            abgelegt.
            <p>
              Bitte beachte, dass eine eventuell zuvor erstellte Agendaübersicht ("Agenda
              Overview.pdf") oder Protokollvorlage ("Protocol DRAFT.docx"){" "}
              <strong>überschrieben</strong> wird.
            </p>
            <p>
              <strong>
                Aus diesem Grund nenne die Protokollvorlage vor dem Bearbeiten/Ausfüllen
                um.
                <br />
                Klicke dazu während des Bearbeitens auf den Dateinamen in der Titelleiste.
              </strong>
            </p>
            <Accordion collapsible>
              <AccordionItem value="1">
                <AccordionHeader>Zeige Screenshot</AccordionHeader>
                <AccordionPanel>
                  <Image
                    alt="Rename File"
                    bordered
                    shadow
                    //src={renameFileImage}
                    width={"400px"}
                  />
                </AccordionPanel>
              </AccordionItem>
            </Accordion>
          </Text>
        </div>
        <div className={styles.buttonRow}>
          <Button
            appearance="primary"
            onClick={() => handleSubmit(true)}
            disabled={loading}
          >
            {"Protokoll-Vorlage erstellen (DOCX)"}
          </Button>
          <Button
            appearance="primary"
            onClick={() => handleSubmit(false)}
            disabled={loading}
          >
            {"Agendaübersicht erstellen (PDF)"}
          </Button>
        </div>
        {!loading && successMessage != "" && (
          <div style={{ bottom: "10px", color: "green" }}>{successMessage}</div>
        )}
        {!loading && error && <div style={{ bottom: "10px", color: "red" }}>{error}</div>}
      </div>
    </div>
  );
}