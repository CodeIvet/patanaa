import { useContext, useEffect, useState } from "react";
import {
  DataGrid,
  Text,
  DataGridHeader,
  DataGridRow,
  Spinner,
  DataGridBody,
  TableColumnDefinition,
  DataGridHeaderCell,
  DataGridCell,
  makeStyles,
  tokens,
  ToolbarButton,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTrigger,
  Button,
  DialogTitle,
  DialogContent,
} from "@fluentui/react-components";
import { useMediaQuery } from "react-responsive";
import "./Dashboard.css";
import { TeamsFxContext } from "../Context";
import * as helper from "./lib/helper";
import { MgtTemplateProps, People } from "@microsoft/mgt-react";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import {
  IcCalendarAgenda,
  IcEdit,
  IcFiles,
  IcCreateProtocol,
  IcInvites,
  IcJoinMeeting,
} from "./lib/Icons";
import {
  AddCircle24Filled,
  Dismiss24Regular,
  Settings24Regular,
} from "@fluentui/react-icons";
import { CreateEditBoardMeeting } from "./CreateEditBoardMeeting";
import { EditUserMappings } from "./EditUserMappings";
import { ManageAgenda } from "./ManageAgenda";
import { CreateProtocolAgenda } from "./CreateProtocolAgenda";
import { Invite } from "./Invite";
import { DateTime } from "luxon";
import { timeZoneOptions } from "./lib/helper";

const useStyles = makeStyles({
  gridRow: {
    backgroundColor: tokens.colorNeutralBackground4,
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground4,
      fontWeight: "bold",
    },
  },
});

export function Dashboard(props: { showFunction?: boolean; environment?: string }) {
  const [boardMeetings, setBoardMeetings] = useState<helper.BoardMeeting[]>([]);
  const [rawBoardMeetings, setRawBoardMeetings] = useState<any>();
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [openDialogCreateEditBoardMeeting, setOpenDialogCreateEditBoardMeeting] =
    useState(false);
  const [openDialogCreateProtocolAgenda, setOpenDialogCreateProtocolAgenda] =
    useState(false);
  const [openDialogManageAgenda, setOpenDialogManageAgenda] = useState(false);
  const [openDialogInvites, setOpenDialogInvites] = useState(false);
  const [shouldReloadBoardMeetings, setShouldReloadBoardMeetings] =
    useState<boolean>(false);
  const [isBoardMeetingEditScreen, setIsBoardMeetingEditScreen] =
    useState<boolean>(false);
  const [currentMeetingItem, setCurrentMeetingItem] = useState<helper.BoardMeeting>();
  const [preventDialogClose, setPreventDialogClose] = useState<boolean>(false);
  const [openDialogEditUserMappings, setOpenDialogEditUserMappings] = useState(false);

  const isSmallScreen = useMediaQuery({ query: "(max-width:960px)" });
  const { teamsUserCredential } = useContext(TeamsFxContext);

  const styles = useStyles();
  initializeIcons();

  // Custom Overflow Template Component
  const OverflowTemplate = (props: MgtTemplateProps) => {
    const { dataContext } = props;
    const extra = dataContext.extra; // Number of extra users
    const people = dataContext.people; // List of all people

    return (
      <span style={{ marginLeft: "8px", fontWeight: "bold", color: "#0078D4" }}>
        +{extra} more
      </span>
    );
  };

  const columnsFull: TableColumnDefinition<helper.BoardMeeting>[] = [
    {
      columnId: "id",
      compare: (a, b) => a.id - b.id,
      renderHeaderCell: () => (
        <DataGridHeaderCell
          style={{
            fontWeight: "bold",
            minWidth: "50px",
            maxWidth: "50px",
            paddingLeft: "5px",
          }}
        >
          <DataGridCell className="cell">Id</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => {
        return (
          <DataGridCell
            className="cell"
            style={{
              minWidth: "50px",
              maxWidth: "50px",
              paddingLeft: "10px",
              cursor: "pointer",
            }}
            onClick={() => console.log("rowId: ", item.id)}
          >
            {item.id}
          </DataGridCell>
        );
      },
    },
    {
      columnId: "date",
      compare: (a, b) => {
        return a.startTime.toMillis() - b.startTime.toMillis();
      },
      renderHeaderCell: () => (
        <DataGridHeaderCell
          style={{
            fontWeight: "bold",
            minWidth: "250px",
            maxWidth: "250px",
            paddingLeft: "0px",
          }}
        >
          <DataGridCell className="cell">Datum</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item: helper.BoardMeeting) => {
        return (
          <DataGridCell
            className="cell"
            style={{
              minWidth: "250px",
              maxWidth: "250px",
              cursor: "pointer",
            }}
            onClick={() => console.log("rowId: ", item.id)}
          >
            {item.startTime.setLocale("de-DE").toFormat("dd.MM.yyyy")}{" "}
            {item.startTime.toFormat("HH:mm")}
            {" ("}
            {timeZoneOptions.find((tz) => tz.value === item.timeZone)?.text}
            {")"}
            {item.timeZone !== "Europe/Berlin" && (
              <>
                <br />
                {item.startTime
                  .setZone("Europe/Berlin")
                  .setLocale("de-DE")
                  .toFormat("dd.MM.yyyy")}{" "}
                {item.startTime.setZone("Europe/Berlin").toFormat("HH:mm")}
                {" (Berlin)"}
              </>
            )}
          </DataGridCell>
        );
      },
    },
    {
      columnId: "title",
      compare: (a, b) => a.title.localeCompare(b.title),
      renderHeaderCell: () => (
        <DataGridHeaderCell
          style={{
            fontWeight: "bold",
            flexGrow: 1,
            flexShrink: 0,
            flexBasis: "0%",
            overflow: "hidden",
            paddingLeft: "0px",
          }}
        >
          <DataGridCell className="cell">Titel</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => {
        return (
          <DataGridCell
            className="cell"
            style={{
              flexGrow: 1,
              flexShrink: 0,
              flexBasis: "0%",
              overflow: "hidden",
              display: "block",
              cursor: "pointer",
            }}
            onClick={() => console.log("rowId: ", item.id)}
          >
            {item.title}
          </DataGridCell>
        );
      },
    },
    {
      columnId: "fixedParticipants",
      compare: (a, b) => a.title.localeCompare(b.title),
      renderHeaderCell: () => (
        <DataGridHeaderCell
          style={{
            fontWeight: "bold",
            minWidth: "150px",
            maxWidth: "350px",
            paddingLeft: "0px",
          }}
        >
          <DataGridCell className="cell">Feste Teilnehmer</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => {
        const upns = item.fixedParticipants ? item.fixedParticipants.split(";") : [];
        // Split into groups of 5
        const chunkSize = isSmallScreen ? 5 : 80;
        const upnChunks = helper.chunkArray(upns, chunkSize);

        const externalUsers = upns
          .filter((upn) => !upn.endsWith("@axelspringer.com"))
          .map((email, index) => ({
            displayName: email,
            email: email,
            id: `external-${index}`, // unique identifier for each entry
          }));

        return (
          <DataGridCell
            className="cell"
            style={{
              minWidth: "150px",
              maxWidth: "350px",
              paddingLeft: "0px",
            }}
          >
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                // gap: "10px", // Space between rows
              }}
            >
              {upnChunks.map((chunk, index) => (
                <div key={index} style={{ display: "flex" }}>
                  <People
                    key={`${item.id}-${chunk.join(",")}`}
                    userIds={chunk}
                    showMax={chunk.length}
                    personCardInteraction="hover"
                    fallbackDetails={externalUsers}
                  />
                </div>
              ))}
            </div>
          </DataGridCell>
        );
      },
    },
    {
      columnId: "remarks",
      compare: (a, b) => a.remarks.localeCompare(b.remarks),
      renderHeaderCell: () => (
        <DataGridHeaderCell
          style={{
            fontWeight: "bold",
            flexGrow: 1,
            flexShrink: 0,
            flexBasis: "0%",
            overflow: "hidden",
            paddingLeft: "0px",
          }}
        >
          <DataGridCell className="cell">Bemerkungen</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => {
        return (
          <DataGridCell
            className="cell"
            style={{
              flexGrow: 1,
              flexShrink: 0,
              flexBasis: "0%",
              overflow: "hidden",
              display: "block",
              cursor: "pointer",
            }}
            onClick={() => console.log("rowId: ", item.id)}
          >
            {item.remarks}
          </DataGridCell>
        );
      },
    },
    {
      columnId: "actions",
      compare: (a, b) => a.remarks.localeCompare(b.remarks),
      renderHeaderCell: () => (
        <DataGridHeaderCell
          style={{
            fontWeight: "bold",
            minWidth: "300px",
            maxWidth: "300px",
            paddingLeft: "0px",
          }}
        >
          <DataGridCell className="cell">Aktionen</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => {
        return (
          <DataGridCell
            className="cell"
            style={{
              minWidth: "300px",
              maxWidth: "300px",
              paddingLeft: "0px",
            }}
          >
            <IcEdit
              onClick={() => {
                setCurrentMeetingItem(item);
                setIsBoardMeetingEditScreen(true);
                setOpenDialogCreateEditBoardMeeting(true);
              }}
              disabled={false}
            />
            <IcCalendarAgenda
              onClick={() => {
                setCurrentMeetingItem(item);
                setOpenDialogManageAgenda(true);
              }}
              disabled={false}
            />
            <IcCreateProtocol
              onClick={() => {
                setCurrentMeetingItem(item);
                setOpenDialogCreateProtocolAgenda(true);
              }}
              disabled={false}
            />
            <IcFiles
              onClick={async () => {
                setIsLoading(true);
                const folderLink = await helper.callBackend(
                  "getFolderWebUrl",
                  "GET",
                  teamsUserCredential,
                  undefined,
                  ["fileLocationId=" + item.fileLocationId, "driveName=Meetings"]
                );

                setIsLoading(false);
                window.open(folderLink, "_blank");
              }}
              disabled={item.fileLocationId === null}
            />
            <IcInvites
              onClick={() => {
                setCurrentMeetingItem(item);
                setOpenDialogInvites(true);
              }}
              disabled={false}
            />
            <IcJoinMeeting
              onClick={() => window.open(item.meetingLink, "_blank")}
              disabled={false}
            />
          </DataGridCell>
        );
      },
    },
    // ... other columns as needed
  ];

  const columnsRedux: TableColumnDefinition<helper.BoardMeeting>[] = columnsFull.filter(
    (column) => column.columnId !== "remarks"
  );

  const fetchBoardMeetings = async () => {
    setIsLoading(true);
    try {
      const boardmeetings = await helper.callBackend(
        "getBoardMeetings",
        "GET",
        teamsUserCredential,
        undefined,
        undefined
      );
      setRawBoardMeetings(JSON.parse(boardmeetings));
    } catch (error) {
      setError(JSON.stringify(error));
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    const fetchData = async () => {
      await fetchBoardMeetings(); // Await the call inside the async function
      if (shouldReloadBoardMeetings) {
        setShouldReloadBoardMeetings(false);
      }
    };

    fetchData(); // Invoke the async function
  }, [shouldReloadBoardMeetings]);

  useEffect(() => {
    if (rawBoardMeetings) {
      // Transform the JSON data to an array of BoardMeeting objects
      const transformedData: helper.BoardMeeting[] = rawBoardMeetings.map(
        (item: any) => ({
          id: item.ID.toString(), // Convert ID to string
          startTime: DateTime.fromISO(item.StartTime, { zone: item.TimeZone }), // Convert StartTime to DateTime
          title: item.Title, // Map Title directly
          fixedParticipants: item.FixedParticipants, // Map FixedParticipants directly
          remarks: item.Remarks, // Map Remarks directly
          location: item.Location, // Map Location directly
          meetingLink: item.MeetingLink, // Map MeetingLink directly
          fileLocationId: item.FileLocationId, // Map FileLocationId directly
          eventId: item.EventId, // Map EventId directly
          timeZone: item.TimeZone, // Map TimeZone directly
          room: item.Room, // Map Room directly
        })
      );
      setBoardMeetings(transformedData);
    }
  }, [rawBoardMeetings]);

  return (
    <div className="page-padding">
      <div style={{ display: "flex", flexDirection: "column", padding: "10px" }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            paddingBottom: "10px",
          }}
        >
          <ToolbarButton
            appearance={"primary"}
            icon={<AddCircle24Filled />}
            style={{ marginRight: "10px" }}
            onClick={() => {
              setIsBoardMeetingEditScreen(false);
              setCurrentMeetingItem(undefined);
              setOpenDialogCreateEditBoardMeeting(true);
            }}
          >
            {!isSmallScreen && "Sitzung erstellen"}
          </ToolbarButton>
          <ToolbarButton
            appearance={"primary"}
            icon={<Settings24Regular />}
            style={{ marginRight: "10px" }}
            onClick={() => {
              setOpenDialogEditUserMappings(true);
            }}
          ></ToolbarButton>
        </div>
      </div>
      <DataGrid
        items={boardMeetings}
        columns={isSmallScreen ? columnsRedux : columnsFull}
        sortable
        getRowId={(item) => item.id}
        focusMode="composite"
        defaultSortState={{
          sortColumn: "date",
          sortDirection: "descending",
        }}
        subtleSelection={true}
        selectionAppearance={"neutral"}
      >
        <DataGridHeader>
          <DataGridRow>{({ renderHeaderCell }) => <>{renderHeaderCell()}</>}</DataGridRow>
        </DataGridHeader>
        {isLoading ? (
          <Spinner style={{ padding: 10 }} />
        ) : (
          <DataGridBody<helper.BoardMeeting>>
            {({ item, rowId }) => (
              <DataGridRow<helper.BoardMeeting> key={rowId} className={styles.gridRow}>
                {({ renderCell }) => <>{renderCell(item as any)}</>}
              </DataGridRow>
            )}
          </DataGridBody>
        )}
      </DataGrid>

      <Text size={100} id="appVer">
        1.4.2
      </Text>
      <Dialog
        open={openDialogCreateEditBoardMeeting}
        onOpenChange={(event, data) => {
          setOpenDialogCreateEditBoardMeeting(data.open);
          if (!data.open) {
            //reloadUser();
          }
        }}
        modalType="alert"
      >
        <DialogSurface
          style={isSmallScreen ? { maxWidth: "90%" } : { maxWidth: "900px" }}
        >
          <DialogBody>
            <DialogTitle
              action={
                <DialogTrigger action="close">
                  <Button
                    appearance="subtle"
                    aria-label="close"
                    icon={<Dismiss24Regular />}
                    disabled={preventDialogClose}
                  />
                </DialogTrigger>
              }
            >
              {isBoardMeetingEditScreen ? "Sitzung bearbeiten" : "Neue Sitzung erstellen"}
            </DialogTitle>
            <DialogContent>
              <CreateEditBoardMeeting
                isSmallScreen={isSmallScreen}
                isEditScreen={isBoardMeetingEditScreen}
                setOpenDialogCreateEditBoardMeeting={setOpenDialogCreateEditBoardMeeting}
                currentMeetingItem={currentMeetingItem}
                setShouldReloadBoardMeetings={setShouldReloadBoardMeetings}
                setPreventDialogClose={setPreventDialogClose}
              />
            </DialogContent>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <Dialog
        open={openDialogManageAgenda}
        onOpenChange={(event, data) => {
          setOpenDialogManageAgenda(data.open);
          if (!data.open) {
            //reloadUser();
          }
        }}
        modalType="alert"
      >
        <DialogSurface style={isSmallScreen ? { maxWidth: "90%" } : {}}>
          <DialogBody>
            <DialogTitle
              action={
                <DialogTrigger action="close">
                  <Button
                    appearance="subtle"
                    aria-label="close"
                    icon={<Dismiss24Regular />}
                    disabled={preventDialogClose}
                  />
                </DialogTrigger>
              }
            >
              Agenda bearbeiten
              <br />
              {currentMeetingItem?.startTime
                .setLocale("de")
                .toLocaleString(DateTime.DATE_SHORT)}{" "}
              {currentMeetingItem?.startTime.toFormat("HH:mm")}:{" "}
              {currentMeetingItem?.title}
            </DialogTitle>
            <DialogContent>
              <ManageAgenda
                //isSmallScreen={isSmallScreen}
                //setOpenDialogManageAgenda={setOpenDialogManageAgenda}
                //currentMeetingItem={currentMeetingItem}
                //setShouldReloadBoardMeetings={setShouldReloadBoardMeetings}
                //setPreventDialogClose={setPreventDialogClose}
              />
            </DialogContent>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <Dialog
        open={openDialogCreateProtocolAgenda}
        onOpenChange={(event, data) => {
          setOpenDialogCreateProtocolAgenda(data.open);
        }}
        modalType="alert"
      >
        <DialogSurface style={isSmallScreen ? { maxWidth: "90%" } : {}}>
          <DialogBody>
            <DialogTitle
              action={
                <DialogTrigger action="close">
                  <Button
                    appearance="subtle"
                    aria-label="close"
                    icon={<Dismiss24Regular />}
                    disabled={preventDialogClose}
                  />
                </DialogTrigger>
              }
            >
              {"Agendaübersicht / Protokollvorlage erstellen"}
            </DialogTitle>
            <DialogContent>
              <CreateProtocolAgenda
                setOpenDialogCreateProtocolAgenda={setOpenDialogCreateProtocolAgenda}
                currentMeetingItem={currentMeetingItem}
                setPreventDialogClose={setPreventDialogClose}
              />
            </DialogContent>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <Dialog
        open={openDialogInvites}
        onOpenChange={(event, data) => {
          setOpenDialogInvites(data.open);
          if (!data.open) {
            //reloadUser();
          }
        }}
        modalType="alert"
      >
        <DialogSurface
          style={isSmallScreen ? { maxWidth: "90%" } : { maxWidth: "900px" }}
        >
          <DialogBody>
            <DialogTitle
              action={
                <DialogTrigger action="close">
                  <Button
                    appearance="subtle"
                    aria-label="close"
                    icon={<Dismiss24Regular />}
                    disabled={preventDialogClose}
                  />
                </DialogTrigger>
              }
            >
              Einladungen verwalten
            </DialogTitle>
            <DialogContent>
              <Invite
                isSmallScreen={isSmallScreen}
                setOpenDialogInvites={setOpenDialogInvites}
                currentMeetingItem={currentMeetingItem}
                setPreventDialogClose={setPreventDialogClose}
                setShouldReloadBoardMeetings={setShouldReloadBoardMeetings}
              />
            </DialogContent>
          </DialogBody>
        </DialogSurface>
      </Dialog>
      <Dialog
        open={openDialogEditUserMappings}
        onOpenChange={(event, data) => {
          setOpenDialogEditUserMappings(data.open);
          if (!data.open) {
            //reloadUser();
          }
        }}
        modalType="alert"
      >
        <DialogSurface
          style={isSmallScreen ? { maxWidth: "90%" } : { maxWidth: "900px" }}
        >
          <DialogBody>
            <DialogTitle
              action={
                <DialogTrigger action="close">
                  <Button
                    appearance="subtle"
                    aria-label="close"
                    icon={<Dismiss24Regular />}
                    disabled={preventDialogClose}
                  />
                </DialogTrigger>
              }
            >
              Anzeigenamen der Teilnehmer verwalten
            </DialogTitle>
            <DialogContent>
              <EditUserMappings
                isSmallScreen={isSmallScreen}
                setOpenDialogEditUserMappings={setOpenDialogEditUserMappings}
                setPreventDialogClose={setPreventDialogClose}
              />
            </DialogContent>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}