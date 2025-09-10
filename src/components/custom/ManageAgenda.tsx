import {
  Button,
  Field,
  makeStyles,
  ProgressBar,
  Spinner,
  Accordion,
  SearchBox,
  Text,
  SearchBoxChangeEvent,
  InputOnChangeData,
  PopoverSurface,
  Popover,
  PopoverTrigger,
  Tooltip,
  PopoverProps,
  tokens,
  AccordionToggleEventHandler,
} from "@fluentui/react-components";
import React, { useContext, useEffect } from "react";
import { useState } from "react";
import { TeamsFxContext } from "../Context";
import { AgendaItem, BoardMeeting } from "./lib/helper";
import * as helper from "./lib/helper";
import {
  DndContext,
  useSensor,
  useSensors,
  PointerSensor,
  KeyboardSensor,
  closestCenter,
  DragEndEvent,
} from "@dnd-kit/core";
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  rectSortingStrategy,
} from "@dnd-kit/sortable";
import { restrictToVerticalAxis } from "@dnd-kit/modifiers";
import {
  AddCircle24Filled,
  CollectionsAdd24Filled,
  Delete24Regular,
} from "@fluentui/react-icons";

import SortableAgendaItem from "./lib/SortableAgendaItem";
import { DateTime } from "luxon";

// import "./CreateBoardMeeting.css";

const useStyles = makeStyles({
  root: {
    display: "grid",
    columnGap: "20px",
    gridTemplateColumns: "repeat(2, 1fr)",
  },
  columnStyle: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  buttonContainerBottom: {
    display: "flex",
    justifyContent: "flex-end",
    paddingTop: "20px",
  },
  buttonSpacingUnderAccordion: {
    display: "flex",
    alignItems: "center", // Aligns the buttons vertically
    gap: "8px", // Adds space between the buttons
    marginTop: "8px",
  },
  searchResultsContainer: {
    marginTop: "10px",
    borderTop: "1px solid #ccc",
    paddingTop: "10px",
  },

  moreResults: {
    marginTop: "8px",
    fontStyle: "italic",
    color: "gray",
  },
  deleteOverlay: {
    position: "absolute",
    top: "0",
    left: "0",
    width: "100%",
    height: "100%",
    backgroundColor: "rgba(0, 0, 0, 0.3)", // Dark overlay effect
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    zIndex: 10, // Ensure it appears on top of the popover content
  },
  deleteBox: {
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "16px",
    borderRadius: "8px",
    boxShadow: "0px 4px 10px rgba(0, 0, 0, 0.1)",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
    width: "280px",
  },
  buttonGroup: {
    display: "flex",
    gap: "8px",
    justifyContent: "flex-end",
    width: "100%",
  },
  resultItem: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "8px",
    borderBottom: "1px solid #ddd",
  },
  buttonDanger: {
    backgroundColor: tokens.colorStatusDangerBackground3,
    "&:hover": {
      backgroundColor: tokens.colorStatusDangerBackground3Pressed,
    },
  },
  title: {
    flexGrow: 1, // Makes the title take available space, pushing the bin to the right
    textAlign: "left",
    cursor: "pointer",
  },
  deleteIcon: {
    flexShrink: 0, // Prevents icon from shrinking
    marginLeft: "8px", // Adds space between title and trash bin
  },
});

export function ManageAgenda(props: {
  isSmallScreen: boolean;
  setOpenDialogManageAgenda: (open: boolean) => void;
  currentMeetingItem?: BoardMeeting;
  setShouldReloadBoardMeetings: (open: boolean) => void;
  setPreventDialogClose: (open: boolean) => void;
}) {
  const [loading, setLoading] = useState(false);
  const [purgeLoading, setPurgeLoading] = useState(false);
  const [savingProgress, setSavingProgress] = useState("");
  const savingProgressPercent = React.useRef(0.0);
  const [error, setError] = useState<string>("");
  const [agendaItems, setAgendaItems] = useState<AgendaItem[]>([]);
  const [unassignedAgendaItems, setUnassignedAgendaItems] = useState<AgendaItem[]>(); // Items without an assigned boardmeeting
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filteredItems, setFilteredItems] = useState<AgendaItem[]>([]);
  const [itemToBePurged, setItemToBePurged] = useState<AgendaItem | null>(null);
  const [openItems, setOpenItems] = useState<(unknown)[]>([]);

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const styles = useStyles();

  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const handleToggle: AccordionToggleEventHandler<unknown> = (event, data) => {
    setOpenItems(data.openItems);
  };

  const handleDragStart = () => {
    setOpenItems([]); // Collapse all accordion panels
  };

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;

    if (active.id !== over?.id) {
      setAgendaItems((items: AgendaItem[]) => {
        const oldIndex = items.findIndex((item) => item.id === active.id);
        const newIndex = items.findIndex((item) => item.id === over?.id);

        // Calculate and update agenda items with timestamps
        const updatedAgendaItems: AgendaItem[] = helper.calculateTimestamps(
          props.currentMeetingItem?.startTime ?? DateTime.now(),
          arrayMove(items, oldIndex, newIndex)
        );

        return updatedAgendaItems;
      });
    }
  };

  // Handle search input change
  const handleSearchChange = (event: SearchBoxChangeEvent, data: InputOnChangeData) => {
    const query = data.value;
    setSearchQuery(query);

    // Filter the unassigned agenda items based on the search query
    const filtered = unassignedAgendaItems?.filter((item) =>
      item.title.toLowerCase().includes(query.toLowerCase())
    );
    setFilteredItems(filtered ?? []);
  };

  const handleSubmit = async () => {
    setLoading(true);
    props.setPreventDialogClose(true);
    let progressInfo = "";
    const progressSteps = 2;
    const progressIncrement = 1 / progressSteps;

    // Create Board Meeting in Database
    progressInfo = "Speichern läuft...";
    savingProgressPercent.current = progressIncrement;
    setSavingProgress(progressInfo);

    try {
      await helper.callBackend("updateAgenda", "POST", teamsUserCredential, {
        agendaItems: agendaItems,
        unassignedAgendaItems: unassignedAgendaItems,
        boardMeetingId: props.currentMeetingItem?.id,
        meetingFolderId: props.currentMeetingItem?.fileLocationId,
      });
      progressInfo = "Fertig! Das Fenster wird nun automatisch geschlossen...";
      savingProgressPercent.current += progressIncrement;
      setSavingProgress(progressInfo);

      props.setShouldReloadBoardMeetings(true);
      setTimeout(() => {
        props.setOpenDialogManageAgenda(false);
        setLoading(false);
        props.setPreventDialogClose(false);
      }, 2000);
    } catch (error) {
      setSavingProgress("Fehler beim Speichern");
      setError(JSON.stringify(error));
      setLoading(false);
      props.setPreventDialogClose(false);
    }
  };

  const fetchAgendaItems = async () => {
    setLoading(true);
    try {
      const agendaItems = await helper.callBackend(
        "getAgendaItems",
        "GET",
        teamsUserCredential,
        undefined,
        ["boardmeeting=" + props.currentMeetingItem?.id]
      );
      const agendaItemsWithStartDates = helper.calculateTimestamps(
        props.currentMeetingItem?.startTime ?? DateTime.now(),
        JSON.parse(agendaItems) as AgendaItem[]
      );
      setAgendaItems(agendaItemsWithStartDates);

      const unassignedItems = await helper.callBackend(
        "getAgendaItems",
        "GET",
        teamsUserCredential,
        undefined,
        []
      );
      const unassignedAgendaItemsJson = JSON.parse(unassignedItems) as AgendaItem[];
      setUnassignedAgendaItems(unassignedAgendaItemsJson);
      setFilteredItems(unassignedAgendaItemsJson);
    } catch (error) {
      setError(JSON.stringify(error));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchAgendaItems();
  }, []);

  const [openPopover, setOpenPopover] = React.useState(false);
  const handleOpenPopoverChange: PopoverProps["onOpenChange"] = (e, data) =>
    setOpenPopover(data.open || false);

  const handleUpdateAgendaItem = (
    id: number,
    newTitle: string,
    newDurationInMinutes: number,
    newIsMisc: boolean,
    newIsDecisionNeeded: boolean,
    newSelectedParticipants: string,
    startTime: DateTime,
    newRemarks?: string
  ) => {
    setAgendaItems((prevItems) => {
      const updatedAgendaItemArray = prevItems.map((item) =>
        item.id === id
          ? {
              ...item,
              title: newTitle,
              durationInMinutes: newDurationInMinutes,
              isMisc: newIsMisc,
              needsDecision: newIsDecisionNeeded,
              additionalParticipants: newSelectedParticipants,
              startTime: startTime,
              remarks: newRemarks || "",
            }
          : item
      );
      const newAgendaItems: AgendaItem[] = helper.calculateTimestamps(
        props.currentMeetingItem?.startTime ?? DateTime.now(),
        updatedAgendaItemArray
      );

      return newAgendaItems;
    });
  };

  // Function to handle the selection of an item
  const handleExistingItemClick = (item: AgendaItem) => {
    // item.boardMeeting = props.currentMeetingItem?.id ?? 0;

    // Calculate the new agenda items with start dates
    const agendaItemsWithStartDates = helper.calculateTimestamps(
      props.currentMeetingItem?.startTime ?? DateTime.now(),
      [...agendaItems, item] // Include the new item to calculate timestamps
    );

    setAgendaItems(agendaItemsWithStartDates);

    const updatedUnassignedAgendaItems = unassignedAgendaItems?.filter(
      (unassignedItem) => unassignedItem.id !== item.id
    );
    setUnassignedAgendaItems(updatedUnassignedAgendaItems ?? []);
    setFilteredItems(updatedUnassignedAgendaItems ?? []);
    setSearchQuery("");

    setOpenPopover(false);
  };

  const handleAddAgendaItem = () => {
    const newAgendaItem: AgendaItem = {
      id: Math.floor(Math.random() * 1000000), // this is a temporary id, it will be replaced upon saving
      title: "Neuer Agendapunkt",
      durationInMinutes: 30,
      isMisc: false,
      needsDecision: false,
      additionalParticipants: "",
      orderIndex: agendaItems.length,
      isNew: true,
    };

    const updatedAgendaItems = helper.calculateTimestamps(
      props.currentMeetingItem?.startTime ?? DateTime.now(),
      [...agendaItems, newAgendaItem]
    );

    setAgendaItems(updatedAgendaItems);
  };

  const handleDeleteAgendaItem = (id: number) => {
    const agendaItemToDelete = agendaItems.find((item) => item.id === id);

    // Build new agenda items array without the deleted item
    setAgendaItems((prevItems) => {
      const updatedAgendaItemArray = prevItems.filter((item) => item.id !== id);
      const newAgendaItems: AgendaItem[] = helper.calculateTimestamps(
        props.currentMeetingItem?.startTime ?? DateTime.now(),
        updatedAgendaItemArray
      );

      return newAgendaItems;
    });

    // Put the deleted item back into the unassigned items and filtered Items
    const unassignedItems = unassignedAgendaItems
      ? [...unassignedAgendaItems, agendaItemToDelete!]
      : [agendaItemToDelete!];

    setUnassignedAgendaItems(unassignedItems);
    setFilteredItems(unassignedItems);
    setSearchQuery("");
  };

  const handlePurgeClick = (item: AgendaItem) => {
    setItemToBePurged(item);
  };

  const handleConfirmPurge = async () => {
    if (itemToBePurged) {
      try {
        console.error("Purging agenda item: ", itemToBePurged);
        setPurgeLoading(true);
        await helper.callBackend("deleteAgendaItem", "POST", teamsUserCredential, {
          itemId: itemToBePurged.id,
          eventId: itemToBePurged.eventId,
          fileLocationId: itemToBePurged.fileLocationId,
        });

        const unassignedItems = await helper.callBackend(
          "getAgendaItems",
          "GET",
          teamsUserCredential,
          undefined,
          []
        );
        const unassignedAgendaItemsJson = JSON.parse(unassignedItems) as AgendaItem[];
        setUnassignedAgendaItems(unassignedAgendaItemsJson);
        setFilteredItems(unassignedAgendaItemsJson);
      } catch (error) {
        console.error("Error purging agenda item: ", error);
      } finally {
        setPurgeLoading(false);
      }
    }
    setItemToBePurged(null);
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
        {/* {agendaItems.length > 0 && ( */}
        <div>
          <div>
            <DndContext
              sensors={sensors}
              collisionDetection={closestCenter}
              onDragStart={handleDragStart}
              onDragEnd={handleDragEnd}
              modifiers={[restrictToVerticalAxis]}
            >
              <SortableContext
                items={agendaItems.map((item) => item.id)}
                strategy={rectSortingStrategy}
              >
                <Accordion collapsible openItems={openItems} onToggle={handleToggle}>
                  {agendaItems.map((item) => (
                    <SortableAgendaItem
                      key={item.id}
                      id={item.id}
                      title={item.title}
                      durationInMinutes={item.durationInMinutes}
                      isMisc={item.isMisc}
                      isDecisionNeeded={item.needsDecision}
                      selectedParticipants={item.additionalParticipants}
                      remarks={item.remarks || ""}
                      startTime={item.startTime ?? DateTime.now()}
                      onUpdate={handleUpdateAgendaItem}
                      handleDeleteAgendaItem={handleDeleteAgendaItem}

                    />
                  ))}
                </Accordion>
              </SortableContext>
            </DndContext>
          </div>
          {/* Button and Popover for SearchBox and results */}
          <div className={styles.buttonSpacingUnderAccordion}>
            <Tooltip
              positioning={"below"}
              content={"Neuen Agendapunkt erstellen"}
              relationship="label"
            >
              <Button
                appearance="primary"
                onClick={handleAddAgendaItem}
                disabled={loading}
                icon={<AddCircle24Filled />}
              ></Button>
            </Tooltip>
            <Popover open={openPopover} onOpenChange={handleOpenPopoverChange}>
              <PopoverTrigger>
                <Tooltip
                  positioning="below"
                  content="Vorhandenen Agendapunkt hinzufügen"
                  relationship="label"
                >
                  <Button
                    appearance="primary"
                    disabled={loading}
                    icon={<CollectionsAdd24Filled />}
                  />
                </Tooltip>
              </PopoverTrigger>
              <PopoverSurface>
                <SearchBox
                  placeholder="Suche nach Titel"
                  value={searchQuery}
                  onChange={handleSearchChange}
                />
                <div className="searchResultsContainer">
                  {filteredItems.length > 0 ? (
                    <>
                      {filteredItems
                        .sort((a, b) => a.title.localeCompare(b.title))
                        .slice(0, 7)
                        .map((item) => (
                          <div key={item.id} className={styles.resultItem}>
                            <Text
                              className={styles.title}
                              onClick={() => handleExistingItemClick(item)}
                            >
                              {item.title}
                            </Text>
                            <Tooltip content="Löschen" relationship="label">
                              <Button
                                appearance="subtle"
                                icon={<Delete24Regular />}
                                className={styles.deleteIcon}
                                onClick={() => handlePurgeClick(item)}
                              />
                            </Tooltip>
                          </div>
                        ))}
                      {filteredItems.length > 7 && (
                        <Text className={styles.moreResults}>
                          Mehr TOPs vorhanden, bitte Suche benutzen!
                        </Text>
                      )}
                    </>
                  ) : (
                    <Text>Nichts gefunden!</Text>
                  )}
                </div>
                {/* Delete Confirmation Overlay (inside Popover) */}
                {itemToBePurged && (
                  <div className={styles.deleteOverlay}>
                    <div className={styles.deleteBox}>
                      {purgeLoading && <Spinner />}
                      <Text>
                        <h2>Achtung</h2>
                        TOP "{itemToBePurged.title}" wirklich endgültig löschen?
                        <br />
                        Alle Dokumente in der Dateiablage des TOPs werden ebenfalls
                        gelöscht!
                      </Text>
                      <div className={styles.buttonGroup}>
                        <Button
                          onClick={() => setItemToBePurged(null)}
                          appearance="primary"
                        >
                          Abbrechen
                        </Button>
                        <Button
                          appearance="primary"
                          onClick={handleConfirmPurge}
                          className={styles.buttonDanger}
                        >
                          Löschen
                        </Button>
                      </div>
                    </div>
                  </div>
                )}
              </PopoverSurface>
            </Popover>
          </div>
          {/* Second button at the bottom right */}
          <div className={styles.buttonContainerBottom}>
            <Button appearance="primary" onClick={handleSubmit} disabled={loading}>
              Speichern
            </Button>
          </div>
        </div>
        {/* )} */}
        {!loading && error && <div style={{ bottom: "10px" }}>{error}</div>}
      </div>
    </>
  );
}