// SortableAgendaItem.tsx
import React, { useState } from "react";
import { useSortable } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import {
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Button,
  Input,
  Switch,
  Field,
  tokens,
  makeStyles,
  Tooltip,
  Textarea,
  Label,
} from "@fluentui/react-components";
import { ReOrderRegular, BinRecycleRegular } from "@fluentui/react-icons";
import { PeoplePicker } from "@microsoft/mgt-react";
import { DateTime } from "luxon";

interface SortableAgendaItemProps {
  id: number;
  title: string;
  durationInMinutes: number;
  isMisc: boolean;
  isDecisionNeeded: boolean;
  selectedParticipants: string;
  startTime: DateTime;
  remarks: string;
  onUpdate: (
    id: number,
    newTitle: string,
    newDurationInMinutes: number,
    newIsMisc: boolean,
    newIsDecisionNeeded: boolean,
    newSelectedParticipants: string,
    startTime: DateTime,
    remarks?: string
  ) => void;
  handleDeleteAgendaItem: (id: number) => void;
}

const useStyles = makeStyles({
  peoplePicker: {
    "--people-picker-input-background": tokens.colorNeutralBackground1,
  },
});

const SortableAgendaItem: React.FC<SortableAgendaItemProps> = ({
  id,
  title,
  durationInMinutes,
  isMisc,
  isDecisionNeeded,
  selectedParticipants,
  startTime,
  remarks,
  onUpdate,
  handleDeleteAgendaItem,
}) => {
  const [isEditingTitle, setIsEditingTitle] = useState(false);
  const [isEditingDurationInMinutes, setIsEditingDurationInMinutes] = useState(false);
  const [newTitle, setNewTitle] = useState(title);
  const [newDurationInMinutes, setNewDurationInMinutes] = useState(durationInMinutes);
  const [newIsMisc, setNewIsMisc] = useState(isMisc);
  const [newIsDecisionNeeded, setNewIsDecisionNeeded] = useState(isDecisionNeeded);
  const [newSelectedParticipants, setNewSelectedParticipants] =
    useState(selectedParticipants);
  const [newRemarks, setNewRemarks] = useState(remarks);

  const { setNodeRef, transform, transition, isDragging, attributes, listeners } =
    useSortable({ id });

  const style: React.CSSProperties = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
    border: "1px solid #ccc",
    padding: "0px", // Remove padding here
    marginBottom: "8px",
    backgroundColor: tokens.colorNeutralBackground4, //"#999",
    borderRadius: "4px",
  };

  const styles = useStyles();

  const handleTitleClick = (e: React.MouseEvent) => {
    e.stopPropagation();
    setIsEditingTitle(true);
  };

  const handleDurationInMinutesClick = (e: React.MouseEvent) => {
    setIsEditingDurationInMinutes(true);
  };

  const handleTitleBlur = () => {
    setIsEditingTitle(false);
    if (newTitle !== title) {
      onUpdate(
        id,
        newTitle,
        newDurationInMinutes,
        newIsMisc,
        newIsDecisionNeeded,
        newSelectedParticipants,
        startTime,
        newRemarks
      );
    }
  };

  const handleRemarksBlur = () => {
    onUpdate(
      id,
      newTitle,
      newDurationInMinutes,
      newIsMisc,
      newIsDecisionNeeded,
      newSelectedParticipants,
      startTime,
      newRemarks
    );
  };

  const handleDurationInMinutesBlur = () => {
    setIsEditingDurationInMinutes(false);
    if (newDurationInMinutes !== durationInMinutes) {
      onUpdate(
        id,
        title,
        newDurationInMinutes,
        newIsMisc,
        newIsDecisionNeeded,
        newSelectedParticipants,
        startTime,
        newRemarks
      );
    }
  };

  const handleIsMiscChange = (
    event: React.ChangeEvent<HTMLInputElement>,
    data: { checked: boolean }
  ) => {
    setNewIsMisc(data.checked);
    onUpdate(
      id,
      newTitle,
      newDurationInMinutes,
      data.checked,
      newIsDecisionNeeded,
      newSelectedParticipants,
      startTime,
      newRemarks
    );
  };

  const handleIsDecisionNeededChange = (
    event: React.ChangeEvent<HTMLInputElement>,
    data: { checked: boolean }
  ) => {
    setNewIsDecisionNeeded(data.checked);
    onUpdate(
      id,
      newTitle,
      newDurationInMinutes,
      newIsMisc,
      data.checked,
      newSelectedParticipants,
      startTime,
      newRemarks
    );
  };

  const handleOnKeyDown = (e: any) => {
    if (e.key === " ") {
      if (e.target.tagName == "INPUT" || e.target.tagName == "TEXTAREA") {
        // If in an input field, add space manually after preventDefault
        e.preventDefault();
        const target = e.target;
        const cursorPos = target.selectionStart;
        target.value =
          target.value.slice(0, cursorPos) + " " + target.value.slice(cursorPos);
        target.setSelectionRange(cursorPos + 1, cursorPos + 1);
      }
    }
  };

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
    if (participantsString != newSelectedParticipants) {
      setNewSelectedParticipants(participantsString);
      onUpdate(
        id,
        newTitle,
        newDurationInMinutes,
        newIsMisc,
        newIsDecisionNeeded,
        participantsString,
        startTime,
        newRemarks
      );
    }
  };

  return (
    <div ref={setNodeRef} style={style}>
      <AccordionItem value={id} onKeyDownCapture={handleOnKeyDown}>
        <AccordionHeader>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              width: "100%",
            }}
          >
            {isEditingTitle ? (
              <Input
                value={newTitle}
                onChange={(e) => setNewTitle(e.target.value.trim())}
                onBlur={handleTitleBlur}
                autoFocus
                style={{ flexGrow: 1 }}
              />
            ) : (
              <span
                style={{ flexGrow: 1, cursor: "text" }}
                onClick={handleTitleClick}
                onMouseDown={(e) => e.stopPropagation()}
              >
                <strong>
                  {startTime.toLocaleString({
                    hour: "2-digit",
                    minute: "2-digit",
                    hour12: false,
                  })}
                </strong>
                {" - " + title}
              </span>
            )}
            <Tooltip
              positioning={"below"}
              content={"Agendapunkt entfernen"}
              relationship="label"
            >
              <Button
                icon={<BinRecycleRegular />}
                appearance="subtle"
                aria-label="Delete Item"
                style={{ marginLeft: "8px" }}
                onClick={() => handleDeleteAgendaItem(id)}
              />
            </Tooltip>
            <Button
              icon={<ReOrderRegular />}
              appearance="subtle"
              aria-label="Drag handle"
              {...attributes}
              {...listeners}
              style={{ marginLeft: "8px" }}
            />
          </div>
        </AccordionHeader>
        <AccordionPanel>
          <Field required orientation="horizontal" label="Dauer in Minuten:">
            {isEditingDurationInMinutes ? (
              <>
                <Input
                  value={newDurationInMinutes.toString()}
                  onChange={(e) => {
                    if (!isNaN(Number(e.target.value))) {
                      setNewDurationInMinutes(Number(e.target.value));
                    }
                  }}
                  onBlur={handleDurationInMinutesBlur}
                  autoFocus
                  style={{ width: "100%", alignItems: "center" }}
                  type="number"
                />
              </>
            ) : (
              <>
                <div
                  onClick={handleDurationInMinutesClick}
                  style={{
                    cursor: "pointer",
                    paddingBottom: "6px",
                    paddingTop: "6px",
                    paddingLeft: "8px",
                  }}
                >
                  {newDurationInMinutes}
                </div>
              </>
            )}
          </Field>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              width: "100%",
              paddingBottom: "6px",
            }}
          >
            <div
              id="eins"
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between", // Ensures internal alignment
                gap: "45px",
              }}
            >
              <Label>Entscheidung nötig:</Label>
              <div style={{ marginLeft: "auto" }}>
                <Switch
                  checked={newIsDecisionNeeded}
                  onChange={handleIsDecisionNeededChange}
                />
              </div>
            </div>
            <div
              id="zwei"
              style={{
                display: "flex",
                alignItems: "center",
                gap: "8px",
                justifyContent: "flex-end", // Ensures right alignment
              }}
            >
              <Label>ist "Verschiedenes":</Label>
              <Switch checked={newIsMisc} onChange={handleIsMiscChange} />
            </div>
          </div>
          <Field label="Zusätzliche Teilnehmer:" orientation="horizontal">
            <PeoplePicker
              className={styles.peoplePicker}
              selectionMode="multiple"
              type="person"
              placeholder="Personen wählen"
              allowAnyEmail
              selectionChanged={handlePeoplePickerChanged}
              selectedPeople={
                newSelectedParticipants && newSelectedParticipants.trim() !== ""
                  ? newSelectedParticipants.split(";").map((email) => ({
                      displayName: email.split("@")[0], // Extracts the part before the '@' symbol
                      email: email.trim(), // Trims any whitespace
                      userPrincipalName: email.trim(), // Trims any whitespace
                    }))
                  : []
              }
            />
          </Field>
          <Field label="Anmerkungen" style={{ paddingBottom: "10px" }}>
            <Textarea
              style={{ height: "100px" }}
              value={newRemarks}
              onChange={(_, data) => {
                setNewRemarks(data.value);
              }}
              onBlur={handleRemarksBlur}
            />
          </Field>
        </AccordionPanel>
      </AccordionItem>
    </div>
  );
};

export default SortableAgendaItem;