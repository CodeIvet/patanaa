import * as React from "react";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { Icon } from "@fluentui/react/lib/Icon"; // https://developer.microsoft.com/en-us/fluentui#/styles/web/icons
import { Tooltip, tokens } from "@fluentui/react-components";

const iconClassFont = mergeStyles({
  fontSize: 25,
  height: 25,
  width: 25,
  margin: "10px",
  color: tokens.colorBrandBackground,
  cursor: "pointer",
  selectors: {
    ":hover": {
      color: tokens.colorBrandBackgroundHover, // Fill color on hover
      fontWeight: "bold",
    },
  },
});

const iconClassFontDisabled = mergeStyles(
  iconClassFont, // Inherit everything from iconClassFont
  {
    color: "grey", // New fill color, overwrite the existing fill color
    cursor: "default",
    selectors: {
      ":hover": { color: "grey" }, // Fill color on hover},
    },
  }
);

export const IcJoinMeeting: React.FunctionComponent<{
  onClick: () => void;
  disabled?: boolean;
}> = ({ onClick, disabled = false }) => {
  const className = disabled ? iconClassFontDisabled : iconClassFont;
  return (
    <Tooltip
      positioning={"below"}
      content={
        disabled
          ? "Teams-Meeting wurde noch nicht erstellt"
          : "An Teams-Meeting teilnehmen"
      }
      relationship="label"
    >
      <span onClick={!disabled ? onClick : undefined}>
        <Icon iconName="JoinOnlineMeeting" className={className} />
      </span>
    </Tooltip>
  );
};

export const IcCalendarAgenda: React.FunctionComponent<{
  onClick: () => void;
  disabled?: boolean;
}> = ({ onClick, disabled = false }) => {
  const className = disabled ? iconClassFontDisabled : iconClassFont;
  return (
    <Tooltip
      positioning={"below"}
      content={disabled ? "n/a" : "Agenda bearbeiten"}
      relationship="label"
    >
      <span onClick={!disabled ? onClick : undefined}>
        <Icon iconName="FileCSS" className={className} />
      </span>
    </Tooltip>
  );
};

export const IcEdit: React.FunctionComponent<{
  onClick: () => void;
  disabled?: boolean;
}> = ({ onClick, disabled = false }) => {
  const className = disabled ? iconClassFontDisabled : iconClassFont;
  return (
    <Tooltip
      positioning={"below"}
      content={disabled ? "n/a" : "Sitzung bearbeiten"}
      relationship="label"
    >
      <span onClick={!disabled ? onClick : undefined}>
        <Icon iconName="Edit" className={className} />
      </span>
    </Tooltip>
  );
};

export const IcInvites: React.FunctionComponent<{
  onClick: () => void;
  disabled?: boolean;
}> = ({ onClick, disabled = false }) => {
  const className = disabled ? iconClassFontDisabled : iconClassFont;
  return (
    <Tooltip
      positioning={"below"}
      content={disabled ? "n/a" : "Outlook-Einladungen verwalten"}
      relationship="label"
    >
      <span onClick={!disabled ? onClick : undefined}>
        <Icon iconName="Mail" className={className} />
      </span>
    </Tooltip>
  );
};

export const IcFiles: React.FunctionComponent<{
  onClick: () => void;
  disabled?: boolean;
}> = ({ onClick, disabled = false }) => {
  const className = disabled ? iconClassFontDisabled : iconClassFont;
  return (
    <Tooltip
      positioning={"below"}
      content={disabled ? "n/a" : "Dateien verwalten"}
      relationship="label"
    >
      <span onClick={!disabled ? onClick : undefined}>
        <Icon iconName="Attach" className={className} />
      </span>
    </Tooltip>
  );
};

export const IcCreateProtocol: React.FunctionComponent<{
  onClick: () => void;
  disabled?: boolean;
}> = ({ onClick, disabled = false }) => {
  const className = disabled ? iconClassFontDisabled : iconClassFont;
  return (
    <Tooltip
      positioning={"below"}
      content={disabled ? "n/a" : "Protokoll (Word) / Agenda (PDF) erstellen"}
      relationship="label"
    >
      <span onClick={!disabled ? onClick : undefined}>
        <Icon iconName="PDF" className={className} />
      </span>
    </Tooltip>
  );
};