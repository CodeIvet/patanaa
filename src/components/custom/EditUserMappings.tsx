import * as React from "react";
import {
  DataGrid,
  DataGridHeader,
  DataGridRow,
  DataGridHeaderCell,
  DataGridBody,
  DataGridCell,
  TableColumnDefinition,
  Input,
  Button,
  makeStyles,
  Tooltip,
  Spinner,
  Text,
} from "@fluentui/react-components";
import { AddCircle24Filled, BinRecycleRegular } from "@fluentui/react-icons";
import { useContext } from "react";
import { TeamsFxContext } from "../Context";
import { callBackend } from "./lib/helper";

type UserMapping = {
  upn: string;
  displayName: string;
  id: string;
};

const useStyles = makeStyles({
  buttonContainerBottom: {
    display: "flex",
    justifyContent: "flex-end",
    paddingTop: "20px",
  },
});

const getUpnDuplicates = (users: UserMapping[]): string[] => {
  const upnCounts: Record<string, number> = {};

  for (const user of users) {
    const upn = user.upn.toLowerCase(); // Case-insensitive check
    upnCounts[upn] = (upnCounts[upn] || 0) + 1;
  }

  const duplicates = Object.entries(upnCounts)
    .filter(([_, count]) => count > 1)
    .map(([upn]) => upn);

  return duplicates;
};

// This Input keeps focus and cursor by using local state
const EditableCell = React.memo(
  ({
    initialValue,
    onCommit,
    loading,
  }: {
    initialValue: string;
    onCommit: (value: string) => void;
    loading: boolean;
  }) => {
    const [localValue, setLocalValue] = React.useState(initialValue);

    // Sync when value changes externally
    React.useEffect(() => {
      setLocalValue(initialValue);
    }, [initialValue]);

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
      setLocalValue(e.target.value);
    };

    const handleBlur = () => {
      if (localValue !== initialValue) {
        onCommit(localValue);
      }
    };

    return (
      <Input
        value={localValue}
        onChange={handleChange}
        onBlur={handleBlur}
        disabled={loading}
        style={{ width: "100%" }}
      />
    );
  }
);

export function EditUserMappings(props: {
  isSmallScreen: boolean;
  setOpenDialogEditUserMappings: (open: boolean) => void;
  setPreventDialogClose: (open: boolean) => void;
}) {
  const [data, setData] = React.useState<UserMapping[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [upnDuplicates, setUpnDuplicates] = React.useState<string[]>([]);
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const styles = useStyles();

  React.useEffect(() => {
    async function fetchData() {
      try {
        setLoading(true);
        const userMappings = await callBackend(
          "getUserMappings",
          "GET",
          teamsUserCredential
        );

        setData(userMappings);
        setLoading(false);

        console.log("Groups Response:", userMappings);
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    }

    fetchData();
  }, []);

  React.useEffect(() => {
    setUpnDuplicates(getUpnDuplicates(data));
  }, [data]);

  const handleEdit = React.useCallback(
    (id: string, key: keyof UserMapping, value: string) => {
      setData((prev) =>
        prev.map((item) =>
          item.id === id && item[key] !== value ? { ...item, [key]: value } : item
        )
      );
    },
    []
  );

  const handleAdd = () => {
    const newEntry: UserMapping = {
      upn: "max.mustermann@beispiel.com",
      displayName: "Max Mustermann",
      id: crypto.randomUUID(),
    };

    setData([...data, newEntry]);
  };

  const handleDelete = (id: string): void => {
    setData((prev) => prev.filter((item) => item.id !== id));
  };

  const handleSubmit = async () => {
    try {
      setLoading(true);
      props.setPreventDialogClose(true);
      const bulkInsert = await callBackend(
        "updateUserMappings",
        "POST",
        teamsUserCredential,
        data
      );
      setLoading(false);
    } catch (error) {
      console.error("Error submitting data:", error);
    } finally {
      setLoading(false);
      props.setPreventDialogClose(false);
      props.setOpenDialogEditUserMappings(false);
    }
  };

  const columns: TableColumnDefinition<UserMapping>[] = [
    {
      columnId: "UPN",
      compare: (a, b) => a.upn.localeCompare(b.upn),
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
          <DataGridCell className="cell">UPN</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => (
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
        >
          <EditableCell
            initialValue={item.upn}
            onCommit={(val) => handleEdit(item.id, "upn", val)}
            loading={loading}
          />
        </DataGridCell>
      ),
    },
    {
      columnId: "displayName",
      compare: (a, b) => a.displayName.localeCompare(b.displayName),
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
          <DataGridCell className="cell">Anzeigename</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => (
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
        >
          <EditableCell
            initialValue={item.displayName}
            onCommit={(val) => handleEdit(item.id, "displayName", val)}
            loading={loading}
          />
        </DataGridCell>
      ),
    },
    {
      columnId: "deleteEntry",
      compare: (a, b) => a.displayName.localeCompare(b.displayName),
      renderHeaderCell: () => (
        <DataGridHeaderCell
          style={{
            fontWeight: "bold",
            minWidth: "80px",
            maxWidth: "80px",
            paddingLeft: "5px",
          }}
        >
          <DataGridCell className="cell">Löschen</DataGridCell>
        </DataGridHeaderCell>
      ),
      renderCell: (item) => (
        <DataGridCell
          className="cell"
          style={{
            minWidth: "80px",
            maxWidth: "80px",
            paddingLeft: "10px",
            cursor: "pointer",
          }}
        >
          <Tooltip positioning="below" content="Eintrag löschen" relationship="label">
            <Button
              icon={<BinRecycleRegular />}
              appearance="subtle"
              aria-label="Delete Item"
              onClick={() => handleDelete(item.id)}
            />
          </Tooltip>
        </DataGridCell>
      ),
    },
  ];

  return (
    <div>
      {loading && <Spinner />}
      <DataGrid items={data} columns={columns} getRowId={(item) => item.id}>
        <DataGridHeader>
          <DataGridRow>{({ renderHeaderCell }) => <>{renderHeaderCell()}</>}</DataGridRow>
        </DataGridHeader>
        <DataGridBody<UserMapping>>
          {({ item }) => (
            <DataGridRow<UserMapping> key={item.id}>
              {({ renderCell }) => <>{renderCell(item)}</>}
            </DataGridRow>
          )}
        </DataGridBody>
      </DataGrid>

      {upnDuplicates.length > 0 && (
        <Text style={{ color: "red" }}>
          UPN bereits vorhanden! Bitte eindeutige UPNs verwenden: <br />
          {upnDuplicates.join("; ")}
        </Text>
      )}

      {/* Add New Entry */}
      <div style={{ marginTop: "10px" }}>
        <Button
          appearance="primary"
          onClick={handleAdd}
          disabled={loading}
          icon={<AddCircle24Filled />}
        />
      </div>

      <div className={styles.buttonContainerBottom}>
        <Button
          appearance="primary"
          disabled={loading || upnDuplicates.length > 0}
          onClick={handleSubmit}
        >
          Speichern
        </Button>
      </div>
    </div>
  );
}