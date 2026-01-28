import * as React from "react";
import {
  makeStyles,
  tokens,
  Button,
  Text,
  Card,
  Badge,
  shorthands,
  Divider,
} from "@fluentui/react-components";
import {
  ChevronLeft24Regular,
  ChevronRight24Regular,
  Checkmark24Regular,
  Dismiss24Regular,
  CheckmarkCircle24Regular,
  DismissCircle24Regular,
} from "@fluentui/react-icons";
import {
  getTrackedChanges,
  navigateToChange,
  acceptChange,
  rejectChange,
  acceptAllChanges,
  rejectAllChanges,
  TrackedChangeInfo,
} from "../services/word";

const useStyles = makeStyles({
  container: {
    ...shorthands.padding("12px"),
    backgroundColor: tokens.colorNeutralBackground3,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    maxHeight: "50vh",
    overflowY: "auto",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
    flexWrap: "wrap",
    ...shorthands.gap("8px"),
  },
  navigation: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("4px"),
  },
  changeCard: {
    ...shorthands.padding("10px"),
    marginBottom: "8px",
  },
  changeText: {
    fontFamily: "monospace",
    fontSize: tokens.fontSizeBase200,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxHeight: "120px",
    overflowY: "auto",
    ...shorthands.padding("8px"),
    ...shorthands.borderRadius("4px"),
  },
  actions: {
    display: "flex",
    ...shorthands.gap("8px"),
    marginTop: "8px",
    flexWrap: "wrap",
  },
  bulkActions: {
    display: "flex",
    ...shorthands.gap("8px"),
    justifyContent: "flex-end",
    flexWrap: "wrap",
  },
  emptyState: {
    textAlign: "center",
    color: tokens.colorNeutralForeground2,
    ...shorthands.padding("8px"),
  },
  changeHeader: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("8px"),
    marginBottom: "8px",
    flexWrap: "wrap",
  },
});

interface ChangeReviewerProps {
  onClose: () => void;
  onChangesResolved: () => void;
}

const ChangeReviewer: React.FC<ChangeReviewerProps> = ({ onClose, onChangesResolved }) => {
  const styles = useStyles();
  const [changes, setChanges] = React.useState<TrackedChangeInfo[]>([]);
  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [isLoading, setIsLoading] = React.useState(true);

  const loadChanges = React.useCallback(async () => {
    setIsLoading(true);
    try {
      const trackedChanges = await getTrackedChanges();
      setChanges(trackedChanges);
      if (trackedChanges.length === 0) {
        onChangesResolved();
      } else if (currentIndex >= trackedChanges.length) {
        setCurrentIndex(Math.max(0, trackedChanges.length - 1));
      }
    } catch (err) {
      console.error("Failed to load tracked changes:", err);
    }
    setIsLoading(false);
  }, [currentIndex, onChangesResolved]);

  React.useEffect(() => {
    loadChanges();
  }, [loadChanges]);

  React.useEffect(() => {
    if (changes.length > 0 && currentIndex < changes.length) {
      navigateToChange(currentIndex).catch(console.error);
    }
  }, [currentIndex, changes.length]);

  const handlePrevious = () => {
    if (currentIndex > 0) {
      setCurrentIndex(currentIndex - 1);
    }
  };

  const handleNext = () => {
    if (currentIndex < changes.length - 1) {
      setCurrentIndex(currentIndex + 1);
    }
  };

  const handleAccept = async () => {
    try {
      await acceptChange(currentIndex);
      await loadChanges();
    } catch (err) {
      console.error("Failed to accept change:", err);
    }
  };

  const handleReject = async () => {
    try {
      await rejectChange(currentIndex);
      await loadChanges();
    } catch (err) {
      console.error("Failed to reject change:", err);
    }
  };

  const handleAcceptAll = async () => {
    try {
      await acceptAllChanges();
      await loadChanges();
    } catch (err) {
      console.error("Failed to accept all changes:", err);
    }
  };

  const handleRejectAll = async () => {
    try {
      await rejectAllChanges();
      await loadChanges();
    } catch (err) {
      console.error("Failed to reject all changes:", err);
    }
  };

  if (isLoading) {
    return (
      <div className={styles.container}>
        <Text>Loading changes...</Text>
      </div>
    );
  }

  if (changes.length === 0) {
    return (
      <div className={styles.container}>
        <div className={styles.header}>
          <Text weight="semibold">Review Changes</Text>
          <Button appearance="subtle" icon={<Dismiss24Regular />} onClick={onClose} />
        </div>
        <div className={styles.emptyState}>
          <Text>No tracked changes to review.</Text>
        </div>
      </div>
    );
  }

  const currentChange = changes[currentIndex];
  // Check for deletion instead - Word API may return "Deleted", "deleted", or enum
  const typeStr = String(currentChange.type).toLowerCase();
  const isInsertion = !typeStr.includes("delete"); // catching the insertion was harder than deletion (maybe it was addition it sends back?)

  const textStyle: React.CSSProperties = isInsertion
    ? {
        backgroundColor: "#dff6dd",
        color: "#107c10",
        borderLeft: "3px solid #107c10",
      }
    : {
        backgroundColor: "#fde7e9",
        color: "#a80000",
        textDecoration: "line-through",
        borderLeft: "3px solid #a80000",
      };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text weight="semibold">Review Changes</Text>
        <div className={styles.navigation}>
          <Button
            appearance="subtle"
            size="small"
            icon={<ChevronLeft24Regular />}
            onClick={handlePrevious}
            disabled={currentIndex === 0}
          />
          <Text size={200}>
            {currentIndex + 1}/{changes.length}
          </Text>
          <Button
            appearance="subtle"
            size="small"
            icon={<ChevronRight24Regular />}
            onClick={handleNext}
            disabled={currentIndex === changes.length - 1}
          />
          <Button appearance="subtle" size="small" icon={<Dismiss24Regular />} onClick={onClose} />
        </div>
      </div>

      <Card className={styles.changeCard}>
        <div className={styles.changeHeader}>
          <Badge
            appearance="filled"
            color={isInsertion ? "success" : "danger"}
          >
            {isInsertion ? "Addition" : "Deletion"}
          </Badge>
          <Text size={100} style={{ color: "#666" }}>
            {currentChange.author}
          </Text>
        </div>
        <div className={styles.changeText} style={textStyle}>
          {currentChange.text || "(empty)"}
        </div>
        <div className={styles.actions}>
          <Button
            appearance="primary"
            size="small"
            icon={<Checkmark24Regular />}
            onClick={handleAccept}
          >
            Accept
          </Button>
          <Button
            appearance="secondary"
            size="small"
            icon={<Dismiss24Regular />}
            onClick={handleReject}
          >
            Reject
          </Button>
        </div>
      </Card>

      <Divider />

      <div className={styles.bulkActions} style={{ marginTop: "8px" }}>
        <Button
          appearance="subtle"
          size="small"
          icon={<CheckmarkCircle24Regular />}
          onClick={handleAcceptAll}
        >
          Accept All
        </Button>
        <Button
          appearance="subtle"
          size="small"
          icon={<DismissCircle24Regular />}
          onClick={handleRejectAll}
        >
          Reject All
        </Button>
      </div>
    </div>
  );
};

export default ChangeReviewer;
