import * as React from "react";
import {
  makeStyles,
  tokens,
  Textarea,
  Button,
  Spinner,
  Card,
  Text,
  Badge,
  shorthands,
  Tooltip,
} from "@fluentui/react-components";
import {
  Send24Regular,
  Sparkle24Regular,
  DocumentArrowRight24Regular,
  ArrowSwap24Regular,
  Checkmark24Regular,
  DocumentEdit24Regular,
  Attach24Regular,
  Dismiss12Regular,
  Document24Regular,
  ChevronDown12Regular,
  ChevronUp12Regular,
} from "@fluentui/react-icons";
import {
  sendToOpenAI,
  Message as OpenAIMessage,
  StructuredResponse,
  FileContext,
} from "../services/openai";
import {
  insertTextIntoWord,
  replaceWithTrackedChanges,
  getSelectedText,
  getTrackedChanges,
} from "../services/word";
import ChangeReviewer from "./ChangeReviewer";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  header: {
    display: "flex",
    alignItems: "center",
    ...shorthands.padding("16px"),
    ...shorthands.gap("12px"),
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  headerIcon: {
    color: tokens.colorBrandForeground1,
  },
  headerText: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
    flex: 1,
  },
  reviewButton: {
    position: "relative" as const,
  },
  changeBadge: {
    position: "absolute" as const,
    top: "-4px",
    right: "-4px",
    minWidth: "18px",
    height: "18px",
    borderRadius: "9px",
    backgroundColor: tokens.colorPaletteRedBackground3,
    color: tokens.colorNeutralForegroundOnBrand,
    fontSize: "10px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  selectionBanner: {
    ...shorthands.padding("8px", "16px"),
    backgroundColor: tokens.colorNeutralBackground4,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("8px"),
  },
  selectionText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    flex: 1,
  },
  messagesArea: {
    flex: 1,
    overflowY: "auto",
    ...shorthands.padding("16px"),
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("12px"),
  },
  messageCard: {
    maxWidth: "90%",
    ...shorthands.padding("12px", "16px"),
  },
  userMessage: {
    alignSelf: "flex-end",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  assistantMessage: {
    alignSelf: "flex-start",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  messageContent: {
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  messageContentCollapsed: {
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxHeight: "40px",
    overflow: "hidden",
    position: "relative" as const,
  },
  collapsedOverlay: {
    position: "absolute" as const,
    bottom: 0,
    left: 0,
    right: 0,
    height: "20px",
    background: "linear-gradient(transparent, inherit)",
  },
  messageCardClickable: {
    cursor: "pointer",
    ":hover": {
      opacity: 0.9,
    },
  },
  expandIndicator: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    ...shorthands.gap("4px"),
    marginTop: "4px",
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
  messageHeader: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("8px"),
    marginBottom: "8px",
  },
  messageActions: {
    display: "flex",
    justifyContent: "flex-end",
    ...shorthands.gap("8px"),
    marginTop: "8px",
    paddingTop: "8px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  emptyState: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    ...shorthands.padding("32px"),
    textAlign: "center",
    ...shorthands.gap("16px"),
  },
  emptyIcon: {
    fontSize: "48px",
    color: tokens.colorBrandForeground2,
  },
  emptyText: {
    color: tokens.colorNeutralForeground2,
    maxWidth: "280px",
  },
  inputArea: {
    ...shorthands.padding("16px"),
    backgroundColor: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  inputWrapper: {
    display: "flex",
    ...shorthands.gap("8px"),
    alignItems: "flex-end",
  },
  textarea: {
    flex: 1,
  },
  sendButton: {
    minWidth: "40px",
    height: "40px",
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  appliedBadge: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
  },
  filesBanner: {
    ...shorthands.padding("8px", "16px"),
    backgroundColor: tokens.colorNeutralBackground4,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    ...shorthands.gap("8px"),
  },
  fileChip: {
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("4px"),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.padding("4px", "8px"),
    ...shorthands.borderRadius("4px"),
    fontSize: tokens.fontSizeBase200,
  },
  fileChipRemove: {
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    ":hover": {
      color: tokens.colorPaletteRedForeground1,
    },
  },
  hiddenInput: {
    display: "none",
  },
});

interface Message {
  role: "user" | "assistant";
  content: string;
  documentContent?: string;
  action?: "none" | "insert" | "replace";
  applied?: boolean;
}

interface UploadedFile {
  name: string;
  content: string;
}

const PromptPane: React.FC = () => {
  const styles = useStyles();
  const [prompt, setPrompt] = React.useState("");
  const [messages, setMessages] = React.useState<Message[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [selectedText, setSelectedText] = React.useState<string>("");
  const [showReviewer, setShowReviewer] = React.useState(false);
  const [pendingChangesCount, setPendingChangesCount] = React.useState(0);
  const [uploadedFiles, setUploadedFiles] = React.useState<UploadedFile[]>([]);
  const [expandedMessages, setExpandedMessages] = React.useState<Set<number>>(new Set());
  const messagesEndRef = React.useRef<HTMLDivElement>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // Auto-collapse threshold - messages beyond this count get collapsed
  const AUTO_COLLAPSE_THRESHOLD = 4;
  const COLLAPSE_CHAR_THRESHOLD = 100; // Only collapse messages longer than this

  const apiKey = process.env.OPENAI_API_KEY;

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  React.useEffect(() => {
    scrollToBottom();
  }, [messages]);

  // Poll for selected text and tracked changes
  React.useEffect(() => {
    const checkDocumentState = async () => {
      try {
        const text = await getSelectedText();
        setSelectedText(text);

        const changes = await getTrackedChanges();
        setPendingChangesCount(changes.length);
      } catch {
        // Ignore errors when document isn't ready
      }
    };

    checkDocumentState();
    const interval = setInterval(checkDocumentState, 1000);
    return () => clearInterval(interval);
  }, []);

  const applyAction = async (structured: StructuredResponse): Promise<boolean> => {
    if (!structured.content) return false;

    try {
      if (structured.action === "insert") {
        await insertTextIntoWord(structured.content);
        return true;
      } else if (structured.action === "replace") {
        await replaceWithTrackedChanges(structured.content);
        // Auto-open reviewer after replace action
        setShowReviewer(true);
        return true;
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to modify document");
    }
    return false;
  };

  const handleSend = async () => {
    if (!prompt.trim() || isLoading) return;

    if (!apiKey) {
      setError("OPENAI_API_KEY is not configured. Please add it to your .env.local file.");
      return;
    }

    const userMessage = prompt.trim();
    const currentSelection = selectedText;
    setPrompt("");
    setError(null);
    setMessages((prev) => [...prev, { role: "user", content: userMessage }]);
    setIsLoading(true);

    // Build conversation history for context (use display messages only)
    const conversationHistory: OpenAIMessage[] = messages.map((m) => ({
      role: m.role,
      content: m.content,
    }));
    conversationHistory.push({ role: "user", content: userMessage });

    // Convert uploaded files to FileContext format
    const fileContext: FileContext[] = uploadedFiles.map((f) => ({
      name: f.name,
      content: f.content,
    }));

    const response = await sendToOpenAI(
      apiKey,
      conversationHistory,
      currentSelection,
      fileContext.length > 0 ? fileContext : undefined
    );

    if (response.error) {
      setError(response.error);
      setIsLoading(false);
      return;
    }

    let displayMessage: Message;

    if (response.structured) {
      const applied = await applyAction(response.structured);

      displayMessage = {
        role: "assistant",
        content: response.structured.message,
        documentContent: response.structured.content,
        action: response.structured.action,
        applied,
      };
    } else {
      // Fallback if parsing failed
      displayMessage = {
        role: "assistant",
        content: response.rawContent,
        action: "none",
      };
    }

    setMessages((prev) => [...prev, displayMessage]);
    setIsLoading(false);
  };

  const handleManualInsert = async (content: string, index: number) => {
    try {
      await insertTextIntoWord(content);
      setMessages((prev) =>
        prev.map((m, i) => (i === index ? { ...m, applied: true } : m))
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to insert into document");
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const content = event.target?.result as string;
        setUploadedFiles((prev) => [
          ...prev,
          { name: file.name, content },
        ]);
      };
      reader.readAsText(file);
    });

    // Reset input so same file can be uploaded again
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const handleRemoveFile = (index: number) => {
    setUploadedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const toggleMessageExpanded = (index: number) => {
    setExpandedMessages((prev) => {
      const next = new Set(prev);
      if (next.has(index)) {
        next.delete(index);
      } else {
        next.add(index);
      }
      return next;
    });
  };

  const shouldCollapseMessage = (index: number, content: string): boolean => {
    // Don't collapse if explicitly expanded
    if (expandedMessages.has(index)) return false;
    // Don't collapse short messages
    if (content.length < COLLAPSE_CHAR_THRESHOLD) return false;
    // Don't collapse the last 2 messages
    if (index >= messages.length - 2) return false;
    // Collapse older messages when we have many
    return messages.length > AUTO_COLLAPSE_THRESHOLD;
  };

  const getActionBadge = (message: Message) => {
    if (message.applied) {
      return (
        <Badge appearance="filled" color="success" icon={<Checkmark24Regular />}>
          Applied
        </Badge>
      );
    }
    if (message.action === "insert") {
      return (
        <Badge appearance="outline" color="informative">
          Insert
        </Badge>
      );
    }
    if (message.action === "replace") {
      return (
        <Badge appearance="outline" color="warning">
          Replace
        </Badge>
      );
    }
    return null;
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Sparkle24Regular className={styles.headerIcon} />
        <Text className={styles.headerText}>PatentAid Assistant</Text>
        {pendingChangesCount > 0 && (
          <Tooltip content="Review tracked changes" relationship="label">
            <Button
              className={styles.reviewButton}
              appearance="subtle"
              icon={<DocumentEdit24Regular />}
              onClick={() => setShowReviewer(true)}
            >
              <span className={styles.changeBadge}>{pendingChangesCount}</span>
            </Button>
          </Tooltip>
        )}
      </div>

      {selectedText && (
        <div className={styles.selectionBanner}>
          <ArrowSwap24Regular />
          <Text className={styles.selectionText}>
            Selected: "{selectedText.substring(0, 50)}{selectedText.length > 50 ? "..." : ""}"
          </Text>
        </div>
      )}

      {uploadedFiles.length > 0 && (
        <div className={styles.filesBanner}>
          <Document24Regular />
          <Text size={200}>Context files:</Text>
          {uploadedFiles.map((file, index) => (
            <span key={index} className={styles.fileChip}>
              {file.name}
              <span
                className={styles.fileChipRemove}
                onClick={() => handleRemoveFile(index)}
              >
                <Dismiss12Regular />
              </span>
            </span>
          ))}
        </div>
      )}

      {messages.length === 0 ? (
        <div className={styles.emptyState}>
          <Sparkle24Regular className={styles.emptyIcon} />
          <Text size={400} weight="semibold">
            How can I help you today?
          </Text>
          <Text className={styles.emptyText}>
            Describe your invention and I'll help draft the patent structure, claims, and detailed
            description. Select text in your document to modify it.
          </Text>
        </div>
      ) : (
        <div className={styles.messagesArea}>
          {messages.map((message, index) => {
            const isCollapsed = shouldCollapseMessage(index, message.content);
            const isCollapsible = message.content.length >= COLLAPSE_CHAR_THRESHOLD && index < messages.length - 2;

            return (
              <Card
                key={index}
                className={`${styles.messageCard} ${
                  message.role === "user" ? styles.userMessage : styles.assistantMessage
                } ${isCollapsible ? styles.messageCardClickable : ""}`}
                onClick={isCollapsible ? () => toggleMessageExpanded(index) : undefined}
              >
                {message.role === "assistant" && message.action !== "none" && (
                  <div className={styles.messageHeader}>{getActionBadge(message)}</div>
                )}
                <div style={{ position: "relative" }}>
                  <Text className={isCollapsed ? styles.messageContentCollapsed : styles.messageContent}>
                    {message.content}
                  </Text>
                </div>
                {isCollapsed && (
                  <div className={styles.expandIndicator}>
                    <ChevronDown12Regular />
                    <span>Click to expand</span>
                  </div>
                )}
                {!isCollapsed && isCollapsible && expandedMessages.has(index) && (
                  <div className={styles.expandIndicator}>
                    <ChevronUp12Regular />
                    <span>Click to collapse</span>
                  </div>
                )}
                {message.role === "assistant" && message.documentContent && !message.applied && !isCollapsed && (
                  <div className={styles.messageActions} onClick={(e) => e.stopPropagation()}>
                    <Tooltip content="Insert content into document" relationship="label">
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<DocumentArrowRight24Regular />}
                        onClick={() => handleManualInsert(message.documentContent!, index)}
                      >
                        Insert
                      </Button>
                    </Tooltip>
                  </div>
                )}
              </Card>
            );
          })}
          {isLoading && (
            <Card className={`${styles.messageCard} ${styles.assistantMessage}`}>
              <Spinner size="tiny" label="Drafting..." />
            </Card>
          )}
          <div ref={messagesEndRef} />
        </div>
      )}

      {error && (
        <div className={styles.inputArea} style={{ paddingBottom: 0 }}>
          <Text className={styles.errorText}>{error}</Text>
        </div>
      )}

      {showReviewer && (
        <ChangeReviewer
          onClose={() => setShowReviewer(false)}
          onChangesResolved={() => {
            setShowReviewer(false);
            setPendingChangesCount(0);
          }}
        />
      )}

      <div className={styles.inputArea}>
        <input
          type="file"
          ref={fileInputRef}
          className={styles.hiddenInput}
          onChange={handleFileUpload}
          multiple
          accept=".txt,.md,.pdf,.doc,.docx,.json,.xml,.csv"
        />
        <div className={styles.inputWrapper}>
          <Tooltip content="Attach reference files" relationship="label">
            <Button
              appearance="subtle"
              icon={<Attach24Regular />}
              onClick={() => fileInputRef.current?.click()}
              disabled={isLoading}
            />
          </Tooltip>
          <Textarea
            className={styles.textarea}
            placeholder={
              selectedText
                ? "What would you like to do with the selected text?"
                : "Describe your invention or ask for help..."
            }
            value={prompt}
            onChange={(_e, data) => setPrompt(data.value)}
            onKeyDown={handleKeyDown}
            resize="vertical"
            disabled={isLoading}
          />
          <Button
            className={styles.sendButton}
            appearance="primary"
            icon={<Send24Regular />}
            onClick={handleSend}
            disabled={!prompt.trim() || isLoading}
          />
        </div>
      </div>
    </div>
  );
};

export default PromptPane;
