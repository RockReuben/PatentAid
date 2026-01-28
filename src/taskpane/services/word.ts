export interface TrackedChangeInfo {
  id: number;
  type: string;
  text: string;
  author: string;
  date: Date;
}

export async function getSelectedText(): Promise<string> {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text || "";
  });
}

export async function insertTextIntoWord(text: string): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertText(text, Word.InsertLocation.end);
    await context.sync();
  });
}

export async function replaceSelectedText(text: string): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text, Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function replaceWithTrackedChanges(newText: string): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");

    const document = context.document;
    document.load("changeTrackingMode");
    await context.sync();

    const previousMode = document.changeTrackingMode;

    document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();

    selection.insertText(newText, Word.InsertLocation.replace);
    await context.sync();

    document.changeTrackingMode = previousMode;
    await context.sync();
  });
}

export async function getTrackedChanges(): Promise<TrackedChangeInfo[]> {
  return await Word.run(async (context) => {
    const body = context.document.body;
    const trackedChanges = body.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();

    const changes: TrackedChangeInfo[] = [];
    for (let i = 0; i < trackedChanges.items.length; i++) {
      const change = trackedChanges.items[i];
      change.load(["type", "text", "author", "date"]);
    }
    await context.sync();

    for (let i = 0; i < trackedChanges.items.length; i++) {
      const change = trackedChanges.items[i];
      changes.push({
        id: i,
        type: change.type,
        text: change.text || "",
        author: change.author || "Unknown",
        date: change.date,
      });
    }

    return changes;
  });
}

export async function navigateToChange(index: number): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    const trackedChanges = body.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();

    if (index >= 0 && index < trackedChanges.items.length) {
      const change = trackedChanges.items[index];
      const range = change.getRange();
      range.select();
      await context.sync();
    }
  });
}

export async function acceptChange(index: number): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    const trackedChanges = body.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();

    if (index >= 0 && index < trackedChanges.items.length) {
      const change = trackedChanges.items[index];
      change.accept();
      await context.sync();
    }
  });
}

export async function rejectChange(index: number): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    const trackedChanges = body.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();

    if (index >= 0 && index < trackedChanges.items.length) {
      const change = trackedChanges.items[index];
      change.reject();
      await context.sync();
    }
  });
}

export async function acceptAllChanges(): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    const trackedChanges = body.getTrackedChanges();
    trackedChanges.acceptAll();
    await context.sync();
  });
}

export async function rejectAllChanges(): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    const trackedChanges = body.getTrackedChanges();
    trackedChanges.rejectAll();
    await context.sync();
  });
}
