# PatentAid
A Microsoft Word Add-in to empower law gods to patent with the best of 'em. Implemented with Office.js, scaffolding generated via [MS Yeoman Generator](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview).

![](https://github.com/RockReuben/PatentAid/blob/main/gif.gif)
## Setup

1. Create `.env.local` with your OpenAI key:
   ```
   OPENAI_API_KEY=sk-...
   ```

2. Install and run:
   ```bash
   pnpm install
   pnpm start
   ```

This starts word and the add-in server, then sideloads the add-in in Word.


## Features

### 1. Patent Aid
Send prompts to generate patent content following standard USPTO format (Title, Field, Background, Summary, Claims, Abstract, etc.).

**Code Flow:** `User prompt → sendToOpenAI() → JSON response {action, content, message} → Display in chat`

Responses then simply respond, insert, or replace content in the word doc for you (see `src/taskpane/services/openai.ts` for the prompt and its action/content/message json return format)

**Insertion Code Flow:** `action: "insert" → insertTextIntoWord() → Word.body.insertText()`

### 2. Review Changes
Navigate and accept/reject tracked changes individually or in bulk. See `ChangeReviewer.tsx`.

**Flow:** `getTrackedChanges() → Load items → navigateToChange() selects range → acceptChange()/rejectChange() → Refresh list`

### 3. Modify Selected Text 
When text is selected, it's added as context to the prompt sent to the LLM API; AI can replace it with improvements using Word's Track Changes.

**Flow:** `User selects text → action: "replace" → replaceWithTrackedChanges() → Enables trackAll mode → Replaces selection → Restores previous mode → Opens ChangeReviewer`


### 4. File Upload Context (only works with .txt files for now!)
Attach reference files (one-pagers, specs) to provide context for all prompts.

**Flow:** `FileReader.readAsText() → Store in uploadedFiles[] → Prepend to prompt as [REFERENCE FILES] block`

## Files of note

| File | Purpose |
|------|---------|
| `openai.ts` | OpenAI API calls, structured response parsing |
| `word.ts` | Word.run() operations: insert, replace, track changes |
| `PromptPane.tsx` | Chat interface, file upload, message collapse logic |
| `ChangeReviewer.tsx` | Iterate tracked changes with accept/reject controls |
| `webpack.config.js` | Bundles app, injects `OPENAI_API_KEY` from `.env.local` |
