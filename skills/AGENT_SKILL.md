---
name: google-docs
description: |
  Google Docs & Sheets MCP server (local, via mcporter). 48 tools for document reading, editing, formatting, table operations, undo/redo snapshots, comments, Drive file management, and Sheets. Use when users want to read, create, edit, or manage Google Docs/Sheets/Drive content.
compatibility: Requires google-docs MCP server configured in your MCP client
metadata:
  author: thebtf
  version: "2.0"
---

# Google Docs MCP — Agent Skill

Local MCP server running via stdio (`node dist/server.js`). Configure in your MCP client (Claude Code, OpenClaw mcporter, etc.) as `google-docs`.

All tools are MCP tools called via the standard tool interface (not HTTP). The `documentId` parameter is always the Google Doc ID from the URL.

---

## Tool Catalog (48 tools)

### 1. Reading & Navigation

| Tool | Purpose | Key Params |
|------|---------|------------|
| `readGoogleDoc` | Read document content (text, markdown, or structured) | `documentId`, `outputFormat?` ('text'\|'markdown'\|'structured'), `maxLength?` |
| `listDocumentTabs` | List all tabs with hierarchy | `documentId` |
| `getDocumentInfo` | Document metadata (owner, dates, sharing) | `documentId` |
| `getTableStructure` | All tables: dimensions, headers, indices | `documentId` |
| `readTableCells` | 2D array of cell values + metadata | `documentId`, `tableIndex` |
| `readTableCellsFormatted` | Per-cell FormattedRun[] + ImageInfo[] with dimensions | `documentId`, `tableIndex` |

### 2. Writing & Editing

| Tool | Purpose | Key Params |
|------|---------|------------|
| `insertText` | Insert text at index | `documentId`, `text`, `index` |
| `appendToGoogleDoc` | Append text to end | `documentId`, `text` |
| `deleteRange` | Delete content range | `documentId`, `startIndex`, `endIndex` |
| `replaceDocumentWithMarkdown` | Replace entire doc with Markdown | `documentId`, `markdown` |
| `appendMarkdownToGoogleDoc` | Append Markdown to end | `documentId`, `markdown` |

### 3. Formatting

| Tool | Purpose | Key Params |
|------|---------|------------|
| `applyTextStyle` | Bold, italic, colors, fonts, links | `documentId`, `target` (range or text find), `style` |
| `applyParagraphStyle` | Alignment, spacing, heading styles | `documentId`, `target`, `style` |
| `formatMatchingText` | Find text + apply char formatting | `documentId`, `textToFind`, formatting params |
| `fixListFormatting` | EXPERIMENTAL: convert text lists to native lists | `documentId`, `range?` |

### 4. Table Operations

| Tool | Purpose | Key Params |
|------|---------|------------|
| `insertTable` | Create new table | `documentId`, `rows`, `columns`, `index` |
| `editTableCell` | Replace text in one cell | `documentId`, `tableIndex`, `row`, `col`, `text` |
| `batchEditTableCells` | Bulk edit up to 500 cells | `documentId`, `tableIndex`, `edits[]` |
| `fillTableFromData` | Fill table from 2D string array | `documentId`, `tableIndex`, `data[][]`, `startRow?`, `startCol?`, `skipEmpty?` |
| `batchEditTableCellsFormatted` | Write cells with per-run formatting | `documentId`, `tableIndex`, `cells[]` with `runs[]` |
| `insertImageInTableCell` | Image in cell with optional size | `documentId`, `tableIndex`, `row`, `col`, `imageUrl`, `width?`, `height?` |
| `batchInsertImagesInTable` | Batch insert up to 50 images | `documentId`, `tableIndex`, `images[]` |
| `findTableRow` | Search rows by column content | `documentId`, `tableIndex`, `searchColumn`, `searchText` |
| `addTableRow` | Insert row after index | `documentId`, `tableIndex`, `insertBelow` |

### 5. Undo/Redo Snapshots

| Tool | Purpose | Key Params |
|------|---------|------------|
| `createDocumentSnapshot` | Save current state to undo stack | `documentId`, `label?` |
| `undoLastChange` | Restore most recent snapshot | `documentId` |
| `redoLastChange` | Re-apply last undone change | `documentId` |
| `listDocumentSnapshots` | List undo/redo stacks | `documentId` |

### 6. Images & Page Layout

| Tool | Purpose | Key Params |
|------|---------|------------|
| `insertImageFromUrl` | Insert image from public URL | `documentId`, `imageUrl`, `index`, `width?`, `height?` |
| `insertLocalImage` | Upload local file + insert | `documentId`, `localImagePath`, `index`, `width?`, `height?` |
| `insertPageBreak` | Insert page break | `documentId`, `index` |

### 7. Comments

| Tool | Purpose | Key Params |
|------|---------|------------|
| `listComments` | All comments with anchors | `documentId` |
| `getComment` | Single comment + replies | `documentId`, `commentId` |
| `addComment` | Add comment to text range | `documentId`, `startIndex`, `endIndex`, `commentText` |
| `replyToComment` | Reply to comment | `documentId`, `commentId`, `replyText` |
| `resolveComment` | Mark comment resolved | `documentId`, `commentId` |
| `deleteComment` | Delete comment | `documentId`, `commentId` |

### 8. Google Drive

| Tool | Purpose | Key Params |
|------|---------|------------|
| `listGoogleDocs` | List docs with filtering | `maxResults?`, `query?`, `orderBy?` |
| `searchGoogleDocs` | Search by name/content | `searchQuery`, `searchIn?`, `maxResults?` |
| `getRecentGoogleDocs` | Recently modified docs | `maxResults?`, `daysBack?` |
| `createDocument` | Create new doc | `title`, `parentFolderId?`, `initialContent?` |
| `createFromTemplate` | Copy template + replacements | `templateId`, `newTitle`, `replacements?` |
| `createFolder` | Create Drive folder | `name`, `parentFolderId?` |
| `listFolderContents` | List folder items | `folderId` |
| `getFolderInfo` | Folder metadata | `folderId` |
| `moveFile` | Move file/folder | `fileId`, `newParentId` |
| `copyFile` | Copy file | `fileId`, `newName?`, `parentFolderId?` |
| `renameFile` | Rename file/folder | `fileId`, `newName` |
| `deleteFile` | Delete (trash or permanent) | `fileId`, `skipTrash?` |

### 9. Google Sheets

| Tool | Purpose | Key Params |
|------|---------|------------|
| `readSpreadsheet` | Read range | `spreadsheetId`, `range` |
| `writeSpreadsheet` | Write to range | `spreadsheetId`, `range`, `values[][]` |
| `appendSpreadsheetRows` | Append rows | `spreadsheetId`, `range`, `values[][]` |
| `clearSpreadsheetRange` | Clear range | `spreadsheetId`, `range` |
| `getSpreadsheetInfo` | Spreadsheet + sheet metadata | `spreadsheetId` |
| `addSpreadsheetSheet` | Add sheet/tab | `spreadsheetId`, `sheetTitle` |
| `createSpreadsheet` | Create new spreadsheet | `title`, `parentFolderId?`, `initialData?` |
| `listGoogleSheets` | List spreadsheets | `maxResults?`, `query?` |

---

## Critical Workflows

### Workflow 1: Safe Destructive Editing (ALWAYS use this pattern)

Before ANY destructive operation, create a snapshot:

```
1. createDocumentSnapshot(documentId, label="before <operation>")
2. <destructive operation> (replaceDocumentWithMarkdown, deleteRange, batchEditTableCells, etc.)
3. If result is wrong → undoLastChange(documentId)
4. If undo was wrong → redoLastChange(documentId)
```

**Destructive operations that REQUIRE a prior snapshot:**
- `replaceDocumentWithMarkdown`
- `deleteRange`
- `batchEditTableCells` (bulk overwrite)
- `batchEditTableCellsFormatted`
- `fillTableFromData`

### Workflow 2: Table Discovery → Read → Edit

```
1. getTableStructure(documentId)          → find tableIndex, row/col counts
2. readTableCells(documentId, tableIndex) → see current values
3. editTableCell / batchEditTableCells    → make changes
```

For formatted content:
```
1. readTableCellsFormatted(documentId, tableIndex) → FormattedRun[] per cell
2. batchEditTableCellsFormatted(documentId, tableIndex, cells) → write with formatting
```

### Workflow 3: Copy Table Formatting Between Docs

```
1. readTableCellsFormatted(sourceDocId, tableIndex)
   → returns { cells: [[{ runs, imageInfo }]] }
2. For each cell with imageInfo: note width/height (PT units)
3. batchEditTableCellsFormatted(targetDocId, tableIndex, cells)
   → writes text + formatting
4. batchInsertImagesInTable(targetDocId, tableIndex, images)
   → images[] with { row, col, imageUrl, width, height }
```

### Workflow 4: Markdown Document Replacement

```
1. createDocumentSnapshot(documentId, label="before markdown replace")
2. replaceDocumentWithMarkdown(documentId, markdown)
3. Verify with readGoogleDoc(documentId)
4. If bad → undoLastChange(documentId)
```

### Workflow 5: Spreadsheet Operations

```
1. getSpreadsheetInfo(spreadsheetId)           → see sheets/tabs
2. readSpreadsheet(spreadsheetId, "Sheet1!A1:Z100") → read data
3. writeSpreadsheet(spreadsheetId, "A1:C10", values) → overwrite range
4. appendSpreadsheetRows(spreadsheetId, "A1", values) → add rows at end
```

---

## Key Constraints

### Snapshot System
- **In-memory only** — snapshots are lost when MCP server restarts
- **Max 10 per document** — oldest evicted automatically
- **Image URIs expire** — Google contentUri has limited lifetime; restore soon after snapshot
- **Lists not preserved** — bullet/numbered list native formatting is not restored
- **Table cell backgrounds/borders not preserved** — only text content and inline formatting

### Table Operations
- All indices are **0-based** (tableIndex, row, col)
- `batchEditTableCells` max 500 edits, executed in chunks of 50
- `batchInsertImagesInTable` max 50 images per call
- Images require **public URLs** (not Google Drive private links)
- `readTableCellsFormatted` returns `imageInfo[]` with `{ uri, width?, height? }` in PT

### Document Indices
- Text indices are **1-based** (document starts at index 1)
- Use `readGoogleDoc` with `outputFormat: 'structured'` to see indices
- Always get fresh indices before editing (indices shift after inserts/deletes)

### API Limits
- Google Docs batch update: max ~50 requests per call (auto-chunked)
- Rate limit: ~10 requests/second
- Large documents may need multiple API calls for complex operations

---

## Error Recovery

| Error | Cause | Fix |
|-------|-------|-----|
| "Table index N out of range" | Wrong tableIndex | Use `getTableStructure` first |
| "Row/Column out of range" | Bad row/col | Use `readTableCells` to check dimensions |
| "Document not found (404)" | Bad documentId | Verify ID from URL |
| "Permission denied (403)" | No edit access | Check sharing settings |
| "No snapshots available" | No prior `createDocumentSnapshot` | Always snapshot before destructive ops |
| "Invalid request (400)" | Stale indices after edit | Re-read document to get fresh indices |

---

## Examples

### Read and summarize a doc
```
readGoogleDoc(documentId="1abc...", outputFormat="markdown")
```

### Edit specific table cell
```
getTableStructure(documentId="1abc...")
→ Table 0: 10 rows x 5 cols

editTableCell(documentId="1abc...", tableIndex=0, row=3, col=2, text="Updated value")
```

### Bulk populate a table
```
fillTableFromData(documentId="1abc...", tableIndex=0, data=[
  ["Name", "Role", "Email"],
  ["Alice", "Dev", "alice@ex.com"],
  ["Bob", "PM", "bob@ex.com"]
], startRow=0)
```

### Safe markdown replacement
```
createDocumentSnapshot(documentId="1abc...", label="before update")
replaceDocumentWithMarkdown(documentId="1abc...", markdown="# New Content\n\nUpdated document.")
readGoogleDoc(documentId="1abc...", outputFormat="text")
// If wrong:
undoLastChange(documentId="1abc...")
```
