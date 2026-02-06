// src/snapshotManager.ts
// Document snapshot management for undo/redo functionality.
// Stores full document body snapshots in-memory and restores via API.

import { docs_v1 } from 'googleapis';
import { UserError } from 'fastmcp';
import { TextStyleArgs } from './types.js';
import { buildUpdateTextStyleRequest, executeBatchUpdate, executeBatchUpdateChunked } from './googleDocsApiHelpers.js';
import {
  extractTableElements,
  extractFormattedCellContent,
  buildBatchEditCellRequests,
  buildFormattedCellFormatRequests,
  buildBatchInsertImageRequests,
  getCellElement,
  getCellInsertionPoint,
  rgbToHex,
  type FormattedRun,
  type ImageInfo,
  type CellEdit,
  type CellImageInsert,
} from './tableHelpers.js';

// --- Types ---

export interface DocumentSnapshot {
  id: string;
  documentId: string;
  timestamp: number;
  label: string;
  body: docs_v1.Schema$StructuralElement[];
  inlineObjects: Record<string, docs_v1.Schema$InlineObject>;
}

interface SnapshotStack {
  undoStack: DocumentSnapshot[];
  redoStack: DocumentSnapshot[];
}

// --- State ---

const MAX_SNAPSHOTS = 10;
const snapshotStacks = new Map<string, SnapshotStack>();

function getStack(documentId: string): SnapshotStack {
  if (!snapshotStacks.has(documentId)) {
    snapshotStacks.set(documentId, { undoStack: [], redoStack: [] });
  }
  return snapshotStacks.get(documentId)!;
}

function generateId(): string {
  return `snap_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
}

// --- Paragraph Content Extraction ---

export interface ParagraphContent {
  runs: FormattedRun[];
  imageInfo: Array<ImageInfo & { offset: number }>;
  namedStyle?: string;
}

/**
 * Extracts formatted content from a paragraph element.
 * TypeScript port of extractParagraphRuns from copy-doc.mjs.
 */
export function extractParagraphFormattedContent(
  paragraph: docs_v1.Schema$Paragraph,
  inlineObjects: Record<string, docs_v1.Schema$InlineObject>
): ParagraphContent {
  const runs: FormattedRun[] = [];
  const imageInfo: Array<ImageInfo & { offset: number }> = [];
  let textOffset = 0;

  if (!paragraph.elements) return { runs, imageInfo };

  for (const pe of paragraph.elements) {
    // Handle inline images
    if (pe.inlineObjectElement?.inlineObjectId) {
      const objId = pe.inlineObjectElement.inlineObjectId;
      const obj = inlineObjects?.[objId];
      if (obj) {
        const embeddedObj = obj.inlineObjectProperties?.embeddedObject;
        const uri = embeddedObj?.imageProperties?.contentUri
          ?? embeddedObj?.imageProperties?.sourceUri;
        if (uri) {
          const size = embeddedObj?.size;
          imageInfo.push({
            uri,
            width: size?.width?.magnitude ?? undefined,
            height: size?.height?.magnitude ?? undefined,
            offset: textOffset,
          });
        }
      }
      continue;
    }

    // Handle text runs
    if (pe.textRun?.content) {
      let content = pe.textRun.content;
      if (content.endsWith('\n')) content = content.slice(0, -1);
      if (content.length === 0) continue;

      const ts = pe.textRun.textStyle;
      const style: TextStyleArgs = {};
      let hasStyle = false;

      if (ts?.bold) { style.bold = true; hasStyle = true; }
      if (ts?.italic) { style.italic = true; hasStyle = true; }
      if (ts?.underline) { style.underline = true; hasStyle = true; }
      if (ts?.strikethrough) { style.strikethrough = true; hasStyle = true; }
      if (ts?.foregroundColor?.color?.rgbColor) {
        const hex = rgbToHex(ts.foregroundColor.color.rgbColor);
        if (hex !== '#000000') { style.foregroundColor = hex; hasStyle = true; }
      }
      if (ts?.backgroundColor?.color?.rgbColor) {
        style.backgroundColor = rgbToHex(ts.backgroundColor.color.rgbColor);
        hasStyle = true;
      }
      if (ts?.fontSize?.magnitude) {
        style.fontSize = ts.fontSize.magnitude;
        hasStyle = true;
      }
      if (ts?.weightedFontFamily?.fontFamily) {
        style.fontFamily = ts.weightedFontFamily.fontFamily;
        hasStyle = true;
      }
      if (ts?.link?.url) {
        style.linkUrl = ts.link.url;
        hasStyle = true;
      }

      runs.push({ text: content, ...(hasStyle ? { style } : {}) });
      textOffset += content.length;
    }
  }

  const namedStyle = paragraph.paragraphStyle?.namedStyleType ?? undefined;
  return { runs, imageInfo, namedStyle };
}

// --- Restore Pipeline ---

const log = {
  info: (msg: string) => console.error('[snapshot]', msg),
};

/**
 * Restores document content from a snapshot.
 * Pipeline: clear doc → insert paragraphs/tables → apply formatting → insert images.
 */
export async function restoreDocumentContent(
  docs: docs_v1.Docs,
  documentId: string,
  snapshot: DocumentSnapshot,
): Promise<void> {
  // Step 1: Clear document
  const currentDoc = await docs.documents.get({ documentId });
  const currentBody = currentDoc.data.body?.content || [];
  const endIndex = currentBody[currentBody.length - 1]?.endIndex;
  if (endIndex && endIndex > 2) {
    await executeBatchUpdate(docs, documentId, [{
      deleteContentRange: { range: { startIndex: 1, endIndex: endIndex - 1 } }
    }]);
  }

  // Step 2: Group elements into consecutive paragraph batches and tables
  const elementGroups: Array<
    | { type: 'paragraphs'; elements: docs_v1.Schema$StructuralElement[] }
    | { type: 'table'; table: docs_v1.Schema$Table }
  > = [];
  let currentParaGroup: { type: 'paragraphs'; elements: docs_v1.Schema$StructuralElement[] } | null = null;

  for (const el of snapshot.body) {
    if (el.sectionBreak) continue;

    if (el.paragraph) {
      if (!currentParaGroup) {
        currentParaGroup = { type: 'paragraphs', elements: [] };
      }
      currentParaGroup.elements.push(el);
    } else {
      if (currentParaGroup) {
        elementGroups.push(currentParaGroup);
        currentParaGroup = null;
      }
      if (el.table) {
        elementGroups.push({ type: 'table', table: el.table });
      }
    }
  }
  if (currentParaGroup) {
    elementGroups.push(currentParaGroup);
  }

  // Step 3: Copy each group
  for (const group of elementGroups) {
    if (group.type === 'paragraphs') {
      await restoreParagraphBatch(docs, documentId, group.elements, snapshot.inlineObjects);
    } else {
      await restoreTable(docs, documentId, group.table, snapshot.inlineObjects);
    }
  }
}

/**
 * Gets current end-of-document insertion point.
 */
async function getInsertionPoint(docs: docs_v1.Docs, documentId: string): Promise<number> {
  const doc = await docs.documents.get({ documentId });
  const body = doc.data.body?.content || [];
  return body[body.length - 1].endIndex! - 1;
}

/**
 * Restores a batch of consecutive paragraphs.
 */
async function restoreParagraphBatch(
  docs: docs_v1.Docs,
  documentId: string,
  paragraphElements: docs_v1.Schema$StructuralElement[],
  inlineObjects: Record<string, docs_v1.Schema$InlineObject>,
): Promise<void> {
  const insertPoint = await getInsertionPoint(docs, documentId);

  // Extract all paragraph data
  const paraData: ParagraphContent[] = [];
  for (const el of paragraphElements) {
    const para = el.paragraph!;
    paraData.push(extractParagraphFormattedContent(para, inlineObjects));
  }

  // Build combined text
  const allText = paraData.map(p => p.runs.map(r => r.text).join('') + '\n').join('');

  if (allText.trim().length === 0 && paraData.every(p => p.imageInfo.length === 0)) {
    if (allText.length > 0) {
      await executeBatchUpdate(docs, documentId, [{
        insertText: { location: { index: insertPoint }, text: allText }
      }]);
    }
    return;
  }

  // Phase 1: Insert all text
  const insertRequests: docs_v1.Schema$Request[] = [{
    insertText: { location: { index: insertPoint }, text: allText }
  }];

  // Phase 2: Paragraph styles and text formatting
  const formatRequests: docs_v1.Schema$Request[] = [];
  const paraStyleRequests: docs_v1.Schema$Request[] = [];
  let offset = insertPoint;

  // Store startIndex for each paragraph to compute image positions later
  const paraDataWithIndex: Array<ParagraphContent & { startIndex: number }> = [];

  for (const pd of paraData) {
    const paraStart = offset;
    paraDataWithIndex.push({ ...pd, startIndex: paraStart });
    const paraTextLen = pd.runs.map(r => r.text).join('').length;
    const paraEnd = offset + paraTextLen + 1; // +1 for \n

    if (pd.namedStyle && pd.namedStyle !== 'NORMAL_TEXT') {
      paraStyleRequests.push({
        updateParagraphStyle: {
          range: { startIndex: paraStart, endIndex: paraEnd },
          paragraphStyle: { namedStyleType: pd.namedStyle },
          fields: 'namedStyleType',
        }
      });
    }

    let runOffset = paraStart;
    for (const run of pd.runs) {
      const runEnd = runOffset + run.text.length;
      if (run.style && run.text.length > 0) {
        const result = buildUpdateTextStyleRequest(runOffset, runEnd, run.style);
        if (result) formatRequests.push(result.request);
      }
      runOffset = runEnd;
    }

    offset = paraEnd;
  }

  // Execute: insert → paragraph styles → text formatting
  const allRequests = [...insertRequests, ...paraStyleRequests, ...formatRequests];
  if (allRequests.length > 0) {
    // Use chunked execution to stay within API limits
    await executeBatchUpdateChunked(docs, documentId, allRequests, 50, log);
  }

  // Phase 3: Insert images at correct positions
  const imageInserts: Array<ImageInfo & { offset: number; index: number }> = [];
  for (const pd of paraDataWithIndex) {
    for (const img of pd.imageInfo) {
      imageInserts.push({ ...img, index: pd.startIndex + img.offset });
    }
  }

  if (imageInserts.length > 0) {
    // Sort in reverse order to avoid shifting subsequent indices
    imageInserts.sort((a, b) => b.index - a.index);

    for (const img of imageInserts) {
      try {
        const req: docs_v1.Schema$Request = {
          insertInlineImage: {
            location: { index: img.index },
            uri: img.uri,
            ...(img.width && img.height && {
              objectSize: {
                height: { magnitude: img.height, unit: 'PT' },
                width: { magnitude: img.width, unit: 'PT' },
              },
            }),
          }
        };
        await executeBatchUpdate(docs, documentId, [req]);
      } catch {
        // Image insertion can fail if URI expired — continue with rest
      }
    }
  }
}

/**
 * Restores a table from snapshot data.
 */
async function restoreTable(
  docs: docs_v1.Docs,
  documentId: string,
  table: docs_v1.Schema$Table,
  inlineObjects: Record<string, docs_v1.Schema$InlineObject>,
): Promise<void> {
  const tableRows = table.tableRows || [];
  const numCols = table.columns || 0;
  const numRows = tableRows.length;

  if (numRows === 0 || numCols === 0) return;

  // Step A: Insert empty table
  const insertPoint = await getInsertionPoint(docs, documentId);
  await executeBatchUpdate(docs, documentId, [{
    insertTable: {
      location: { index: insertPoint },
      rows: numRows,
      columns: numCols,
    }
  }]);

  // Step B: Extract formatted content from source cells
  const cellData: Array<{
    row: number;
    col: number;
    runs: FormattedRun[];
    imageInfo: ImageInfo[];
  }> = [];

  for (let r = 0; r < tableRows.length; r++) {
    const cells = tableRows[r].tableCells || [];
    for (let c = 0; c < cells.length; c++) {
      const { runs, imageInfo } = extractFormattedCellContent(cells[c], inlineObjects);
      if (runs.length > 0 || imageInfo.length > 0) {
        cellData.push({ row: r, col: c, runs, imageInfo });
      }
    }
  }

  if (cellData.length === 0) return;

  // Step C: Fill text into cells
  let doc2 = await docs.documents.get({ documentId });
  let body2 = doc2.data.body?.content || [];
  let tables2 = extractTableElements(body2);
  const newTableIndex = tables2.length - 1;

  const textEdits: CellEdit[] = cellData
    .filter(cd => cd.runs.length > 0)
    .map(cd => ({
      row: cd.row,
      col: cd.col,
      text: cd.runs.map(r => r.text).join(''),
    }));

  if (textEdits.length > 0) {
    const textRequests = buildBatchEditCellRequests(body2, newTableIndex, textEdits);
    if (textRequests.length > 0) {
      await executeBatchUpdateChunked(docs, documentId, textRequests, 50, log);
    }
  }

  // Step D: Apply formatting
  const cellsWithFormatting = cellData.filter(cd => cd.runs.some(r => r.style));
  if (cellsWithFormatting.length > 0) {
    doc2 = await docs.documents.get({ documentId });
    body2 = doc2.data.body?.content || [];
    tables2 = extractTableElements(body2);
    const targetTable = tables2[newTableIndex]?.table;

    if (targetTable) {
      const formatRequests: docs_v1.Schema$Request[] = [];
      for (const cd of cellsWithFormatting) {
        const targetCell = getCellElement(targetTable, cd.row, cd.col);
        const cellStart = getCellInsertionPoint(targetCell);
        const cellFormatReqs = buildFormattedCellFormatRequests(cellStart, cd.runs);
        formatRequests.push(...cellFormatReqs);
      }
      if (formatRequests.length > 0) {
        await executeBatchUpdateChunked(docs, documentId, formatRequests, 50, log);
      }
    }
  }

  // Step E: Insert images
  const cellsWithImages = cellData.filter(cd => cd.imageInfo.length > 0);
  if (cellsWithImages.length > 0) {
    doc2 = await docs.documents.get({ documentId });
    body2 = doc2.data.body?.content || [];

    const imageInserts: CellImageInsert[] = [];
    for (const cd of cellsWithImages) {
      for (const img of cd.imageInfo) {
        imageInserts.push({
          row: cd.row,
          col: cd.col,
          imageUrl: img.uri,
          width: img.width,
          height: img.height,
        });
      }
    }

    if (imageInserts.length > 0) {
      const imageRequests = buildBatchInsertImageRequests(body2, newTableIndex, imageInserts);
      if (imageRequests.length > 0) {
        await executeBatchUpdateChunked(docs, documentId, imageRequests, 50, log);
      }
    }
  }
}

// --- Snapshot CRUD ---

/**
 * Captures current document state and pushes to undo stack.
 * Clears the redo stack (new change branch).
 */
export async function createSnapshot(
  docs: docs_v1.Docs,
  documentId: string,
  label: string = 'snapshot',
): Promise<DocumentSnapshot> {
  const doc = await docs.documents.get({ documentId });
  const body = doc.data.body?.content || [];
  const inlineObjects = (doc.data as any).inlineObjects || {};

  const snapshot: DocumentSnapshot = {
    id: generateId(),
    documentId,
    timestamp: Date.now(),
    label,
    body: JSON.parse(JSON.stringify(body)),
    inlineObjects: JSON.parse(JSON.stringify(inlineObjects)),
  };

  const stack = getStack(documentId);
  stack.undoStack.push(snapshot);
  stack.redoStack = []; // New change branch

  // Enforce max snapshots
  while (stack.undoStack.length > MAX_SNAPSHOTS) {
    stack.undoStack.shift();
  }

  return snapshot;
}

/**
 * Undoes the last change by restoring the most recent snapshot.
 * Current state is saved to redo stack.
 */
export async function undoLastChange(
  docs: docs_v1.Docs,
  documentId: string,
): Promise<{ restored: DocumentSnapshot; message: string }> {
  const stack = getStack(documentId);

  if (stack.undoStack.length === 0) {
    throw new UserError('No snapshots available to undo. Create a snapshot before making changes.');
  }

  // Save current state to redo stack
  const currentDoc = await docs.documents.get({ documentId });
  const currentBody = currentDoc.data.body?.content || [];
  const currentInlineObjects = (currentDoc.data as any).inlineObjects || {};

  const currentSnapshot: DocumentSnapshot = {
    id: generateId(),
    documentId,
    timestamp: Date.now(),
    label: 'pre-undo state',
    body: JSON.parse(JSON.stringify(currentBody)),
    inlineObjects: JSON.parse(JSON.stringify(currentInlineObjects)),
  };
  stack.redoStack.push(currentSnapshot);

  // Pop and restore from undo stack
  const snapshot = stack.undoStack.pop()!;
  await restoreDocumentContent(docs, documentId, snapshot);

  return {
    restored: snapshot,
    message: `Restored snapshot "${snapshot.label}" (${new Date(snapshot.timestamp).toISOString()}). ${stack.undoStack.length} undo(s) remaining.`,
  };
}

/**
 * Redoes the last undone change.
 * Current state is saved to undo stack.
 */
export async function redoLastChange(
  docs: docs_v1.Docs,
  documentId: string,
): Promise<{ restored: DocumentSnapshot; message: string }> {
  const stack = getStack(documentId);

  if (stack.redoStack.length === 0) {
    throw new UserError('No redo states available.');
  }

  // Save current state to undo stack
  const currentDoc = await docs.documents.get({ documentId });
  const currentBody = currentDoc.data.body?.content || [];
  const currentInlineObjects = (currentDoc.data as any).inlineObjects || {};

  const currentSnapshot: DocumentSnapshot = {
    id: generateId(),
    documentId,
    timestamp: Date.now(),
    label: 'pre-redo state',
    body: JSON.parse(JSON.stringify(currentBody)),
    inlineObjects: JSON.parse(JSON.stringify(currentInlineObjects)),
  };
  stack.undoStack.push(currentSnapshot);

  // Pop and restore from redo stack
  const snapshot = stack.redoStack.pop()!;
  await restoreDocumentContent(docs, documentId, snapshot);

  return {
    restored: snapshot,
    message: `Redone to state "${snapshot.label}" (${new Date(snapshot.timestamp).toISOString()}). ${stack.redoStack.length} redo(s) remaining.`,
  };
}

/**
 * Lists all snapshots for a document.
 */
export function listSnapshots(documentId: string): Array<{
  id: string;
  label: string;
  timestamp: number;
  stack: 'undo' | 'redo';
}> {
  const stack = getStack(documentId);
  const result: Array<{ id: string; label: string; timestamp: number; stack: 'undo' | 'redo' }> = [];

  for (const s of stack.undoStack) {
    result.push({ id: s.id, label: s.label, timestamp: s.timestamp, stack: 'undo' });
  }
  for (const s of stack.redoStack) {
    result.push({ id: s.id, label: s.label, timestamp: s.timestamp, stack: 'redo' });
  }

  return result;
}
