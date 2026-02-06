// src/markdownToGoogleDocs.ts
import { docs_v1 } from 'googleapis';
import type Token from 'markdown-it/lib/token.mjs';
import { parseMarkdown, getLinkHref, getHeadingLevel } from './markdownParser.js';
import { buildUpdateTextStyleRequest, buildUpdateParagraphStyleRequest } from './googleDocsApiHelpers.js';
import { MarkdownConversionError } from './types.js';

// --- Internal Types ---

interface TextRange {
  startIndex: number;
  endIndex: number;
  formatting: FormattingState;
}

interface FormattingState {
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  link?: string;
}

interface ParagraphRange {
  startIndex: number;
  endIndex: number;
  namedStyleType?: string;
}

interface ListState {
  type: 'bullet' | 'ordered';
  level: number;
  listId?: string;
}

interface PendingListItem {
  startIndex: number;
  endIndex?: number;
  listId: string;
  nestingLevel: number;
  isOrdered: boolean;
}

// --- Table State Types ---

interface TableBuildState {
  insertIndex: number;
  rows: number;
  columns: number;
  cells: string[][];
  headerCells: string[];
  currentRow: string[];
  inHead: boolean;
  inBody: boolean;
  currentCellText: string;
  /** Line range from markdown-it token.map [startLine, endLine) */
  sourceMap?: [number, number];
}

/**
 * Describes a table that needs to be filled after insertTable is executed.
 */
export interface PendingTableFill {
  insertIndex: number;
  data: string[][];
  rows: number;
  columns: number;
  hasBoldHeaders: boolean;
}

interface ConversionContext {
  currentIndex: number;
  insertRequests: docs_v1.Schema$Request[];
  formatRequests: docs_v1.Schema$Request[];
  textRanges: TextRange[];
  formattingStack: FormattingState[];
  listStack: ListState[];
  paragraphRanges: ParagraphRange[];
  pendingListItems: PendingListItem[];
  listIds: Map<string, string>;
  tabId?: string;
  currentParagraphStart?: number;
  currentHeadingLevel?: number;
  tableState?: TableBuildState;
  pendingTableFills: PendingTableFill[];
  /** Set to true when a table is closed — stops further token processing */
  stopAfterTable: boolean;
  /** Original markdown text for splitting at table boundaries */
  originalMarkdown: string;
  /** Line number where the last table ended (exclusive) */
  lastTableEndLine?: number;
}

export interface MarkdownConversionResult {
  requests: docs_v1.Schema$Request[];
  pendingTableFills: PendingTableFill[];
  /** Remaining markdown content after the last table (for multi-phase execution) */
  postTableContent?: string;
}

// --- Main Conversion Functions ---

/**
 * Converts markdown text to Google Docs API batch update requests.
 * (Backward-compatible — ignores table fill data and post-table content)
 */
export function convertMarkdownToRequests(
  markdown: string,
  startIndex: number = 1,
  tabId?: string
): docs_v1.Schema$Request[] {
  return convertMarkdownToRequestsWithTables(markdown, startIndex, tabId).requests;
}

/**
 * Converts markdown text to Google Docs API requests WITH table fill information.
 *
 * When a GFM table is encountered, processing stops at the table boundary.
 * The insertTable request is included in `requests`, cell data in `pendingTableFills`,
 * and any remaining markdown after the table in `postTableContent`.
 *
 * The caller must:
 * 1. Execute `requests` (via executeBatchUpdateWithSplitting)
 * 2. Fill tables using `pendingTableFills`
 * 3. If `postTableContent` exists, re-read the document end index and
 *    recursively call this function for the remaining content.
 */
export function convertMarkdownToRequestsWithTables(
  markdown: string,
  startIndex: number = 1,
  tabId?: string
): MarkdownConversionResult {
  if (!markdown || markdown.trim().length === 0) {
    return { requests: [], pendingTableFills: [] };
  }

  const parsed = parseMarkdown(markdown);

  const context: ConversionContext = {
    currentIndex: startIndex,
    insertRequests: [],
    formatRequests: [],
    textRanges: [],
    formattingStack: [],
    listStack: [],
    paragraphRanges: [],
    pendingListItems: [],
    listIds: new Map(),
    tabId,
    pendingTableFills: [],
    stopAfterTable: false,
    originalMarkdown: markdown,
  };

  try {
    for (const token of parsed.tokens) {
      if (context.stopAfterTable) break;
      processToken(token, context);
    }

    finalizeFormatting(context);

    // Determine remaining markdown after the last table
    let postTableContent: string | undefined;
    if (context.stopAfterTable && context.lastTableEndLine != null) {
      const lines = markdown.split('\n');
      const remaining = lines.slice(context.lastTableEndLine).join('\n').trim();
      if (remaining.length > 0) {
        postTableContent = remaining;
      }
    }

    return {
      requests: [...context.insertRequests, ...context.formatRequests],
      pendingTableFills: context.pendingTableFills,
      postTableContent,
    };
  } catch (error) {
    if (error instanceof MarkdownConversionError) {
      throw error;
    }
    throw new MarkdownConversionError(
      `Failed to convert markdown: ${error instanceof Error ? error.message : 'Unknown error'}`
    );
  }
}

// --- Token Processing ---

function processToken(token: Token, context: ConversionContext): void {
  switch (token.type) {
    // Headings
    case 'heading_open':
      handleHeadingOpen(token, context);
      break;
    case 'heading_close':
      handleHeadingClose(context);
      break;

    // Paragraphs
    case 'paragraph_open':
      handleParagraphOpen(context);
      break;
    case 'paragraph_close':
      handleParagraphClose(context);
      break;

    // Text content
    case 'text':
    case 'code_inline':
      handleTextToken(token, context);
      break;

    // Inline formatting
    case 'strong_open':
      context.formattingStack.push({ bold: true });
      break;
    case 'strong_close':
      popFormatting(context, 'bold');
      break;

    case 'em_open':
      context.formattingStack.push({ italic: true });
      break;
    case 'em_close':
      popFormatting(context, 'italic');
      break;

    case 's_open':
      context.formattingStack.push({ strikethrough: true });
      break;
    case 's_close':
      popFormatting(context, 'strikethrough');
      break;

    // Links
    case 'link_open': {
      const href = getLinkHref(token);
      if (href) {
        context.formattingStack.push({ link: href });
      }
      break;
    }
    case 'link_close':
      popFormatting(context, 'link');
      break;

    // Lists
    case 'bullet_list_open':
      context.listStack.push({
        type: 'bullet',
        level: context.listStack.length
      });
      break;
    case 'bullet_list_close':
      context.listStack.pop();
      break;

    case 'ordered_list_open':
      context.listStack.push({
        type: 'ordered',
        level: context.listStack.length
      });
      break;
    case 'ordered_list_close':
      context.listStack.pop();
      break;

    case 'list_item_open':
      handleListItemOpen(context);
      break;
    case 'list_item_close':
      handleListItemClose(context);
      break;

    // Soft breaks and hard breaks
    case 'softbreak':
      insertText(' ', context);
      break;
    case 'hardbreak':
      insertText('\n', context);
      break;

    // Inline elements (like inline code)
    case 'inline':
      if (token.children) {
        for (const child of token.children) {
          processToken(child, context);
        }
      }
      break;

    // --- Table handling ---
    case 'table_open':
      handleTableOpen(token, context);
      break;
    case 'table_close':
      handleTableClose(context);
      break;
    case 'thead_open':
      if (context.tableState) context.tableState.inHead = true;
      break;
    case 'thead_close':
      if (context.tableState) context.tableState.inHead = false;
      break;
    case 'tbody_open':
      if (context.tableState) context.tableState.inBody = true;
      break;
    case 'tbody_close':
      if (context.tableState) context.tableState.inBody = false;
      break;
    case 'tr_open':
      if (context.tableState) context.tableState.currentRow = [];
      break;
    case 'tr_close':
      handleTableRowClose(context);
      break;
    case 'th_open':
    case 'td_open':
      if (context.tableState) context.tableState.currentCellText = '';
      break;
    case 'th_close':
    case 'td_close':
      handleTableCellClose(context);
      break;

    // Other structural tokens
    case 'fence':
    case 'code_block':
    case 'blockquote_open':
    case 'blockquote_close':
    case 'hr':
      break;

    default:
      break;
  }
}

// --- Table Handlers ---

function handleTableOpen(token: Token, context: ConversionContext): void {
  context.tableState = {
    insertIndex: context.currentIndex,
    rows: 0,
    columns: 0,
    cells: [],
    headerCells: [],
    currentRow: [],
    inHead: false,
    inBody: false,
    currentCellText: '',
    sourceMap: token.map as [number, number] | undefined,
  };
}

function handleTableCellClose(context: ConversionContext): void {
  if (!context.tableState) return;
  context.tableState.currentRow.push(context.tableState.currentCellText);
  context.tableState.currentCellText = '';
}

function handleTableRowClose(context: ConversionContext): void {
  if (!context.tableState) return;
  const ts = context.tableState;

  if (ts.inHead) {
    ts.headerCells = [...ts.currentRow];
    ts.columns = Math.max(ts.columns, ts.currentRow.length);
  } else {
    ts.cells.push([...ts.currentRow]);
    ts.columns = Math.max(ts.columns, ts.currentRow.length);
  }

  ts.currentRow = [];
}

function handleTableClose(context: ConversionContext): void {
  if (!context.tableState) return;
  const ts = context.tableState;

  // Build full data: header row + body rows
  const allData: string[][] = [];
  if (ts.headerCells.length > 0) {
    allData.push(ts.headerCells);
  }
  for (const row of ts.cells) {
    allData.push(row);
  }

  const totalRows = allData.length;
  const totalColumns = ts.columns;

  if (totalRows === 0 || totalColumns === 0) {
    context.tableState = undefined;
    return;
  }

  // Pad rows to equal column count
  for (const row of allData) {
    while (row.length < totalColumns) {
      row.push('');
    }
  }

  // Generate insertTable request
  const location: any = { index: ts.insertIndex };
  if (context.tabId) {
    location.tabId = context.tabId;
  }

  context.insertRequests.push({
    insertTable: {
      location,
      rows: totalRows,
      columns: totalColumns,
    },
  });

  context.pendingTableFills.push({
    insertIndex: ts.insertIndex,
    data: allData,
    rows: totalRows,
    columns: totalColumns,
    hasBoldHeaders: ts.headerCells.length > 0,
  });

  // Record source line for splitting remaining markdown
  if (ts.sourceMap) {
    context.lastTableEndLine = ts.sourceMap[1];
  }

  // STOP processing further tokens — we can't predict post-table indices.
  // Remaining content will be handled in a second pass by the caller.
  context.stopAfterTable = true;
  context.tableState = undefined;
}

// --- Heading Handlers ---

function handleHeadingOpen(token: Token, context: ConversionContext): void {
  const level = getHeadingLevel(token);
  if (level) {
    context.currentHeadingLevel = level;
    context.currentParagraphStart = context.currentIndex;
  }
}

function handleHeadingClose(context: ConversionContext): void {
  if (context.currentHeadingLevel && context.currentParagraphStart !== undefined) {
    const headingStyleType = `HEADING_${context.currentHeadingLevel}`;
    context.paragraphRanges.push({
      startIndex: context.currentParagraphStart,
      endIndex: context.currentIndex,
      namedStyleType: headingStyleType
    });

    // Add newline after heading
    insertText('\n', context);

    context.currentHeadingLevel = undefined;
    context.currentParagraphStart = undefined;
  }
}

// --- Paragraph Handlers ---

function handleParagraphOpen(context: ConversionContext): void {
  // Skip if we're in a list or table
  if (context.listStack.length === 0 && !context.tableState) {
    context.currentParagraphStart = context.currentIndex;
  }
}

function handleParagraphClose(context: ConversionContext): void {
  // Skip if we're in a list or table
  if (context.listStack.length === 0 && !context.tableState) {
    insertText('\n\n', context);
    context.currentParagraphStart = undefined;
  }
}

// --- List Handlers ---

function handleListItemOpen(context: ConversionContext): void {
  if (context.listStack.length === 0) {
    throw new MarkdownConversionError('List item found outside of list context');
  }

  const currentList = context.listStack[context.listStack.length - 1];
  const listKey = `${currentList.type}_${currentList.level}`;

  if (!context.listIds.has(listKey)) {
    const listId = `list_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
    context.listIds.set(listKey, listId);
  }

  const listId = context.listIds.get(listKey)!;

  const itemStart = context.currentIndex;
  context.pendingListItems.push({
    startIndex: itemStart,
    listId,
    nestingLevel: currentList.level,
    isOrdered: currentList.type === 'ordered'
  });
}

function handleListItemClose(context: ConversionContext): void {
  if (context.pendingListItems.length > 0) {
    const lastItem = context.pendingListItems[context.pendingListItems.length - 1];
    lastItem.endIndex = context.currentIndex;
    insertText('\n', context);
  }
}

// --- Text Handling ---

function handleTextToken(token: Token, context: ConversionContext): void {
  const text = token.content;
  if (!text) return;

  // If inside a table cell, accumulate text instead of inserting
  if (context.tableState) {
    context.tableState.currentCellText += text;
    return;
  }

  const startIndex = context.currentIndex;
  const endIndex = startIndex + text.length;

  insertText(text, context);

  const currentFormatting = mergeFormattingStack(context.formattingStack);
  if (hasFormatting(currentFormatting)) {
    context.textRanges.push({
      startIndex,
      endIndex,
      formatting: currentFormatting
    });
  }
}

function insertText(text: string, context: ConversionContext): void {
  const location: any = { index: context.currentIndex };
  if (context.tabId) {
    location.tabId = context.tabId;
  }

  context.insertRequests.push({
    insertText: {
      location,
      text
    }
  });

  context.currentIndex += text.length;
}

// --- Formatting Stack Management ---

function mergeFormattingStack(stack: FormattingState[]): FormattingState {
  const merged: FormattingState = {};

  for (const state of stack) {
    if (state.bold !== undefined) merged.bold = state.bold;
    if (state.italic !== undefined) merged.italic = state.italic;
    if (state.strikethrough !== undefined) merged.strikethrough = state.strikethrough;
    if (state.link !== undefined) merged.link = state.link;
  }

  return merged;
}

function hasFormatting(formatting: FormattingState): boolean {
  return formatting.bold === true ||
         formatting.italic === true ||
         formatting.strikethrough === true ||
         formatting.link !== undefined;
}

function popFormatting(context: ConversionContext, type: keyof FormattingState): void {
  for (let i = context.formattingStack.length - 1; i >= 0; i--) {
    if (context.formattingStack[i][type] !== undefined) {
      context.formattingStack.splice(i, 1);
      break;
    }
  }
}

// --- Finalization ---

function finalizeFormatting(context: ConversionContext): void {
  for (const range of context.textRanges) {
    if (range.formatting.bold || range.formatting.italic || range.formatting.strikethrough) {
      const styleRequest = buildUpdateTextStyleRequest(
        range.startIndex,
        range.endIndex,
        {
          bold: range.formatting.bold,
          italic: range.formatting.italic,
          strikethrough: range.formatting.strikethrough
        },
        context.tabId
      );
      if (styleRequest) {
        context.formatRequests.push(styleRequest.request);
      }
    }

    if (range.formatting.link) {
      const linkRequest = buildUpdateTextStyleRequest(
        range.startIndex,
        range.endIndex,
        { linkUrl: range.formatting.link },
        context.tabId
      );
      if (linkRequest) {
        context.formatRequests.push(linkRequest.request);
      }
    }
  }

  for (const paraRange of context.paragraphRanges) {
    if (paraRange.namedStyleType) {
      const paraRequest = buildUpdateParagraphStyleRequest(
        paraRange.startIndex,
        paraRange.endIndex,
        { namedStyleType: paraRange.namedStyleType as any },
        context.tabId
      );
      if (paraRequest) {
        context.formatRequests.push(paraRequest.request);
      }
    }
  }

  for (const listItem of context.pendingListItems) {
    if (listItem.endIndex !== undefined) {
      const rangeLocation: docs_v1.Schema$Range = {
        startIndex: listItem.startIndex,
        endIndex: listItem.endIndex
      };
      if (context.tabId) {
        rangeLocation.tabId = context.tabId;
      }

      const bulletPreset = listItem.isOrdered
        ? 'NUMBERED_DECIMAL_ALPHA_ROMAN'
        : 'BULLET_DISC_CIRCLE_SQUARE';

      context.formatRequests.push({
        createParagraphBullets: {
          range: rangeLocation,
          bulletPreset,
        }
      });
    }
  }
}
