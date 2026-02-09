// src/tableHelpers.ts
// Table-specific helpers for Google Docs table editing
import { docs_v1 } from 'googleapis';
import { UserError } from 'fastmcp';
import { TextStyleArgs } from './types.js';
import { buildUpdateTextStyleRequest, buildObjectSize } from './googleDocsApiHelpers.js';

// --- Types ---

export interface TableInfo {
  tableIndex: number;
  rows: number;
  columns: number;
  startIndex: number;
  endIndex: number;
  headerRow: string[];
}

export interface CellMetadata {
  text: string;
  hasImage: boolean;
  startIndex: number;
  endIndex: number;
  colSpan?: number;
}

export interface TableCellsResult {
  tableIndex: number;
  rows: number;
  columns: number;
  values: string[][];
  metadata: CellMetadata[][];
}

// --- Core Helpers ---

/**
 * Extracts all table structural elements from a document's body content.
 * Returns the raw content elements that contain tables (with startIndex/endIndex).
 */
export function extractTableElements(bodyContent: docs_v1.Schema$StructuralElement[]): docs_v1.Schema$StructuralElement[] {
  return bodyContent.filter(el => el.table != null);
}

/**
 * Gets a specific table element by index (0-based).
 * Throws UserError if index is out of range.
 */
export function getTableElement(bodyContent: docs_v1.Schema$StructuralElement[], tableIndex: number): docs_v1.Schema$StructuralElement {
  const tables = extractTableElements(bodyContent);
  if (tableIndex < 0 || tableIndex >= tables.length) {
    throw new UserError(
      `Table index ${tableIndex} out of range. Document has ${tables.length} table(s) (0-indexed).`
    );
  }
  return tables[tableIndex];
}

/**
 * Gets cell content element from a table.
 * Validates row/col bounds and throws UserError on out-of-range.
 */
export function getCellElement(
  table: docs_v1.Schema$Table,
  row: number,
  col: number
): docs_v1.Schema$TableCell {
  const tableRows = table.tableRows;
  if (!tableRows || row < 0 || row >= tableRows.length) {
    throw new UserError(
      `Row ${row} out of range. Table has ${tableRows?.length ?? 0} rows (0-indexed).`
    );
  }
  const tableCells = tableRows[row].tableCells;
  if (!tableCells || col < 0 || col >= tableCells.length) {
    throw new UserError(
      `Column ${col} out of range. Row ${row} has ${tableCells?.length ?? 0} columns (0-indexed).`
    );
  }
  return tableCells[col];
}

/**
 * Extracts text content from a cell, concatenating all text runs across all paragraphs.
 */
export function getCellText(cell: docs_v1.Schema$TableCell): string {
  if (!cell.content) return '';
  let text = '';
  for (const contentEl of cell.content) {
    if (contentEl.paragraph?.elements) {
      for (const pe of contentEl.paragraph.elements) {
        if (pe.textRun?.content) {
          text += pe.textRun.content;
        }
      }
    }
  }
  // Trim trailing newline that Google Docs adds to every cell
  return text.replace(/\n$/, '');
}

/**
 * Checks if a cell contains any inline images.
 */
export function cellHasImage(cell: docs_v1.Schema$TableCell): boolean {
  if (!cell.content) return false;
  for (const contentEl of cell.content) {
    if (contentEl.paragraph?.elements) {
      for (const pe of contentEl.paragraph.elements) {
        if (pe.inlineObjectElement) return true;
      }
    }
  }
  return false;
}

/**
 * Gets the full content range of a cell (start to end index).
 * The endIndex is the last character position (excludes the cell's trailing structural newline).
 */
export function getCellRange(cell: docs_v1.Schema$TableCell): { startIndex: number; endIndex: number } {
  if (!cell.content || cell.content.length === 0) {
    throw new UserError('Cell has no content elements.');
  }
  const firstContent = cell.content[0];
  const lastContent = cell.content[cell.content.length - 1];
  const startIndex = firstContent.startIndex ?? 0;
  // endIndex from Docs API points AFTER the trailing \n of the last paragraph
  const endIndex = lastContent.endIndex ?? startIndex;
  return { startIndex, endIndex };
}

/**
 * Gets the text-only range within a cell for replacement.
 * Skips paragraphs that contain only inline images.
 * Returns the range of text that can be safely deleted and replaced.
 *
 * For a cell with content like:
 *   [paragraph with text] [paragraph with image] [paragraph with text]
 * This returns ranges for text paragraphs only.
 *
 * Returns null if cell has no text content (only images).
 */
export function getCellTextRanges(cell: docs_v1.Schema$TableCell): Array<{ startIndex: number; endIndex: number }> {
  if (!cell.content) return [];

  const textRanges: Array<{ startIndex: number; endIndex: number }> = [];

  for (const contentEl of cell.content) {
    if (!contentEl.paragraph?.elements) continue;

    // Check if this paragraph has any text runs (not just images or structural elements)
    const hasTextRun = contentEl.paragraph.elements.some(
      pe => pe.textRun?.content && pe.textRun.content !== '\n'
    );

    if (hasTextRun) {
      const pStart = contentEl.startIndex ?? 0;
      // endIndex - 1: preserve the paragraph's trailing newline for structural integrity
      const pEnd = (contentEl.endIndex ?? pStart) - 1;
      if (pEnd > pStart) {
        textRanges.push({ startIndex: pStart, endIndex: pEnd });
      }
    }
  }

  return textRanges;
}

/**
 * Gets the insertion point for new text in a cell.
 * Returns the startIndex of the first paragraph in the cell.
 */
export function getCellInsertionPoint(cell: docs_v1.Schema$TableCell): number {
  if (!cell.content || cell.content.length === 0) {
    throw new UserError('Cell has no content elements.');
  }
  return cell.content[0].startIndex ?? 0;
}

// --- High-level functions ---

/**
 * Returns structure info for all tables in the document.
 */
export function getTablesInfo(bodyContent: docs_v1.Schema$StructuralElement[]): TableInfo[] {
  const result: TableInfo[] = [];
  let tableIdx = 0;

  for (const element of bodyContent) {
    if (!element.table) continue;
    const table = element.table;
    const rows = table.tableRows?.length ?? 0;
    const columns = table.columns ?? 0;

    // Extract header row text (first row)
    const headerRow: string[] = [];
    if (rows > 0 && table.tableRows) {
      const firstRow = table.tableRows[0];
      if (firstRow.tableCells) {
        for (const cell of firstRow.tableCells) {
          headerRow.push(getCellText(cell));
        }
      }
    }

    result.push({
      tableIndex: tableIdx,
      rows,
      columns,
      startIndex: element.startIndex ?? 0,
      endIndex: element.endIndex ?? 0,
      headerRow,
    });
    tableIdx++;
  }

  return result;
}

/**
 * Reads all cell values and metadata from a specific table.
 */
export function readTableCells(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
): TableCellsResult {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;
  const rows = table.tableRows?.length ?? 0;
  const columns = table.columns ?? 0;

  const values: string[][] = [];
  const metadata: CellMetadata[][] = [];

  if (table.tableRows) {
    for (let r = 0; r < table.tableRows.length; r++) {
      const rowValues: string[] = [];
      const rowMeta: CellMetadata[] = [];
      const tableRow = table.tableRows[r];

      if (tableRow.tableCells) {
        for (let c = 0; c < tableRow.tableCells.length; c++) {
          const cell = tableRow.tableCells[c];
          const text = getCellText(cell);
          const hasImg = cellHasImage(cell);
          const range = getCellRange(cell);

          rowValues.push(text);
          rowMeta.push({
            text,
            hasImage: hasImg,
            startIndex: range.startIndex,
            endIndex: range.endIndex,
            colSpan: cell.tableCellStyle?.columnSpan ?? undefined,
          });
        }
      }

      values.push(rowValues);
      metadata.push(rowMeta);
    }
  }

  return { tableIndex, rows, columns, values, metadata };
}

/**
 * Builds batchUpdate requests to replace text in a specific cell.
 * Preserves inline images — only deletes and replaces text runs.
 *
 * IMPORTANT: Requests are returned in reverse index order
 * (higher indices first) so they can be executed in a single batch
 * without index shifting issues.
 */
export function buildEditCellRequests(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  row: number,
  col: number,
  newText: string,
): docs_v1.Schema$Request[] {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;
  const cell = getCellElement(table, row, col);

  const requests: docs_v1.Schema$Request[] = [];
  const textRanges = getCellTextRanges(cell);

  if (textRanges.length === 0) {
    // Cell has no text (only images or empty). Insert text at the beginning of the cell.
    const insertPoint = getCellInsertionPoint(cell);
    requests.push({
      insertText: {
        location: { index: insertPoint },
        text: newText,
      },
    });
    return requests;
  }

  // Strategy: delete all existing text ranges (in reverse order), then insert new text
  // at the position of the first (lowest index) text range.

  // Sort text ranges by startIndex descending for safe deletion
  const sortedRanges = [...textRanges].sort((a, b) => b.startIndex - a.startIndex);

  // Delete existing text ranges (reverse order — highest index first)
  for (const range of sortedRanges) {
    requests.push({
      deleteContentRange: {
        range: {
          startIndex: range.startIndex,
          endIndex: range.endIndex,
        },
      },
    });
  }

  // Insert new text at the position of the first text range (lowest startIndex)
  const insertPoint = textRanges.reduce(
    (min, r) => Math.min(min, r.startIndex),
    Infinity,
  );

  if (newText) {
    requests.push({
      insertText: {
        location: { index: insertPoint },
        text: newText,
      },
    });
  }

  return requests;
}

/**
 * Builds a request to insert an inline image into a specific table cell.
 * Image is inserted at the beginning of the cell (before any existing content).
 */
export function buildInsertImageInCellRequest(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  row: number,
  col: number,
  imageUrl: string,
  width?: number,
  height?: number,
): docs_v1.Schema$Request {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;
  const cell = getCellElement(table, row, col);
  const insertPoint = getCellInsertionPoint(cell);

  const request: docs_v1.Schema$Request = {
    insertInlineImage: {
      location: { index: insertPoint },
      uri: imageUrl,
      ...buildObjectSize(width, height),
    },
  };

  return request;
}

// --- Formatted Run Types ---

export interface ImageInfo {
  uri: string;
  width?: number;   // PT
  height?: number;  // PT
}

export interface FormattedRun {
  text: string;
  style?: import('./types.js').TextStyleArgs;
}

export interface FormattedCellEdit {
  row: number;
  col: number;
  runs: FormattedRun[];
}

// --- Batch Operations ---

export interface CellEdit {
  row: number;
  col: number;
  text: string;
}

export interface CellImageInsert {
  row: number;
  col: number;
  imageUrl: string;
  width?: number;
  height?: number;
}

/**
 * Builds batch edit requests for multiple cells in a single table.
 * Fetches cell data once from the provided body content and generates
 * all delete+insert requests sorted by descending startIndex.
 *
 * @param bodyContent - Document body content (fetched once by caller)
 * @param tableIndex - Table index (0-based)
 * @param edits - Array of {row, col, text} edits
 * @returns Array of requests sorted by descending index, safe for sequential execution
 */
export function buildBatchEditCellRequests(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  edits: CellEdit[],
  tabId?: string,
): docs_v1.Schema$Request[] {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;

  // Build requests per cell, tracking the cell's startIndex for sorting
  const cellGroups: Array<{ sortIndex: number; requests: docs_v1.Schema$Request[] }> = [];

  for (const edit of edits) {
    const cell = getCellElement(table, edit.row, edit.col);
    const textRanges = getCellTextRanges(cell);
    const requests: docs_v1.Schema$Request[] = [];

    if (textRanges.length === 0) {
      // Empty cell — just insert
      const insertPoint = getCellInsertionPoint(cell);
      if (edit.text) {
        requests.push({
          insertText: {
            location: { index: insertPoint, ...(tabId && { tabId }) },
            text: edit.text,
          },
        });
      }
      cellGroups.push({ sortIndex: insertPoint, requests });
      continue;
    }

    // Delete existing text ranges in reverse order
    const sortedRanges = [...textRanges].sort((a, b) => b.startIndex - a.startIndex);
    for (const range of sortedRanges) {
      requests.push({
        deleteContentRange: {
          range: { startIndex: range.startIndex, endIndex: range.endIndex, ...(tabId && { tabId }) },
        },
      });
    }

    // Insert new text at lowest index
    const insertPoint = textRanges.reduce((min, r) => Math.min(min, r.startIndex), Infinity);
    if (edit.text) {
      requests.push({
        insertText: {
          location: { index: insertPoint, ...(tabId && { tabId }) },
          text: edit.text,
        },
      });
    }

    cellGroups.push({ sortIndex: insertPoint, requests });
  }

  // Sort cell groups by descending sortIndex to prevent index shifting
  cellGroups.sort((a, b) => b.sortIndex - a.sortIndex);

  // Flatten into single array preserving group order
  return cellGroups.flatMap(g => g.requests);
}

/**
 * Builds batch image insert requests for multiple cells in a single table.
 * Images are inserted at the beginning of each cell.
 * Requests are sorted by descending insertion index.
 *
 * @param bodyContent - Document body content
 * @param tableIndex - Table index (0-based)
 * @param images - Array of {row, col, imageUrl, width?, height?}
 * @returns Array of insertInlineImage requests sorted by descending index
 */
export function buildBatchInsertImageRequests(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  images: CellImageInsert[],
  tabId?: string,
): docs_v1.Schema$Request[] {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;

  const requests: Array<{ insertIndex: number; request: docs_v1.Schema$Request }> = [];

  for (const img of images) {
    const cell = getCellElement(table, img.row, img.col);
    const insertPoint = getCellInsertionPoint(cell);

    const request: docs_v1.Schema$Request = {
      insertInlineImage: {
        location: { index: insertPoint, ...(tabId && { tabId }) },
        uri: img.imageUrl,
        ...buildObjectSize(img.width, img.height),
      },
    };

    requests.push({ insertIndex: insertPoint, request });
  }

  // Sort by descending insertion index
  requests.sort((a, b) => b.insertIndex - a.insertIndex);

  return requests.map(r => r.request);
}

/**
 * Finds rows in a table where a specific column contains the search text.
 * Returns matching row indices and the full row data.
 */
export function findTableRows(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  searchColumn: number,
  searchText: string,
  caseSensitive: boolean = false,
): Array<{ rowIndex: number; values: string[] }> {
  const tableData = readTableCells(bodyContent, tableIndex);
  const results: Array<{ rowIndex: number; values: string[] }> = [];

  const normalizeText = (t: string) => caseSensitive ? t : t.toLowerCase();
  const needle = normalizeText(searchText);

  for (let r = 0; r < tableData.values.length; r++) {
    if (searchColumn >= tableData.values[r].length) continue;
    const cellText = normalizeText(tableData.values[r][searchColumn]);
    if (cellText.includes(needle)) {
      results.push({
        rowIndex: r,
        values: tableData.values[r],
      });
    }
  }

  return results;
}

/**
 * Builds requests to add a new row at the end of a table.
 * Uses insertTableRow request.
 */
export function buildAddTableRowRequest(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  insertBelow: number, // row index after which to insert (0-based)
): docs_v1.Schema$Request {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;
  const rows = table.tableRows?.length ?? 0;

  if (insertBelow < 0 || insertBelow >= rows) {
    throw new UserError(
      `insertBelow ${insertBelow} out of range. Table has ${rows} rows (0-indexed, max: ${rows - 1}).`
    );
  }

  // Get the endIndex of the last cell in the target row to use as insertion reference
  const targetRow = table.tableRows![insertBelow];
  const lastCell = targetRow.tableCells![targetRow.tableCells!.length - 1];
  const cellRange = getCellRange(lastCell);

  return {
    insertTableRow: {
      tableCellLocation: {
        tableStartLocation: { index: tableEl.startIndex ?? 0 },
        rowIndex: insertBelow,
        columnIndex: 0,
      },
      insertBelow: true,
    },
  };
}

// --- Rich Formatting Helpers ---

/**
 * Converts Google Docs API RGB color (0.0-1.0 floats) to hex string.
 */
export function rgbToHex(rgb: docs_v1.Schema$RgbColor): string {
  const r = Math.round((rgb.red ?? 0) * 255);
  const g = Math.round((rgb.green ?? 0) * 255);
  const b = Math.round((rgb.blue ?? 0) * 255);
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

/**
 * Extracts formatted content from a source cell, preserving per-run text styles
 * and inline image URIs. Handles \u000b (in-cell line breaks) properly.
 *
 * @param cell - Source table cell
 * @param inlineObjects - Document's inlineObjects map for resolving image URIs
 * @returns runs (text with style) and imageUris found in the cell
 */
export function extractFormattedCellContent(
  cell: docs_v1.Schema$TableCell,
  inlineObjects?: Record<string, docs_v1.Schema$InlineObject>
): { runs: FormattedRun[]; imageInfo: ImageInfo[] } {
  const runs: FormattedRun[] = [];
  const imageInfo: ImageInfo[] = [];

  if (!cell.content) return { runs, imageInfo };

  const paragraphs = cell.content.filter(el => el.paragraph);

  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const para = paragraphs[pIdx].paragraph!;
    if (!para.elements) continue;

    // Add \u000b between paragraphs (not before first, not after last)
    if (pIdx > 0 && runs.length > 0) {
      runs.push({ text: '\u000b' });
    }

    for (const pe of para.elements) {
      // Handle inline images
      if (pe.inlineObjectElement?.inlineObjectId) {
        const objId = pe.inlineObjectElement.inlineObjectId;
        if (inlineObjects?.[objId]) {
          const embeddedObj = inlineObjects[objId].inlineObjectProperties?.embeddedObject;
          const uri = embeddedObj?.imageProperties?.contentUri
            ?? embeddedObj?.imageProperties?.sourceUri;
          if (uri) {
            const size = embeddedObj?.size;
            imageInfo.push({
              uri,
              width: size?.width?.magnitude ?? undefined,
              height: size?.height?.magnitude ?? undefined,
            });
          }
        }
        continue;
      }

      // Handle text runs
      if (pe.textRun?.content) {
        let content = pe.textRun.content;

        // Strip trailing \n only from the last element of the last paragraph
        const isLastParagraph = pIdx === paragraphs.length - 1;
        const isLastElement = pe === para.elements[para.elements.length - 1];
        if (isLastParagraph && isLastElement && content.endsWith('\n')) {
          content = content.slice(0, -1);
        }

        if (content.length === 0) continue;

        // Extract text style
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
      }
    }
  }

  return { runs, imageInfo };
}

/**
 * Given a target cell's startIndex and the FormattedRun[] that was inserted,
 * calculates run boundaries and generates updateTextStyle requests.
 *
 * @param cellStartIndex - The startIndex of the cell content after text insertion
 * @param runs - The FormattedRun[] used during text insertion
 * @param tabId - Optional tab ID
 * @returns Array of updateTextStyle requests
 */
export function buildFormattedCellFormatRequests(
  cellStartIndex: number,
  runs: FormattedRun[],
  tabId?: string
): docs_v1.Schema$Request[] {
  const requests: docs_v1.Schema$Request[] = [];
  let offset = cellStartIndex;

  for (const run of runs) {
    const runEnd = offset + run.text.length;

    if (run.style && run.text.length > 0) {
      const result = buildUpdateTextStyleRequest(offset, runEnd, run.style, tabId);
      if (result) {
        requests.push(result.request);
      }
    }

    offset = runEnd;
  }

  return requests;
}

/**
 * Reads all cells from a table and returns formatted content (runs + images) per cell.
 */
export function readTableCellsFormatted(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  inlineObjects?: Record<string, docs_v1.Schema$InlineObject>
): {
  tableIndex: number;
  rows: number;
  columns: number;
  cells: Array<Array<{ runs: FormattedRun[]; imageInfo: ImageInfo[] }>>;
} {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;
  const rows = table.tableRows?.length ?? 0;
  const columns = table.columns ?? 0;

  const cells: Array<Array<{ runs: FormattedRun[]; imageInfo: ImageInfo[] }>> = [];

  if (table.tableRows) {
    for (const tableRow of table.tableRows) {
      const rowCells: Array<{ runs: FormattedRun[]; imageInfo: ImageInfo[] }> = [];
      if (tableRow.tableCells) {
        for (const cell of tableRow.tableCells) {
          rowCells.push(extractFormattedCellContent(cell, inlineObjects));
        }
      }
      cells.push(rowCells);
    }
  }

  return { tableIndex, rows, columns, cells };
}

// --- Table Cell Style ---

export interface CellStyleUpdate {
  paddingTop?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  paddingRight?: number;
  contentAlignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
  backgroundColor?: { red?: number; green?: number; blue?: number };
}

/**
 * Build updateTableCellStyle requests for a range of cells.
 * If no row/col range is specified, applies to the entire table.
 */
export function buildUpdateTableCellStyleRequests(
  bodyContent: docs_v1.Schema$StructuralElement[],
  tableIndex: number,
  style: CellStyleUpdate,
  rowRange?: { startRow: number; rowSpan?: number },
  colRange?: { startCol: number; colSpan?: number },
): docs_v1.Schema$Request[] {
  const tableEl = getTableElement(bodyContent, tableIndex);
  const table = tableEl.table!;
  const tableStartIndex = tableEl.startIndex ?? 0;

  const totalRows = table.rows ?? 0;
  const totalCols = table.columns ?? 0;

  const startRow = rowRange?.startRow ?? 0;
  const rowSpan = rowRange?.rowSpan ?? totalRows - startRow;
  const startCol = colRange?.startCol ?? 0;
  const colSpan = colRange?.colSpan ?? totalCols - startCol;

  // Build tableCellStyle object
  const tableCellStyle: any = {};
  const fields: string[] = [];

  if (style.paddingTop != null) {
    tableCellStyle.paddingTop = { magnitude: style.paddingTop, unit: 'PT' };
    fields.push('paddingTop');
  }
  if (style.paddingBottom != null) {
    tableCellStyle.paddingBottom = { magnitude: style.paddingBottom, unit: 'PT' };
    fields.push('paddingBottom');
  }
  if (style.paddingLeft != null) {
    tableCellStyle.paddingLeft = { magnitude: style.paddingLeft, unit: 'PT' };
    fields.push('paddingLeft');
  }
  if (style.paddingRight != null) {
    tableCellStyle.paddingRight = { magnitude: style.paddingRight, unit: 'PT' };
    fields.push('paddingRight');
  }
  if (style.contentAlignment != null) {
    tableCellStyle.contentAlignment = style.contentAlignment;
    fields.push('contentAlignment');
  }
  if (style.backgroundColor != null) {
    tableCellStyle.backgroundColor = { color: { rgbColor: style.backgroundColor } };
    fields.push('backgroundColor');
  }

  if (fields.length === 0) {
    throw new UserError('No style properties specified.');
  }

  // Build one request per row-chunk (Google API applies to a row-range of cells)
  const requests: docs_v1.Schema$Request[] = [];

  for (let r = startRow; r < startRow + rowSpan; r++) {
    requests.push({
      updateTableCellStyle: {
        tableCellStyle,
        tableRange: {
          tableCellLocation: {
            tableStartLocation: { index: tableStartIndex },
            rowIndex: r,
            columnIndex: startCol,
          },
          rowSpan: 1,
          columnSpan: colSpan,
        },
        fields: fields.join(','),
      },
    });
  }

  return requests;
}
