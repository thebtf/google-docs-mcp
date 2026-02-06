/**
 * Copy content from reference doc to test doc preserving rich formatting.
 * Works directly with Google Docs API structure — no markdown intermediary.
 * Preserves: bold, italic, underline, colors, fonts, images, \u000b line breaks.
 *
 * Optimized: batches all paragraph inserts into minimal API calls.
 */

import { google } from 'googleapis';
import { authorize } from './dist/auth.js';
import * as GDocsHelpers from './dist/googleDocsApiHelpers.js';
import * as TableHelpers from './dist/tableHelpers.js';

const REF_DOC_ID = process.env.REF_DOC_ID;
const TEST_DOC_ID = process.env.TEST_DOC_ID;
if (!REF_DOC_ID || !TEST_DOC_ID) {
  throw new Error('Set REF_DOC_ID and TEST_DOC_ID env vars before running copy-doc.mjs');
}

const log = {
  info: (msg) => console.log('[INFO]', msg),
  error: (msg) => console.error('[ERROR]', msg),
};

async function main() {
  const auth = await authorize();
  const docs = google.docs({ version: 'v1', auth });
  console.log('Authenticated.\n');

  // --- Step 1: Read reference document (full structure + inlineObjects) ---
  console.log('Step 1: Reading reference document...');
  const refDoc = await docs.documents.get({ documentId: REF_DOC_ID });
  const refBody = refDoc.data.body?.content || [];
  const inlineObjects = refDoc.data.inlineObjects || {};
  console.log(`  ${refBody.length} body elements, ${Object.keys(inlineObjects).length} inline objects`);

  // --- Step 2: Clear test document ---
  console.log('\nStep 2: Clearing test document...');
  const testDoc = await docs.documents.get({ documentId: TEST_DOC_ID });
  const testBody = testDoc.data.body?.content || [];
  const endIndex = testBody[testBody.length - 1]?.endIndex - 1;
  if (endIndex > 1) {
    await GDocsHelpers.executeBatchUpdate(docs, TEST_DOC_ID, [{
      deleteContentRange: { range: { startIndex: 1, endIndex } }
    }]);
    console.log(`  Deleted indices 1-${endIndex}`);
  }

  // --- Step 3: Analyze reference structure ---
  // Split into consecutive paragraph groups and tables
  console.log('\nStep 3: Analyzing reference structure...');
  const elementGroups = [];
  let currentParaGroup = null;

  for (const el of refBody) {
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
        elementGroups.push({ type: 'table', element: el });
      }
    }
  }
  if (currentParaGroup) {
    elementGroups.push(currentParaGroup);
  }

  console.log(`  ${elementGroups.length} element groups (paragraphs batches + tables)`);

  // --- Step 4: Copy each group ---
  console.log('\nStep 4: Copying content...');
  for (const group of elementGroups) {
    if (group.type === 'paragraphs') {
      await copyParagraphBatch(docs, group.elements, inlineObjects);
    } else if (group.type === 'table') {
      await copyTable(docs, group.element.table, inlineObjects);
    }
  }

  // --- Step 5: Verify ---
  console.log('\nStep 5: Verifying...');
  const verifyDoc = await docs.documents.get({ documentId: TEST_DOC_ID });
  const vBody = verifyDoc.data.body?.content || [];
  let tableElements = 0;
  let paragraphElements = 0;
  for (const v of vBody) {
    if (v.table) tableElements++;
    if (v.paragraph) paragraphElements++;
  }
  console.log(`  ${vBody.length} elements: ${paragraphElements} paragraphs, ${tableElements} tables`);

  if (tableElements > 0) {
    const tables = TableHelpers.extractTableElements(vBody);
    for (let i = 0; i < tables.length; i++) {
      const t = tables[i];
      console.log(`  Table ${i}: ${t.table?.tableRows?.length || 0} rows x ${t.table?.columns || 0} cols`);
    }
  }
}

/**
 * Gets the current end-of-document insertion point.
 */
async function getInsertionPoint(docs) {
  const doc = await docs.documents.get({ documentId: TEST_DOC_ID });
  const body = doc.data.body?.content || [];
  return body[body.length - 1].endIndex - 1;
}

/**
 * Extracts FormattedRun[] from a paragraph's elements.
 */
function extractParagraphRuns(paragraph, inlineObjects) {
  const runs = [];
  const imageInfo = [];
  let textOffset = 0;

  if (!paragraph.elements) return { runs, imageInfo };

  for (const pe of paragraph.elements) {
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
            width: size?.width?.magnitude,
            height: size?.height?.magnitude,
            offset: textOffset,
          });
        }
      }
      continue;
    }

    if (pe.textRun?.content) {
      let content = pe.textRun.content;
      if (content.endsWith('\n')) content = content.slice(0, -1);
      if (content.length === 0) continue;

      const ts = pe.textRun.textStyle;
      const style = {};
      let hasStyle = false;

      if (ts?.bold) { style.bold = true; hasStyle = true; }
      if (ts?.italic) { style.italic = true; hasStyle = true; }
      if (ts?.underline) { style.underline = true; hasStyle = true; }
      if (ts?.strikethrough) { style.strikethrough = true; hasStyle = true; }
      if (ts?.foregroundColor?.color?.rgbColor) {
        const hex = TableHelpers.rgbToHex(ts.foregroundColor.color.rgbColor);
        if (hex !== '#000000') { style.foregroundColor = hex; hasStyle = true; }
      }
      if (ts?.backgroundColor?.color?.rgbColor) {
        style.backgroundColor = TableHelpers.rgbToHex(ts.backgroundColor.color.rgbColor);
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

  return { runs, imageInfo };
}

/**
 * Copy a batch of consecutive paragraphs in one API operation.
 * 1) Insert all text at once
 * 2) Apply paragraph styles
 * 3) Apply text formatting
 * 4) Insert images
 */
async function copyParagraphBatch(docs, paragraphElements, inlineObjects) {
  const insertPoint = await getInsertionPoint(docs);

  // Extract all paragraph data
  const paraData = [];
  for (const el of paragraphElements) {
    const para = el.paragraph;
    const { runs, imageInfo } = extractParagraphRuns(para, inlineObjects);
    const namedStyle = para.paragraphStyle?.namedStyleType;
    const fullText = runs.map(r => r.text).join('');
    paraData.push({ runs, imageInfo, namedStyle, fullText });
  }

  // Build combined text: each paragraph's text + \n
  // All paragraphs concatenated, inserted at insertPoint
  const allText = paraData.map(p => p.fullText + '\n').join('');

  if (allText.trim().length === 0 && paraData.every(p => p.imageInfo.length === 0)) {
    // All empty paragraphs — insert newlines
    if (allText.length > 0) {
      await GDocsHelpers.executeBatchUpdate(docs, TEST_DOC_ID, [{
        insertText: { location: { index: insertPoint }, text: allText }
      }]);
    }
    return;
  }

  // Phase 1: Insert all text at once
  const insertRequests = [];
  insertRequests.push({
    insertText: { location: { index: insertPoint }, text: allText }
  });

  // Phase 2: Calculate paragraph ranges and build style requests
  const formatRequests = [];
  const paraStyleRequests = [];
  let offset = insertPoint;

  for (const pd of paraData) {
    const paraStart = offset;
    pd.startIndex = paraStart; // Store for image insertion
    const paraTextLen = pd.fullText.length;
    const paraEnd = offset + paraTextLen + 1; // +1 for \n

    // Paragraph style (headings)
    if (pd.namedStyle && pd.namedStyle !== 'NORMAL_TEXT') {
      paraStyleRequests.push({
        updateParagraphStyle: {
          range: { startIndex: paraStart, endIndex: paraEnd },
          paragraphStyle: { namedStyleType: pd.namedStyle },
          fields: 'namedStyleType',
        }
      });
    }

    // Per-run text formatting
    let runOffset = paraStart;
    for (const run of pd.runs) {
      const runEnd = runOffset + run.text.length;
      if (run.style && run.text.length > 0) {
        const result = GDocsHelpers.buildUpdateTextStyleRequest(runOffset, runEnd, run.style);
        if (result) formatRequests.push(result.request);
      }
      runOffset = runEnd;
    }

    offset = paraEnd;
  }

  // Execute: insert text, then paragraph styles, then text formatting
  const allRequests = [...insertRequests, ...paraStyleRequests, ...formatRequests];
  await GDocsHelpers.executeBatchUpdateWithSplitting(docs, TEST_DOC_ID, allRequests, log);

  log.info(`  Copied ${paraData.length} paragraphs (${allText.length} chars, ${formatRequests.length} format ops)`);

  // Phase 3: Insert images at correct positions
  const imageInserts = [];
  for (const pd of paraData) {
    for (const img of pd.imageInfo) {
      imageInserts.push({ ...img, index: pd.startIndex + img.offset });
    }
  }

  if (imageInserts.length > 0) {
    // Sort in reverse order to avoid shifting subsequent indices
    imageInserts.sort((a, b) => b.index - a.index);

    for (const img of imageInserts) {
      try {
        const req = {
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
        await GDocsHelpers.executeBatchUpdate(docs, TEST_DOC_ID, [req]);
        log.info(`  Inserted image at index ${img.index}: ${img.uri.substring(0, 60)}... (${img.width || '?'}x${img.height || '?'}pt)`);
      } catch (e) {
        log.error(`  Failed to insert image ${img.uri.substring(0, 60)}: ${e.message}`);
      }
    }
  }
}

/**
 * Copy a table to the test document, preserving cell formatting.
 */
async function copyTable(docs, table, inlineObjects) {
  const tableRows = table.tableRows || [];
  const numCols = table.columns || 0;
  const numRows = tableRows.length;

  if (numRows === 0 || numCols === 0) return;

  // Step A: Insert empty table
  const insertPoint = await getInsertionPoint(docs);
  log.info(`  Inserting table ${numRows}x${numCols} at index ${insertPoint}`);

  await GDocsHelpers.executeBatchUpdate(docs, TEST_DOC_ID, [{
    insertTable: {
      location: { index: insertPoint },
      rows: numRows,
      columns: numCols,
    }
  }]);

  // Step B: Extract formatted content from source cells
  const cellData = [];
  for (let r = 0; r < tableRows.length; r++) {
    const cells = tableRows[r].tableCells || [];
    for (let c = 0; c < cells.length; c++) {
      const { runs, imageInfo } = TableHelpers.extractFormattedCellContent(cells[c], inlineObjects);
      if (runs.length > 0 || imageInfo.length > 0) {
        cellData.push({ row: r, col: c, runs, imageInfo });
      }
    }
  }

  if (cellData.length === 0) return;

  // Step C: Fill text into cells (phase 1)
  let doc2 = await docs.documents.get({ documentId: TEST_DOC_ID });
  let body2 = doc2.data.body?.content || [];
  let tables2 = TableHelpers.extractTableElements(body2);
  const newTableIndex = tables2.length - 1;

  const textEdits = cellData
    .filter(cd => cd.runs.length > 0)
    .map(cd => ({
      row: cd.row,
      col: cd.col,
      text: cd.runs.map(r => r.text).join(''),
    }));

  if (textEdits.length > 0) {
    const textRequests = TableHelpers.buildBatchEditCellRequests(body2, newTableIndex, textEdits);
    if (textRequests.length > 0) {
      await GDocsHelpers.executeBatchUpdateChunked(docs, TEST_DOC_ID, textRequests, 50, log);
      log.info(`  Filled ${textEdits.length} cells with text`);
    }
  }

  // Step D: Apply formatting (phase 2)
  const cellsWithFormatting = cellData.filter(cd => cd.runs.some(r => r.style));
  if (cellsWithFormatting.length > 0) {
    doc2 = await docs.documents.get({ documentId: TEST_DOC_ID });
    body2 = doc2.data.body?.content || [];
    tables2 = TableHelpers.extractTableElements(body2);
    const targetTable = tables2[newTableIndex]?.table;

    if (targetTable) {
      const formatRequests = [];
      for (const cd of cellsWithFormatting) {
        const targetCell = TableHelpers.getCellElement(targetTable, cd.row, cd.col);
        const cellStart = TableHelpers.getCellInsertionPoint(targetCell);
        const cellFormatReqs = TableHelpers.buildFormattedCellFormatRequests(cellStart, cd.runs);
        formatRequests.push(...cellFormatReqs);
      }
      if (formatRequests.length > 0) {
        await GDocsHelpers.executeBatchUpdateChunked(docs, TEST_DOC_ID, formatRequests, 50, log);
        log.info(`  Applied formatting to ${cellsWithFormatting.length} cells (${formatRequests.length} style requests)`);
      }
    }
  }

  // Step E: Insert images into cells (phase 3)
  const cellsWithImages = cellData.filter(cd => cd.imageInfo.length > 0);
  if (cellsWithImages.length > 0) {
    doc2 = await docs.documents.get({ documentId: TEST_DOC_ID });
    body2 = doc2.data.body?.content || [];

    const imageInserts = [];
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
      const imageRequests = TableHelpers.buildBatchInsertImageRequests(body2, newTableIndex, imageInserts);
      if (imageRequests.length > 0) {
        await GDocsHelpers.executeBatchUpdateChunked(docs, TEST_DOC_ID, imageRequests, 50, log);
        log.info(`  Inserted ${imageInserts.length} images in table cells`);
      }
    }
  }
}

await main().catch(e => {
  console.error('FATAL:', e.message);
  if (e.response?.data?.error) console.error('API:', JSON.stringify(e.response.data.error));
  process.exit(1);
});
