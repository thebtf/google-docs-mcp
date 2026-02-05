// Quick test script for table tools
import { google } from 'googleapis';
import * as fs from 'fs/promises';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const creds = JSON.parse(await fs.readFile(path.join(__dirname, 'credentials.json'), 'utf8'));
const token = JSON.parse(await fs.readFile(path.join(__dirname, 'token.json'), 'utf8'));
const key = creds.installed || creds.web;

const oauth2 = new google.auth.OAuth2(key.client_id, key.client_secret, key.redirect_uris?.[0]);
oauth2.setCredentials(token);
const docs = google.docs({ version: 'v1', auth: oauth2 });

const DOC_ID = process.argv[2] || '1bI1auzYY018Urt6TIHvAoemWYr30zLc-UNSuJ-X1RlA';

console.log(`\n=== Fetching document ${DOC_ID} ===\n`);
const res = await docs.documents.get({ documentId: DOC_ID });
const body = res.data.body?.content || [];

// Import helpers dynamically
const { getTablesInfo, readTableCells } = await import('./dist/tableHelpers.js');

// Test 1: getTableStructure
console.log('=== TABLE STRUCTURE ===');
const tables = getTablesInfo(body);
console.log(JSON.stringify(tables, null, 2));

// Test 2: readTableCells for each table
for (const t of tables) {
  console.log(`\n=== TABLE ${t.tableIndex} CELLS ===`);
  const cells = readTableCells(body, t.tableIndex);
  console.log(JSON.stringify(cells, null, 2));
}
