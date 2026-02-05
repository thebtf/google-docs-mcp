// Standalone OAuth2 authorization script
// Usage: node auth-standalone.mjs [paste-url-with-code]
import { google } from 'googleapis';
import * as fs from 'fs/promises';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
const TOKEN_PATH = path.join(__dirname, 'token.json');
const SCOPES = [
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/spreadsheets',
];

const raw = await fs.readFile(CREDENTIALS_PATH, 'utf8');
const keys = JSON.parse(raw);
const key = keys.installed || keys.web;
const redirectUri = key.redirect_uris?.[0] || 'http://localhost';

const oauth2 = new google.auth.OAuth2(key.client_id, key.client_secret, redirectUri);

const pastedUrl = process.argv[2];

if (!pastedUrl) {
  // Step 1: print auth URL
  const url = oauth2.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
  console.log('\n=== Open this URL in browser ===\n');
  console.log(url);
  console.log('\nAfter authorization, browser will redirect to http://localhost?code=...');
  console.log('Copy the FULL URL from the address bar and run:');
  console.log(`  node auth-standalone.mjs "http://localhost?code=XXXXX..."\n`);
  process.exit(0);
}

// Step 2: extract code from pasted URL and exchange for tokens
let code;
try {
  const parsed = new URL(pastedUrl);
  code = parsed.searchParams.get('code');
} catch {
  // Maybe user pasted just the code
  code = pastedUrl;
}

if (!code) {
  console.error('ERROR: Could not extract authorization code from the URL.');
  process.exit(1);
}

console.log('Exchanging code for tokens...');
const { tokens } = await oauth2.getToken(code);
oauth2.setCredentials(tokens);

const payload = {
  type: 'authorized_user',
  client_id: key.client_id,
  client_secret: key.client_secret,
  refresh_token: tokens.refresh_token,
};

await fs.writeFile(TOKEN_PATH, JSON.stringify(payload, null, 2));
console.log(`Token saved to ${TOKEN_PATH}`);
console.log('Authorization complete!');
