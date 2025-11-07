/**
 * Microsoft Graph API Example: Send an email via Outlook (delegated)
 *
 * How to run (Node.js 18+):
 *   On Windows CMD:    set GRAPH_TOKEN=eyJ... && node graph-send-mail.js
 *   On PowerShell:     $env:GRAPH_TOKEN="eyJ..." ; node graph-send-mail.js
 *   On macOS/Linux:    export GRAPH_TOKEN=eyJ... && node graph-send-mail.js
 *
 * IMPORTANT:
 * - Get a short-lived delegated access token from Graph Explorer:
 *     https://developer.microsoft.com/graph/graph-explorer
 *   Sign in with your DEV tenant -> run any query -> "Access token" tab -> Copy.
 * - Required delegated Graph permissions (consent in Graph Explorer if prompted):
 *     Mail.Send  (to send email)
 *     Mail.Read  (optional; not needed to send)
 *
 * NOTE: Node 18+ includes global fetch. If you’re on Node < 18, install node-fetch@2:
 *   npm init -y && npm install node-fetch@2
 *   Then uncomment the require('node-fetch') line below.
 */

// For Node < 18, uncomment the next line:
// const fetch = require('node-fetch');

const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0/me/sendMail";

// ---- Configuration (edit these values or provide via env vars) ----
const TO_EMAIL = process.env.GRAPH_TO || "your.recipient@yourdomain.com";
const SUBJECT  = process.env.GRAPH_SUBJECT || "Hello from Microsoft Graph API ✅";
const BODY_HTML = process.env.GRAPH_BODY_HTML || `
  <p>Hello from <strong>Microsoft Graph</strong> via Node.js! 🚀</p>
  <p>This message was sent using the <code>/me/sendMail</code> endpoint.</p>
  <p>— UC Dev Lab</p>
`;

// Optional: CC and BCC (comma-separated emails in env vars)
const CC_EMAILS  = (process.env.GRAPH_CC  || "").split(",").map(s => s.trim()).filter(Boolean);
const BCC_EMAILS = (process.env.GRAPH_BCC || "").split(",").map(s => s.trim()).filter(Boolean);

// Attachments example (disabled by default). To enable, set GRAPH_ATTACH_PATH to a file path.
// The script will attach that file as a Base64 contentBytes.
const ATTACH_PATH = process.env.GRAPH_ATTACH_PATH || ""; // e.g., "./docs/hello.txt"

const TOKEN = process.env.GRAPH_TOKEN || "PASTE_TEMP_TOKEN_FOR_LOCAL_TEST_ONLY";

// ---- Helpers ----
const fs = require('fs');
const path = require('path');

/**
 * Build the Graph sendMail payload.
 * Docs: https://learn.microsoft.com/graph/api/user-sendmail
 */
function buildMessagePayload() {
  const toRecipients = [{ emailAddress: { address: TO_EMAIL } }];

  const ccRecipients = CC_EMAILS.map(addr => ({ emailAddress: { address: addr } }));
  const bccRecipients = BCC_EMAILS.map(addr => ({ emailAddress: { address: addr } }));

  const message = {
    subject: SUBJECT,
    body: {
      contentType: "HTML",
      content: BODY_HTML
    },
    toRecipients,
  };

  if (ccRecipients.length)  message.ccRecipients  = ccRecipients;
  if (bccRecipients.length) message.bccRecipients = bccRecipients;

  // Optional attachment (single file)
  if (ATTACH_PATH) {
    const full = path.resolve(ATTACH_PATH);
    if (!fs.existsSync(full)) {
      console.warn(`⚠️  Attachment not found at ${full} — sending without attachment.`);
    } else {
      const bytes = fs.readFileSync(full);
      const base64 = bytes.toString('base64');
      const name = path.basename(full);

      message.attachments = [{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name,
        contentType: guessContentType(name),
        contentBytes: base64
      }];
    }
  }

  return { message, saveToSentItems: true };
}

/**
 * Lightweight content-type guess based on file extension.
 */
function guessContentType(filename) {
  const ext = path.extname(filename).toLowerCase();
  switch (ext) {
    case '.txt':  return 'text/plain';
    case '.html': return 'text/html';
    case '.json': return 'application/json';
    case '.pdf':  return 'application/pdf';
    case '.png':  return 'image/png';
    case '.jpg':
    case '.jpeg': return 'image/jpeg';
    default:      return 'application/octet-stream';
  }
}

async function sendMail() {
  try {
    if (!TOKEN || TOKEN.startsWith("PASTE_")) {
      console.error("⚠️  No valid token found. Set GRAPH_TOKEN env var with a delegated access token.");
      process.exitCode = 1;
      return;
    }

    const payload = buildMessagePayload();

    const res = await fetch(GRAPH_ENDPOINT, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${TOKEN}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    if (res.status === 202) {
      console.log("✅ Email accepted for delivery (202 Accepted). Check your Sent Items.");
      return;
    }

    // If not 202, print diagnostic info
    const text = await res.text();
    console.error("❌ Request failed:", res.status, text);
    if (res.status === 403) {
      console.error("Hint: Ensure 'Mail.Send' (delegated) is granted in Graph Explorer → Modify permissions.");
    } else if (res.status === 401) {
      console.error("Hint: Token expired. Grab a fresh token from Graph Explorer and set GRAPH_TOKEN again.");
    }
  } catch (err) {
    console.error("❌ Error sending mail:", err);
  }
}

sendMail();
