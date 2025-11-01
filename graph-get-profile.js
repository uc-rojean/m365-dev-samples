/**
 * Simple Microsoft Graph GET /me example
 * IMPORTANT: Do not commit real access tokens to GitHub.
 * Use this file for learning/demo only.
 */

// 1) Get a temporary access token from Graph Explorer:
// https://developer.microsoft.com/graph/graph-explorer
// Sign in with your DEV tenant, run any sample, then click "Access token" to copy it.
// (Tokens expire quickly; never commit a real token to your repo.)

const endpoint = "https://graph.microsoft.com/v1.0/me";

// Paste a temporary token to test locally (DO NOT COMMIT the token)
const token = process.env.GRAPH_TOKEN || "PASTE_TEMP_TOKEN_FOR_LOCAL_TEST_ONLY";

// Node 18+ includes fetch; for older Node versions install node-fetch.
// npm install node-fetch@2  (if needed)  and then uncomment the line below:
// const fetch = require('node-fetch');

async function run() {
  try {
    if (!token || token.startsWith("PASTE_")) {
      console.error("⚠️ No token found. Set GRAPH_TOKEN env var or paste a temporary token for local testing.");
      return;
    }

    const res = await fetch(endpoint, {
      headers: {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json"
      }
    });

    if (!res.ok) {
      console.error("Request failed:", res.status, await res.text());
      return;
    }

    const data = await res.json();
    console.log("✅ User Profile:", {
      displayName: data.displayName,
      userPrincipalName: data.userPrincipalName,
      id: data.id
    });
  } catch (err) {
    console.error("Error:", err);
  }
}

run();
