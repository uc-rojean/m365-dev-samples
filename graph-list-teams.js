/**
 * Microsoft Graph API Example: List Teams the user has joined
 * IMPORTANT: Never commit real tokens to GitHub.
 */

const endpoint = "https://graph.microsoft.com/v1.0/me/joinedTeams";

// Use environment variable for token (recommended)
const token = process.env.GRAPH_TOKEN || "PASTE_TEMP_TOKEN_FOR_LOCAL_TEST_ONLY";

// Node 18+ includes fetch; for older Node versions, install node-fetch
// npm install node-fetch@2
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
    console.log("✅ Teams Joined:");
    data.value.forEach(team => {
      console.log(`- ${team.displayName} (ID: ${team.id})`);
    });
  } catch (err) {
    console.error("Error:", err);
  }
}

run();
