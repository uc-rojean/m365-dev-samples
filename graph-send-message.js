/**
 * Microsoft Graph API Example: Send a message to a Teams channel
 * Requires Channel ID and Team ID
 */
const endpoint = "https://graph.microsoft.com/v1.0/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages";
const token = process.env.GRAPH_TOKEN || "PASTE_TEMP_TOKEN_FOR_LOCAL_TEST_ONLY";

async function run() {
  try {
    if (!token || token.startsWith("PASTE_")) {
      console.error("⚠️ No token found. Set GRAPH_TOKEN env var or paste a temporary token for local testing.");
      return;
    }

    const messageBody = {
      body: {
        content: "Hello from Microsoft Graph API! 🚀"
      }
    };

    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(messageBody)
    });

    if (!res.ok) {
      console.error("Request failed:", res.status, await res.text());
      return;
    }

    const data = await res.json();
    console.log("✅ Message sent:", data.id);
  } catch (err) {
    console.error("Error:", err);
  }
}

run();
