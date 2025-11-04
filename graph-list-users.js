/**
 * Microsoft Graph API Example: List all users in the tenant
 */
const endpoint = "https://graph.microsoft.com/v1.0/users";
const token = process.env.GRAPH_TOKEN || "PASTE_TEMP_TOKEN_FOR_LOCAL_TEST_ONLY";

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
    console.log("✅ Users in Tenant:");
    data.value.forEach(user => {
      console.log(`- ${user.displayName} (${user.userPrincipalName})`);
    });
  } catch (err) {
    console.error("Error:", err);
  }
}

run();
