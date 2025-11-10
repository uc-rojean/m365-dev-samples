const fetch = require('node-fetch');

const token = process.env.GRAPH_TOKEN;

const event = {
  subject: "Day 09 – Dev Calendar Event via Graph ✅",
  body: {
    contentType: "HTML",
    content: "<p>Created via Microsoft Graph API as part of Day 09 activities.</p>"
  },
  start: {
    dateTime: "2025-11-11T09:00:00",
    timeZone: "Asia/Manila"
  },
  end: {
    dateTime: "2025-11-11T10:00:00",
    timeZone: "Asia/Manila"
  },
  location: {
    displayName: "Universal Collab HQ"
  },
  attendees: [
    {
      emailAddress: {
        address: "rojeandeleon@universalcollab.com",
        name: "Rojean B. De Leon"
      },
      type: "required"
    }
  ],
  isOnlineMeeting: true,
  onlineMeetingProvider: "teamsForBusiness"
};

fetch("https://graph.microsoft.com/v1.0/me/events", {
  method: "POST",
  headers: {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json"
  },
  body: JSON.stringify(event)
})
.then(res => res.json())
.then(data => console.log("✅ Event created:", data))
.catch(err => console.error("❌ Error:", err));
