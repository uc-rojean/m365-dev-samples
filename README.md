# 🌟 Microsoft 365 Developer Samples 🚀

Welcome to **Microsoft 365 Developer Samples** — a curated set of practical examples to showcase **real development activity** across Microsoft 365 services.  
This helps keep my **Microsoft 365 Developer E5 subscription active** while exploring Teams, SharePoint, Outlook, Planner, Power Automate, and Graph API.

---

## ✅ Included Samples
### 🔹 Power Automate Flow
- **Daily Greeting Bot**  
  Posts scheduled greetings (with optional random Bible verses) to a Microsoft Teams channel.

### 🔹 Microsoft Graph API Scripts
- `graph-get-profile.js` → Fetch signed-in user profile details.
- `graph-list-teams.js` → List Teams the user has joined.
- `graph-list-drives.js` → Enumerate OneDrive and SharePoint document libraries.
- `graph-list-users.js` → List all users in the tenant.
- `graph-send-message.js` → Send a message to a Teams channel.
- `graph-send-mail.js` → Send an email via Outlook using Graph API.
- *(Upcoming)* Planner API examples → Create and list tasks.
- *(Upcoming)* Calendar API examples → Create events programmatically.

---

## 🛠 How to Use
1. **Power Automate Flow**
   - Import the `.zip` package into your environment.
   - Configure your Team and Channel for scheduled posts.

2. **Graph API Scripts**
   - Requires Node.js (v18+ recommended).
   - Set your `GRAPH_TOKEN` environment variable with a valid OAuth token from Graph Explorer.
   - Run scripts:
     ```bash
     node graph-get-profile.js
     node graph-list-teams.js
     node graph-send-mail.js
     ```

3. **Graph Explorer Quick Calls**
   - Try endpoints like:
     - `GET /me`
     - `GET /me/joinedTeams`
     - `GET /groups/{group-id}/planner/plans`

---

## 📌 Why This Repo Exists
To demonstrate **real Microsoft 365 development activity** across multiple workloads:
- ✅ Teams
- ✅ SharePoint / OneDrive
- ✅ Outlook
- ✅ Planner
- ✅ Power Automate
- ✅ Microsoft Graph API

---

### 🌐 Stay Connected
⭐ If this repo helps you, **give it a star** and share your feedback!  
📬 Questions? Reach out via Microsoft Graph Explorer.

---
