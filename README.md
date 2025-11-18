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
- `graph-create-event.js` → Create a calendar event with Teams link.
- **Planner API examples** → Create and list tasks.
- **Calendar API examples** → Create events programmatically.

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
     node graph-create-event.js
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

## 🌍 Stay Connected
⭐ If this repo helps you, **give it a star** and share your feedback!  
📬 Questions? Reach out via Microsoft Graph Explorer.

---

# 📅 Activity Log (Day 01 → Day 15)

### ✅ Day 01 – Kickoff
- Created Power Automate flow for Daily Greeting Bot
- Tested basic Graph API calls: `/me`, `/me/joinedTeams`

### ✅ Day 02 – Teams API Basics
- Listed Teams and channels via Graph API
- Updated README.md with Day 02 summary

### ✅ Day 03 – SharePoint Read Signal
- Listed document libraries via `/sites/{site-id}/drives`
- Confirmed OneDrive provisioning status

### ✅ Day 04 – Teams Message Post
- Posted a message to a Teams channel via `/teams/{team-id}/channels/{channel-id}/messages`

### ✅ Day 05 – SharePoint Upload
- Uploaded a file to Group drive root via `/groups/{group-id}/drive/root:/...:/content`

### ✅ Day 06 – Power Automate Flow Trigger
- Manually triggered Bible Verse / Daily Greeting flow (Succeeded)

### ✅ Day 07 – GitHub Activity
- Committed updated scripts and README.md
- Added tags: `Microsoft365`, `GraphAPI`, `PowerAutomate`

### ✅ Day 08 – Bonus Graph Signals
- Queried `/me/messages`, `/me/events`, `/me/drive/recent`

### ✅ Day 09 – Advanced Graph API Activities
- 📂 **Planner**: Created tasks via `/planner/tasks`
- ✉️ **Outlook Mail**: Sent email using `/me/sendMail`
- 📅 **Calendar**: Created event with Teams link via `/me/events`
- 📝 **GitHub**: Committed updated scripts and README.md

### ✅ Day 10 – High Signal Boost
- **SharePoint**: Listed and uploaded files via Group drive
- **Teams**: Posted a channel message via Graph API
- **Power Automate**: Triggered flow manually
- **GitHub**: Updated README.md with Day 10 summary

### ✅ Day 11 – Microsoft Graph API Activities
- **Users**: Listed all users via `/users`
- **Teams**: Listed all channels via `/teams/{team-id}/channels`
- **SharePoint**: Listed drives via `/sites/{site-id}/drives`
- **GitHub**: Updated README.md with Day 11 summary

### ✅ Day 12 – Microsoft 365 Developer Activities
- **Teams**: Posted a channel message via Graph API
- **SharePoint**: Uploaded a file to Group drive root
- **Power Automate**: Triggered Bible Verse flow (Succeeded)
- **Bonus**: Continued daily Graph signals

### ✅ Day 13 – Microsoft 365 Developer Activities
- Created Planner task via Graph API
- Sent Outlook email via Graph API
- Optimized GitHub repo visibility and tags

### ✅ Day 14 – Microsoft 365 Developer Activities
- **Teams**: Posted a channel message (with/without @mention)
- **SharePoint**: Uploaded a text file with meaningful content
- **Power Automate**: Triggered Bible Verse flow (Succeeded)
- **GitHub**: Made repo public and added Microsoft 365/Graph API tags

### ✅ Day 15 – Microsoft 365 Developer Activities
- **Planner**: Created a task in “Tasks by Planner and To Do” and assigned it to myself
- **Outlook**: Sent an email via `/me/sendMail`
- **SharePoint**: Created `Day15/` and uploaded `Day15-Update.txt`
- **Bonus**: Queried Planner tasks, Messages, and Recent files for telemetry

---

## 📂 New Scripts Added
- `graph-send-mail.js` — Sends Outlook email via Graph API
- `graph-create-event.js` — Creates calendar event with Teams link

---

### ✅ Commit Message Suggestions
```bash
feat: Day 15 Graph API activities – Planner, Mail, SharePoint
