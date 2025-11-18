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
- *(Upcoming)* Planner API examples → Create and list tasks.

---

## 🔹 Day 09 – Microsoft Graph API Activities
As part of my Microsoft 365 Developer Program renewal, I completed the following advanced activities using Graph API:

- 🗂 **Planner**: Created tasks via `/planner/tasks` endpoint
- ✉️ **Outlook Mail**: Sent email using `/me/sendMail`
- 📅 **Calendar**: Created event with Teams link via `/me/events`
- 📝 **GitHub**: Committed updated scripts and README to track dev signals

These activities demonstrate real usage of Microsoft 365 workloads and help maintain my Developer E5 subscription.

### 📂 New Scripts Added
- `graph-send-mail.js` — Sends Outlook email via Graph API
- `graph-create-event.js` — Creates calendar event with Teams link

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

### 🌐 Stay Connected
⭐ If this repo helps you, **give it a star** and share your feedback!  
📬 Questions? Reach out via Microsoft Graph Explorer.


## 📌 Day 11 – Microsoft Graph API Activities

Continuing my renewal journey, here are the Day 11 activities:

👥 **Users**:  
- Listed all users in the tenant via `/users` endpoint

💬 **Teams**:  
- Listed all channels in the Team via `/teams/{team-id}/channels`

📁 **SharePoint**:  
- Listed drives in the SharePoint site via `/sites/{site-id}/drives`

📁 **GitHub**:  
- Updated README.md with Day 11 activity summary

## 📌 Day 12 – Microsoft 365 Developer Activities

💬 **Teams**  
- Posted a channel message via `/teams/{team-id}/channels/{channel-id}/messages`

📁 **SharePoint**  
- Uploaded a file to Group drive root via `/groups/{group-id}/drive/root:/...:/content`

🔁 **Power Automate**  
- Manually triggered Bible Verse / Daily Greeting flow (succeeded)

🔎 **Bonus**  
- Continued daily Graph signals to maintain renewal activity

## 📌 Day 13 – Microsoft 365 Developer Activities
✅ Created Planner task via Graph API  
✅ Sent Outlook email via Graph API  
✅ Optimized GitHub repo visibility and tags  

## 📌 Day 14 – Microsoft 365 Developer Activities
**Teams** — Posted a channel message (with/without @mention)  
**SharePoint** — Uploaded a text file with meaningful content  
**Power Automate** — Manually triggered Bible Verse / Greeting flow (Succeeded)  
**GitHub** — Made repo public and added Microsoft 365/Graph API tags

## 📌 Day 15 – Microsoft 365 Developer Activities
**Planner**
- Created a task in “Tasks by Planner and To Do” and assigned it to myself

**Outlook**
- Sent an email via `/me/sendMail`

**SharePoint**
- Created `Day15/` and uploaded `Day15-Update.txt`

**Bonus**
- Queried Planner tasks, Messages, and Recent files for telemetry
