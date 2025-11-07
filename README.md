# Microsoft 365 Developer Samples 🚀

This repository contains **practical examples** for Microsoft 365 development, designed to keep my Microsoft 365 Developer E5 subscription active by demonstrating real development activity.

## ✅ Included Samples
- **Power Automate Flow**  
  A scheduled flow that posts daily greetings (with optional random Bible verses) to a Microsoft Teams channel.
- **Microsoft Graph API Scripts**  
  - `graph-get-profile.js` → Fetch signed-in user profile details.
  - `graph-list-teams.js` → List Teams the user has joined.
  - `graph-list-drives.js` → Enumerate OneDrive and SharePoint document libraries.
  - `graph-list-users.js` → List all users in the tenant.
  - `graph-send-message.js` → Send a message to a Teams channel.
  - *(Upcoming)* `graph-send-mail.js` → Send an email via Outlook using Graph API.
  - *(Upcoming)* Planner API examples → Create and list tasks.

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
     ```
3. **Graph Explorer Quick Calls**  
   - Try endpoints like:
     - `GET /me`
     - `GET /me/joinedTeams`
     - `GET /groups/{group-id}/planner/plans`

## 📌 Why This Repo Exists
To showcase **real Microsoft 365 development activity** across multiple workloads:
- Teams
- SharePoint / OneDrive
- Outlook
- Planner
- Power Automate
- Microsoft Graph API

---
⭐ If this helps you, give it a star!
