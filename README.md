# Project Management Hub (Google Apps Script)

A lightweight, secure, and dynamic Project Management tool built entirely on Google Workspace. It leverages Google Sheets as the database, Google Apps Script (GAS) for backend API and logic, and HTML/CSS/JS for a reactive frontend UI.

## 📋 Features

### 1. Robust Authentication & Security
- **Secure Access**: Integrated with Google Workspace. Users automatically authenticate via their Google accounts seamlessly without external logins.
- **Role-Based Access Control (RBAC)**:
  - **Admins**: Full control over all system modules.
  - **Project Owners**: Full CRUD (Create, Read, Update, Delete) capabilities for the projects they own and ability to assign actions.
  - **Managers**: Read-only oversight access to all projects and actions within their designated department.
  - **Action Owners**: Task-level access. Can view their assigned tasks, update statuses, and log progress notes.

### 2. Real-Time Dynamic Dashboard
- The web interface dynamically filters and renders project and action data based precisely on the logged-in user's email and role.
- Single-page application (SPA) feel, utilizing asynchronous `google.script.run` calls to fetch and save data without page reloads.

### 3. Workflow Automation
- **Automated Notifications**: Utilizes GAS `MailApp` to automatically send email triggers when new tasks are assigned.
- **Deadline Reminders**: Time-driven triggers to alert Action Owners and Project Owners of approaching deadlines.

---

## 🏗 System Architecture

### Backend (Google Apps Script)
The GAS environment acts as the server. 
- Serves the frontend via the `HtmlService`.
- Contains controller logic functions (e.g., `getProjects()`, `updateTask()`) that the frontend calls asynchronously.
- Manages security validation (verifying user session email against the database records before executing writes).

### Frontend (HtmlService Web App)
- Clean, responsive HTML/JS interface.
- Built utilizing templated HTML (`index.html`, `css.html`, `js.html`) to keep code modular within the GAS environment.

### Database (Google Sheets)
Acts as a relational database. It is structured into three primary tables:

#### 1. `Projects`
Tracks high-level project information and ownership.
*Columns*: `ProjectID` | `Name` | `OwnerEmail` | `ManagerEmail` | `Status` | `Deadline` | `CreatedAt`

#### 2. `Actions`
Tracks granular tasks associated with specific projects.
*Columns*: `ActionID` | `ProjectID` | `Description` | `ActionOwnerEmail` | `Status` | `Priority` | `UpdateLog`

#### 3. `Users`
Central directory mapping emails to permissions.
*Columns*: `Email` | `Name` | `Role` | `Department`

---

## 🚀 Development & Setup Guide

### Prerequisites
1. A Google Workspace or standard Google account.
2. [Clasp (Command Line Apps Script Projects)](https://github.com/google/clasp) installed locally if developing outside the browser editor.

### Initial Setup
1. **Create the Database**: 
   - Create a new Google Sheet.
   - Create 3 tabs named exactly: `Projects`, `Actions`, and `Users`.
   - Add the column headers defined in the Database section above to row 1 of each tab.
2. **Initialize the Script**:
   - Go to `Extensions > Apps Script` in the Google Sheet.
   - (Or use `clasp clone <script-id>` to develop locally).
3. **Deploy as Web App**:
   - In the Apps Script editor, click **Deploy > New deployment**.
   - Select **Web app**.
   - Execute as: **User accessing the web app**.
   - Who has access: **Anyone with a Google account** (or restrict to your Workspace domain).

### Interacting with the API
Frontend JS communicates with the backend exclusively via `google.script.run`:
```javascript
// Example: Fetching user dashboard data
google.script.run
  .withSuccessHandler(function(data) {
    renderDashboard(data);
  })
  .withFailureHandler(function(error) {
    console.error("Failed to load dashboard:", error);
  })
  .getDashboardData(); // Server-side function in Code.gs
```

---

## 🔮 Future Enhancements (Roadmap)
- **File Attachments**: Integration with Google Drive API to attach mockups/specs to Projects.
- **Calendar Integration**: Syncing project deadlines automatically to the Project Owner's Google Calendar.
- **Advanced Charting**: Adding Chart.js to the frontend to visualize completion metrics and bottleneck analysis.
