# Project Management Hub (Google Apps Script)

A lightweight, secure, and dynamic Project Management tool built entirely on Google Workspace. It leverages Google Sheets as the database, Google Apps Script (GAS) for backend API and logic, and HTML/CSS/JS for a reactive frontend UI.

## 📋 Features

### 1. Robust Authentication & Security

- **Secure Access**: Integrated with Google Workspace. Users automatically authenticate via their Google accounts seamlessly without external logins.
- **Role-Based Access Control (RBAC)**:
  - **Admins**: Full control over all system modules.
  - **Project Owners**: Full CRUD (Create, Read, Update, Delete) capabilities for the projects they own and ability to assign actions.
  - **Administrators**: Can now permanently remove projects (along with their linked actions) and delete individual actions or project notes directly from the UI. Deletion buttons are only visible to admin users.
  - **Managers**: Read-only oversight access to all projects and actions within their designated department.
  - **Action Owners**: Task-level access. Can view their assigned tasks, update statuses, and log progress notes.

### 2. Real-Time Dynamic Dashboard

- The web interface dynamically filters and renders project and action data based precisely on the logged-in user's email and role.
- Dashboard filters can now operate on both **status** and **phase**; both values are driven from the `Settings` sheet for consistency.
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

### Developer Notes 🔧

A quick reference to some of the key server‑side routines you’ll find in `Code.gs`:

- `initializeDatabase()` – one‑time utility that creates/clears the four sheets (`Projects`, `Actions`, `Users`, `Settings`), sets header formatting, seeds default dropdown values (project types, statuses, phases), and writes the current session user as an Admin in `Users`.
- `getDashboardData()` – the main read endpoint used by the frontend. It applies role‑based filters, builds downstream employee chains, parses JSON logs, and returns projects, actions, metrics, and settings to render the UI.
- CRUD operations:
  - `createProject(...)`, `updateProject(...)` – handle project creation (now accepting an initial status and phase) and status/phase/percentage/date updates with authorization checks. UI dropdowns for status and phase are driven by the `Settings` sheet so that admins can define new values.
  - `createAction(...)`, `updateActionStatus(...)` – actions are always assigned to the owning project’s owner; update routines append JSON logs and send assignment emails.
  - Admin-only deletion helpers: `deleteProject(...)`, `deleteAction(...)`, `deleteProjectComment(...)`.
- Notification helpers:
  - `sendTaskAssignmentEmail(...)` – standardizes the assignment message, replies go to a no‑reply address.
  - `sendDailySummaryEmails()` – time‑driven trigger to email opted‑in users a digest of their active projects/actions.
- Utility functions:
  - `getUserRole()` / `getDownstreamEmployees()` – used for access control and manager roll‑ups.
  - `generateId(prefix)` – simple random identifier generator for projects/actions.

Understanding these functions will help you customize the behavior, add new endpoints, or troubleshoot permission issues.

### Database (Google Sheets)

Acts as a relational database. Sheets are created and formatted by the `initializeDatabase()` helper function (see **Step 1.3** below). The script also adds a `Settings` sheet for dropdown options and seeds some sensible defaults, and inserts the executing user as an **Admin**.

It is composed of four tables:

#### 1. `Projects`

Tracks high-level project information and ownership.
*Columns* (in order):
`ProjectID` | `Name` | `OwnerEmail` | `ManagerEmail` | `Status` | `Phase` | `PercentageCompleted` | `StartDate` | `Deadline` | `BusinessOutcomes` | `KeyRisks` | `LastUpdatedText` | `CreatedAt` | `LastUpdated` | `Comments` | `ProjectType`

#### 2. `Actions`

Tracks granular tasks associated with specific projects.
*Columns*:
`ActionID` | `ProjectID` | `Description` | `ActionOwnerEmail` | `Status` | `PercentageCompleted` | `Priority` | `LastUpdated` | `Updates`

#### 3. `Users`

Central directory mapping emails to permissions.
*Columns*:
`Email` | `Name` | `Role` | `ManagerEmail` | `EmailNotifications`

#### 4. `Settings`

Holds key/value pairs used to populate dropdowns for project types, statuses, and phases. The initialization routine populates this sheet with a standard set of values but administrators can add/remove as needed.

---

## 🚀 Installation & Setup Guide (Step-by-Step for New Users)

### Prerequisites

- A Google Workspace or standard Google account
- Access to Google Drive, Google Sheets, and Google Apps Script
- (Optional) [Clasp](https://github.com/google/clasp) installed for local development

---

### Step 1: Create the Google Sheet Database

1. **Create a new Google Sheet**:
   - Go to [Google Drive](https://drive.google.com)
   - Click `+ Create` → `Google Sheets` → `Blank spreadsheet`
   - Name it: **"Project Management Hub Database"**

2. **Set up the required sheets manually** (optional):
   - Delete the default "Sheet1" tab
   - Add new tabs named exactly `Projects`, `Actions`, `Users`, and `Settings`.

3. **Run the initializer (optional but recommended)**:
   - Open the Apps Script editor after you've copied `Code.gs` (see Step 2 below).
   - Choose the `initializeDatabase` function from the dropdown at the top and click the ▶️ run button.
   - Approve any authorization requests that appear. The script will create/clear all four sheets, set bold headers, add a row of default settings values, and insert your email as an Admin user.

> **Tip:** if you run this function you can skip steps **3–6** below since the required columns are already populated.

> **Tip:** the repo contains a handy Apps Script helper called `initializeDatabase()` that will create and format all four sheets for you with the correct column headers, seed default dropdown entries, and add the current user as an Admin.
### Recent Updates

* Modal dialogs are now vertically centered, capped at 90% of the viewport height, and the body becomes scrollable on smaller screens.
* Email notifications use a consistent display name and reply‑to address so replies are routed correctly.  Name/`from` fields are fixed for both task assignment and daily summary messages.
* Admin users can remove projects, actions and notes (see above).

3. **(Manual alternative) Configure the Projects sheet**:
   - Click the `Projects` tab
   - In row 1, add these column headers (A to P):
     - `ProjectID`, `Name`, `OwnerEmail`, `ManagerEmail`, `Status`, `Phase`, `PercentageCompleted`, `StartDate`, `Deadline`, `BusinessOutcomes`, `KeyRisks`, `LastUpdatedText`, `CreatedAt`, `LastUpdated`, `Comments`, `ProjectType`

4. **(Manual alternative) Configure the Actions sheet**:
   - Click the `Actions` tab
   - In row 1, add these column headers (A to I):
     - `ActionID`, `ProjectID`, `Description`, `ActionOwnerEmail`, `Status`, `PercentageCompleted`, `Priority`, `LastUpdated`, `Updates`

5. **(Manual alternative) Configure the Users sheet**:
   - Click the `Users` tab
   - In row 1, add these column headers (A to E):
     - `Email`, `Name`, `Role`, `ManagerEmail`, `EmailNotifications`
   - Add at least one user (your email):
     - Example: `your.email@company.com` | `Your Name` | `Admin` | `IT`
     - Roles can be: `Admin`, `ProjectOwner`, `Manager`, or `ActionOwner`

6. **(Manual alternative) Configure the Settings sheet**:
   - Click the `Settings` tab
   - In row 1, add headers `SettingKey` and `SettingValue`. This sheet is used by the app to produce dropdown lists for statuses, phases and project types. The initializer function will populate a useful starter set of values.

---

### Step 2: Create the Google Apps Script Project

1. **Open Apps Script editor**:
   - In your newly created Google Sheet, click `Extensions` → `Apps Script`
   - This will open the Apps Script editor in a new tab

2. **Copy the code files**:
   - You'll need three files: `Code.gs`, and HTML/CSS/JS files
   - In the Apps Script editor, create files as follows:

   **Step 2a: Create `Code.gs`** (Backend Logic)
   - In the left panel, click `+ New` → `File`
   - Name it: `Code.gs`
   - Copy and paste the entire contents of the `Code.js` file from this repository
   - Save (Ctrl+S)

   **Step 2b: Create HTML/CSS/JS files** (Frontend UI)
   - Create a file named `index.html`
   - Copy contents from the repository's `index.html`
   - Create a file named `styles.html`
   - Copy contents from the repository's `styles.html`
   - Create a file named `scripts.html`
   - Copy contents from the repository's `scripts.html`
   - Create a file named `appsscript.json`
   - Copy contents from the repository's `appsscript.json`
   - Save all files (Ctrl+S)

3. **No manual sheet ID required (container‑bound script)**:
   - This project is designed to be copied directly into the Apps Script editor opened from the spreadsheet itself. The backend uses `SpreadsheetApp.getActiveSpreadsheet()` so you don't need to hard‑code a sheet ID. Just make sure you open the script by navigating to `Extensions → Apps Script` from the sheet you created.

---

**Note:** When you open the web app the first time after deployment, you'll be able to create new projects using the modal. The form now includes a **Project Name** field and a **Project Type** dropdown (driven by the Settings sheet). Likewise, the **Edit Project** dialog exposes the project type so it can be adjusted later. If you run `initializeDatabase()` this form will be pre‑populated with default project types, statuses and phases.

### Step 3: Deploy as a Web App

1. **Create a deployment**:
   - In the Apps Script editor, click the `Deploy` button (top right, next to `+ New`)
   - Select `New deployment`

2. **Configure deployment settings**:
   - Click the dropdown and select `Web app`
   - **Execute as**: Select your Google account (appears as your email)
   - **Who has access**:
     - For testing: Select `Only myself`
     - For production: Select `Anyone with a Google account` or restrict to your Google Workspace domain
   - Click `Deploy`

3. **Grant permissions**:
   - A popup may ask for Google permissions
   - Click `Review permissions` → `Select your account` → `Allow`
   - You'll see a message: "New deployment created successfully"

4. **Copy the Web App URL**:
   - A link starting with `https://script.google.com/macros/d/...` will appear
   - Copy this URL and save it somewhere (this is your app URL)

---

### Step 4: Test the Installation

1. **Open the web app**:
   - Paste the deployment URL into a new browser tab
   - You should see the Project Management Hub dashboard

2. **Initial setup test**:
   - Check if you can see the dashboard
   - If you see an error, check the browser console (F12 → Console tab) for error messages

3. **Add sample data**:
   - Return to your Google Sheet
   - Click the `Projects` tab
   - Add a sample project in row 2:
     - `ProjectID`: `P001`
     - `Name`: `Test Project`
     - `OwnerEmail`: Your email address
     - `Status`: `Planning`
     - And fill in other fields as needed
   - Refresh the web app to see the project appear

---

### Step 5: Add Users (Critical Step)

1. **Open the Users sheet** in your Google Sheet
2. **Add all users who will access the app**:
   - Each row should have: `Email`, `Name`, `Role`, `Department`
   - **Roles**:
     - `Admin`: Full system access
     - `ProjectOwner`: Can create/edit projects and assign actions
     - `Manager`: Read-only access to projects in their department
     - `ActionOwner`: Can only view and update their assigned actions

   **Example**:

   ```markdown
   Email                    | Name           | Role         | Department
   admin@company.com        | Admin User     | Admin        | IT
   john.owner@company.com   | John Owner     | ProjectOwner | Engineering
   jane.manager@company.com | Jane Manager   | Manager      | Engineering
   ```

---

### Step 6: Ongoing Management

**To update the app code**:

1. Return to the Apps Script editor
2. Edit the files as needed
3. Save (Ctrl+S)
4. Click `Deploy` → `Manage deployments`
5. Click the edit icon (pencil) on your web app deployment
6. Change the version or make other edits, then save
7. Users will see the updated version on refresh

**To add new users**:

- Simply add rows to the `Users` sheet
- They'll have access once their email is in the system

**To modify project/action data**:

- Edit directly in the Google Sheets tabs
- Changes appear in the app after a refresh

---

### Troubleshooting

**"Error: Cannot read properties of undefined"**

- Make sure the Sheet ID in `Code.gs` matches your actual Google Sheet ID
- Verify all three sheets (`Projects`, `Actions`, `Users`) exist with correct names
- Check that column headers are exactly as specified above

**Web app shows blank page**

- Open browser console (F12) to see error messages
- Click the Apps Script deployment link and re-authorize permissions
- Check that `index.html` contains the full HTML structure

**Users can't see projects**

- Verify their email is in the `Users` sheet
- Check the `ManagerEmail` or `OwnerEmail` fields match exactly (case-sensitive)
- Ensure their `Role` is set correctly

**Can't deploy as web app**

- Make sure you have a Google Workspace account or standard Google account
- Try going to `Extensions > Apps Script` from within the Google Sheet
- Ensure all code files are properly saved before deploying

---

### Development & Setup Guide

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
