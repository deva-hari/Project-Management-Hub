// --- CONSTANTS ---
const SHEET_NAMES = {
    PROJECTS: "Projects",
    ACTIONS: "Actions",
    USERS: "Users",
    SETTINGS: "Settings"
};

// --- ROUTING ---
function doGet(e) {
    const userEmail = Session.getActiveUser().getEmail();
    const template = HtmlService.createTemplateFromFile('index');
    template.userEmail = userEmail;

    return template.evaluate()
        .setTitle('Project Management Hub')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // Necessary if embedding in Google Sites
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- DATABASE INITIALIZATION ---
// Run this function ONCE from the Apps Script editor to set up the sheets
function initializeDatabase() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Create / Format Projects Sheet
    let projectsSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    if (!projectsSheet) {
        projectsSheet = ss.insertSheet(SHEET_NAMES.PROJECTS);
    }
    projectsSheet.clear(); // Reset 
    projectsSheet.appendRow(["ProjectID", "Name", "OwnerEmail", "ManagerEmail", "Status", "Phase", "PercentageCompleted", "StartDate", "Deadline", "BusinessOutcomes", "KeyRisks", "LastUpdatedText", "CreatedAt", "LastUpdated", "Comments", "ProjectType"]);
    projectsSheet.getRange("A1:P1").setFontWeight("bold").setBackground("#d9ead3");
    projectsSheet.setFrozenRows(1);

    // Create / Format Actions Sheet
    let actionsSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    if (!actionsSheet) {
        actionsSheet = ss.insertSheet(SHEET_NAMES.ACTIONS);
    }
    actionsSheet.clear();
    actionsSheet.appendRow(["ActionID", "ProjectID", "Description", "ActionOwnerEmail", "Status", "PercentageCompleted", "Priority", "LastUpdated", "Updates"]);
    actionsSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#cfe2f3");
    actionsSheet.setFrozenRows(1);

    // Create / Format Users Sheet
    let usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    if (!usersSheet) {
        usersSheet = ss.insertSheet(SHEET_NAMES.USERS);
    }
    usersSheet.clear();
    usersSheet.appendRow(["Email", "Name", "Role", "ManagerEmail", "EmailNotifications"]);
    usersSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#fff2cc");
    usersSheet.setFrozenRows(1);

    // Create / Format Settings Sheet
    let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (!settingsSheet) {
        settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    }
    settingsSheet.clear();
    settingsSheet.appendRow(["SettingKey", "SettingValue"]);
    settingsSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#ead1dc");
    settingsSheet.setFrozenRows(1);

    // Seed Default Settings
    const defaultSettings = [
        ["ProjectType", "Infrastructure"],
        ["ProjectType", "Software Development"],
        ["ProjectType", "Cloud Migration"],
        ["Status", "Not Started"],
        ["Status", "In Progress"],
        ["Status", "On Hold"],
        ["Status", "Closure"],
        ["Phase", "Open"],
        ["Phase", "In Progress"],
        ["Phase", "Execution"],
        ["Phase", "UAT"],
        ["Phase", "Monitoring"],
        ["Phase", "Closure"]
    ];
    settingsSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);

    // Add Current User as Admin
    const email = Session.getActiveUser().getEmail();
    usersSheet.appendRow([email, "System Admin", "Admin", "", true]);

    return "Database Initialized Successfully!";
}

// --- DATA FETCHING ---
function getUserRole(email) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const data = usersSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === email) {
            return data[i][2]; // Return Role
        }
    }
    return "None";
}

// Recursively finds all emails that report up to the given managerEmail
function getDownstreamEmployees(managerEmail, usersData) {
    let downstream = new Set();

    // Catch infinite loop bug on blank manager assignments
    if (!managerEmail || managerEmail.toString().trim() === "") {
        return [];
    }

    // Prevent infinite loops - add a max depth limit
    const MAX_DEPTH = 10;
    let depth = 0;

    function findReports(manager) {
        if (!manager || depth >= MAX_DEPTH) return;
        depth++;

        for (let i = 1; i < usersData.length; i++) {
            const empEmail = usersData[i][0];
            const empManager = usersData[i][3];

            if (!empEmail) continue; // Skip blank rows
            if (empEmail === manager) continue; // Skip self-reference

            if (empManager === manager) {
                if (!downstream.has(empEmail)) {
                    downstream.add(empEmail);
                    findReports(empEmail); // Recurse
                }
            }
        }
    }

    try {
        findReports(managerEmail);
    } catch (e) {
        Logger.log("Error in getDownstreamEmployees: " + e.message);
        return [];
    }
    
    return Array.from(downstream);
}

function getDashboardData() {
    const email = Session.getActiveUser().getEmail();
    const role = getUserRole(email);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log("=== getDashboardData START for " + email + " ===");
    
    // Fetch Users for hierarchy
    const usersData = ss.getSheetByName(SHEET_NAMES.USERS).getDataRange().getValues();
    
    let downstreamEmails = [];
    try {
        downstreamEmails = getDownstreamEmployees(email, usersData);
        Logger.log("getDownstreamEmployees completed successfully. Count: " + downstreamEmails.length);
    } catch (e) {
        Logger.log("ERROR in getDownstreamEmployees: " + e.message);
        downstreamEmails = [];
    }

    // Fetch Settings (for UI dropdowns)
    const settingsData = ss.getSheetByName(SHEET_NAMES.SETTINGS).getDataRange().getValues();
    settingsData.shift(); // remove headers
    let settings = {
        projectTypes: [],
        statuses: [],
        phases: []
    };
    settingsData.forEach(row => {
        if (row[0] === "ProjectType") settings.projectTypes.push(row[1]);
        if (row[0] === "Status") settings.statuses.push(row[1]);
        if (row[0] === "Phase") settings.phases.push(row[1]);
    });

    Logger.log("User: " + email + ", Role: " + role + ", Downstream: " + JSON.stringify(downstreamEmails));
    Logger.log("Settings loaded: ProjectTypes=" + settings.projectTypes.length + ", Statuses=" + settings.statuses.length + ", Phases=" + settings.phases.length);

    // 1. Fetch Projects
    const projectsSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const pData = projectsSheet.getDataRange().getValues();
    const pHeaders = pData.shift();

    let projects = [];
    pData.forEach((row, idx) => {
        try {
            if (!row[0]) return; // Skip empty rows
            
            const projectOwner = row[2] || "";
            const projectManager = row[3] || "";
            
            // Admin, Project Owner, Direct Manager, or Manager anywhere up the chain
            let canSeeProject = false;
            
            if (role === "Admin") {
                canSeeProject = true; // Admins see everything
            } else if (projectOwner === email) {
                canSeeProject = true; // Owner can see own projects
            } else if (projectManager === email) {
                canSeeProject = true; // Manager can see their projects
            } else if (projectOwner && downstreamEmails.includes(projectOwner)) {
                canSeeProject = true; // Manager of owner can see
            }
            
            if (canSeeProject) {
                // Safeguard JSON parsing
                let parsedComments = [];
                if (row[14] && typeof row[14] === 'string' && row[14].trim() !== "") {
                    try { parsedComments = JSON.parse(row[14]); } catch (e) { 
                        Logger.log("Error parsing comments for project " + row[0] + ": " + e.message);
                    }
                }

                projects.push({
                    id: row[0],
                    name: row[1],
                    owner: row[2],
                    manager: row[3],
                    status: row[4],
                    phase: row[5],
                    percentageCompleted: row[6],
                    startDate: row[7] ? new Date(row[7]).toLocaleDateString() : "",
                    deadline: row[8] ? new Date(row[8]).toLocaleDateString() : "",
                    outcomes: row[9],
                    risks: row[10],
                    lastUpdatedText: row[11],
                    createdAt: row[12] ? new Date(row[12]).toLocaleString() : "",
                    lastUpdatedDate: row[13] ? new Date(row[13]).toLocaleString() : "",
                    comments: parsedComments,
                    projectType: row[15] || "Other",
                    updates: parsedComments  // Use comments as update history
                });
            }
        } catch (err) {
            Logger.log("Error parsing project row " + idx + ": " + err.message);
        }
    });

    // 2. Fetch Actions
    const actionsSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const aData = actionsSheet.getDataRange().getValues();
    const aHeaders = aData.shift();

    let actions = [];
    aData.forEach((row, idx) => {
        try {
            // Admin, Action Owner, or Manager in the chain. 
            // Also visible if the user can see the linked project.
            if (!row[0]) return; // Skip empty rows

            const linkedProject = projects.find(p => p.id === row[1]);
            const actionOwner = row[3] || "";

            // FIXED: Check if user can actually access the linked project
            let canSeeAction = false;
            if (role === "Admin") {
                canSeeAction = true;
            } else if (actionOwner === email) {
                canSeeAction = true;
            } else if (actionOwner && downstreamEmails.includes(actionOwner)) {
                canSeeAction = true;
            } else if (linkedProject) {
                // Only if user can see the linked project
                canSeeAction = (
                    (linkedProject.owner === email) ||
                    (linkedProject.manager === email) ||
                    (linkedProject.owner && downstreamEmails.includes(linkedProject.owner))
                );
            }

            if (canSeeAction) {
                // Safeguard JSON parsing
                let parsedUpdates = [];
                if (row[8] && typeof row[8] === 'string' && row[8].trim() !== "") {
                    try { parsedUpdates = JSON.parse(row[8]); } catch (e) { }
                }

                actions.push({
                    id: row[0],
                    projectId: row[1],
                    desc: row[2],
                    owner: row[3],
                    status: row[4],
                    percentageCompleted: row[5],
                    priority: row[6],
                    lastUpdatedDate: row[7] ? new Date(row[7]).toLocaleString() : "",
                    updates: parsedUpdates
                });
            }
        } catch (err) {
            console.error("Error parsing action row " + idx + ": " + err.message);
        }
    });

    // Calculate Manager Roll-Up Metrics
    let totalProjects = projects.length;
    let activeProjects = projects.filter(p => p.status !== "Completed" && p.status !== "Closure" && p.status !== "On Hold");
    let avgCompletion = totalProjects > 0 ? Math.round(projects.reduce((acc, curr) => acc + (Number(curr.percentageCompleted) || 0), 0) / totalProjects) : 0;
    let blockedTasks = actions.filter(a => a.status === "Blocked").length;

    let statusCounts = {};
    projects.forEach(p => {
        statusCounts[p.status] = (statusCounts[p.status] || 0) + 1;
    });

    const metrics = {
        totalActive: activeProjects.length,
        averageCompletion: avgCompletion,
        blockedActions: blockedTasks,
        statusCounts: statusCounts
    };

    Logger.log("Dashboard data ready: Projects=" + projects.length + ", Actions=" + actions.length);
    Logger.log("=== getDashboardData END ===");

    // Convert users data to format for dropdown
    const userList = usersData.slice(1).map(r => ({
        email: r[0],
        name: r[1],
        role: r[2]
    }));

    return {
        email: email,
        role: role,
        projects: projects,
        actions: actions,
        settings: settings, // Provide global settings to the UI
        metrics: metrics,
        users: userList // Include users list for assignee dropdown
    };
}

// --- HELPER ENDPOINTS ---
function getScriptUrl() {
    return ScriptApp.getService().getUrl();
}

// --- ADMIN ENDPOINTS ---

function getAdminData() {
    const email = Session.getActiveUser().getEmail();
    const role = getUserRole(email);
    if (role !== "Admin") throw new Error("Unauthorized");

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Fetch All Users
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const uData = userSheet.getDataRange().getValues();
    uData.shift();
    let users = uData.map(r => ({
        email: r[0],
        name: r[1],
        role: r[2],
        manager: r[3],
        emailEnabled: r[4] === true || String(r[4]).toLowerCase() === "true"
    }));

    // 2. Fetch All Settings 
    const setSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const sData = setSheet.getDataRange().getValues();
    sData.shift();
    let settingsList = sData.map(r => ({
        key: r[0],
        value: r[1]
    }));

    return { users, settingsList };
}

function saveUser(userEmail, name, role, manager, emailEnabled) {
    const email = Session.getActiveUser().getEmail();
    if (getUserRole(email) !== "Admin") throw new Error("Unauthorized");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();

    // Update if exists
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === userEmail) {
            sheet.getRange(i + 1, 2).setValue(name);
            sheet.getRange(i + 1, 3).setValue(role);
            sheet.getRange(i + 1, 4).setValue(manager);
            sheet.getRange(i + 1, 5).setValue(emailEnabled);
            return "User updated";
        }
    }
    // Create if new
    sheet.appendRow([userEmail, name, role, manager, emailEnabled]);
    return "User added";
}

function deleteUser(userEmail) {
    const email = Session.getActiveUser().getEmail();
    if (getUserRole(email) !== "Admin") throw new Error("Unauthorized");
    if (email === userEmail) throw new Error("Cannot delete yourself");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === userEmail) {
            sheet.deleteRow(i + 1);
            return "User deleted";
        }
    }
}

function saveSetting(key, value) {
    const email = Session.getActiveUser().getEmail();
    if (getUserRole(email) !== "Admin") throw new Error("Unauthorized");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);

    // Avoid exact duplicates
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key && data[i][1] === value) {
            return "Setting already exists";
        }
    }

    sheet.appendRow([key, value]);
    return "Setting added";
}

function deleteSetting(key, value) {
    const email = Session.getActiveUser().getEmail();
    if (getUserRole(email) !== "Admin") throw new Error("Unauthorized");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i > 0; i--) { // Reverse loop for safe deletion
        if (data[i][0] === key && data[i][1] === value) {
            sheet.deleteRow(i + 1);
            return "Setting removed";
        }
    }
}

// --- DATA MUTATION ---
function generateId(prefix) {
    return prefix + "-" + Math.random().toString(36).substr(2, 6).toUpperCase();
}

function createProject(name, startDate, deadline, phase, projType, outcomes = "", risks = "") {
    const email = Session.getActiveUser().getEmail();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const newId = generateId("PRJ");
    const timestamp = new Date().toISOString();

    sheet.appendRow([
        newId,
        name,
        email,
        "", // Manager field - no longer used, role-based access from Users sheet
        "Not Started", // Status defaults here
        phase || "Open",
        0, // percentage completed
        startDate,
        deadline,
        outcomes,
        risks,
        "Project Initialized", // Last Updated Text
        timestamp, // CreatedAt
        timestamp, // LastUpdated
        "[]", // Comments JSON
        projType
    ]);

    return "Project created successfully!";
}

function createAction(projectId, desc, actionOwner, priority) {
    const email = Session.getActiveUser().getEmail();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const newId = generateId("ACT");
    const timestamp = new Date().toISOString();

    const initialLog = [{
        user: email,
        timestamp: timestamp,
        status: "Pending",
        text: "Action created and assigned."
    }];

    sheet.appendRow([
        newId,
        projectId,
        desc,
        actionOwner,
        "Pending",
        0, // Percentage completed
        priority,
        timestamp, // LastUpdated
        JSON.stringify(initialLog)
    ]);

    // Trigger Notification
    sendTaskAssignmentEmail({
        id: newId,
        desc: desc,
        owner: actionOwner
    });

    return "Action assigned successfully!";
}

function updateActionStatus(actionId, newStatus, pctComplete, updateNote) {
    const email = Session.getActiveUser().getEmail();
    const role = getUserRole(email);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const actionSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const projectSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);

    const actionData = actionSheet.getDataRange().getValues();
    const projectData = projectSheet.getDataRange().getValues();
    const usersData = usersSheet.getDataRange().getValues();

    // FIXED: Find the action first
    let actionRowIndex = -1;
    let actionRow = null;
    for (let i = 1; i < actionData.length; i++) {
        if (actionData[i][0] === actionId) {
            actionRow = actionData[i];
            actionRowIndex = i;
            break;
        }
    }

    if (!actionRow) throw new Error("Action not found");

    // FIXED: Check authorization
    const actionOwner = actionRow[3];
    const projectId = actionRow[1];
    const downstreamEmails = getDownstreamEmployees(email, usersData);

    // Find linked project
    let linkedProject = null;
    for (let i = 1; i < projectData.length; i++) {
        if (projectData[i][0] === projectId) {
            linkedProject = projectData[i];
            break;
        }
    }

    // FIXED: Authorization check
    const canUpdate = (
        role === "Admin" ||
        actionOwner === email ||
        downstreamEmails.includes(actionOwner) ||
        (linkedProject && (
            linkedProject[3] === email ||
            downstreamEmails.includes(linkedProject[3])
        ))
    );

    if (!canUpdate) {
        throw new Error("You don't have permission to update this action");
    }

    // Now safe to update
    const timestamp = new Date().toISOString();

    actionSheet.getRange(actionRowIndex + 1, 5).setValue(newStatus);
    actionSheet.getRange(actionRowIndex + 1, 6).setValue(pctComplete);
    actionSheet.getRange(actionRowIndex + 1, 8).setValue(timestamp);

    // Append to JSON Log 
    if (updateNote || newStatus) {
        let currentLog;
        try {
            currentLog = JSON.parse(actionData[actionRowIndex][8] || "[]");
        } catch (e) {
            currentLog = [];
        }

        const newLogEntry = {
            user: email,
            timestamp: timestamp,
            status: newStatus,
            text: updateNote || "Status updated."
        };
        currentLog.push(newLogEntry);

        actionSheet.getRange(actionRowIndex + 1, 9).setValue(JSON.stringify(currentLog));
    }

    return "Action updated successfully!";
}

function addProjectComment(projectId, commentText) {
    const email = Session.getActiveUser().getEmail();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === projectId) {

            const timestamp = new Date().toISOString();
            // FIXED: Column 14 (1-based) = index [13] (0-based) = LastUpdated
            sheet.getRange(i + 1, 14).setValue(timestamp);

            let currentLog;
            try {
                // FIXED: Column 15 (1-based) = index [14] (0-based) = Comments
                currentLog = JSON.parse(data[i][14] || "[]");
            } catch (e) {
                currentLog = [];
            }

            const newEntry = {
                user: email,
                timestamp: timestamp,
                text: commentText
            };
            currentLog.push(newEntry);

            // FIXED: Write to Column 15
            sheet.getRange(i + 1, 15).setValue(JSON.stringify(currentLog));
            return "Comment added to project!";
        }
    }
    throw new Error("Project ID not found.");
}

// --- NOTIFICATIONS ---
function sendTaskAssignmentEmail(action) {
    const subject = `New Task Assigned: [${action.id}]`;
    const body = `
    You have been assigned a new task.
    Task ID: ${action.id}
    Description: ${action.desc}
    
    Please log in to the Project Management Hub to view and update this task.
  `;

    try {
        //MailApp.sendEmail(action.owner, subject, body);
        MailApp.sendEmail({
                to: action.owner,
                name:     'Project Management System',     // this is the “from” name
                replyTo:  'no-reply@example.com',   // replies will go here
                subject: subject,
                htmlBody: body
            });
    } catch (e) {
        console.error("Failed to send email to: " + action.owner);
    }
}

// Automatically triggers (e.g., Daily at 5 PM)
function sendDailySummaryEmails() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const uData = ss.getSheetByName(SHEET_NAMES.USERS).getDataRange().getValues();
    uData.shift(); // Remove header

    // Find all users who opted in
    const optedInUsers = uData.filter(r => r[4] === true || String(r[4]).toLowerCase() === "true");
    if (optedInUsers.length === 0) return;

    // Load full dataset once to save time
    const pData = ss.getSheetByName(SHEET_NAMES.PROJECTS).getDataRange().getValues();
    pData.shift();
    const headersP = ["ProjectID", "Name", "OwnerEmail", "ManagerEmail", "Status", "Phase", "PercentageCompleted", "StartDate", "Deadline", "BusinessOutcomes", "KeyRisks", "LastUpdatedText", "CreatedAt", "LastUpdated", "Comments", "ProjectType"];

    const aData = ss.getSheetByName(SHEET_NAMES.ACTIONS).getDataRange().getValues();
    aData.shift();

    optedInUsers.forEach(u => {
        const userEmail = u[0];
        const role = u[2];
        const downstreamEmails = getDownstreamEmployees(userEmail, uData);

        // Find Active Projects for this user
        let activeProjectsHTML = "";
        pData.forEach(r => {
            const p = {};
            headersP.forEach((k, i) => p[k] = r[i]);

            const isRelated = p.OwnerEmail === userEmail || p.ManagerEmail === userEmail || downstreamEmails.includes(p.OwnerEmail);
            const isActive = p.Status !== "Completed" && p.Status !== "Closure" && p.Status !== "On Hold";

            if (isActive && (role === "Admin" || isRelated)) {
                activeProjectsHTML += `
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;"><b>${p.Name}</b></td>
            <td style="border: 1px solid #ddd; padding: 8px;">${p.Status} / ${p.Phase}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">${p.PercentageCompleted}%</td>
            <td style="border: 1px solid #ddd; padding: 8px;">${p.LastUpdatedText || "No updates"}</td>
          </tr>
        `;
            }
        });

        // Find Active Assigned Actions 
        let activeActionsHTML = "";
        aData.forEach(r => {
            if (r[3] === userEmail && r[4] !== "Completed") {
                activeActionsHTML += `
          <tr>
            <td style="border: 1px solid #ddd; padding: 8px;">${r[2]}</td>
            <td style="border: 1px solid #ddd; padding: 8px; color: ${r[6] === 'High' || r[6] === 'Critical' ? 'red' : 'black'}">${r[6]}</td>
            <td style="border: 1px solid #ddd; padding: 8px;">${r[4]}</td>
          </tr>
        `;
            }
        });

        // Compile Email
        if (activeProjectsHTML === "" && activeActionsHTML === "") return; // Skip if nothing to report

        const emailHtml = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <h2 style="color: #0d6efd;">Daily Project Hub Summary</h2>
        <p>Hello ${u[1]}, here is your end-of-day digest.</p>
        
        ${activeActionsHTML ? `
        <h3>Your Pending Actions</h3>
        <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
          <tr style="background-color: #f8f9fa;">
             <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Task</th>
             <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Priority</th>
             <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Status</th>
          </tr>
          ${activeActionsHTML}
        </table>` : ""}

        ${activeProjectsHTML ? `
        <h3>Active Projects Overview</h3>
        <table style="border-collapse: collapse; width: 100%;">
          <tr style="background-color: #f8f9fa;">
             <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Project</th>
             <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Status / Phase</th>
             <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">% Complete</th>
             <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Latest Note</th>
          </tr>
          ${activeProjectsHTML}
        </table>` : ""}
        
        <br>
        <p style="font-size: 12px; color: #777;">You are receiving this because your Email Notifications are enabled in the Project Hub Admin settings.</p>
      </div>
    `;

        try {
            MailApp.sendEmail({
                to: userEmail,
                name:     'Project Management System',     // this is the “from” name
                replyTo:  'no-reply@example.com',   // replies will go here
                subject: "Project Hub: Daily Wrap-Up",
                htmlBody: emailHtml
            });
        } catch (e) {
            console.error(`Failed to send digest to ${userEmail}: ` + e.message);
        }
    });
}

// --- UPDATE PROJECT ---
function updateProject(projectId, newStatus, newPercentage, updateNote) {
    const email = Session.getActiveUser().getEmail();
    const role = getUserRole(email);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projectSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const projectData = projectSheet.getDataRange().getValues();
    
    // Find project
    let projectRowIndex = -1;
    for (let i = 1; i < projectData.length; i++) {
        if (projectData[i][0] === projectId) {
            projectRowIndex = i;
            break;
        }
    }
    
    if (projectRowIndex === -1) {
        throw new Error("Project not found");
    }
    
    const projectOwner = projectData[projectRowIndex][2];
    const projectManager = projectData[projectRowIndex][3];
    
    // Authorization: Admin, Owner, or Manager
    const canUpdate = (
        role === "Admin" ||
        projectOwner === email ||
        projectManager === email
    );
    
    if (!canUpdate) {
        throw new Error("You don't have permission to update this project");
    }
    
    // Update the fields
    const timestamp = new Date().toISOString();
    
    if (newStatus) {
        projectSheet.getRange(projectRowIndex + 1, 5).setValue(newStatus);
    }
    
    if (newPercentage !== undefined && newPercentage !== null) {
        projectSheet.getRange(projectRowIndex + 1, 7).setValue(parseInt(newPercentage) || 0);
    }
    
    // Update lastUpdatedDate (column 14)
    projectSheet.getRange(projectRowIndex + 1, 14).setValue(timestamp);
    
    // Update lastUpdatedText (column 12) if note provided
    if (updateNote) {
        projectSheet.getRange(projectRowIndex + 1, 12).setValue(updateNote);
        
        // ALSO add to comments array (column 15) so it shows in history
        let currentComments = [];
        try {
            currentComments = JSON.parse(projectData[projectRowIndex][14] || "[]");
        } catch (e) {
            currentComments = [];
        }
        
        const newEntry = {
            user: email,
            timestamp: timestamp,
            text: updateNote
        };
        currentComments.push(newEntry);
        
        // Write updated comments to column 15
        projectSheet.getRange(projectRowIndex + 1, 15).setValue(JSON.stringify(currentComments));
    }
    
    Logger.log("Project " + projectId + " updated by " + email + ": Status=" + newStatus + ", %=" + newPercentage);
    
    return "Project updated successfully";
}
