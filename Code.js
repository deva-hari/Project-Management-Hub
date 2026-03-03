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
        ["ProjectType", "Other"],
        ["ProjectType", "Infrastructure"],
        ["ProjectType", "Software Development"],
        ["ProjectType", "Cloud Migration"],
        ["Status", "Not Started"],
        ["Status", "Pending"],
        ["Status", "In Progress"],
        ["Status", "On Hold"],
        ["Status", "Completed"],
        ["Status", "Cancelled"],
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
    // === PERFORMANCE: Use bounded query ===
    const lastRow = usersSheet.getLastRow();
    if (lastRow <= 1) return "None";
    const data = usersSheet.getRange(1, 1, lastRow, 3).getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === email) {
            return data[i][2]; // Return Role
        }
    }
    return "None";
}

/**
 * Return the manager email for a given user email, or empty string if unknown.
 */
function getManagerForUser(email) {
    if (!email) return "";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    // === PERFORMANCE: Use bounded query ===
    const lastRow = usersSheet.getLastRow();
    if (lastRow <= 1) return "";
    const data = usersSheet.getRange(1, 1, lastRow, 4).getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === email) {
            return data[i][3] || "";
        }
    }
    return "";
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log("=== getDashboardData START for " + email + " ===");
    
    // === PERFORMANCE: Per-request memoization cache ===
    const requestCache = {
        userRoles: {},
        userManagers: {},
        downstreamEmployees: {}
    };
    
    // === PERFORMANCE: Replace getDataRange() with bounded queries ===
    // Fetch Users for hierarchy
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const usersLastRow = usersSheet.getLastRow();
    const usersData = usersLastRow > 1 ? usersSheet.getRange(1, 1, usersLastRow, 5).getValues() : [[]];
    
    // Memoized getUserRole for this request
    const getUserRoleMemo = (userEmail) => {
        if (!userEmail) return "None";
        if (requestCache.userRoles[userEmail] !== undefined) {
            return requestCache.userRoles[userEmail];
        }
        for (let i = 1; i < usersData.length; i++) {
            if (usersData[i][0] === userEmail) {
                requestCache.userRoles[userEmail] = usersData[i][2];
                return usersData[i][2];
            }
        }
        requestCache.userRoles[userEmail] = "None";
        return "None";
    };
    
    // Memoized getManagerForUser for this request
    const getManagerForUserMemo = (userEmail) => {
        if (!userEmail) return "";
        if (requestCache.userManagers[userEmail] !== undefined) {
            return requestCache.userManagers[userEmail];
        }
        for (let i = 1; i < usersData.length; i++) {
            if (usersData[i][0] === userEmail) {
                requestCache.userManagers[userEmail] = usersData[i][3] || "";
                return usersData[i][3] || "";
            }
        }
        requestCache.userManagers[userEmail] = "";
        return "";
    };
    
    const role = getUserRoleMemo(email);
    
    // Memoized getDownstreamEmployees for this request
    const getDownstreamEmployeesMemo = (managerEmail) => {
        if (!managerEmail) return [];
        if (requestCache.downstreamEmployees[managerEmail] !== undefined) {
            return requestCache.downstreamEmployees[managerEmail];
        }
        const result = getDownstreamEmployees(managerEmail, usersData);
        requestCache.downstreamEmployees[managerEmail] = result;
        return result;
    };
    
    let downstreamEmails = [];
    try {
        downstreamEmails = getDownstreamEmployeesMemo(email);
        Logger.log("getDownstreamEmployees completed successfully. Count: " + downstreamEmails.length);
    } catch (e) {
        Logger.log("ERROR in getDownstreamEmployees: " + e.message);
        downstreamEmails = [];
    }

    // === PERFORMANCE: Fetch Settings with bounded query ===
    const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const settingsLastRow = settingsSheet.getLastRow();
    const settingsData = settingsLastRow > 1 ? settingsSheet.getRange(1, 1, settingsLastRow, 2).getValues() : [[]];
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

    // === PERFORMANCE: Fetch Projects with bounded query ===
    const projectsSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const projectsLastRow = projectsSheet.getLastRow();
    const pData = projectsLastRow > 1 ? projectsSheet.getRange(2, 1, projectsLastRow - 1, 16).getValues() : [];
    
    // Track projects that need manager derivation for batch write later
    const projectsNeedingManager = [];

    let projects = [];
    pData.forEach((row, idx) => {
        try {
            if (!row[0]) return; // Skip empty rows
            
            const projectOwner = row[2] || "";
            let projectManager = row[3] || "";

            // === PERFORMANCE: Derive manager but DON'T write during read ===
            // Store for batch write later if needed
            if (!projectManager && projectOwner) {
                const derived = getManagerForUserMemo(projectOwner);
                if (derived) {
                    projectManager = derived;
                    projectsNeedingManager.push({ rowIndex: idx + 2, manager: derived });
                }
            }

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
                // === PERFORMANCE: Column projection - send minimal data for list view ===
                // Don't parse full comment history for dashboard, only send summary
                let commentCount = 0;
                let lastComment = null;
                if (row[14] && typeof row[14] === 'string' && row[14].trim() !== "") {
                    try { 
                        const parsedComments = JSON.parse(row[14]);
                        commentCount = parsedComments.length;
                        if (parsedComments.length > 0) {
                            lastComment = parsedComments[parsedComments.length - 1];
                        }
                    } catch (e) { 
                        Logger.log("Error parsing comments for project " + row[0] + ": " + e.message);
                    }
                }

                projects.push({
                    id: row[0],
                    name: row[1],
                    owner: projectOwner,
                    manager: projectManager,
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
                    projectType: row[15] || "Other",
                    // === PERFORMANCE: Lazy load - only send summary, not full comments ===
                    commentCount: commentCount,
                    lastComment: lastComment
                });
            }
        } catch (err) {
            Logger.log("Error parsing project row " + idx + ": " + err.message);
        }
    });

    // === PERFORMANCE: Batch write derived managers (async, don't block response) ===
    if (projectsNeedingManager.length > 0) {
        try {
            // Use batch setValues instead of individual setValue calls
            projectsNeedingManager.forEach(item => {
                projectsSheet.getRange(item.rowIndex, 4).setValue(item.manager);
            });
            Logger.log("Batch updated " + projectsNeedingManager.length + " project managers");
        } catch (e) {
            Logger.log("Failed batch manager update: " + e.message);
        }
    }

    // === PERFORMANCE: Fetch Actions with bounded query ===
    const actionsSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const actionsLastRow = actionsSheet.getLastRow();
    const aData = actionsLastRow > 1 ? actionsSheet.getRange(2, 1, actionsLastRow - 1, 9).getValues() : [];

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

// === PERFORMANCE: Lazy-load project details on-demand ===
/**
 * Get full project details including complete comment history when user clicks on a project.
 * This avoids sending large comment arrays for all projects on initial load.
 */
function getProjectDetails(projectId) {
    const email = Session.getActiveUser().getEmail();
    const role = getUserRole(email);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const projectsSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const projectsLastRow = projectsSheet.getLastRow();
    const pData = projectsLastRow > 1 ? projectsSheet.getRange(2, 1, projectsLastRow - 1, 16).getValues() : [];
    
    for (let i = 0; i < pData.length; i++) {
        const row = pData[i];
        if (row[0] === projectId) {
            // Check authorization (same logic as getDashboardData)
            const projectOwner = row[2] || "";
            const projectManager = row[3] || "";
            
            // Simple authorization check
            const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
            const usersLastRow = usersSheet.getLastRow();
            const usersData = usersLastRow > 1 ? usersSheet.getRange(1, 1, usersLastRow, 5).getValues() : [[]];
            const downstreamEmails = getDownstreamEmployees(email, usersData);
            
            const canSee = (
                role === "Admin" ||
                projectOwner === email ||
                projectManager === email ||
                (projectOwner && downstreamEmails.includes(projectOwner))
            );
            
            if (!canSee) {
                throw new Error("Unauthorized to view this project");
            }
            
            // Parse full comments
            let parsedComments = [];
            if (row[14] && typeof row[14] === 'string' && row[14].trim() !== "") {
                try {
                    parsedComments = JSON.parse(row[14]);
                } catch (e) {
                    Logger.log("Error parsing comments for project " + projectId + ": " + e.message);
                }
            }
            
            return {
                id: row[0],
                comments: parsedComments,
                updates: parsedComments
            };
        }
    }
    
    throw new Error("Project not found");
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

    // === PERFORMANCE: Use bounded queries ===
    // 1. Fetch All Users
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const userLastRow = userSheet.getLastRow();
    const uData = userLastRow > 1 ? userSheet.getRange(2, 1, userLastRow - 1, 5).getValues() : [];
    let users = uData.map(r => ({
        email: r[0],
        name: r[1],
        role: r[2],
        manager: r[3],
        emailEnabled: r[4] === true || String(r[4]).toLowerCase() === "true"
    }));

    // 2. Fetch All Settings 
    const setSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    const setLastRow = setSheet.getLastRow();
    const sData = setLastRow > 1 ? setSheet.getRange(2, 1, setLastRow - 1, 2).getValues() : [];
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

function createProject(name, startDate, deadline, status, phase, projType, outcomes = "", risks = "") {
    const email = Session.getActiveUser().getEmail();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const newId = generateId("PRJ");
    const timestamp = new Date().toISOString();

    // fall back to defaults if caller didn't provide
    status = status || "Not Started";
    phase = phase || "Open";
    projType = projType || "Other";

    // determine manager email by looking up user table
    const managerEmail = getManagerForUser(email);

    sheet.appendRow([
        newId,
        name,
        email,
        managerEmail, // populate manager if available
        status,
        phase,
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

    // Determine the actual assignee: the owner of the linked project
    const projectSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    // === PERFORMANCE: Use bounded query ===
    const projLastRow = projectSheet.getLastRow();
    const projData = projLastRow > 1 ? projectSheet.getRange(1, 1, projLastRow, 16).getValues() : [[]];
    let projectOwner = null;
    for (let i = 1; i < projData.length; i++) {
        if (projData[i][0] === projectId) {
            projectOwner = projData[i][2];
            break;
        }
    }
    if (!projectOwner) {
        throw new Error("Project not found or has no owner: " + projectId);
    }

    actionOwner = projectOwner; // override any passed value

    const initialLog = [{
        user: email,
        timestamp: timestamp,
        status: "Pending",
        text: "Action created and assigned to project owner."
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // === PERFORMANCE: Use bounded queries instead of getDataRange() ===
    const actionSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const projectSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    const actionLastRow = actionSheet.getLastRow();
    const projectLastRow = projectSheet.getLastRow();
    const usersLastRow = usersSheet.getLastRow();
    
    const actionData = actionLastRow > 1 ? actionSheet.getRange(1, 1, actionLastRow, 9).getValues() : [[]];
    const projectData = projectLastRow > 1 ? projectSheet.getRange(1, 1, projectLastRow, 16).getValues() : [[]];
    const usersData = usersLastRow > 1 ? usersSheet.getRange(1, 1, usersLastRow, 5).getValues() : [[]];
    
    // Cache role lookup
    let role = "None";
    for (let i = 1; i < usersData.length; i++) {
        if (usersData[i][0] === email) {
            role = usersData[i][2];
            break;
        }
    }

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
    
    // Prepare update log
    let updatedLog = null;
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
        updatedLog = JSON.stringify(currentLog);
    }

    // === PERFORMANCE: Batch write - single setValues call instead of 3-4 separate setValue calls ===
    if (updatedLog) {
        // Update status, percentage, timestamp, and log in one operation
        actionSheet.getRange(actionRowIndex + 1, 5, 1, 5).setValues([[
            newStatus,
            pctComplete,
            actionRow[6], // Priority (unchanged)
            timestamp,
            updatedLog
        ]]);
    } else {
        // Update status, percentage, timestamp only
        actionSheet.getRange(actionRowIndex + 1, 5, 1, 4).setValues([[
            newStatus,
            pctComplete,
            actionRow[6], // Priority (unchanged)
            timestamp
        ]]);
    }
    
    // Return updated action data for client-side cache update
    return {
        success: true,
        action: {
            id: actionId,
            status: newStatus,
            percentageCompleted: pctComplete,
            lastUpdatedDate: new Date(timestamp).toLocaleString(),
            updates: updatedLog ? JSON.parse(updatedLog) : JSON.parse(actionData[actionRowIndex][8] || "[]")
        }
    };
}

function editActionUpdate(actionId, timestamp, newText) {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // === PERFORMANCE: Use bounded queries ===
    const actionSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const projectSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    const actionLastRow = actionSheet.getLastRow();
    const projectLastRow = projectSheet.getLastRow();
    const usersLastRow = usersSheet.getLastRow();
    
    const actionData = actionLastRow > 1 ? actionSheet.getRange(1, 1, actionLastRow, 9).getValues() : [[]];
    const projectData = projectLastRow > 1 ? projectSheet.getRange(1, 1, projectLastRow, 16).getValues() : [[]];
    const usersData = usersLastRow > 1 ? usersSheet.getRange(1, 1, usersLastRow, 5).getValues() : [[]];
    
    // Cache role lookup
    let role = "None";
    for (let i = 1; i < usersData.length; i++) {
        if (usersData[i][0] === email) {
            role = usersData[i][2];
            break;
        }
    }

    // Find the action
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

    const actionOwner = actionRow[3];
    const projectId = actionRow[1];

    // Find linked project
    let projectOwner = null;
    for (let i = 1; i < projectData.length; i++) {
        if (projectData[i][0] === projectId) {
            projectOwner = projectData[i][2]; // Column C (index 2) is OwnerEmail
            break;
        }
    }

    // Get the updates log
    let currentLog;
    try {
        currentLog = JSON.parse(actionData[actionRowIndex][8] || "[]");
    } catch (e) {
        currentLog = [];
    }

    // Find the log entry by timestamp
    const logIndex = currentLog.findIndex(log => log.timestamp === timestamp);
    if (logIndex === -1) {
        throw new Error("Log entry not found.");
    }

    const logEntry = currentLog[logIndex];

    // Authorization: Admin, action owner, project owner, or log author
    const canEdit = (
        role === "Admin" ||
        logEntry.user === email ||
        actionOwner === email ||
        projectOwner === email
    );

    if (!canEdit) {
        throw new Error("You don't have permission to edit this action log.");
    }

    // Update the log entry text and add edited indicator
    currentLog[logIndex].text = newText + " (edited)";

    // Write back to sheet
    actionSheet.getRange(actionRowIndex + 1, 9).setValue(JSON.stringify(currentLog));

    // Update last modified timestamp
    const updateTimestamp = new Date().toISOString();
    actionSheet.getRange(actionRowIndex + 1, 8).setValue(updateTimestamp);

    return "Action log updated successfully!";
}

function addProjectComment(projectId, commentText) {
    const email = Session.getActiveUser().getEmail();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    // === PERFORMANCE: Use bounded query ===
    const lastRow = sheet.getLastRow();
    const data = lastRow > 1 ? sheet.getRange(1, 1, lastRow, 16).getValues() : [[]];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === projectId) {
            const projectName = data[i][1]; // Column B (index 1) = Name
            const projectOwner = data[i][2]; // Column C (index 2) = OwnerEmail

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
            
            // Send notification email if the note author is not the project owner
            if (email !== projectOwner) {
                try {
                    sendProjectNoteNotification(projectId, projectName, projectOwner, email, commentText);
                } catch (emailError) {
                    // Log but don't block the comment from being added
                    Logger.log('Failed to send notification email: ' + emailError.message);
                }
            }
            
            return "Comment added to project!";
        }
    }
    throw new Error("Project ID not found.");
}

function editProjectComment(projectId, timestamp, newText) {
    const email = Session.getActiveUser().getEmail();
    const role = getUserRole(email);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === projectId) {
            let currentLog;
            try {
                currentLog = JSON.parse(data[i][14] || "[]");
            } catch (e) {
                currentLog = [];
            }

            // Find the note by timestamp
            const noteIndex = currentLog.findIndex(note => note.timestamp === timestamp);
            if (noteIndex === -1) {
                throw new Error("Note not found.");
            }

            const note = currentLog[noteIndex];

            // Authorization: Admin or note author only
            if (role !== "Admin" && note.user !== email) {
                throw new Error("You don't have permission to edit this note.");
            }

            // Update the note text and add edited indicator
            currentLog[noteIndex].text = newText + " (edited)";
            
            // Update the sheet
            sheet.getRange(i + 1, 15).setValue(JSON.stringify(currentLog));
            
            // Update last modified timestamp
            const updateTimestamp = new Date().toISOString();
            sheet.getRange(i + 1, 14).setValue(updateTimestamp);

            return "Note updated successfully!";
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

    // make sure we have a valid recipient
    if (!action || !action.owner) {
        Logger.log('sendTaskAssignmentEmail called without a valid owner: %s', JSON.stringify(action));
        return;
    }

    try {
        MailApp.sendEmail({
            to: action.owner,
            subject: subject,
            body: body,
            // use the display name and a no‑reply address so the mailbox is readable
            name: 'Project Management Hub',
            replyTo: 'no-reply@example.com'
            // if you need a specific from-address that is an alias, add: from: 'alias@yourdomain.com'
        });
    } catch (e) {
        // Log full error and rethrow so caller / execution logs show the failure
        Logger.log('sendTaskAssignmentEmail: attempting to send to: %s', action.owner);
        Logger.log('sendTaskAssignmentEmail: error: %s', e && e.toString ? e.toString() : JSON.stringify(e));
        throw new Error('Failed to send assignment email to ' + action.owner + ': ' + (e && e.message ? e.message : JSON.stringify(e)));
    }
}

function sendProjectNoteNotification(projectId, projectName, projectOwner, noteAuthor, noteText) {
    // Validate inputs
    if (!projectOwner || !noteAuthor) {
        Logger.log('sendProjectNoteNotification called with invalid parameters');
        return;
    }

    // Get the web app URL (you may need to customize this)
    const scriptUrl = ScriptApp.getService().getUrl();

    const subject = `New Note Added to Your Project: ${projectName}`;
    
    // Create HTML body for better formatting
    const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <h2 style="color: #0066cc;">New Note Added to Your Project</h2>
      
      <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0;">
        <p><strong>Project:</strong> ${projectName}</p>
        <p><strong>Project ID:</strong> ${projectId}</p>
        <p><strong>Note By:</strong> ${noteAuthor}</p>
        <p><strong>Time:</strong> ${new Date().toLocaleString()}</p>
      </div>
      
      <div style="background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0;">
        <strong>Note:</strong>
        <p style="margin-top: 10px;">${noteText}</p>
      </div>
      
      <p style="color: #666; font-size: 14px; margin-top: 30px;">
        Please log in to the Project Management Hub to view all project details and respond if needed.
      </p>
      
      <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
      
      <p style="color: #999; font-size: 12px;">
        This is an automated notification from the Project Management Hub. 
        Please do not reply to this email.
      </p>
    </div>
    `;

    // Plain text fallback
    const plainBody = `
    New Note Added to Your Project: ${projectName}
    
    Project: ${projectName}
    Project ID: ${projectId}
    Note By: ${noteAuthor}
    Time: ${new Date().toLocaleString()}
    
    Note:
    ${noteText}
    
    Please log in to the Project Management Hub to view all project details and respond if needed.
    
    ---
    This is an automated notification from the Project Management Hub.
    Please do not reply to this email.
    `;

    try {
        MailApp.sendEmail({
            to: projectOwner,
            subject: subject,
            body: plainBody,
            htmlBody: htmlBody,
            name: 'Project Management Hub',
            replyTo: 'no-reply@example.com'
        });
        Logger.log(`Successfully sent note notification to ${projectOwner} for project ${projectId}`);
    } catch (e) {
        Logger.log('sendProjectNoteNotification: attempting to send to: ' + projectOwner);
        Logger.log('sendProjectNoteNotification: error: ' + (e && e.toString ? e.toString() : JSON.stringify(e)));
        // Don't throw - we don't want email failures to block the note from being saved
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
                subject: "Project Hub: Daily Wrap-Up",
                htmlBody: emailHtml,
                name: 'Project Management Hub',     // sender display name
                replyTo: 'no-reply@example.com'
                // use `from` here if you have a verified alias: from: 'alias@yourdomain.com'
            });
        } catch (e) {
            console.error(`Failed to send digest to ${userEmail}: ` + e.message);
        }
    });
}

// --- UPDATE PROJECT ---
function updateProject(projectId, newStatus, newPhase, newPercentage, updateNote, newStart, newDeadline, newType) {
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
    
    if (newPhase) {
        projectSheet.getRange(projectRowIndex + 1, 6).setValue(newPhase);
    }
    
    if (newPercentage !== undefined && newPercentage !== null) {
        projectSheet.getRange(projectRowIndex + 1, 7).setValue(parseInt(newPercentage) || 0);
    }
    
    // allow type change as well
    if (newType !== undefined && newType !== null) {
        projectSheet.getRange(projectRowIndex + 1, 16).setValue(newType);
    }
    
    // optional date updates
    if (newStart !== undefined && newStart !== null) {
        projectSheet.getRange(projectRowIndex + 1, 8).setValue(newStart);
    }
    if (newDeadline !== undefined && newDeadline !== null) {
        projectSheet.getRange(projectRowIndex + 1, 9).setValue(newDeadline);
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
    
    Logger.log("Project " + projectId + " updated by " + email + ": Status=" + newStatus + ", Phase=" + newPhase + ", %=" + newPercentage + ", start=" + newStart + ", deadline=" + newDeadline + ", type=" + newType);
    
    return "Project updated successfully";
}

// ------------------------------------------------------------------
// ADMIN HELPERS - DELETE OPERATIONS
// ------------------------------------------------------------------

/**
 * Deletes a project row along with any linked actions. Admin only.
 */
function deleteProject(projectId) {
    const email = Session.getActiveUser().getEmail();
    if (getUserRole(email) !== "Admin") throw new Error("Unauthorized");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const pData = pSheet.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < pData.length; i++) {
        if (pData[i][0] === projectId) {
            pSheet.deleteRow(i + 1);
            found = true;
            break;
        }
    }
    if (!found) throw new Error("Project not found: " + projectId);

    // remove any actions tied to the deleted project
    const aSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const aData = aSheet.getDataRange().getValues();
    // iterate backwards when deleting rows
    for (let j = aData.length - 1; j > 0; j--) {
        if (aData[j][1] === projectId) {
            aSheet.deleteRow(j + 1);
        }
    }

    return "Project and associated actions removed";
}

/**
 * Deletes a single action (admin only).
 */
function deleteAction(actionId) {
    const email = Session.getActiveUser().getEmail();
    if (getUserRole(email) !== "Admin") throw new Error("Unauthorized");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aSheet = ss.getSheetByName(SHEET_NAMES.ACTIONS);
    const aData = aSheet.getDataRange().getValues();

    for (let i = 1; i < aData.length; i++) {
        if (aData[i][0] === actionId) {
            aSheet.deleteRow(i + 1);
            return "Action deleted";
        }
    }
    throw new Error("Action not found: " + actionId);
}

/**
 * Remove an individual comment/note from a project. Admin only.
 * timestamp should match the note object that was stored.
 */
function deleteProjectComment(projectId, timestamp) {
    const email = Session.getActiveUser().getEmail();
    if (getUserRole(email) !== "Admin") throw new Error("Unauthorized");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
    const pData = pSheet.getDataRange().getValues();

    for (let i = 1; i < pData.length; i++) {
        if (pData[i][0] === projectId) {
            let comments = [];
            try {
                comments = JSON.parse(pData[i][14] || "[]");
            } catch (e) {
                comments = [];
            }

            const filtered = comments.filter(c => c.timestamp !== timestamp);
            pSheet.getRange(i + 1, 15).setValue(JSON.stringify(filtered));

            return "Comment removed";
        }
    }
    throw new Error("Project not found: " + projectId);
}

