/**
 * @file Google Apps Script for TCP FREEZER Web Application Backend
 * @overview This script provides the backend API for the TCP FREEZER web application,
 * handling data storage and retrieval from Google Sheets.
 */

// Replace with your Google Sheet ID
const SPREADSHEET_ID = '1w8mYaPwZiYomYNYRcI3emAlfbew99XscqNHgCTi4H9w';

// Sheet Names
const USERS_SHEET_NAME = 'Users';
const FREEZERS_SHEET_NAME = 'Freezers';
const FREEZE_HISTORY_SHEET_NAME = 'FreezeHistory';
const LOGIN_HISTORY_SHEET_NAME = 'LoginHistory';

/**
 * Handles GET requests to the web app.
 * Primarily used for JSONP callbacks or serving initial HTML if deployed as a standalone web app.
 * @param {GoogleAppsScript.Events.DoGet} e The event object from the GET request.
 * @returns {GoogleAppsScript.Content.TextOutput} JSONP response or HTML content.
 */
function doGet(e) {
    // Call the centralized command handler
    return handleCommands(e);
}

/**
 * Handles POST requests to the web app.
 * Primarily used for logging login attempts and recording freeze events.
 * Note: For JSONP, POST requests are not direct. The client-side `sendRequestToGS`
 * is currently using GET for all interactions. If a true POST were needed,
 * a different client-side approach (e.g., XMLHttpRequest or Fetch API without JSONP)
 * would be required, and `doPost` would parse the request body.
 * Given the JSONP requirement, all data transfer is currently through `doGet` parameters.
 * This `doPost` function is kept as a placeholder if the client-side changes.
 * For this specific project, `doGet` will handle all operations, as per JSONP convention.
 * @param {GoogleAppsScript.Events.DoPost} e The event object from the POST request.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON response.
 */
function doPost(e) {
    // This function is not used by the current JSONP client-side implementation.
    // All operations are handled via doGet parameters as JSONP doesn't directly support POST.
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'doPost is not used with current JSONP client setup.' }))
        .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Logs a login attempt to the LoginHistory sheet.
 * This is called via doGet with 'command=logLogin'.
 * @param {object} params - Object containing login details: email, staffCode, timestamp.
 * @returns {object} Success status.
 */
function logLogin(params) {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGIN_HISTORY_SHEET_NAME);
    if (!sheet) {
        return { success: false, message: `Sheet '${LOGIN_HISTORY_SHEET_NAME}' not found.` };
    }
    try {
        sheet.appendRow([params.timestamp, params.email, params.staffCode]);
        return { success: true };
    } catch (error) {
        return { success: false, message: `Failed to log login: ${error.message}` };
    }
}

/**
 * Authenticates a user against the Users sheet.
 * This is handled implicitly by `getUsers` on the client side.
 * @returns {object} Data from the Users sheet.
 */
function getUsers() {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET_NAME);
    if (!sheet) {
        return { success: false, message: `Sheet '${USERS_SHEET_NAME}' not found.` };
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Get headers and remove from data
    const users = data.map(row => {
        let obj = {};
        headers.forEach((header, i) => {
            obj[header] = row[i];
        });
        return obj;
    });
    return { success: true, data: users };
}

/**
 * Retrieves all freezer data from the Freezers sheet.
 * @returns {object} Data from the Freezers sheet.
 */
function getFreezers() {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FREEZERS_SHEET_NAME);
    if (!sheet) {
        return { success: false, message: `Sheet '${FREEZERS_SHEET_NAME}' not found.` };
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const freezers = data.map(row => {
        let obj = {};
        headers.forEach((header, i) => {
            obj[header] = row[i];
        });
        // Ensure stdTime is treated as a string for "HH:MM" display
        obj['stdTime'] = obj['STD.time (hours)'] || '';
        obj['id'] = obj['ID'];
        obj['name'] = obj['Name'];
        obj['details'] = obj['Details'];
        obj['status'] = obj['Status'];
        obj['estimatedCompletion'] = obj['EstimatedCompletion'];
        obj['currentMeter'] = obj['CurrentMeter'];
        obj['lastEmployee'] = obj['LastEmployee'];
        obj['lastEventTimestamp'] = obj['LastEventTimestamp'];
        obj['currentStartTimestamp'] = obj['CurrentStartTimestamp'];
        obj['refreezeCount'] = obj['RefreezeCount'];
        obj['totalFreezeTime'] = obj['TotalFreezeTime']; // Aggregate total
        obj['notes'] = obj['Notes'];
        return obj;
    });
    return { success: true, data: freezers };
}

/**
 * Retrieves all freeze history records from the FreezeHistory sheet.
 * @returns {object} Data from the FreezeHistory sheet.
 */
function getFreezeHistory() {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FREEZE_HISTORY_SHEET_NAME);
    if (!sheet) {
        return { success: false, message: `Sheet '${FREEZE_HISTORY_SHEET_NAME}' not found.` };
    }
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return { success: true, data: [] }; // Handle empty sheet
    const headers = data.shift();
    const history = data.map(row => {
        let obj = {};
        headers.forEach((header, i) => {
            obj[header] = row[i];
        });
        return obj;
    });
    return { success: true, data: history };
}

/**
 * Records a freeze event (Start, Stop, Clear, Refreeze) and updates freezer status.
 * This is called via doGet with 'command=recordFreezeLog'.
 * @param {object} params - Event details: freezerId, event, timestamp, meterStart, meterEnd, employee, estimatedCompletion, notes, totalFreezeTime.
 * @returns {object} Success status.
 */
function recordFreezeLog(params) {
    const freezerSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FREEZERS_SHEET_NAME);
    const historySheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FREEZE_HISTORY_SHEET_NAME);

    if (!freezerSheet) {
        return { success: false, message: `Sheet '${FREEZERS_SHEET_NAME}' not found.` };
    }
    if (!historySheet) {
        return { success: false, message: `Sheet '${FREEZE_HISTORY_SHEET_NAME}' not found.` };
    }

    try {
        const freezerData = freezerSheet.getDataRange().getValues();
        const freezerHeaders = freezerData[0];
        let freezerRowIndex = -1;
        let currentFreezer = null;

        // Find the freezer row
        for (let i = 1; i < freezerData.length; i++) {
            if (freezerData[i][freezerHeaders.indexOf('ID')] === params.freezerId) {
                freezerRowIndex = i;
                currentFreezer = {};
                freezerHeaders.forEach((header, idx) => {
                    currentFreezer[header] = freezerData[i][idx];
                });
                break;
            }
        }

        if (freezerRowIndex === -1) {
            return { success: false, message: `Freezer with ID '${params.freezerId}' not found.` };
        }

        // Record to history sheet
        const historyHeaders = historySheet.getDataRange().getValues()[0];
        const newHistoryRow = [];
        historyHeaders.forEach(header => {
            switch (header) {
                case 'Timestamp':
                    newHistoryRow.push(params.timestamp);
                    break;
                case 'FreezerID':
                    newHistoryRow.push(params.freezerId);
                    break;
                case 'Event':
                    newHistoryRow.push(params.event);
                    break;
                case 'MeterStart':
                    newHistoryRow.push(params.meterStart);
                    break;
                case 'MeterEnd':
                    newHistoryRow.push(params.meterEnd);
                    break;
                case 'Employee':
                    newHistoryRow.push(params.employee);
                    break;
                case 'EstimatedCompletion':
                    newHistoryRow.push(params.estimatedCompletion);
                    break;
                case 'Notes':
                    newHistoryRow.push(params.notes || '');
                    break;
                case 'TotalFreezeTime':
                    newHistoryRow.push(params.totalFreezeTime || ''); // GS will calculate for Refreeze
                    break;
                case 'RefreezeCount':
                    newHistoryRow.push(params.refreezeCount || ''); // GS will calculate for Refreeze
                    break;
                default:
                    newHistoryRow.push('');
                    break;
            }
        });
        historySheet.appendRow(newHistoryRow);

        // Update freezer status in Freezers sheet
        let newStatus = currentFreezer['Status'];
        let newEstimatedCompletion = currentFreezer['EstimatedCompletion'];
        let newCurrentMeter = currentFreezer['CurrentMeter'];
        let newCurrentStartTimestamp = currentFreezer['CurrentStartTimestamp'];
        let newRefreezeCount = currentFreezer['RefreezeCount'] || 0;
        let newTotalFreezeTime = currentFreezer['TotalFreezeTime'] || 0;

        switch (params.event) {
            case 'Start':
                newStatus = 'กำลัง Freeze';
                newEstimatedCompletion = params.estimatedCompletion;
                newCurrentMeter = params.meterStart;
                newCurrentStartTimestamp = params.timestamp;
                newTotalFreezeTime = 0; // Reset for a new cycle
                newRefreezeCount = 0; // Reset refreeze count
                break;
            case 'Stop':
                newStatus = 'มีสินค้าในตู้';
                newCurrentMeter = params.meterEnd;
                // Calculate total freeze time for this segment and add to cumulative
                if (currentFreezer['CurrentStartTimestamp'] && params.timestamp) {
                    const startTime = new Date(currentFreezer['CurrentStartTimestamp']).getTime();
                    const stopTime = new Date(params.timestamp).getTime();
                    const segmentDurationMs = stopTime - startTime;
                    newTotalFreezeTime = (parseFloat(currentFreezer['TotalFreezeTime'] || 0) + (segmentDurationMs / (1000 * 60 * 60))).toFixed(2);
                }
                newCurrentStartTimestamp = ''; // Clear start timestamp after stop
                break;
            case 'Clear':
                newStatus = 'พร้อมใช้งาน';
                newEstimatedCompletion = '';
                newCurrentMeter = '';
                newCurrentStartTimestamp = '';
                newRefreezeCount = 0;
                newTotalFreezeTime = 0; // Completely clear
                break;
            case 'Refreeze':
                newStatus = 'กำลัง Freeze';
                newEstimatedCompletion = ''; // No estimated completion on re-freeze until started again
                newCurrentMeter = params.meterStart; // Meter from previous stop becomes start
                newCurrentStartTimestamp = params.timestamp; // New start timestamp for refreeze
                newRefreezeCount = parseInt(newRefreezeCount) + 1; // Increment refreeze count
                // TotalFreezeTime will continue to accumulate from the previous cycle
                break;
        }

        // Update the freezer row in the Freezers sheet
        const statusCol = freezerHeaders.indexOf('Status');
        const estimatedCompletionCol = freezerHeaders.indexOf('EstimatedCompletion');
        const currentMeterCol = freezerHeaders.indexOf('CurrentMeter');
        const lastEmployeeCol = freezerHeaders.indexOf('LastEmployee');
        const lastEventTimestampCol = freezerHeaders.indexOf('LastEventTimestamp');
        const currentStartTimestampCol = freezerHeaders.indexOf('CurrentStartTimestamp');
        const refreezeCountCol = freezerHeaders.indexOf('RefreezeCount');
        const totalFreezeTimeCol = freezerHeaders.indexOf('TotalFreezeTime');
        const notesCol = freezerHeaders.indexOf('Notes'); // For general notes if any

        freezerSheet.getRange(freezerRowIndex + 1, statusCol + 1).setValue(newStatus);
        freezerSheet.getRange(freezerRowIndex + 1, estimatedCompletionCol + 1).setValue(newEstimatedCompletion);
        freezerSheet.getRange(freezerRowIndex + 1, currentMeterCol + 1).setValue(newCurrentMeter);
        freezerSheet.getRange(freezerRowIndex + 1, lastEmployeeCol + 1).setValue(params.employee);
        freezerSheet.getRange(freezerRowIndex + 1, lastEventTimestampCol + 1).setValue(params.timestamp);
        freezerSheet.getRange(freezerRowIndex + 1, currentStartTimestampCol + 1).setValue(newCurrentStartTimestamp);
        freezerSheet.getRange(freezerRowIndex + 1, refreezeCountCol + 1).setValue(newRefreezeCount);
        freezerSheet.getRange(freezerRowIndex + 1, totalFreezeTimeCol + 1).setValue(newTotalFreezeTime);
        freezerSheet.getRange(freezerRowIndex + 1, notesCol + 1).setValue(params.notes || '');


        return { success: true };
    } catch (error) {
        return { success: false, message: `Failed to record freeze log or update freezer status: ${error.message}` };
    }
}

/**
 * Retrieves data for the overview charts based on month and year.
 * @param {number} month - The month (1-12).
 * @param {number} year - The year.
 * @returns {object} Chart data including avgFreezeTimes, usageCounts, efficiency.
 */
function getChartData(month, year) {
    const historySheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FREEZE_HISTORY_SHEET_NAME);
    const freezerSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FREEZERS_SHEET_NAME);

    if (!historySheet || !freezerSheet) {
        return { success: false, message: `Required sheets not found for chart data.` };
    }

    const historyData = historySheet.getDataRange().getValues();
    const freezerData = freezerSheet.getDataRange().getValues();

    if (historyData.length < 2 || freezerData.length < 2) {
        return { success: true, data: { avgFreezeTimes: {}, usageCounts: {}, efficiency: {} } }; // Return empty if sheets are empty
    }

    const historyHeaders = historyData[0];
    const historyRecords = historyData.slice(1).map(row => {
        let obj = {};
        historyHeaders.forEach((header, i) => {
            obj[header] = row[i];
        });
        return obj;
    });

    const freezerHeaders = freezerData[0];
    const freezers = freezerData.slice(1).map(row => {
        let obj = {};
        freezerHeaders.forEach((header, i) => {
            obj[header] = row[i];
        });
        obj['stdHours'] = parseDurationToHours(obj['STD.time (hours)'] || '00:00');
        return obj;
    });

    const filteredHistory = historyRecords.filter(record => {
        if (record.Timestamp instanceof Date) {
            return record.Timestamp.getMonth() + 1 === month && record.Timestamp.getFullYear() === year;
        }
        // Attempt to parse string timestamps if they are not Date objects
        try {
            const recordDate = new Date(record.Timestamp);
            return recordDate.getMonth() + 1 === month && recordDate.getFullYear() === year;
        } catch (e) {
            return false;
        }
    });


    const freezerAvgFreezeTimes = {};
    const freezerUsageCounts = {};
    const freezerTotalFreezeTimes = {}; // To sum up times for each freezer
    const freezerSTDHours = {};

    freezers.forEach(f => {
        freezerAvgFreezeTimes[f['Name']] = 0;
        freezerUsageCounts[f['Name']] = 0;
        freezerTotalFreezeTimes[f['Name']] = 0;
        freezerSTDHours[f['Name']] = f.stdHours;
    });

    // Calculate total freeze times and usage counts
    const currentFreezerSessions = {}; // Tracks ongoing freeze sessions

    filteredHistory.forEach(record => {
        const freezerId = record.FreezerID;
        const freezerName = freezers.find(f => f.ID === freezerId)?.Name || freezerId; // Fallback to ID if name not found

        if (record.Event === 'Start') {
            currentFreezerSessions[freezerId] = new Date(record.Timestamp).getTime();
        } else if (record.Event === 'Stop' && currentFreezerSessions[freezerId]) {
            const startTime = currentFreezerSessions[freezerId];
            const stopTime = new Date(record.Timestamp).getTime();
            const durationHours = (stopTime - startTime) / (1000 * 60 * 60);

            if (!isNaN(durationHours) && durationHours > 0) {
                freezerTotalFreezeTimes[freezerName] = (freezerTotalFreezeTimes[freezerName] || 0) + durationHours;
                freezerUsageCounts[freezerName] = (freezerUsageCounts[freezerName] || 0) + 1;
            }
            delete currentFreezerSessions[freezerId]; // End the session
        } else if (record.Event === 'Refreeze') {
            // If refreeze, it also counts as a "start" of a new segment towards total time
            currentFreezerSessions[freezerId] = new Date(record.Timestamp).getTime();
            // For refreeze, we might not count as a new "usage" unless it was specifically a "stop" then "refreeze"
            // Let's count it as a usage for simplicity if it starts a new cycle segment
            // This might need more precise definition based on how "usage" is defined.
            // For now, usage count increases only on "Stop" events from "Start" events.
            // The totalFreezeTime in the sheet should contain the accumulated time for refreeze cycles
            if (record.TotalFreezeTime && !isNaN(record.TotalFreezeTime)) {
                freezerTotalFreezeTimes[freezerName] = parseFloat(record.TotalFreezeTime);
            }
            // Increment usage count for refreeze only if it represents a distinct operation after a stop
            // This logic is complex and would need clarification. Sticking to simple "Stop" count for now.
        } else if (record.Event === 'Clear') {
            delete currentFreezerSessions[freezerId]; // Clear ends any session
        }
    });


    // Calculate average freeze times and efficiency
    const avgFreezeTimes = {};
    const efficiency = {};

    for (const name in freezerTotalFreezeTimes) {
        if (freezerUsageCounts[name] > 0) {
            avgFreezeTimes[name] = freezerTotalFreezeTimes[name] / freezerUsageCounts[name];
        } else {
            avgFreezeTimes[name] = 0; // No usage for this period
        }

        const std = freezerSTDHours[name];
        if (std > 0) {
            efficiency[name] = std - avgFreezeTimes[name]; // Positive means faster, negative means slower
        } else {
            efficiency[name] = 0; // No standard time to compare
        }
    }

    return {
        success: true,
        data: {
            avgFreezeTimes: avgFreezeTimes,
            usageCounts: freezerUsageCounts,
            efficiency: efficiency
        }
    };
}

/**
 * Parses duration string "HH:MM" into hours.
 * Helper function for GS.
 * @param {string} durationStr - Duration string (e.g., "48:00").
 * @returns {number} Total hours.
 */
function parseDurationToHours(durationStr) {
    if (!durationStr || typeof durationStr !== 'string') return 0;
    const parts = durationStr.split(':');
    if (parts.length === 2) {
        const hours = parseInt(parts[0], 10);
        const minutes = parseInt(parts[1], 10);
        if (!isNaN(hours) && !isNaN(minutes)) {
            return hours + (minutes / 60);
        }
    }
    return 0;
}

/**
 * Global function to handle all incoming commands from the client-side
 * @param {GoogleAppsScript.Events.DoGet} e The event object from the GET request.
 * @returns {GoogleAppsScript.Content.TextOutput} JSONP response.
 */
function handleCommands(e) {
    // Check if e.parameter is defined. This is crucial when testing the script directly
    // from the Apps Script editor, as 'e' might be undefined or lack 'parameter'.
    if (!e || !e.parameter) {
        // Return a response that indicates a missing parameter, useful for debugging.
        return ContentService.createTextOutput(
            JSON.stringify({ success: false, message: 'Invalid request: missing event parameters. (Likely direct execution without parameters)' })
        ).setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    const callbackName = e.parameter.callback;
    const command = e.parameter.command;
    const params = e.parameter; // All parameters are available here

    let result = { success: false, message: 'Unknown command' };

    try {
        switch (command) {
            case 'getUsers':
                result = getUsers();
                break;
            case 'getFreezers':
                result = getFreezers();
                break;
            case 'getFreezeHistory':
                result = getFreezeHistory();
                break;
            case 'logLogin':
                result = logLogin(params);
                break;
            case 'recordFreezeLog':
                result = recordFreezeLog(params);
                break;
            case 'getChartData':
                const month = parseInt(params.month);
                const year = parseInt(params.year);
                result = getChartData(month, year);
                break;
            default:
                result = { success: false, message: 'Invalid command' };
                break;
        }
    } catch (error) {
        result = { success: false, message: `Server error: ${error.message}` };
    }

    // Return JSONP response
    return ContentService.createTextOutput(callbackName + '(' + JSON.stringify(result) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
}


// To deploy this script as a web app:
// 1. In Google Apps Script editor, click "Deploy" -> "New deployment".
// 2. Select "Web app" as the type.
// 3. Set "Execute as:" to "Me" (your Google account).
// 4. Set "Who has access:" to "Anyone" (or "Anyone, even anonymous" if authentication is handled in-app).
//    For this setup, "Anyone" is fine as we are using JSONP and authentication is custom.
// 5. Copy the "Web app URL". This is your GOOGLE_APP_SCRIPT_URL.
// 6. Replace the placeholder SPREADSHEET_ID at the top of this script with your actual Google Sheet ID.
