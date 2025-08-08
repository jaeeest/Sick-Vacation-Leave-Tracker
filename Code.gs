
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('PIMI-MBI LEAVE TRACKER');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const SPREADSHEET_ID = '1mpl0q1P4NCiD3FdR-h0js8arSXDhC5XvM0ykisYJo0k'; //edit this and replace with your google sheets id
const EMPLOYEES_SHEET_NAME = 'EMPLOYEES';
const BALANCE_SHEET_NAME = 'BALANCE';
const AUDIT_LOG_SHEET_NAME = 'AUDIT LOGS';
const REQUEST_SHEET_NAME = 'REQUEST';
const FORCE_SHEET_NAME = 'FORCE LEAVE HISTORY';

const PARENT_FOLDER_ID = '1az5g6Ezpt1Pv21xsaXujZnh4rKMfHH_k'; //edit this and replace with your own google drive id
const LEAVE_FORM_TEMPLATE_ID = '1JobUmobQ5rrDMqWFnzXCPd-bTpYBqrq4ExqC0FNkkuc';


const DEPARTMENT_EMAIL_RECIPIENTS = {
  "MIS": "christalagmata@gmail.com",
  "HRAD": "angeloramos2324@gmail.com",
  "EDTN": "jnllglng15@gmail.com",
  "EDTB": "angeloramos2324@gmail.com",
  "CIRC": "cjderoxas16@gmail.com",
  "ACCF": "christalagmata@gmail.com",
  "GS": "cjderoxas16@gmail.com",
  "CNC": "sammiespam09@gmail.com",
  "ADT": "christalagmata@gmail.com",
};

const DEPARTMENT_ACRONYM_MAP = {
  "MANAGEMENT INFORMATION SYSTEM": "MIS",
  "HR & ADMINISTRATION": "HRAD",
  "EDITORIAL NEWS": "EDTN",
  "EDITORIAL BUSINESS": "EDTB",
  "CIRCULATION AND PRODUCTION": "CIRC",
  "ACCOUNTING & FINANCE": "ACCF",
  "GENERAL SERVICES": "GS",
  "CREDIT & COLLECTION": "CNC",
  "ADVERTISING & SALES MARKETING": "ADT",
};

/**
 * Converts a department name (potentially full name) to its standardized acronym.
 * @param {string} departmentName The department name from the sheet.
 * @returns {string} The standardized department acronym, or the original name if no mapping is found.
 */
function mapDepartmentToAcronym(departmentName) {
  if (!departmentName) {
    return '';
  }
  const cleanedName = departmentName.trim().toUpperCase(); // Clean and standardize for lookup
  return DEPARTMENT_ACRONYM_MAP[cleanedName] || departmentName; // Return acronym or original if not found
}

/**
 * Sends an email to a specific department head with the updated sick and vacation leave balances
 * for a single affected employee.
 * @param {string} targetDepartmentAcronym The standardized acronym of the department to send the email for.
 * @param {object} affectedEmployeeBalance The employee object with updated balance information.
 */
function sendLeaveBalanceEmailForAffectedEmployee(targetDepartmentAcronym, affectedEmployeeBalance) {
  try {
    const recipients = DEPARTMENT_EMAIL_RECIPIENTS[targetDepartmentAcronym];
    if (!recipients) {
      console.warn(`No email recipient configured for department: ${targetDepartmentAcronym}. Skipping email.`);
      return;
    }

    if (!affectedEmployeeBalance) {
      console.log(`No affected employee balance data provided for department: ${targetDepartmentAcronym}. Skipping email.`);
      return;
    }

    let emailBody = `Dear Department Head,\n\n`;
    emailBody += `The leave balance for ${affectedEmployeeBalance.name} in the ${targetDepartmentAcronym} department has been updated:\n\n`;
    emailBody += `<table style="width:100%; border-collapse: collapse;">`;
    emailBody += `<tr style="background-color:#E2EFFF;">
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Employee Name</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Sick Leave (SL)</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Vacation Leave (VL)</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">M/P</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Emergency</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Birthday</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Absent</th>
                  </tr>`;

    emailBody += `<tr>
                    <td style="border: 1px solid #ddd; padding: 8px;">${affectedEmployeeBalance.name}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">${affectedEmployeeBalance.sickLeave}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">${affectedEmployeeBalance.vacationLeave}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">${affectedEmployeeBalance.maternityPaternity}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">${affectedEmployeeBalance.emergencyLeave}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">${affectedEmployeeBalance.birthdayLeave}</td>
                    <td style="border: 1px solid #ddd; padding: 8px;">${affectedEmployeeBalance.absent}</td>
                  </tr>`;

    emailBody += `</table>\n\n`;
    emailBody += `This is an automated notification based on a recent change in the leave balance sheet.\n\n`;
    emailBody += `Best regards,\nSVL Tracker`;

    const subject = `Updated Leave Balances for ${affectedEmployeeBalance.name} (${targetDepartmentAcronym}) - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy")}`;

    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: emailBody
    });
    console.log(`Email sent to ${recipients} for ${affectedEmployeeBalance.name} in ${targetDepartmentAcronym} with updated leave balances.`);

  } catch (error) {
    console.error(`Error sending leave balance email for affected employee in ${targetDepartmentAcronym}:`, error.message);
  }
}


// SUMMARY CARDS - LEAVE HISTORY
function displayEmployeeCount(columnLetter) {
  const spreadsheet =  SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET_NAME);

  if (!sheet) {
    return 'Sheet ${EMPLOYEES_SHEET_NAME}not found!';
  }

  const lastDataRow = sheet.getLastRow();
    if (lastDataRow < 2) {
        return 0; // No data rows
    }

  const range = sheet.getRange(columnLetter + '2:' + columnLetter + lastDataRow);
  const values = range.getValues();

  let employeeCount = 0;

  for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      if (cellValue !== "" && cellValue !== null && typeof cellValue !== 'undefined') {
          employeeCount++;
      }
  }

  return employeeCount;
}

function displayPendingCount(columnLetter) {
  const spreadsheet =  SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(REQUEST_SHEET_NAME);

  if (!sheet) {
    return 'Sheet ${REQUEST_SHEET_NAME} not found!';
  }

  const lastDataRow = sheet.getLastRow();
    if (lastDataRow < 2) {
        return 0; // No data rows
    }

    const range = sheet.getRange(columnLetter + '2:' + columnLetter + lastDataRow);
    const values = range.getValues();

    let pendingCount = 0;
    for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][0];
        if (typeof cellValue === 'string' && cellValue.trim().toLowerCase() === 'pending') {
            pendingCount++;
        }
    }

    return pendingCount;
}

function displayRequestCount(columnLetter) {
  const spreadsheet =  SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(REQUEST_SHEET_NAME);

  if (!sheet) {
    return 'Sheet ${REQUEST_SHEET_NAME} not found!';
  }

  const lastDataRow = sheet.getLastRow();
    if (lastDataRow < 2) {
        return 0; // No data rows
    }

  const currentYear = new Date().getFullYear();

  const range = sheet.getRange(columnLetter + '2:' + columnLetter + lastDataRow);
  const values = range.getValues();

  let annualRequestCount = 0;

  for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];

      if (cellValue instanceof Date) {
          const fileYear = cellValue.getFullYear();
          if (fileYear === currentYear) {
              annualRequestCount++; 
          }
      }
  }

  return annualRequestCount;
}

/**
 * Fetches employee data from the 'EMPLOYEES' spreadsheet tab.
 * Assumes EMPLOYEES sheet columns: Employee Number, First Name, Middle Name, Last Name, Department, Hire Date, CIVIL STATUS, TENURE IN YEARS
 */
function getEmployeesData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET_NAME);

    if (!sheet) {
      console.error(`Sheet named '${EMPLOYEES_SHEET_NAME}' not found in spreadsheet ID: ${SPREADSHEET_ID}`);
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    console.log("Raw values from EMPLOYEES sheet:", values);

    if (values.length === 0) {
      console.log('No data found in the Employees sheet.');
      return [];
    }

    const headers = values.shift(); // Remove header row
    console.log("Headers from EMPLOYEES sheet:", headers);

    const employees = [];
    values.forEach(row => {
      if (row.length >= 12) { 
        const employeeNumber = row[0] || '';
        const firstName = row[1] || '';
        const middleName = row[2] || '';
        const lastName = row[3] || '';
        const department = row[4] || '';
        const hireDate = row[5] || ''; 
        const civilStatus = row[8] || ''; 
        const tenureInYears = parseFloat(row[10]) || 0; 
        const sex = row[11] || '';
        const fullName = `${firstName} ${middleName ? middleName + ' ' : ''}${lastName}`.trim().replace(/\s\s+/g, ' ');

        employees.push({
          fullName: fullName,
          employeeNumber: employeeNumber,
          department: department,
          hireDate: hireDate, 
          civilStatus: civilStatus, 
          tenureInYears: tenureInYears,
          sex: sex
        });
      } else {
        console.warn('Skipping row in Employees sheet due to insufficient columns or missing Hire Date/Tenure/Civil Status:', row);
      }
    });
    console.log("Parsed employee data:", employees);

    return employees;

  } catch (error) {
    console.error("Error fetching employee data:", error.message);
    return [];
  }
}

/**
 * Fetches a single employee's data by employee number.
 */
function getEmployeeByNumber(employeeNumber) {
  const employees = getEmployeesData();
  return employees.find(emp => emp.employeeNumber === employeeNumber);
}

/**
 * Fetches employee leave balance data from the 'BALANCE' spreadsheet tab.
 * Assumes BALANCE sheet columns: NAME, DEPARTMENT, SL, VL, MATERNITY/PATERNITY, EMERGENCY, BIRTHDAY, ABSENT
 */
function getEmployeeBalanceData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(BALANCE_SHEET_NAME);

    if (!sheet) {
      console.error(`Sheet named '${BALANCE_SHEET_NAME}' not found in spreadsheet ID: ${SPREADSHEET_ID}`);
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    console.log("Raw values from BALANCE sheet:", values);

    if (values.length === 0) {
      console.log('No data found in the Balance sheet.');
      return [];
    }

    // Assuming the first row contains headers, skip it.
    const headers = values.shift();
    console.log("Headers from BALANCE sheet:", headers);

    const employeeBalances = [];
    values.forEach(row => {
      if (row.length >= 8) {
        employeeBalances.push({
          name: row[0] || '',
          department: row[1] || '',
          sickLeave: parseFloat(row[2]) || 0,
          vacationLeave: parseFloat(row[3]) || 0,
          maternityPaternity: parseFloat(row[4]) || 0, 
          emergencyLeave: parseFloat(row[5]) || 0,     
          birthdayLeave: parseFloat(row[6]) || 0,      
          absent: parseFloat(row[7]) || 0              
        });
      } else {
        console.warn('Skipping row in Balance sheet due to insufficient columns:', row);
      }
    });
    console.log("Parsed employee balances:", employeeBalances);

    return employeeBalances;

  } catch (error) {
    console.error("Error fetching employee balance data:", error.message);
    return [];
  }
}

/**
 * Fetches a single employee's leave balance data by full name.
 */
function getEmployeeBalanceByNumber(fullName) {
  const balances = getEmployeeBalanceData();
  return balances.find(balance => balance.name.toLowerCase() === fullName.toLowerCase());
}

/**
 * Updates an employee's leave balance in the 'BALANCE' spreadsheet tab.
 */
function updateEmployeeBalance(fullName, newBalances) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(BALANCE_SHEET_NAME);

    if (!sheet) {
      console.error(`Sheet named '${BALANCE_SHEET_NAME}' not found.`);
      return { success: false, message: `Sheet '${BALANCE_SHEET_NAME}' not found.` };
    }

    const data = sheet.getDataRange().getValues();
    let rowFound = false;
    for (let i = 1; i < data.length; i++) { 
      if (data[i][0].toLowerCase() === fullName.toLowerCase()) { 
        sheet.getRange(i + 1, 3).setValue(newBalances.sickLeave);           
        sheet.getRange(i + 1, 4).setValue(newBalances.vacationLeave);       
        sheet.getRange(i + 1, 5).setValue(newBalances.maternityPaternity); 
        sheet.getRange(i + 1, 6).setValue(newBalances.emergencyLeave);     
        sheet.getRange(i + 1, 7).setValue(newBalances.birthdayLeave);      
        sheet.getRange(i + 1, 8).setValue(newBalances.absent);             
        rowFound = true;
        break;
      }
    }

    if (rowFound) {
      console.log(`Leave balance for ${fullName} updated.`);
      return { success: true, message: `Leave balance for ${fullName} updated.` };
    } else {
      console.warn(`Employee ${fullName} not found in BALANCE sheet for update.`);
      return { success: false, message: `Employee ${fullName} not found in balance sheet.` };
    }

  } catch (error) {
    console.error("Error updating employee balance:", error.message);
    return { success: false, message: `Failed to update leave balance: ${error.message}` };
  }
}

/**
 * Adds an audit log entry to the 'AUDIT LOG' spreadsheet tab.
 */
function addAuditLog(action, userEmail) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(AUDIT_LOG_SHEET_NAME);

    if (!sheet) {
      console.error(`Sheet named '${AUDIT_LOG_SHEET_NAME}' not found in spreadsheet ID: ${SPREADSHEET_ID}. Please create it.`);
      return false;
    }

    // Determine the next incremental LOG ID
    const lastRow = sheet.getLastRow();
    let nextLogId = 'LOG-00001';
    if (lastRow > 0) {
      const lastLogId = sheet.getRange(lastRow, 1).getValue();
      if (lastLogId && typeof lastLogId === 'string' && lastLogId.startsWith('LOG-')) {
        const lastNumber = parseInt(lastLogId.substring(4), 10);
        if (!isNaN(lastNumber)) {
          nextLogId = `LOG-${String(lastNumber + 1).padStart(5, '0')}`;
        }
      }
    }

    const now = new Date();
    const dateTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    sheet.appendRow([nextLogId, dateTime, action, userEmail]);
    console.log(`Audit Log added: ID: ${nextLogId}, DateTime: ${dateTime}, Action: ${action}, User: ${userEmail}`);
    return true;

  } catch (error) {
    console.error("Error adding audit log:", error.message);
    return false;
  }
}


function getAuditLogsData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(AUDIT_LOG_SHEET_NAME);

    if (!sheet) {
      console.error(`Sheet named '${AUDIT_LOG_SHEET_NAME}' not found in spreadsheet ID: ${SPREADSHEET_ID}.`);
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    console.log("Raw values from AUDIT LOG sheet:", values);

    if (values.length === 0) {
      console.log('No data found in the Audit Log sheet.');
      return [];
    }

    const headers = values.shift(); // Remove header row
    console.log("Headers from AUDIT LOG sheet:", headers);

    const auditLogs = [];
    values.forEach(row => {
      if (row.length >= 4) {
        auditLogs.push({
          logId: row[0] || '',
          dateTime: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : '',
          action: row[2] || '',
          user: row[3] || ''
        });
      } else {
        console.warn('Skipping row in Audit Log sheet due to insufficient columns:', row);
      }
    });
    console.log("Parsed audit logs:", auditLogs);

    return auditLogs;

  } catch (error) {
    console.error("Error fetching audit logs data:", error.message);
    return [];
  }
}

/**
 * Saves a leave request to the 'REQUEST' spreadsheet tab.
 * Balances are NOT deducted at this stage; they are deducted upon approval.
 */
function saveLeaveRequest(requestData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requestSheet = spreadsheet.getSheetByName(REQUEST_SHEET_NAME);

    if (!requestSheet) {
      console.error(`Sheet named '${REQUEST_SHEET_NAME}' not found in spreadsheet ID: ${SPREADSHEET_ID}. Please create it.`);
      return { success: false, message: `Sheet '${REQUEST_SHEET_NAME}' not found.` };
    }

    const employeeData = getEmployeeByNumber(requestData.employeeNumber);
    if (!employeeData) {
      return { success: false, message: "Employee not found. Please ensure employee number is correct." };
    }

    if (employeeData.tenureInYears !== undefined && employeeData.tenureInYears < 1) {
      return { success: false, message: "Employees with less than one year of tenure cannot submit a leave request." };
    }

    if (requestData.leaveType === 'Paternity/Maternity' && employeeData.civilStatus && employeeData.civilStatus.toLowerCase() === 'single') {
      return { success: false, message: "Single employees are not eligible for Paternity/Maternity leave." };
    }

    const lastRow = requestSheet.getLastRow();
    let nextLeaveId = 'SVL-00001';
    if (lastRow > 0) {
      const lastLeaveId = requestSheet.getRange(lastRow, 1).getValue();
      if (lastLeaveId && typeof lastLeaveId === 'string' && lastLeaveId.startsWith('SVL-')) {
        const lastNumber = parseInt(lastLeaveId.substring(4), 10);
        if (!isNaN(lastNumber)) {
          nextLeaveId = `SVL-${String(lastNumber + 1).padStart(5, '0')}`;
        }
      }
    }

    const now = new Date();
    const submissionDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy");

    const row = [
      nextLeaveId,
      submissionDate,
      requestData.fullName,
      requestData.employeeNumber,
      requestData.department,
      requestData.leaveType,
      requestData.leaveDuration,
      requestData.startDate,
      requestData.endDate,
      requestData.remarks,
      'Pending'
    ];

    requestSheet.appendRow(row);
    console.log("Leave request saved:", row);
    return { success: true, message: "Leave request submitted successfully!", leaveId: nextLeaveId };

  } catch (error) {
    console.error("Error saving leave request:", error.message);
    return { success: false, message: `Failed to save leave request: ${error.message}` };
  }
}

/**
 * Fetches leave history data from the 'REQUEST' spreadsheet tab.
 */
function getLeaveHistoryData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(REQUEST_SHEET_NAME);

    if (!sheet) {
      console.error(`Sheet named '${REQUEST_SHEET_NAME}' not found in spreadsheet ID: ${SPREADSHEET_ID}.`);
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    console.log("Apps Script: Raw values from REQUEST sheet:", values); // NEW LOG

    if (values.length === 0) {
      console.log('Apps Script: No data found in the Request sheet.'); // NEW LOG
      return [];
    }

    const headers = values.shift();
    console.log("Apps Script: Headers from REQUEST sheet:", headers); // NEW LOG

    const leaveHistory = [];
    values.forEach(row => {
      if (row.length >= 11) {
        leaveHistory.push({
          leaveID: row[0] || '',
           submissionDate: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "MM/dd/yyyy") : '',
          name: row[2] || '',
          employeeNumber: row[3] || '',
          department: row[4] || '',
          leaveType: row[5] || '',
          leaveDuration: row[6] || '',
          startDate: row[7] ? Utilities.formatDate(new Date(row[7]), Session.getScriptTimeZone(), "MM/dd/yyyy") : '',
          endDate: row[8] ? Utilities.formatDate(new Date(row[8]), Session.getScriptTimeZone(), "MM/dd/yyyy") : '',
          remarks: row[9] || '',
          status: row[10] || ''
        });
      } else {
        console.warn('Apps Script: Skipping row in Request sheet due to insufficient columns:', row); // NEW LOG
      }
    });
    console.log("Apps Script: Parsed leave history:", leaveHistory); // NEW LOG

    return leaveHistory;

  } catch (error) {
    console.error("Apps Script: Error fetching leave history data:", error.message); // NEW LOG
    return [];
  }
}

/**
 * Updates the status of a leave request in the 'REQUEST' spreadsheet tab.
 * Deducts leave balance if the new status is 'Approved'.
 * Reverts leave balance if the new status is 'Cancelled' and the previous status was 'Approved'.
 */
function updateLeaveRequestStatus(leaveID, newStatus) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(REQUEST_SHEET_NAME);

    if (!sheet) {
      console.error(`Sheet named '${REQUEST_SHEET_NAME}' not found.`);
      return { success: false, message: `Sheet '${REQUEST_SHEET_NAME}' not found.` };
    }

    const data = sheet.getDataRange().getValues();
    let rowFound = false;
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) { // Start from 1 to skip header
      if (data[i][0] === leaveID) {
        rowIndex = i;
        rowFound = true;
        break;
      }
    }

    if (!rowFound) {
      console.warn(`Leave ID ${leaveID} not found for status update.`);
      return { success: false, message: `Leave ID ${leaveID} not found.` };
    }

    // Get the existing request data
    const requestRow = data[rowIndex];
    const fullName = requestRow[2]; 
    const employeeNumber = requestRow[3]; 
    let leaveType = String(requestRow[5]).trim().toLowerCase(); 
    const leaveDuration = requestRow[6]; 
    const startDate = new Date(requestRow[7]); 
    const endDate = new Date(requestRow[8]); 
    const currentStatus = requestRow[10]; 
    const rawEmployeeDepartment = requestRow[4]; 

    const employeeDepartmentAcronym = mapDepartmentToAcronym(rawEmployeeDepartment);

    const employeeData = getEmployeeByNumber(employeeNumber);
    if (!employeeData) {
      return { success: false, message: "Employee data (including sex) not found for leave eligibility check." };
    }
    const sex = employeeData.sex;

    let daysRequested = 0;
    const diffTime = Math.abs(endDate.getTime() - startDate.getTime());
    daysRequested = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

    if (leaveDuration.includes('Half Day')) {
      daysRequested = 0.5;
    }

    if (newStatus === 'Approved' && currentStatus !== 'Approved') {
      // Get current leave balances
      const currentBalances = getEmployeeBalanceByNumber(fullName);
      if (!currentBalances) {
        return { success: false, message: "Employee leave balance not found for deduction. Please ensure employee name is correct in BALANCE sheet." };
      }

      let newSL = currentBalances.sickLeave;
      let newVL = currentBalances.vacationLeave;
      let newMaternityPaternity = currentBalances.maternityPaternity;
      let newEmergency = currentBalances.emergencyLeave;
      let newBirthday = currentBalances.birthdayLeave;
      let newAbsent = currentBalances.absent;

      let deductionApplied = false;

      switch (leaveType) {
        case 'sick leave': // Updated case
          if (newSL >= daysRequested) {
            newSL -= daysRequested;
            deductionApplied = true;
          } else if (newSL > 0 && newSL < daysRequested) {
            const remainingDays = daysRequested - newSL;
            newSL = 0;
            if (newVL >= remainingDays) {
              newVL -= remainingDays;
              deductionApplied = true;
            } else {
              newAbsent += remainingDays;
              newVL = 0;
              deductionApplied = true;
            }
          } else if (newVL >= daysRequested) {
            newVL -= daysRequested;
            deductionApplied = true;
          } else {
            newAbsent += daysRequested;
            deductionApplied = true;
          }
          break;
        case 'vacation leave': // Updated case
          if (newVL >= daysRequested) {
            newVL -= daysRequested;
            deductionApplied = true;
          } else {
            newAbsent += daysRequested;
            deductionApplied = true;
          }
          break;
        case 'paternity/maternity leave':
              const isFemale = sex.toLowerCase() === 'female';

              const allowedDays = isFemale ? 45 : 7;

              if (daysRequested > allowedDays) {
                  let leaveTypeSpecific = '';
                  if (isFemale) {
                      leaveTypeSpecific = 'Maternity leave';
                  } else {
                      leaveTypeSpecific = 'Paternity leave';
                  }
                  return { success: false, message: `${leaveTypeSpecific} cannot exceed ${allowedDays} days.` };
              }

              if (newMaternityPaternity >= 1) {
                  newMaternityPaternity = 0; // Resetting or decrementing the counter
                  deductionApplied = true; // Marking that a deduction was applied
              } else {
                  // If newMaternityPaternity is less than 1, it means the leave has already been used.
                  return { success: false, message: `You have already used your ${leaveType} leave.` };
              }
              break;
        case 'emergency leave': // Updated case
          if (daysRequested > 1) {
            return { success: false, message: "Emergency leave can only be requested for 1 day." };
          }
          if (newEmergency >= 1) {
            newEmergency = 0;
            deductionApplied = true;
          } else {
            return { success: false, message: "You have already used your Emergency leave." };
          }
          break;
        case 'birthday leave': // Updated case
          if (daysRequested > 1) {
            return { success: false, message: "Birthday leave can only be requested for 1 day."};
          }
          if (newBirthday >= 1) {
            newBirthday = 0;
            deductionApplied = true;
          } else {
            return { success: false, message: "You have already used your Birthday leave." };
          }
          break;
        default:
          return { success: false, message: "Invalid leave type." };
      }

      if (!deductionApplied) {
        return { success: false, message: "Failed to apply leave deduction. Please check employee balances or contact HR." };
      }

      const updateBalanceResult = updateEmployeeBalance(fullName, {
        sickLeave: newSL,
        vacationLeave: newVL,
        maternityPaternity: newMaternityPaternity,
        emergencyLeave: newEmergency,
        birthdayLeave: newBirthday,
        absent: newAbsent
      });

      if (!updateBalanceResult.success) {
        return { success: false, message: `Failed to update leave balance: ${updateBalanceResult.message}` };
      }

      if (employeeDepartmentAcronym) {
        // Fetch the CURRENT balance of the affected employee for the email
        const updatedAffectedEmployeeBalance = getEmployeeBalanceByNumber(fullName);
        if (updatedAffectedEmployeeBalance) {
          sendLeaveBalanceEmailForAffectedEmployee(employeeDepartmentAcronym, updatedAffectedEmployeeBalance);
        } else {
          console.warn(`Could not retrieve updated balance for ${fullName} for email notification.`);
        }
      }
    }

    else if (newStatus === 'Cancelled' && currentStatus === 'Approved') {
        const currentBalances = getEmployeeBalanceByNumber(fullName);
        if (!currentBalances) {
            return { success: false, message: "Employee leave balance not found for reversal. Please ensure employee name is correct in BALANCE sheet." };
        }

        let newSL = currentBalances.sickLeave;
        let newVL = currentBalances.vacationLeave;
        let newMaternityPaternity = currentBalances.maternityPaternity;
        let newEmergency = currentBalances.emergencyLeave;
        let newBirthday = currentBalances.birthdayLeave;
        let newAbsent = currentBalances.absent;

        let reversalApplied = false;
        let daysToRestore = daysRequested;

        switch (leaveType) {
            case 'sick leave': // Updated case
                if (daysToRestore > 0) {
                    newSL += daysToRestore;
                    daysToRestore = 0;
                }
                if (newAbsent > 0) {
                    const absentToRestore = Math.min(daysRequested, newAbsent);
                    newAbsent -= absentToRestore;
                }
                reversalApplied = true;
                break;
            case 'vacation leave': // Updated case
                if (daysToRestore > 0) {
                    newVL += daysToRestore;
                    daysToRestore = 0;
                }
                if (newAbsent > 0) {
                    const absentToRestore = Math.min(daysRequested, newAbsent);
                    newAbsent -= absentToRestore;
                }
                reversalApplied = true;
                break;
            case 'paternity/maternity leave': // Updated case
                newMaternityPaternity = 1; // Assuming it reverts to 1 (available)
                reversalApplied = true;
                break;
            case 'emergency leave': // Updated case
                newEmergency = 1; // Assuming it reverts to 1 (available)
                reversalApplied = true;
                break;
            case 'birthday leave': // Updated case
                newBirthday = 1; // Assuming it reverts to 1 (available)
                reversalApplied = true;
                break;
            default:
                console.warn(`Unknown leave type "${leaveType}" for reversal. No balance change applied.`);
                break;
        }

        if (reversalApplied) {
            const updateBalanceResult = updateEmployeeBalance(fullName, {
                sickLeave: newSL,
                vacationLeave: newVL,
                maternityPaternity: newMaternityPaternity,
                emergencyLeave: newEmergency,
                birthdayLeave: newBirthday,
                absent: newAbsent
            });

            if (!updateBalanceResult.success) {
                return { success: false, message: `Failed to reverse leave balance for ${leaveID}: ${updateBalanceResult.message}` };
            }
            console.log(`Leave ID ${leaveID} balance reversed for ${fullName}.`);
            // Successfully reversed balance, now send email to the specific department using the acronym
            if (employeeDepartmentAcronym) {
              // Fetch the CURRENT balance of the affected employee for the email
              const updatedAffectedEmployeeBalance = getEmployeeBalanceByNumber(fullName);
              if (updatedAffectedEmployeeBalance) {
                sendLeaveBalanceEmailForAffectedEmployee(employeeDepartmentAcronym, updatedAffectedEmployeeBalance);
              } else {
                console.warn(`Could not retrieve updated balance for ${fullName} for email notification.`);
              }
            }
        }
    }

    // Update the status in the REQUEST sheet
    sheet.getRange(rowIndex + 1, 11).setValue(newStatus); // Column K (index 10) for Status
    console.log(`Leave ID ${leaveID} status updated to ${newStatus}`);
    return { success: true, message: `Status for Leave ID ${leaveID} updated to ${newStatus}.` };

  } catch (error) {
    console.error("Error updating leave request status:", error.message);
    return { success: false, message: `Failed to update status: ${error.message}` };
  }
}

/**
 * Uploads a file to Google Drive, creating a monthly folder if it doesn't exist.
 */
function uploadFileToDrive(data, fileName, mimeType) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);

    const today = new Date();
    const monthNames = ["January", "February", "March", "April", "May", "June",
                        "July", "August", "September", "October", "November", "December"];
    const monthName = monthNames[today.getMonth()];
    const year = today.getFullYear();
    const folderName = `${monthName} ${year}`;

    let parentFolder;
    if (PARENT_FOLDER_ID) {
      parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    } else {
      // If no specific parent folder ID is provided, use the root folder
      parentFolder = DriveApp.getRootFolder();
      console.warn("No PARENT_FOLDER_ID specified. Uploading to root of Google Drive.");
    }

    // Check if the monthly folder exists, create it if not
    let monthlyFolder = null;
    const folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      monthlyFolder = folders.next();
    } else {
      monthlyFolder = parentFolder.createFolder(folderName);
      console.log(`Created new folder: ${folderName}`);
    }

    // Upload the file to the monthly folder
    const uploadedFile = monthlyFolder.createFile(blob);
    console.log(`File uploaded: ${uploadedFile.getName()} to folder: ${folderName}. URL: ${uploadedFile.getUrl()}`);

    return { success: true, message: `File "${fileName}" uploaded successfully to "${folderName}".`, fileUrl: uploadedFile.getUrl() };

  } catch (error) {
    console.error("Error uploading file to Google Drive:", error.message);
    return { success: false, message: `Failed to upload file: ${error.message}` };
  }
}

//NEW TEST SCRIPT FOR LINK UPLOAD
/**
 * Processes an uploaded file, uploads it to Google Drive, and links its URL
 * to the corresponding Leave ID in the REQUEST sheet.
 * @param {string} base64Data - The base64 encoded content of the file.
 * @param {string} fileName - The name of the file to upload.
 * @param {string} mimeType - The MIME type of the file.
 * @param {string} leaveIdToMatch - The Leave ID (file number) to match in the REQUEST sheet.
 * @returns {object} An object indicating success or failure.
 */
function processFileUploadAndLinkToSheet(base64Data, fileName, mimeType, leaveIdToMatch) {
  try {
    // Logging the start and input data for debugging
    console.log(`[processFileUploadAndLinkToSheet] Starting process for file: ${fileName}, MIME: ${mimeType}, with Leave ID: '${leaveIdToMatch}'`);

    // 1. Upload the file to Google Drive
    // Ensure uploadFileToDrive is modified to return { success: ..., message: ..., fileUrl: ... }
    const uploadResult = uploadFileToDrive(base64Data, fileName, mimeType);
    console.log(`[processFileUploadAndLinkToSheet] Upload Result:`, uploadResult);

    if (!uploadResult.success) {
      console.error(`[processFileUploadAndLinkToSheet] File upload failed: ${uploadResult.message}`);
      return { success: false, message: `File upload failed: ${uploadResult.message}` };
    }

    const fileUrl = uploadResult.fileUrl;
    console.log(`[processFileUploadAndLinkToSheet] File uploaded. URL: ${fileUrl}`);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requestSheet = spreadsheet.getSheetByName(REQUEST_SHEET_NAME);

    if (!requestSheet) {
      console.error(`[processFileUploadAndLinkToSheet] Sheet named '${REQUEST_SHEET_NAME}' not found.`);
      return { success: false, message: `Sheet '${REQUEST_SHEET_NAME}' not found.` };
    }
    console.log(`[processFileUploadAndLinkToSheet] Accessed sheet: ${REQUEST_SHEET_NAME}`);

    const dataRange = requestSheet.getDataRange();
    const values = dataRange.getValues();
    let rowFound = false;
    let rowIndex = -1;
    const leaveIdColumnIndex = 0; 
    const fileUrlColumnIndex = 12; 

    for (let i = 1; i < values.length; i++) {
      const leaveIdInSheet = String(values[i][leaveIdColumnIndex]).trim();

      if (leaveIdInSheet === leaveIdToMatch.trim()) {
        rowIndex = i;
        rowFound = true;
        console.log(`[processFileUploadAndLinkToSheet] Matching Leave ID found at row index (0-based): ${rowIndex}. Leave ID in sheet: '${leaveIdInSheet}'`);
        break;
      }
    }

    if (!rowFound) {
      console.warn(`[processFileUploadAndLinkToSheet] Leave ID "${leaveIdToMatch}" not found in ${REQUEST_SHEET_NAME} sheet. Row update skipped.`);
      return { success: false, message: `Leave ID "${leaveIdToMatch}" not found in the Request sheet.` };
    }

    requestSheet.getRange(rowIndex + 1, fileUrlColumnIndex).setValue(fileUrl);
    console.log(`[processFileUploadAndLinkToSheet] Successfully set URL: '${fileUrl}' for Leave ID "${leaveIdToMatch}" in row ${rowIndex + 1}, column ${fileUrlColumnIndex}.`);

    // IMPORTANT: Force the spreadsheet to apply pending changes immediately
    SpreadsheetApp.flush();
    console.log(`[processFileUploadAndLinkToSheet] Spreadsheet flushed. Changes applied.`);

    return { success: true, message: `File uploaded and linked successfully to Leave ID ${leaveIdToMatch}.` };

  } catch (error) {
    console.error(`[processFileUploadAndLinkToSheet] Error in processFileUploadAndLinkToSheet: ${error.message}`, error.stack);
    return { success: false, message: `Failed to process file upload and link to sheet: ${error.message}` };
}
}

/**
 * Generates a PDF of the leave request form, populating fields and creating two identical copies on separate pages.
 * @param {Object} requestData - The data from the leave request form.
 * @param {string} leaveId - The generated leave ID for the form.
 * @returns {Object} An object containing base64 encoded PDF data and filename, or an error message.
 */
function generateLeaveRequestPdf(requestData, leaveId) {
  try {
    const templateFile = DriveApp.getFileById(LEAVE_FORM_TEMPLATE_ID);

    const tempDocName = `Leave Request - ${requestData.fullName} - ${leaveId}_temp`;
    const tempDoc = DocumentApp.openById(templateFile.makeCopy(tempDocName).getId());
    const body = tempDoc.getBody();

    // Format dates to MM/DD/YYYY for consistency with the form's display
    const submissionDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
    const startDateFormatted = Utilities.formatDate(new Date(requestData.startDate), Session.getScriptTimeZone(), "MM/dd/yyyy");
    const endDateFormatted = Utilities.formatDate(new Date(requestData.endDate), Session.getScriptTimeZone(), "MM/dd/yyyy");

    // Fetch current leave balances for the employee
    const currentBalances = getEmployeeBalanceByNumber(requestData.fullName);
    if (!currentBalances) {
      console.warn(`Employee balance not found for ${requestData.fullName}. Cannot fill balance placeholders in PDF.`);
    }

    body.replaceText('{{to}}', requestData.to || 'HRD Manager');
    body.replaceText('{{thru}}', requestData.thru || 'Supervisor');
    body.replaceText('{{date}}', submissionDate);
    body.replaceText('{{department}}', requestData.department);
    body.replaceText('{{employeeNumber}}', requestData.employeeNumber);
    body.replaceText('{{fullName}}', requestData.fullName);
    body.replaceText('{{leaveType}}', requestData.leaveType);
    body.replaceText('{{leaveDuration}}', requestData.leaveDuration);
    body.replaceText('{{startDate}}', startDateFormatted);
    body.replaceText('{{endDate}}', endDateFormatted);
    body.replaceText('{{remarks}}', requestData.remarks || 'N/A');
    body.replaceText('{{leaveId}}', leaveId); // For the file number field
    body.replaceText('{{newSL}}', currentBalances ? currentBalances.sickLeave.toString() : 'N/A');
    body.replaceText('{{newVL}}', currentBalances ? currentBalances.vacationLeave.toString() : 'N/A');
    body.replaceText('{{totalDays}}', requestData.totalDays.toString()); // ADDED THIS LINE

    tempDoc.saveAndClose();

    const pdfBlob = tempDoc.getAs(MimeType.PDF);

    const fileName = `${requestData.fullName} - ${leaveId}.pdf`;

    const base64Data = Utilities.base64Encode(pdfBlob.getBytes());

    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

    return { success: true, base64Data: base64Data, fileName: fileName };

  } catch (error) {
    console.error("Error generating leave request PDF:", error.message);
    return { success: false, message: `Failed to generate PDF: ${error.message}` };
  }
}

//FORCE LEAVE
function saveForceLeaveData(scope, employees, startDateStr, endDateStr, durationType) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(FORCE_SHEET_NAME);
    const balanceSheet = spreadsheet.getSheetByName(BALANCE_SHEET_NAME);

    if (!sheet) {
      return { success: false, message: `Sheet '${FORCE_SHEET_NAME}' not found.` };
    }
    if (!balanceSheet) {
      return { success: false, message: `Sheet '${BALANCE_SHEET_NAME}' not found.` };
    }

    const now = new Date();
    const dateSubmitted = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy");

    // Calculate total days for deduction
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    let daysToDeduct = 0;
    if (startDateStr && endDateStr) {
      const diffTime = Math.abs(endDate.getTime() - startDate.getTime());
      daysToDeduct = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 to include both start and end days

      if (durationType.includes('Half Day')) {
        daysToDeduct = 0.5;
      }
    } else {
      console.error("Start date or end date is missing for force leave calculation.");
      return { success: false, message: "Start date and end date are required for force leave." };
    }

    const results = [];
    employees.forEach(employee => {
      try {
        // Append to FORCE LEAVE HISTORY sheet
        const lastRow = sheet.getLastRow();
        let nextForceLeaveId = 'FL-00001';
        if (lastRow > 0) {
          const lastId = sheet.getRange(lastRow, 1).getValue();
          if (lastId && typeof lastId === 'string' && lastId.startsWith('FL-')) {
            const lastNumber = parseInt(lastId.substring(3), 10);
            if (!isNaN(lastNumber)) {
              nextForceLeaveId = `FL-${String(lastNumber + 1).padStart(5, '0')}`;
            }
          }
        }
        
        sheet.appendRow([
          nextForceLeaveId,
          dateSubmitted,
          employee.department,
          employee.name,
          daysToDeduct, // Duration in days
          startDateStr,
          endDateStr
        ]);
        console.log(`Force leave recorded for ${employee.name} in FORCE LEAVE HISTORY sheet.`);

        // Deduct from employee balance
        const currentBalance = getEmployeeBalanceByNumber(employee.name);
        if (currentBalance) {
          let updatedBalance = { ...currentBalance };
          let remainingDeduction = daysToDeduct;

          // Deduct from Vacation Leave first
          if (updatedBalance.vacationLeave >= remainingDeduction) {
            updatedBalance.vacationLeave -= remainingDeduction;
            remainingDeduction = 0;
            console.log(`Deducted ${daysToDeduct} from VL for ${employee.name}. VL remaining: ${updatedBalance.vacationLeave}`);
          } else {
            remainingDeduction -= updatedBalance.vacationLeave;
            console.log(`VL exhausted for ${employee.name}. Remaining to deduct from SL: ${remainingDeduction}`);
            updatedBalance.vacationLeave = 0;
          }

          // If still remaining, deduct from Sick Leave
          if (remainingDeduction > 0) {
            updatedBalance.sickLeave -= remainingDeduction;
            console.log(`Deducted ${remainingDeduction} from SL for ${employee.name}. SL remaining: ${updatedBalance.sickLeave}`);
          }
          
          // Ensure balances don't go below zero
          updatedBalance.sickLeave = Math.max(0, updatedBalance.sickLeave);
          updatedBalance.vacationLeave = Math.max(0, updatedBalance.vacationLeave);

          const updateResult = updateEmployeeBalance(employee.name, updatedBalance);
          if (updateResult.success) {
            console.log(`Successfully deducted force leave for ${employee.name}. New balances: VL=${updatedBalance.vacationLeave}, SL=${updatedBalance.sickLeave}`);
            results.push({ success: true, employee: employee.name, message: `Force leave recorded and balance updated for ${employee.name}.` });
            // Send email notification for the updated balance
            const departmentAcronym = mapDepartmentToAcronym(employee.department);
            sendLeaveBalanceEmailForAffectedEmployee(departmentAcronym, updatedBalance);
          } else {
            console.error(`Failed to update balance for ${employee.name}: ${updateResult.message}`);
            results.push({ success: false, employee: employee.name, message: `Force leave recorded, but failed to update balance for ${employee.name}: ${updateResult.message}` });
          }
        } else {
          console.warn(`Employee balance not found for ${employee.name}. Force leave recorded but balance not updated.`);
          results.push({ success: false, employee: employee.name, message: `Force leave recorded, but employee balance not found for ${employee.name}.` });
        }
      } catch (employeeError) {
        console.error(`Error processing force leave for ${employee.name}:`, employeeError.message);
        results.push({ success: false, employee: employee.name, message: `Error processing force leave for ${employee.name}: ${employeeError.message}` });
      }
    });

    // Check if any deduction was successful
    const overallSuccess = results.some(r => r.success);
    const message = overallSuccess ? "Force leave submitted and balances updated for some employees." : "Failed to submit force leave or update balances for any employee.";
    return { success: overallSuccess, message: message, detailedResults: results };

  } catch (error) {
    console.error("Error in saveForceLeaveData:", error.message);
    return { success: false, message: `An unexpected error occurred: ${error.message}` };
  }
}

/**
 * Fetches force leave history data from the 'FORCE LEAVE HISTORY' spreadsheet tab.
 */
function getForceLeaveHistoryData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(FORCE_SHEET_NAME); 

    if (!sheet) {
      console.error(`Sheet named 'FORCE LEAVE HISTORY' not found in spreadsheet ID: ${SPREADSHEET_ID}`);
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    console.log("Raw values from FORCE LEAVE HISTORY sheet:", values);

    if (values.length === 0) {
      console.log('No data found in the FORCE LEAVE HISTORY sheet.');
      return [];
    }

    // Assuming the first row contains headers, skip it.
    const headers = values.shift();
    console.log("Headers from FORCE LEAVE HISTORY sheet:", headers);

    const forceLeaveHistory = [];
    values.forEach(row => {
      // Sheet format: id, date submitted, department, name, duration, start date, end date
      if (row.length >= 7) { // Still checking for minimum 7 columns, as 'Leave Type' is not in sheet
        forceLeaveHistory.push({
          id: row[0] || '', // Column A: id
          dateSubmitted: row[1] ? new Date(row[1]).toLocaleDateString('en-US') : '', // Column B: date submitted
          department: row[2] || '', // Column C: department
          employeeName: row[3] || '', // Column D: name
          duration: row[4] || '', // Column E: duration
          startDate: row[5] ? new Date(row[5]).toLocaleDateString('en-US') : '', // Column F: start date
          endDate: row[6] ? new Date(row[6]).toLocaleDateString('en-US') : '', // Column G: end date
          //leaveType: '' // 'Leave Type' is in the HTML table but not in your provided sheet format. Set to empty or add to sheet.
        });
      } else {
        console.warn('Skipping row in FORCE LEAVE HISTORY sheet due to insufficient columns:', row);
      }
    });
    console.log("Parsed force leave history:", forceLeaveHistory);

    return forceLeaveHistory;

  } catch (error) {
    console.error("Error fetching force leave history data:", error.message);
    return [];
  }
}
