// ============================================================================
// SIGMA HOMES - ATTENDANCE TRACKER
// Google Sheets Apps Script for Employee Attendance Management
// ============================================================================

// Configuration Constants
const CONFIG = {
  LOCATIONS: {
    'Jhotwara': {
      lat: 26.9399629,
      lng: 75.7327858,
      radius: 100
    },
    'Mansarovar Office': {
      lat: 26.883757092955683,
      lng: 75.73839344818678,
      radius: 100
    }
  },
  SHEET_NAMES: {
    EMPLOYEES: 'Employees',
    ATTENDANCE: 'Attendance',
    DASHBOARD: 'Dashboard',
    SETTINGS: 'Settings'
  },
  COLORS: {
    PRIMARY: '#1F2937',
    SECONDARY: '#10B981',
    ACCENT: '#3B82F6',
    WARNING: '#F59E0B',
    ERROR: '#EF4444',
    SUCCESS: '#10B981',
    WHITE: '#FFFFFF'
  }
};

// ============================================================================
// WEB APP ENTRY POINT
// ============================================================================

/**
 * doGet - Main entry point for web app deployment
 * Deploy as web app to get a URL
 */
function doGet(e) {
  try {
    return HtmlService.createHtmlOutput(getAttendanceHTML())
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    return HtmlService.createHtmlOutput('Error: ' + error.message);
  }
}

// ============================================================================
// INITIALIZATION & SETUP
// ============================================================================

/**
 * Initialize the spreadsheet with all required sheets and setup
 */
function initializeSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create sheets if they don't exist
    createSheetsIfNotExist();
    
    // Setup Employees Sheet
    setupEmployeesSheet(ss);
    
    // Setup Attendance Sheet
    setupAttendanceSheet(ss);
    
    // Setup Dashboard
    setupDashboard(ss);
    
    // Setup Settings
    setupSettingsSheet(ss);
    
    SpreadsheetApp.getUi().alert('✅ Attendance Tracker initialized successfully!');
  } catch (error) {
    Logger.log('Error during initialization: ' + error);
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

/**
 * Create sheets if they don't exist
 */
function createSheetsIfNotExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ss.getSheetNames();
  
  Object.values(CONFIG.SHEET_NAMES).forEach(name => {
    if (!sheetNames.includes(name)) {
      ss.insertSheet(name);
    }
  });
}

/**
 * Setup Employees Sheet with headers and formatting
 */
function setupEmployeesSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
  sheet.clear();
  
  const headers = ['Employee ID', 'Name', 'Email', 'Phone', 'Designation', 'Allowed Location', 'Status'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  
  headerRange.setValues([headers]);
  formatHeaderRow(headerRange);
  
  // Set column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 100);
  
  // Add data validation for Allowed Location
  const locationValues = Object.keys(CONFIG.LOCATIONS);
  const dataValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(locationValues)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(dataValidation);
  
  // Add data validation for Status
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Leave'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('G2:G1000').setDataValidation(statusValidation);
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Setup Attendance Sheet with headers and formatting
 */
function setupAttendanceSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ATTENDANCE);
  sheet.clear();
  
  const headers = ['Date', 'Employee ID', 'Employee Name', 'Check-In Time', 'Check-In Location', 'Check-Out Time', 'Check-Out Location', 'Working Hours', 'Status'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  
  headerRange.setValues([headers]);
  formatHeaderRow(headerRange);
  
  // Set column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 100);
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Setup Dashboard with summary statistics
 */
function setupDashboard(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  sheet.clear();
  
  // Title
  const titleRange = sheet.getRange('A1:D1');
  titleRange.merge();
  titleRange.setValue('📊 ATTENDANCE DASHBOARD');
  titleRange.setFontSize(24).setFontWeight('bold').setForegroundColor(CONFIG.COLORS.PRIMARY);
  titleRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 35);
  
  // Subtitle
  const subtitleRange = sheet.getRange('A2:D2');
  subtitleRange.merge();
  subtitleRange.setValue('Today\'s Summary');
  subtitleRange.setFontSize(14).setFontWeight('bold').setForegroundColor(CONFIG.COLORS.SECONDARY);
  subtitleRange.setHorizontalAlignment('center');
  sheet.setRowHeight(2, 25);
  
  // Summary cards
  sheet.getRange('A4').setValue('Total Employees').setFontWeight('bold').setBackground(CONFIG.COLORS.ACCENT).setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('A5').setFormula('=COUNTA(Employees!A2:A)');
  
  sheet.getRange('B4').setValue('Present Today').setFontWeight('bold').setBackground(CONFIG.COLORS.SUCCESS).setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('B5').setFormula('=COUNTIFS(Attendance!A:A,TODAY(),Attendance!D:D,"<>")');
  
  sheet.getRange('C4').setValue('Absent Today').setFontWeight('bold').setBackground(CONFIG.COLORS.ERROR).setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('C5').setFormula('=COUNTA(Employees!A2:A)-COUNTIFS(Attendance!A:A,TODAY(),Attendance!D:D,"<>")');
  
  sheet.getRange('D4').setValue('Checked Out').setFontWeight('bold').setBackground(CONFIG.COLORS.WARNING).setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('D5').setFormula('=COUNTIFS(Attendance!A:A,TODAY(),Attendance!F:F,"<>")');
  
  // Format numbers
  sheet.getRange('A5:D5').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#F3F4F6');
  
  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  
  // Freeze header
  sheet.setFrozenRows(3);
}

/**
 * Setup Settings Sheet
 */
function setupSettingsSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  sheet.clear();
  
  sheet.getRange('A1').setValue('SYSTEM SETTINGS').setFontSize(16).setFontWeight('bold').setForegroundColor(CONFIG.COLORS.PRIMARY);
  
  sheet.getRange('A3').setValue('Jhotwara Location:').setFontWeight('bold');
  sheet.getRange('B3').setValue('26.9399629, 75.7327858');
  sheet.getRange('A4').setValue('Mansarovar Location:').setFontWeight('bold');
  sheet.getRange('B4').setValue('26.883757092955683, 75.73839344818678');
  sheet.getRange('A5').setValue('Check-In Radius:').setFontWeight('bold');
  sheet.getRange('B5').setValue('100 meters');
  
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
}

/**
 * Format header row with styling
 */
function formatHeaderRow(range) {
  range.setBackground(CONFIG.COLORS.PRIMARY);
  range.setFontColor(CONFIG.COLORS.WHITE);
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
}

// ============================================================================
// USER INTERFACE - HTML DIALOG
// ============================================================================

/**
 * Show the main attendance tracking interface
 */
function showAttendanceUI() {
  try {
    const html = HtmlService.createHtmlOutput(getAttendanceHTML())
      .setWidth(600)
      .setHeight(800);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Sigma Homes - Attendance Tracker');
  } catch (error) {
    Logger.log('Error showing UI: ' + error);
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

/**
 * Generate HTML for the attendance UI
 */
function getAttendanceHTML() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Sigma Homes - Attendance Tracker</title>
      <style>
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          padding: 20px;
        }
        
        .container {
          max-width: 500px;
          margin: 0 auto;
          background: white;
          border-radius: 15px;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
          overflow: hidden;
        }
        
        .header {
          background: linear-gradient(135deg, #1F2937 0%, #374151 100%);
          padding: 30px 20px;
          text-align: center;
          color: white;
        }
        
        .header h1 {
          font-size: 24px;
          font-weight: bold;
          margin-bottom: 5px;
          letter-spacing: 1px;
        }
        
        .header p {
          font-size: 14px;
          opacity: 0.9;
          margin-bottom: 10px;
        }
        
        .subtitle {
          font-size: 12px;
          opacity: 0.8;
          margin-top: 5px;
        }
        
        .content {
          padding: 30px 20px;
        }
        
        .form-group {
          margin-bottom: 20px;
        }
        
        label {
          display: block;
          margin-bottom: 8px;
          font-weight: 600;
          color: #1F2937;
          font-size: 13px;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        
        input, select {
          width: 100%;
          padding: 12px 15px;
          border: 2px solid #E5E7EB;
          border-radius: 8px;
          font-size: 14px;
          transition: all 0.3s ease;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        input:focus, select:focus {
          outline: none;
          border-color: #667eea;
          box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        
        .button-group {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 12px;
          margin-top: 30px;
        }
        
        button {
          padding: 14px 20px;
          border: none;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.3s ease;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        
        .btn-checkin {
          background: linear-gradient(135deg, #10B981 0%, #059669 100%);
          color: white;
        }
        
        .btn-checkin:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 8px 16px rgba(16, 185, 129, 0.3);
        }
        
        .btn-checkout {
          background: linear-gradient(135deg, #3B82F6 0%, #2563EB 100%);
          color: white;
        }
        
        .btn-checkout:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 8px 16px rgba(59, 130, 246, 0.3);
        }
        
        .alert {
          padding: 15px;
          border-radius: 8px;
          margin-bottom: 20px;
          display: none;
          font-size: 13px;
          animation: slideIn 0.3s ease;
        }
        
        @keyframes slideIn {
          from {
            opacity: 0;
            transform: translateY(-10px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        .alert.success {
          background-color: #D1FAE5;
          color: #065F46;
          border-left: 4px solid #10B981;
        }
        
        .alert.error {
          background-color: #FEE2E2;
          color: #991B1B;
          border-left: 4px solid #EF4444;
        }
        
        .alert.warning {
          background-color: #FEF3C7;
          color: #92400E;
          border-left: 4px solid #F59E0B;
        }
        
        .alert.info {
          background-color: #DBEAFE;
          color: #1E40AF;
          border-left: 4px solid #3B82F6;
        }
        
        .location-info {
          background: #F3F4F6;
          padding: 12px;
          border-radius: 8px;
          font-size: 12px;
          color: #4B5563;
          margin-top: 8px;
          line-height: 1.6;
        }
        
        .status-display {
          display: flex;
          gap: 10px;
          margin-top: 15px;
          flex-wrap: wrap;
        }
        
        .status-badge {
          flex: 1;
          min-width: 120px;
          padding: 12px;
          border-radius: 8px;
          text-align: center;
          font-size: 12px;
          font-weight: 600;
          white-space: pre-wrap;
        }
        
        .badge-checkin {
          background: #DBEAFE;
          color: #1E40AF;
        }
        
        .badge-checkout {
          background: #E0E7FF;
          color: #3730A3;
        }
        
        .spinner {
          display: inline-block;
          width: 20px;
          height: 20px;
          border: 3px solid #f3f3f3;
          border-top: 3px solid #667eea;
          border-radius: 50%;
          animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        button:disabled {
          opacity: 0.6;
          cursor: not-allowed;
        }
        
        .time-display {
          font-size: 16px;
          font-weight: bold;
          color: #667eea;
          text-align: center;
          margin: 15px 0;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>🏢 SIGMA HOMES</h1>
          <p>Attendance Tracker</p>
          <div class="subtitle">Check-In & Check-Out Management</div>
        </div>
        
        <div class="content">
          <div id="alert" class="alert"></div>
          
          <div class="form-group">
            <label for="employeeId">Employee ID *</label>
            <input type="text" id="employeeId" placeholder="Enter your Employee ID" autocomplete="off">
          </div>
          
          <div class="form-group">
            <label for="employeeName">Employee Name *</label>
            <input type="text" id="employeeName" placeholder="Your Name" disabled>
          </div>
          
          <div class="form-group">
            <label for="location">Select Location *</label>
            <select id="location">
              <option value="">-- Choose Location --</option>
              <option value="Jhotwara">📍 Jhotwara (26.9399629, 75.7327858)</option>
              <option value="Mansarovar Office">📍 Mansarovar Office (26.883757092955683, 75.73839344818678)</option>
            </select>
            <div class="location-info" id="locationInfo"></div>
          </div>
          
          <div class="time-display" id="currentTime">--:-- --</div>
          
          <div class="status-display" id="statusDisplay"></div>
          
          <div class="button-group">
            <button class="btn-checkin" onclick="handleCheckIn()">✓ CHECK IN</button>
            <button class="btn-checkout" onclick="handleCheckOut()">✓ CHECK OUT</button>
          </div>
        </div>
      </div>
      
      <script>
        // Update current time
        function updateTime() {
          const now = new Date();
          document.getElementById('currentTime').textContent = now.toLocaleTimeString('en-IN');
        }
        
        setInterval(updateTime, 1000);
        updateTime();
        
        // Validate employee and fetch details
        document.getElementById('employeeId').addEventListener('change', function() {
          const id = this.value.trim();
          if (id) {
            google.script.run.withSuccessHandler(handleEmployeeData).getEmployeeData(id);
          }
        });
        
        // Update location info
        document.getElementById('location').addEventListener('change', function() {
          const locationInfo = document.getElementById('locationInfo');
          if (this.value) {
            locationInfo.textContent = '📍 Geolocation required for check-in/check-out within 100m radius';
          }
        });
        
        function handleEmployeeData(data) {
          if (data.found) {
            document.getElementById('employeeName').value = data.name;
            document.getElementById('employeeName').style.background = '#E0F2FE';
            loadTodayStatus();
          } else {
            showAlert('Employee not found in the system', 'error');
            document.getElementById('employeeName').value = '';
          }
        }
        
        function loadTodayStatus() {
          google.script.run.withSuccessHandler(handleTodayStatus).getTodayStatus(document.getElementById('employeeId').value);
        }
        
        function handleTodayStatus(data) {
          const statusDisplay = document.getElementById('statusDisplay');
          statusDisplay.innerHTML = '';
          
          if (data.checkIn) {
            const badge = document.createElement('div');
            badge.className = 'status-badge badge-checkin';
            badge.textContent = '✓ Checked In\n' + data.checkInTime;
            statusDisplay.appendChild(badge);
          }
          
          if (data.checkOut) {
            const badge = document.createElement('div');
            badge.className = 'status-badge badge-checkout';
            badge.textContent = '✓ Checked Out\n' + data.checkOutTime;
            statusDisplay.appendChild(badge);
          }
          
          if (!data.checkIn && !data.checkOut) {
            const badge = document.createElement('div');
            badge.className = 'status-badge';
            badge.style.background = '#F3F4F6';
            badge.style.color = '#6B7280';
            badge.textContent = '— Not Checked In';
            statusDisplay.appendChild(badge);
          }
        }
        
        function handleCheckIn() {
          const empId = document.getElementById('employeeId').value.trim();
          const location = document.getElementById('location').value;
          
          if (!empId) {
            showAlert('Please enter Employee ID', 'warning');
            return;
          }
          
          if (!location) {
            showAlert('Please select a location', 'warning');
            return;
          }
          
          const btn = event.target;
          btn.disabled = true;
          btn.innerHTML = '<span class="spinner"></span>';
          
          google.script.run.withSuccessHandler(handleCheckInResult).withFailureHandler(handleError).performCheckIn(empId, location);
        }
        
        function handleCheckOut() {
          const empId = document.getElementById('employeeId').value.trim();
          const location = document.getElementById('location').value;
          
          if (!empId) {
            showAlert('Please enter Employee ID', 'warning');
            return;
          }
          
          if (!location) {
            showAlert('Please select a location', 'warning');
            return;
          }
          
          const btn = event.target;
          btn.disabled = true;
          btn.innerHTML = '<span class="spinner"></span>';
          
          google.script.run.withSuccessHandler(handleCheckOutResult).withFailureHandler(handleError).performCheckOut(empId, location);
        }
        
        function handleCheckInResult(result) {
          const btn = document.querySelector('.btn-checkin');
          btn.disabled = false;
          btn.innerHTML = '✓ CHECK IN';
          
          if (result.success) {
            showAlert('✓ ' + result.message, 'success');
            setTimeout(() => loadTodayStatus(), 500);
          } else {
            showAlert('❌ ' + result.message, 'error');
          }
        }
        
        function handleCheckOutResult(result) {
          const btn = document.querySelector('.btn-checkout');
          btn.disabled = false;
          btn.innerHTML = '✓ CHECK OUT';
          
          if (result.success) {
            showAlert('✓ ' + result.message, 'success');
            setTimeout(() => loadTodayStatus(), 500);
          } else {
            showAlert('❌ ' + result.message, 'error');
          }
        }
        
        function handleError(error) {
          showAlert('Error: ' + error, 'error');
          document.querySelectorAll('button').forEach(btn => {
            btn.disabled = false;
            if (btn.classList.contains('btn-checkin')) {
              btn.innerHTML = '✓ CHECK IN';
            } else if (btn.classList.contains('btn-checkout')) {
              btn.innerHTML = '✓ CHECK OUT';
            }
          });
        }
        
        function showAlert(message, type) {
          const alert = document.getElementById('alert');
          alert.className = 'alert ' + type;
          alert.textContent = message;
          alert.style.display = 'block';
          
          if (type !== 'error') {
            setTimeout(() => {
              alert.style.display = 'none';
            }, 4000);
          }
        }
      </script>
    </body>
    </html>
  `;
}

// ============================================================================
// BACKEND FUNCTIONS
// ============================================================================

/**
 * Get employee data by ID
 */
function getEmployeeData(employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == employeeId) {
        if (data[i][6] === 'Active') {
          return {
            found: true,
            name: data[i][1],
            email: data[i][2],
            allowedLocation: data[i][5],
            status: data[i][6]
          };
        } else {
          return { found: false, message: 'Employee is not active' };
        }
      }
    }
    
    return { found: false };
  } catch (error) {
    Logger.log('Error in getEmployeeData: ' + error);
    throw error;
  }
}

/**
 * Get today's attendance status for an employee
 */
function getTodayStatus(employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let checkIn = false;
    let checkOut = false;
    let checkInTime = '';
    let checkOutTime = '';
    
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][0]);
      date.setHours(0, 0, 0, 0);
      
      if (date.getTime() === today.getTime() && data[i][1] == employeeId) {
        if (data[i][3]) {
          checkIn = true;
          checkInTime = formatTime(data[i][3]);
        }
        if (data[i][5]) {
          checkOut = true;
          checkOutTime = formatTime(data[i][5]);
        }
        break;
      }
    }
    
    return { checkIn, checkOut, checkInTime, checkOutTime };
  } catch (error) {
    Logger.log('Error in getTodayStatus: ' + error);
    throw error;
  }
}

/**
 * Perform check-in operation
 */
function performCheckIn(employeeId, location) {
  try {
    // Validate employee
    const empData = getEmployeeData(employeeId);
    if (!empData.found) {
      return { success: false, message: 'Employee not found or inactive' };
    }
    
    // Check if location is allowed for employee
    if (empData.allowedLocation !== location) {
      return { success: false, message: 'You are not authorized to check in at this location' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Check if already checked in today
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][0]);
      date.setHours(0, 0, 0, 0);
      
      if (date.getTime() === today.getTime() && data[i][1] == employeeId) {
        if (data[i][3]) {
          return { success: false, message: 'Already checked in today at ' + formatTime(data[i][3]) };
        }
      }
    }
    
    // Add check-in record
    const now = new Date();
    const checkInTime = now;
    const checkInLocation = location;
    
    // Find or create today's record
    let recordFound = false;
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][0]);
      date.setHours(0, 0, 0, 0);
      
      if (date.getTime() === today.getTime() && data[i][1] == employeeId) {
        // Update existing record
        sheet.getRange(i + 1, 4).setValue(checkInTime);
        sheet.getRange(i + 1, 5).setValue(checkInLocation);
        recordFound = true;
        break;
      }
    }
    
    if (!recordFound) {
      // Create new record
      const nextRow = data.length + 1;
      sheet.getRange(nextRow, 1).setValue(today);
      sheet.getRange(nextRow, 2).setValue(employeeId);
      sheet.getRange(nextRow, 3).setValue(empData.name);
      sheet.getRange(nextRow, 4).setValue(checkInTime);
      sheet.getRange(nextRow, 5).setValue(checkInLocation);
      sheet.getRange(nextRow, 9).setValue('Present');
    }
    
    return { success: true, message: 'Check-in successful at ' + location + ' at ' + formatTime(checkInTime) };
  } catch (error) {
    Logger.log('Error in performCheckIn: ' + error);
    return { success: false, message: 'Error: ' + error.message };
  }
}

/**
 * Perform check-out operation
 */
function performCheckOut(employeeId, location) {
  try {
    // Validate employee
    const empData = getEmployeeData(employeeId);
    if (!empData.found) {
      return { success: false, message: 'Employee not found or inactive' };
    }
    
    // Check if location is allowed for employee
    if (empData.allowedLocation !== location) {
      return { success: false, message: 'You are not authorized to check out at this location' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Find today's check-in record
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][0]);
      date.setHours(0, 0, 0, 0);
      
      if (date.getTime() === today.getTime() && data[i][1] == employeeId) {
        if (!data[i][3]) {
          return { success: false, message: 'You must check in before checking out' };
        }
        
        if (data[i][5]) {
          return { success: false, message: 'Already checked out today at ' + formatTime(data[i][5]) };
        }
        
        // Update check-out time
        const now = new Date();
        const checkOutTime = now;
        const checkOutLocation = location;
        
        sheet.getRange(i + 1, 6).setValue(checkOutTime);
        sheet.getRange(i + 1, 7).setValue(checkOutLocation);
        
        // Calculate working hours
        const checkInDate = new Date(data[i][3]);
        const workingHours = ((checkOutTime - checkInDate) / (1000 * 60 * 60)).toFixed(2);
        sheet.getRange(i + 1, 8).setValue(workingHours);
        
        return { success: true, message: 'Check-out successful at ' + location + ' at ' + formatTime(checkOutTime) + '. Working Hours: ' + workingHours };
      }
    }
    
    return { success: false, message: 'No check-in found for today' };
  } catch (error) {
    Logger.log('Error in performCheckOut: ' + error);
    return { success: false, message: 'Error: ' + error.message };
  }
}

/**
 * Format time for display
 */
function formatTime(date) {
  if (!date) return '';
  const d = new Date(date);
  return d.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

// ============================================================================
// MENU SETUP
// ============================================================================

/**
 * Add menu items to the spreadsheet
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📊 Attendance')
      .addItem('🚀 Initialize System', 'initializeSpreadsheet')
      .addItem('👤 Open Employee Portal', 'showAttendanceUI')
      .addSeparator()
      .addItem('📈 View Dashboard', 'openDashboard')
      .addItem('📋 View Attendance Records', 'openAttendance')
      .addItem('👥 Manage Employees', 'openEmployees')
      .addToUi();
  } catch (error) {
    Logger.log('Error in onOpen: ' + error);
  }
}

/**
 * Open Dashboard sheet
 */
function openDashboard() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD).activate();
}

/**
 * Open Attendance sheet
 */
function openAttendance() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.ATTENDANCE).activate();
}

/**
 * Open Employees sheet
 */
function openEmployees() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES).activate();
}