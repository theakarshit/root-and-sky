/**
 * Root & Sky — Google Apps Script for form submissions
 *
 * STANDALONE SCRIPT — works from script.google.com directly.
 * Connected to Sheet: https://docs.google.com/spreadsheets/d/1agI1syfp8eLhZZeqFi1_-4tN1Jllj_3GqY7hcDTGflY/
 *
 * SETUP:
 * 1. Go to script.google.com → New project
 * 2. Paste this entire script
 * 3. Run setupSheet() once (select it from dropdown → Run)
 * 4. Grant permissions when prompted (Sheets + Mail access)
 * 5. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the deployment URL → replace APPS_SCRIPT_URL in build.html and index.html
 */

const SHEET_ID = '1agI1syfp8eLhZZeqFi1_-4tN1Jllj_3GqY7hcDTGflY';

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('Leads') || ss.getSheets()[0];
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    // Format plants list
    const plantsList = Array.isArray(data.plants)
      ? data.plants.map(p => typeof p === 'string' ? p : p.name || p.n || '').join(', ')
      : (data.plants || '');

    // Format timestamp
    const timestamp = new Date(data.submittedAt || new Date());
    const formattedDate = Utilities.formatDate(timestamp, 'Asia/Kolkata', 'dd-MMM-yyyy HH:mm');

    // Append row with all data
    sheet.appendRow([
      data.refId || '',                          // A: Ref ID
      formattedDate,                             // B: Timestamp
      data.name || '',                           // C: Name
      data.phone || '',                          // D: Phone
      data.email || '',                          // E: Email
      data.cityLabel || data.city || '',         // F: City
      data.spaceLabel || data.spaceType || '',   // G: Space Type
      data.theme || '',                          // H: Theme
      data.tierName || data.planName || '',       // I: Plan Tier
      data.price || data.grandTotal || '',       // J: Price
      plantsList,                                // K: Plants
      data.size || '',                           // L: Space Size
      data.area || '',                           // M: Locality
      data.timing || '',                         // N: Visit Timing
      data.notes || '',                          // O: Notes
      'New',                                     // P: Status (default)
      '',                                        // Q: Follow-up Date
    ]);

    // Auto-format: set Status column color for new entries
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 16).setBackground('#e8f5e9').setFontWeight('bold');

    // Auto-resize columns occasionally
    if (lastRow % 10 === 0) {
      for (let i = 1; i <= 17; i++) {
        sheet.autoResizeColumn(i);
      }
    }

    // Send email notification to owner (optional — fill in your email)
    const OWNER_EMAIL = ''; // e.g. 'you@email.com'
    if (OWNER_EMAIL) {
      const subject = '🌿 New Garden Request — #' + (data.refId || 'N/A');
      const body = [
        'New garden request received!',
        '',
        'Name: ' + (data.name || ''),
        'Phone: ' + (data.phone || ''),
        'Email: ' + (data.email || ''),
        'City: ' + (data.cityLabel || ''),
        '',
        'Space: ' + (data.spaceLabel || ''),
        'Theme: ' + (data.theme || ''),
        'Plan: ' + (data.tierName || '') + ' — ₹' + (data.price || ''),
        '',
        'Plants: ' + plantsList,
        '',
        'Size: ' + (data.size || 'Not specified'),
        'Area: ' + (data.area || 'Not specified'),
        'Visit: ' + (data.timing || 'Flexible'),
        'Notes: ' + (data.notes || '—'),
      ].join('\n');

      MailApp.sendEmail(OWNER_EMAIL, subject, body);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', refId: data.refId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET handler (for testing the URL works)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Root & Sky API is live. Use POST to submit.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this ONCE to set up the sheet headers and formatting.
 * Select setupSheet from the dropdown → click Run
 */
function setupSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Leads');

  if (!sheet) {
    sheet = ss.insertSheet('Leads');
  }

  // Headers
  const headers = [
    'Ref ID', 'Timestamp', 'Name', 'Phone', 'Email', 'City',
    'Space Type', 'Theme', 'Plan Tier', 'Price', 'Plants',
    'Space Size', 'Locality', 'Visit Timing', 'Notes', 'Status', 'Follow-up Date'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a2318');
  headerRange.setFontColor('#f5f0e8');
  headerRange.setFontWeight('bold');
  headerRange.setFontFamily('Arial');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  const widths = [90, 140, 150, 130, 200, 90, 100, 160, 80, 90, 400, 80, 120, 160, 250, 100, 110];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Add Status dropdown validation (for up to 500 rows)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Contacted', 'Site Visit Booked', 'Site Visit Done', 'Quote Sent', 'Confirmed', 'Completed', 'Lost'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 16, 500, 1).setDataValidation(statusRule);

  // Add conditional formatting for Status column
  const rules = sheet.getConditionalFormatRules();

  const statusColors = {
    'New': '#e8f5e9',
    'Contacted': '#fff3e0',
    'Site Visit Booked': '#e3f2fd',
    'Site Visit Done': '#f3e5f5',
    'Quote Sent': '#fff8e1',
    'Confirmed': '#c8e6c9',
    'Completed': '#a5d6a7',
    'Lost': '#ffcdd2',
  };

  Object.entries(statusColors).forEach(([status, color]) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(color)
        .setRanges([sheet.getRange(2, 16, 500, 1)])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);

  // Add Follow-up Date validation
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 17, 500, 1).setDataValidation(dateRule);

  SpreadsheetApp.flush();
  Logger.log('✅ Sheet setup complete! Check your sheet: https://docs.google.com/spreadsheets/d/' + SHEET_ID);
}
