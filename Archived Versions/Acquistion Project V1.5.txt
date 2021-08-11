//Header on spreadsheet make it easier to reference.
const ACTION = 'Action';   
const COMMENT = 'Comment'; 
const STATUS = 'Status';   
const TIMESTAMP = 'Timestamp';
const CONTRACT = 'Contract'; 
const REQ = 'Req'; 
const REQUISITIONER = 'Requisitioner'; 
const MARS_ENTRY = 'Mars Entry';
const SUBMIT_TO = 'Submit to';
const NAME = 'Name';
const DESCRIPTION = 'Description';
const EMAIL_ADDRESS = 'Email Address';
const DOLLAR_AMOUNT = 'Dollar Amount';
const PROCUREMENT_REQUEST_SUBMISSION = 'Procurement Request Submission';
const IT_CHECKLIST_SUBMISSION = 'IT Checklist Submission';
const SOLE_SOURCE_JUSTIFICATION_AND_BRAND_NAME_JUSTIFICATION = 'Sole Source Justification and Brand Name Justification';
const OTHER_DOCUMENTS = 'Other Documents';
const TECH_PURCHASE = 'Tech Purchase';
const NOAASTANDARDS_FORM = 'NOAAStandards Form';
const CHOOSE_WHICH_BEST_DESCRIBES_YOUR_PURCHASE = 'Choose which best describes your purchase';
const NOAALINK_WORKSHEET = 'NOAALink Worksheet';
const IGCE = 'IGCE'; 
const STATEMENT_OF_NEED = 'Statement of Need';
const CONTRACT_PO_NUMBER = 'Contract PO Number';
const EXTRA_UPLOADS = 'Extra Uploads';
const NOTES = 'Notes';
const PROJECT_TASK_CODE = 'Project Task Code';
const LOG_ASSIGNED = 'Log Assigned';
const LOG_APPROVED = 'Log Approved';
const LOG_AWARDED = 'Log Awarded';

const TEMPLATES_SHEET = 'Templates'; 
const ss = SpreadsheetApp.getActiveSpreadsheet();

/**
 * Installs a trigger in the Spreadsheet to run upon the Sheet being opened and form submission.
 * To learn more about triggers read:
 * https://developers.google.com/apps-script/guides/triggers
 */

function emailer() {
  GmailApp.sendEmail('joshua.mccluskey@noaa.gov, ben.carlson@noaa.gov', 'NOTIFICATION: Req Submission', 'Hello,\n\nPlease review FY 21 Req Log for submission:\n\n' + ss.getUrl())
}

let sheet = SpreadsheetApp.getActive();

function formTrigger() {
  let builder = ScriptApp.newTrigger("emailer")
  .forSpreadsheet(sheet)
  .onFormSubmit()
  .create();
}


function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('📨 Send Notification Email')
      .addItem('Assigned', 'processAssigned')
      .addItem('Approved', 'processApproved')
      .addItem('Awarded', 'processAwarded')
      .addItem('Info', 'processInfo')
      .addToUi();
}

/**
 * Wrapper function of `processRows` for the 'Assigned' action.
 */
function processAssigned() {
  processRows('Assigned');
}

/**
 * Wrapper function of `processRows` for the 'Approved' action.
 */
function processApproved() {
  processRows('Approved');
}

/**
 * Wrapper function of `processRows` for the 'Awarded' action.
 */
function processAwarded() {
  processRows('Awarded');
}

/**
 * Wrapper function of `processRows` for the 'Awarded' action.
 */
function processInfo() {
  processRows('Info');
}

/**
 * Processes only the rows matching the action.
 * It sends an email if the `STATUS` column is empty.
 * This updates the `STATUS` column in the sheet.
 */
function processRows(action, emailTemplate=null) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get the email template doc URLs into a {key: value} Map 
  // in the format {templateName: templateURL}.
  let templateRows = ss.getSheetByName(TEMPLATES_SHEET).getDataRange().getValues();
  let templates = templateRows
      .reduce((result, row) => result.set(row[0], row[1]), new Map());

  // Load the row data and get its headers.
  let dataRange = ss.getActiveSheet().getDataRange();
  let rows = dataRange.getValues();
  let headers = rows.shift();

  // Get the values from the status column.
  // These are the values that we want to write back to the sheet.
  let statusRange = dataRange.offset(1, headers.indexOf(STATUS), rows.length, 1);
  let statusValues = statusRange.getValues();

  // Process each row, send an email if necessary and update the `statusValues`.
  rows
      // Convert the row arrays into objects.
      // Start with an empty object, then create a new field
      // for each header name using the corresponding row value.
      .map(rowArray => headers.reduce((rowObject, fieldName, i) => {
        rowObject[fieldName] = rowArray[i];
        return rowObject;
      }, {}))

      // Add the row index (0-based) to the row object, this is used to update
      // the status of the rows that were modified.
      // We do this because the indices won't match after the next `filter` operation.
      // We use the spread operator to unpack the `row` object.
      // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Spread_syntax
      .map((row, i) => ({...row, rowIndex: i}))

      // From all the rows, filter out and only keep the ones that match the
      // action and the status is empty.
      .filter(row => row[ACTION] == action && !row[STATUS])

      // Send an email and update the status in `statusValues`.
      // We don't need a return value so we use `forEach` instead of `map`.
      .forEach(row => {
        // We start with the doc template HTML body, and then we replace
        // each '{{fieldName}}' with the row's respective value.
        let emailBody = headers.reduce(
          (result, fieldName) => result.replace(`{{${fieldName.toUpperCase()}}}`, row[fieldName]),
          docToHtml(templates.get(emailTemplate || action))
        );

        // Try to send an email, or get the error if it fails.
        let status;
        try { 
         let ui = SpreadsheetApp.getUi();
         let response = ui.alert('Hold Up!','Are you sure you want to send the email? *Choose "No" for Libby\'s purchases*', ui.ButtonSet.YES_NO);

         // Process the user's response.
         if (response == ui.Button.YES) {
          MailApp.sendEmail({
            to: row[REQUISITIONER],
            cc: row[EMAIL_ADDRESS],
            subject: `Purchase Notification: ${row[ACTION]}`,
            htmlBody: emailBody,
          });} else {}
          status = `${row[ACTION]}: ${new Date}`;
        } catch (e) {
          status = `Error: ${e}`;
        }

        // Update the `statusValues` with the new status or error.
        // We use the `rowIndex` from before to update the correct
        // row in `statusValues`.
        statusValues[row.rowIndex][0] = status;
        Logger.log(`Row ${row.rowIndex+2}: ${status}`);
      });

  // Write statusValues back into the sheet "status" column.
  statusRange.setValues(statusValues);
}

/**
 * Fetches a Google Doc as an HTML string.
 *
 * @param {string} docUrl - The URL of a Google Doc to fetch content from.
 * @return {string} The Google Doc rendered as an HTML string.
 */
function docToHtml(docUrl) {
  let docId = DocumentApp.openByUrl(docUrl).getId();
  return UrlFetchApp.fetch(
    `https://docs.google.com/feeds/download/documents/export/Export?id=${docId}&exportFormat=html`,
    {
      method: 'GET',
      headers: {'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`},
      muteHttpExceptions: true,
    },
  ).getContentText();
}
