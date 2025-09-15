function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("New User")
    .addItem("Open User Form", "openUserForm")
    .addToUi();
}

function openUserForm() {
  var url = "YOUR_WEB_APP_URL_HERE"; // Replace with deployed web app URL
  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>')
      .setWidth(100)
      .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening Form...");
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form')
      .setTitle('Add New User')
      .setWidth(800)
      .setHeight(800);
}

// Fetch data from reference sheet
function getDepartments() {
  return fetchColumnData("LoV", "L2:L");
}
function getCompanies() {
  return fetchColumnData("LoV", "N2:N");
}
function getCountries() {
  return fetchColumnData("LoV", "D2:D9");
}

function fetchColumnData(sheetName, range) {
  try {
    var sheet = SpreadsheetApp.openById('YOUR_SHEET_ID_HERE').getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
    return sheet.getRange(range).getValues().flat().filter(String);
  } catch (e) {
    Logger.log(`Error fetching data from ${sheetName}: ${e.toString()}`);
    return [];
  }
}

function addNewUser(formData) {
  try {
    Logger.log('Form Data Received: ' + JSON.stringify(formData));

    if (!formData.firstName || !formData.lastName || !formData.email || 
        !formData.company || !formData.department || !formData.userType || 
        !formData.country) {
      throw new Error('Missing required fields in form data.');
    }

    var sheetId = 'YOUR_CREATION_SHEET_ID_HERE';
    var sheetName = 'For Creation';
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet not found. Check the sheet name and ID.');

    var fullName = formData.firstName + ' ' + formData.lastName;
    var colMValue = getOrganizationalUnit(formData.office);
    var formattedAccountExpiresDate = getFormattedExpirationDate();

    var newUser = [
      '', '', fullName, fullName, formData.firstName, formData.lastName, formData.initials,
      formData.company, formData.description, formData.department, formData.office,
      formData.position, colMValue, formData.email, formattedAccountExpiresDate, 
      formData.manager, formData.joblevel
    ];

    var nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, newUser.length).setValues([newUser]);

    appendToMasterlist(formData, formattedAccountExpiresDate);

    return 'User added successfully!';
  } catch (e) {
    Logger.log('Error adding new user: ' + e.toString());
    return 'Error adding user: ' + e.toString();
  }
}

// Append to masterlist
function appendToMasterlist(formData, formattedAccountExpiresDate) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var masterlistSheet = SpreadsheetApp.openById("YOUR_MASTERLIST_SHEET_ID_HERE").getSheetByName("Masterlist");
    var creationSheet = spreadsheet.getSheetByName("For Creation");

    if (!masterlistSheet) throw new Error("Sheet 'Masterlist' not found.");
    if (!creationSheet) throw new Error("Sheet 'For Creation' not found.");

    var lastRow = masterlistSheet.getLastRow();
    var range = masterlistSheet.getRange(2, 6, lastRow - 1, 12);
    var values = range.getValues();
    var emptyRow = values.findIndex(row => row.every(cell => cell === "")) + 2;
    if (emptyRow === 1) emptyRow = lastRow + 1;

    var manager_name = formData.manager || "";

    var fullName = formData.firstName + ' ' + formData.lastName;
    var newData = [
      formData.userType || "",
      formData.company || "",
      formData.department || "",
      formData.country || "",
      formData.office || "",
      formData.email || "",
      fullName || "",
      manager_name || "",
      "",
      formattedAccountExpiresDate || "",
      "",
      "Active"
    ];

    masterlistSheet.getRange(emptyRow, 6, 1, newData.length).setValues([newData]);

    var validationRange = masterlistSheet.getRange(emptyRow, 17);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Active", "Disabled", "Terminated"], true)
      .build();
    validationRange.setDataValidation(rule);

  } catch (e) {
    Logger.log("Error in appendToMasterlist: " + e.toString());
    throw e;
  }
}

// Map office to organizational unit (placeholder)
function getOrganizationalUnit(office) {
  var mapping = {
    'Office1': 'OU=Office1,OU=Country1,OU=Enabled Users,OU=User Accounts,DC=example,DC=net',
    'Office2': 'OU=Office2,OU=Country2,OU=Enabled Users,OU=User Accounts,DC=example,DC=net'
  };
  return mapping[office] || 'Invalid Office';
}

// Format expiration date
function getFormattedExpirationDate() {
  var today = new Date();
  today.setDate(today.getDate() + 365);
  return Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function getManagers() {
  const sheet = SpreadsheetApp.openById("YOUR_MASTERLIST_SHEET_ID_HERE").getSheetByName("Masterlist");
  const colE = sheet.getRange("E2:E").getValues().flat();
  const colL = sheet.getRange("L2:L").getValues().flat();

  return colE.reduce((acc, val, i) => {
    if (typeof val === "string" && val.trim().endsWith("1")) {
      const displayName = colL[i] || "";
      if (displayName) {
        acc.push({ name: displayName, id: val.trim() });
      }
    }
    return acc;
  }, []);
}
