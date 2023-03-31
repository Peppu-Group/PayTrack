// Original code from https://github.com/jamiewilson/form-to-google-sheets
// Updated for 2021 and ES6 standards

const sheetName = 'Invoice'
const scriptProp = PropertiesService.getScriptProperties()

function doPost(e) {
  let file = DriveApp.getFilesByName('Accounting').next();
  const activeSpreadsheet = SpreadsheetApp.open(file);
  scriptProp.setProperty('key', activeSpreadsheet.getId())

  const lock = LockService.getScriptLock()
  lock.tryLock(10000)

  try {
    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'))
    const sheet = doc.getSheetByName(sheetName)

    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]
    const nextRow = sheet.getLastRow() + 1

    const newRow = headers.map(function (header) {
      const invNumber = Math.floor(100000 + Math.random() * 900000);
      return header === 'Date' ? new Date()
        : header === 'Invoice No.' ? `INV${invNumber}`
          : header === 'Due Date' ? new Date()
            : header === 'Paid' ? 'Yes' : e.parameter[header]
    })

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow])

    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
      .setMimeType(ContentService.MimeType.JSON)
  }

  catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
      .setMimeType(ContentService.MimeType.JSON)
  }

  finally {
    lock.releaseLock()
  }
}