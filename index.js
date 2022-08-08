function saveDoc(doc, firstForm, secondForm, thirdForm) {
  var body = doc.getBody();

  if (firstForm.firstName) body.replaceText('{{FirstName}}', firstForm.firstName); 
  if (secondForm.middleName) body.replaceText('{{MiddleName}}', secondForm.middleName);  
  if (thirdForm.lastName) body.replaceText('{{LastName}}', thirdForm.lastName);

  doc.saveAndClose();
}

function autoFillDoc(e) {
  // Common field in all the forms
  var formId = e.namedValues['Form Id'][0]

  var firstForm = {
    firstName: e.namedValues['First Name'] && e.namedValues['First Name'][0],
  }

  var secondForm = {
    middleName: e.namedValues['Middle Name'] && e.namedValues['Middle Name'][0],
  }

  var thirdForm = {
    lastName: e.namedValues['Last Name'] && e.namedValues['Last Name'][0],
  }

  var dbSS = SpreadsheetApp.getActiveSpreadsheet()
  var dbSheet = dbSS.getSheetByName('DATABASE')
  var lastRow = dbSheet.getLastRow()
  var dbRange = dbSheet.getRange(2, 1, lastRow > 1 ? lastRow - 1: 1, 3)
  var dbVals = dbRange.getValues()

  //check if db has formId
  var docInfo = dbVals.find(el => el[0].toString() === formId.toString())
  var docName = null

  if (docInfo) {
    //edit document
    var doc = DocumentApp.openById(docInfo[1]);
    docName = doc.getName()
    saveDoc(doc, firstForm, secondForm, thirdForm)
  } else {
    //create document
    var file = DriveApp.getFileById('ID_OF_TEMPLATE_FILE'); 
    var folder = DriveApp.getFolderById('ID_OF_OUTPUT_FOLDER')
    var copy = file.makeCopy(`Document - ${formId}`, folder);

    var doc = DocumentApp.openById(copy.getId())
    docName = doc.getName()
    saveDoc(doc, firstForm, secondForm, thirdForm)

    dbSheet.appendRow([formId, copy.getId()])
  }

  if (firstForm.firstName) {
    var emailBody = 'First form filled successfully\n\n'
    emailBody += `Doc Name: ${docName ? `${docName}` : ''}\n`
    emailBody += `Date: ${new Date()}\n`
    GmailApp.sendEmail('EMAIL_ADDRESS', `First form filled - ${docName ? ` - ${docName}` : ''}`, emailBody)
  }
  if (secondForm.middleName) {
    var emailBody = 'Second form filled successfully\n\n'
    emailBody += `Doc Name: ${docName ? `${docName}` : ''}\n`
    emailBody += `Date: ${new Date()}\n`
    GmailApp.sendEmail('EMAIL_ADDRESS', `Second form filled - ${docName ? ` - ${docName}` : ''}`, emailBody)
  }
  if (thirdForm.lastName) {
    var emailBody = 'Third form filled successfully\n\n'
    emailBody += `Doc Name: ${docName ? `${docName}` : ''}\n`
    emailBody += `Date: ${new Date()}\n`
    GmailApp.sendEmail('EMAIL_ADDRESS', `Third form filled - ${docName ? ` - ${docName}` : ''}`, emailBody)
  }
}
