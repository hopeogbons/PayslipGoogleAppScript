function onOpen() {
  var menuEntries = [ {name: 'Dispatch Now!', functionName: 'AutofillDocFromTemplate'}];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu('Payslips', menuEntries);
}

function AutofillDocFromTemplate() {
  var templateid = 'xxxXXXXXXXXxxxXXXxxxXXXXXxxxxxxxxxxxxxxXXXXXXXXXXXX'; // Payslip Template
  var FOLDER_NAME = 'Company Payslips'; // Folder Name
  var date = new Date();
  var period = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM, YYYY');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
  for (var i in data) {
    var row = data[i];
    if (row[20] == 'No') {
      var userId = row[1]; 
      var userName = row[2]; 
      var newDoc = DocumentApp.create(['Payslip Document for ', userName, ' (', userId, ')'].join(''));
      var file = DriveApp.getFileById(newDoc.getId());
      var folder = DriveApp.getFolderById('xxxXXXXXXXXxxxXXXxxxXXXXXxxxxxxxxxxxxxxXXXXXXXXXXXX'); // Folder Id
      folder.addFile(file)  
      var docid = DriveApp.getFileById(templateid).makeCopy().getId();
      var doc = DocumentApp.openById(docid);
      var body = doc.getActiveSection();
      body.replaceText('%employeeId%', userId);
      body.replaceText('%employeeNames%', userName);
      body.replaceText('%jobTitle%', row[3]);
      body.replaceText('%department%', row[4]);
      body.replaceText('%payStartDate%', row[5]);
      body.replaceText('%payEndDate%', row[6]);
      body.replaceText('%basic%', row[7].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%housing%', row[8].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%transportation%', row[9].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%others%', row[10].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%bonus%', row[11].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%totalGrossSalary%', row[12].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%pension%', row[13].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%loansAndOthers%', row[14].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%rent%', row[15].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%taxPayable%', row[16].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%netPay%', row[17].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      body.replaceText('%totalDeductions%', row[18].toFixed(2).replace(/./g, function(c, i, a) {return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;}));
      
      appendToDoc(doc, newDoc);
      
      doc.saveAndClose();
      newDoc.saveAndClose();
      var message = ['Hello ', userName, ', Attached is your Payslip for the month of ', period].join(''); // Email Message
      var emailTo = row[19] // Send To Email
      var subject = ['Payslip for the month of', period].join(' '); // Email Subject
      var pdf = DriveApp.getFileById(newDoc.getId()).getAs('application/pdf').getBytes();
      var attach = {fileName: [row[2], ' (', row[1], ') ', period, ' Payslip.pdf'].join(''),content:pdf, mimeType:'application/pdf'}; // PDF File Name
      MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
      sheet.getRange(2 + parseInt(i), 21).setValue('Yes');
      SpreadsheetApp.flush();
      
      DriveApp.getFileById(docid).setTrashed(true);
    }
    
  }
  ss.toast(['Successful! Payslips have been dispatched for month of ', period, '. Thank you.'].join(''));
}

function appendToDoc(src, dst) {
  for (var i = 0; i < src.getNumChildren(); i++) {
    appendElementToDoc(dst, src.getChild(i));
  }
}

function appendElementToDoc(doc, object) {
  var type = object.getType();
  var element = object.copy();
  if (type == 'PARAGRAPH') {
    doc.appendParagraph(element);
  } else if (type == 'TABLE') {
    doc.appendTable(element);
  } 
}
