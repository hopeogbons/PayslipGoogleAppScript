/******************************************
 ******************************************
 ** Script Title: PayslipGoogleAppScript **
 ** Author: Hope Ogbons                  **
 ** Email: hopeogbons@gmail.com          **
 ** Phone: 08033644880                   **
 ******************************************
 ******************************************/

function onOpen() {
  var menuEntries = [{ name: 'Recalculate Cumulative', functionName: 'recalculateCumulative' }, { name: 'Dispatch Payslips Now!', functionName: 'dispatchPayslips' }];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu('Payslips', menuEntries);
  zebraColorAllSheets(ss);
}

function onEdit(e) {
  if (e) {
    var ss = e.source.getActiveSheet();
    var r = e.source.getActiveRange();
    var sheetName = ss.getName();

    // Editing a valid month sheet
    if (months().indexOf(sheetName) > -1) {
      var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = activeSpreadsheet.getSheetByName(sheetName);
      zebraColorOneSheet(sheet);
    }

    // Editing column 23 of a valid month sheet
    if (months().indexOf(sheetName) > -1 && r.getRow() !== 1 && r.getColumn() === 23) {
      var status = ss.getRange(r.getRow(), 23).getValue();
      var rowRange = ss.getRange(r.getRow(), 23);

      if (status.trim().toLowerCase() === 'sent') {
        rowRange.setBackgroundColor("#98FB98");
      }
      else if (status.trim().toLowerCase() === 'pending') {
        rowRange.setBackgroundColor("#F4AF60");
      } else {
        rowRange.setBackgroundColor("#FFFFFF");
      }
    }
  }
}

function zebraColorOneSheet(sheet) {
  var rowCount = sheet.getLastRow() - 1;
  var columnCount = sheet.getLastColumn() - 1;
  if (rowCount > 0) setAlternatingRowBackgroundColors(sheet.getRange(2, 1, rowCount, columnCount), '#FFFFFF', '#E9E9E9');
}

function zebraColorAllSheets(ss) {
  months().map(function (month) {
    var sheet = ss.getSheetByName(month);
    zebraColorOneSheet(sheet);
  });
}

function setAlternatingRowBackgroundColors(range, oddColor, evenColor) {
  var backgrounds = [];
  for (var row = 1; row <= range.getNumRows(); row++) {
    var rowBackgrounds = [];
    for (var column = 1; column <= range.getNumColumns(); column++) {
      if (row % 2 == 0) {
        rowBackgrounds.push(evenColor);
      } else {
        rowBackgrounds.push(oddColor);
      }
    }
    backgrounds.push(rowBackgrounds);
  }
  range.setBackgrounds(backgrounds);
}

function recalculateCumulative() {
  var arrValidTabs = months();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var objCurrentSheet = ss.getActiveSheet();
  var strCurrentSheetName = objCurrentSheet.getName()
  var intIndexCurrentSheet = arrValidTabs.indexOf(strCurrentSheetName);

  if (intIndexCurrentSheet > -1) {
    var booIsJanTab = (strCurrentSheetName === arrValidTabs[0]) ? true : false;
    var arrCurrentData = objCurrentSheet.getRange(2, 1, objCurrentSheet.getLastRow() - 1, objCurrentSheet.getLastColumn()).getValues();

    for (var x in arrCurrentData) {
      var arrCurrent = arrCurrentData[x];

      var currentUserId = arrCurrent[1];
      var currentGrossPay = arrCurrent[12];
      var currentTaxPayable = arrCurrent[16];

      if (booIsJanTab) {
        var grossPaid = currentGrossPay;
        var taxPaid = currentTaxPayable;

        objCurrentSheet.getRange(2 + parseInt(x), 20).setValue(grossPaid.toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; })).setBackgroundColor("#E9E9E9");
        objCurrentSheet.getRange(2 + parseInt(x), 21).setValue(taxPaid.toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; })).setBackgroundColor("#E9E9E9");
      } else {
        var objPreviousSheet = ss.getSheetByName(arrValidTabs[parseInt(intIndexCurrentSheet) - 1])
        var arrPreviousData = objPreviousSheet.getRange(2, 1, objPreviousSheet.getLastRow() - 1, objPreviousSheet.getLastColumn()).getValues();

        for (var y in arrPreviousData) {
          var arrPrevious = arrPreviousData[y];

          var previousUserId = arrPrevious[1];
          var previousGrossPay = arrPrevious[12];
          var previousTaxPayable = arrPrevious[16];
          var previousGrossPaid = arrPrevious[19];
          var previousTaxPaid = arrPrevious[20];

          if (currentUserId === previousUserId) {
            var grossPaid = currentGrossPay + previousGrossPaid;
            var taxPaid = currentTaxPayable + previousTaxPaid;

            objCurrentSheet.getRange(2 + parseInt(x), 20).setValue(grossPaid.toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; })).setBackgroundColor("#E9E9E9");
            objCurrentSheet.getRange(2 + parseInt(x), 21).setValue(taxPaid.toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; })).setBackgroundColor("#E9E9E9");
          }
        }
      }
    }

    ss.toast('Successful! The cumulative Gross Paid & Tax Paid have been recalculated. Thank you.');
  } else {
    ss.toast('Failed! The cumulative Gross Paid & Tax Paid can not be recalculated. Please, try again later.');
  }
}

function dispatchPayslips() {
  var payslipFolder = DriveApp.getFolderById('xxxXXXXXXXXxxxXXXxxxXXXXXxxxxxxxxxxxxxxXXXXXXXXXXXX');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var i in data) {
    var docTemplateId = DriveApp.getFileById('xxxXXXXXXXXxxxXXXxxxXXXXXxxxxxxxxxxxxxxXXXXXXXXXXXX').makeCopy().getId();
    var row = data[i];
    var sentStatus = row[22];
    if ('pending' === sentStatus.toLowerCase().trim()) {
      var userId = row[1];
      var userName = row[2];
      var userEmail = row[21];

      var staffDate = processSalaryDate(row[5]);
      var date = new Date(staffDate.year, staffDate.month);
      var period = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM, YYYY');

      var newDoc = DocumentApp.create(['Payslip Document for ', userName, ' (', userId, ')'].join(''));
      var file = DriveApp.getFileById(newDoc.getId());
      payslipFolder.addFile(file);
      var doc = DocumentApp.openById(docTemplateId);
      var body = doc.getActiveSection();
      body.replaceText('%month%', period.split(',')[0]);
      body.replaceText('%employeeId%', userId);
      body.replaceText('%employeeNames%', userName);
      body.replaceText('%jobTitle%', row[3]);
      body.replaceText('%department%', row[4]);
      body.replaceText('%payStartDate%', row[5]);
      body.replaceText('%payEndDate%', row[6]);
      body.replaceText('%basic%', row[7].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%housing%', row[8].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%transportation%', row[9].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%others%', row[10].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%bonus%', row[11].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%grossPay%', row[12].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%pension%', row[13].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%loansAndOthers%', row[14].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%rent%', row[15].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%taxPayable%', row[16].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%netPay%', row[17].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%deductions%', row[18].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%grossPaid%', row[19].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));
      body.replaceText('%taxPaid%', row[20].toFixed(2).replace(/./g, function (c, i, a) { return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c; }));

      appendToDoc(doc, newDoc);

      doc.saveAndClose();
      newDoc.saveAndClose();
      var message = [['Hello ', userName.split(' ')[0], ','].join(''), ' ', ['Please, find attached your Payslip for the month of', period].join(' ')].join('\n'); // Email Message
      var emailTo = userEmail.trim() // Send To Email
      var subject = ['Payslip for the month of', period].join(' '); // Email Subject
      var pdf = DriveApp.getFileById(newDoc.getId()).getAs('application/pdf').getBytes();
      var attach = { fileName: [row[2], ' (', row[1], ') ', period, ' Payslip.pdf'].join(''), content: pdf, mimeType: 'application/pdf' }; // PDF File Name
      MailApp.sendEmail(emailTo, subject, message, { attachments: [attach] });
      sheet.getRange(2 + parseInt(i), 23).setValue('Sent').setBackgroundColor("#98FB98");
      DriveApp.getFileById(docTemplateId).setTrashed(true);

      SpreadsheetApp.flush();
    }
  }
  ss.toast('Successful! Payslips have been dispatched. Thank you.');
}

function processSalaryDate(strDate) {
  var objDate = {};
  var arrDate = strDate.trim().split('/');
  var month = arrDate[1] - 1;

  objDate.year = arrDate[2];
  objDate.month = month.toString();

  return objDate;
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

/**
 * Returns an array of the months structure.
 */
function months() {
  return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
}