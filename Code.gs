function exportDropdownPDFs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const printSheet = ss.getSheetByName('deptLoading');
  const printSheetId = printSheet.getSheetId();
  const dropdownCell = printSheet.getRange('B3');
  
  //Collect all dropdown values
  const refSheet = ss.getSheetByName('deptLookup');
  const dropdownRange = refSheet.getRange('A2:A17');
  const dropdownValues = dropdownRange.getValues();

  //Date and time... filename is rendered in yyyy-MM-dd_hh:mm:ss (easily initialised by 'now')...
  SpreadsheetApp.flush();
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, 'Europe/London', 'yyyy-MM-dd_HH-mm');

  //...but the pdf has a timestamp using the 1900 date system
  //Division by 100 trillion removes floating point errors; the dateValue in the cell has been 
  //multiplied by 100 trillion (one second is 0.00001157407407 (14 decimal places)) 
  const dateValueDivideBy100tr = printSheet.getRange('A1').getValue() / 100000000000000; 
  
  // Loop through each option in the drop-down menu and generate pdfs
  for (let i = 0; i < dropdownValues.length; i++) {
    
    // Set the value of the drop-down menu to the current option
    let dropdownValue = dropdownValues[i][0];
    dropdownCell.setValue(dropdownValue);
    SpreadsheetApp.flush();

    //Modify name for filename purposes
    let modifyDropdownValue = dropdownValue.replace(' ', '-');
    let pdfName = 'deptLoading_' + formattedDate + '_Đ' + (i + 1).toString().padStart(2, '0') + '_' + modifyDropdownValue + '.pdf';

    //Export current sheet as pdf
    Utilities.sleep(i * 600); //Allow for latency, 'exponential backoff'
    ss.toast('Creating ' + pdfName);
    let pdf = createPDF(ssId, printSheetId, dateValueDivideBy100tr, pdfName);
    
  }

}

function removeDuplicates(arr) {
  return arr.filter(function(value, index, self) {
    return self.indexOf(value) === index;
  });
}

function createPDF(ssId, sheetId, dateValue, pdfName) {
  
  const folder = DriveApp.getFolderById('1CAbNl9oRZgPQz3cFR4vtWr8Ee1M4ojDP');
  const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export' +
    '?format=pdf&' +
    'size=7&' + //Size 7 is A4, size 6 is A3
    'fzc=false&' + //Repeat frozen columns
    'fzr=false&' + //Repeat frozen rows
    'portrait=false&' +
    'fitw=true&' +
    'printtitle=true&' +
    'sheetnames=true&' +
    'printdate=true&' +
    'printtime=true&' +
    'timestamp=' + dateValue + '&' +
    'top_margin=0.8&' +
    'bottom_margin=0.2&' +
    'left_margin=0.2&' +
    'right_margin=0.2&' +
    'gid=' + sheetId;
  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName);
  const pdfFile = folder.createFile(blob);
  Logger.log('Creating ' + pdfName);  
}

