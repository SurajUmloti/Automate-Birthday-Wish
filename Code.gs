const TEMPLATE_FILE_ID = '1D3dMhS-0iRf6w7wXjCBBnWp6p0R6b2lHv34pO3z9fSE';
const TEMPLATE_FILE = 'https://docs.google.com/document/d/1D3dMhS-0iRf6w7wXjCBBnWp6p0R6b2lHv34pO3z9fSE/edit';
const SPREADSHEET_FILE = 'https://docs.google.com/spreadsheets/d/1Dz42ddb66ct3be8zJxYPAoBH8bAhbNusAKAZw64fsGw/edit#gid=0';

function sendBdayWishes(){
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);
    var sheet = spreadsheet.getActiveSheet();
    var today = new Date(); 
    var startRow = 5; // Starting row in the excel
    
    for(var i =startRow ;i<=sheet.getLastRow(); i++){
        if (sheet.getRange(i, 4).getValue()==true) {
             var name = sheet.getRange(i,1).getValue();
             var toMail= sheet.getRange(i,3).getValue();
             sendMail(name, toMail);
            //  sheet.getRange(i,5).setValue("Bday wishes sent");
        }
     }
    }

    function sendMail(name, toMail){
    var template = DriveApp.getFileById(TEMPLATE_FILE_ID).makeCopy('temp').getId();
    var doc = DocumentApp.openById(template); // Temporary Copy
    var body = doc.getBody();
    body.replaceText('#name#', name);// Update name variable in template
    doc.saveAndClose();

    var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+ template + "&exportFormat=html";
    var param = {
    method      : "get",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()}
    };
    var htmlBody =  UrlFetchApp.fetch(url,param).getContentText();
    var trashed = DriveApp.getFileById(template).setTrashed(true);// delete temporary copy
    
    var subject = 'Happy Birthday' + name;
    var body = {htmlBody : htmlBody}
    MailApp.sendEmail(toMail, subject,' ' , body);
    }
    