var ss = SpreadsheetApp.getActive()
var formR = ss.getSheetByName('Form Responses')
var pdf = ss.getSheetByName('PDF')
var customPdf = ss.getSheetByName('Custom PDF')
var setting = ss.getSheetByName('Setting')
var Master = ss.getSheetByName('Master')
var approved = ss.getSheetByName('Approved')
var disapproved = ss.getSheetByName('Disapproved')


function getPDF(sheetName,lastRow){
var ss = SpreadsheetApp.getActive();
sheet = ss.getSheetByName(sheetName);
const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&portrait=true&top_margin=0.42&bottom_margin=0.42&left_margin=0.42&right_margin=0.42&range=A1:J${lastRow}&gid=${sheet.getSheetId()}`;
const pdfBlob = UrlFetchApp.fetch(url, { headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() } }).getBlob().setName("CustomPDF");
var file = DriveApp.createFile(pdfBlob);
var fileUrl = file.getUrl();
var fileId = file.getId()
// Logger.log(fileUrl);
return [fileUrl,fileId];
}


function send([recipientEmail,subject,mergedBody,fileID,nameIN,ccEmail,bccEmail,replyTo,recipientNo,newMSG,url,api])
{
  var file = DriveApp.getFileById(fileID)
  try{
  GmailApp.sendEmail(recipientEmail, subject, mergedBody, {
  attachments: [file.getAs(MimeType.PDF)],
  name: nameIN,
  cc: ccEmail,
  bcc:bccEmail,
  replyTo:replyTo,
  });
  console.log("url :" + url )
WhatsappAutomation.sendMessage(recipientNo,newMSG,url,[api])
  }catch{
    console.log(error);
  }
}