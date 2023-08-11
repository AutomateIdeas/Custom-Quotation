function onOpen(){
 var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
      .addItem('Create Trigger', 'createTrigger')
      .addSeparator()
          .addItem('Custom PDF', 'customPDFF')
      .addToUi();
}

function createTrigger(){
  deleteTrigger()
ScriptApp.newTrigger("onEditBy").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
ScriptApp.newTrigger("recall").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
ScriptApp.newTrigger("getFormData").timeBased().everyMinutes(5).create()
SpreadsheetApp.getActiveSpreadsheet().toast("Trigger Set","Notification",3);
}

function deleteTrigger(){
  ScriptApp.getScriptTriggers().forEach(function(trigger){
    ScriptApp.deleteTrigger(trigger);
    SpreadsheetApp.getActiveSpreadsheet().toast("Trigger Deleted","Notification",2);
    });
}

function recall(e){
var sourceCell = customPdf.getRange("H36");
var targetRange = customPdf.getRange("C16:I35");
sourceCell.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
var setData = customPdf.getRange("D16:D35").getValues().filter(e=>e[0] !='').flat()
customPdf.showRows(16 , 35)
customPdf.getRange("C16:I35").clearContent()
SpreadsheetApp.flush()
var twoDArray = []
var sheetName = "Custom PDF";
var range = e.range;
var row = e.range.getRow();
var col = e.range.getColumn();
if (sheetName === e.source.getActiveSheet().getName()  && col === 12 && row === 2){
var cName = customPdf.getRange('L2').getValue()
var clientData = Master.getRange(2,1,Master.getLastRow(),12).getValues().filter(f=>f[0] == cName);
customPdf.getRange("B8").setValue(clientData[0][3])

// Logger.log(clientData)
const obj = JSON.parse(clientData[0][7])
for (var i = 0; i < obj.length; i++) {
  var crow = [];
  crow.push(obj[i].row1);
  crow.push(obj[i].row2);
  crow.push(obj[i].row3);
  crow.push(obj[i].row4);
  crow.push(obj[i].row5);
  crow.push(obj[i].row6);
  crow.push(obj[i].row7);
  twoDArray.push(crow);
}

// Logger.log(twoDArray)


var uniqueElements = Array.from(new Set(twoDArray.map(subarray => subarray[0])));
Logger.log(uniqueElements);


var uniqlen = [];
for (let q = 0; q < uniqueElements.length; q++) {
  var f = twoDArray.filter(e => e[0] === uniqueElements[q]);
  // Logger.log(`Filtered for ${uniqueElements[q]}:`, f); 
  uniqlen.push(f.length);
}

Logger.log(uniqlen);

var mergeRow = 16;
  for (var r = 0; r < uniqlen.length; r++){
if(uniqlen[r] > 1){
customPdf.getRange(mergeRow,3,uniqlen[r],1).merge();
mergeRow = mergeRow + uniqlen[r]
}else{
  mergeRow = mergeRow +1
}
Logger.log(uniqlen[r])
}
customPdf.getRange(16,3,twoDArray.length,7).setValues(twoDArray)
customPdf.getRange("C11").setValue(clientData[0][3])
customPdf.getRange("B8").setValue(clientData[0][8])
customPdf.getRange("L3").setValue(clientData[0][10])
customPdf.getRange("L4").setValue(clientData[0][11])

var tda = twoDArray.filter(e=>e[0]!="").length

customPdf.hideRows(16+tda,35-16-tda)
}else {
  customPdf.showRows(16 , 35)
}
}

function customPDFF(){
  var url = getPDF("Custom PDF",62)
  const invNo = setting.getRange("E2").getValue();
  const count = setting.getRange("E3").getValue() ;
  const senderName = setting.getRange("B6").getValue() ;
  const recipientEmail = setting.getRange("B7").getValue();
  const nameIN = setting.getRange("H2").getValue() ;
  const ccEmail = setting.getRange("H4").getValue();
  const bccEmail = setting.getRange("H5").getValue();
  const subject = setting.getRange("H6").getValue();
  const body = setting.getRange("H8").getValue()
  const replyTo = setting.getRange("H3").getValue();
  const mergedBody = body.replace("[Recipient]", recipientEmail).replace("[Sender]", senderName);
  var file = DriveApp.getFileById(url[1]);
  var ivn = invNo+count;
  const cName = customPdf.getRange("C11").getValue()
  const oldPDFID = customPdf.getRange("L2").getValue()
  const cEmail = customPdf.getRange("L3").getValue()
  const cNumber = customPdf.getRange("L4").getValue()
  const cName1 = customPdf.getRange("B8").getValue()

  // GmailApp.sendEmail(recipientEmail, subject, mergedBody, {
  // attachments: [file.getAs(MimeType.PDF)],
  // name: nameIN,
  // cc: ccEmail,
  // bcc:bccEmail,
  // replyTo:replyTo,
  // });
  
  // const recipientNo = setting.getRange("B8").getValue();
  // const whatsAppMsg = setting.getRange("H9").getValue();
  // const api = setting.getRange("B4").getValue();

  // WhatsappAutomation.sendMessage(recipientNo,whatsAppMsg,url[0],[api])

var datatoJSON = customPdf.getRange("C16:I35").getValues().map(row=>{
return{
row1 :row[0],
row2 :row[1],
row3 :row[2],
row4 :row[3],
row5 :row[4],
row6 :row[5],
row7 :row[6],
row8 :row[7],
row9 :row[8],
row10 :row[9],
row11 :row[10],
row12 :row[11],
row13 :row[12],
row14 :row[13],
row15 :row[14],
row16 :row[15],
row17 :row[16],
row18 :row[17],
row19 :row[18],
row20 :row[19]
}
})

var jsonString = JSON.stringify(datatoJSON, null, 2);
setting.getRange("E3").setValue(count + 1);

var idData = Master.getRange("A2:L").getValues().filter(e=>e[0] == oldPDFID).flat()
Logger.log(idData)
var countOf = Master.getRange("I2:I").getValues().filter(e=>e[0] == cName1).length;
Logger.log(countOf)
var updateID = oldPDFID+"."+countOf
var lastRow = Master.getRange("A1:A").getValues().filter(e => e[0] != "")
Master.getRange(lastRow.length + 1, 1, 1, 13).setValues([[ivn,idData[1],idData[2],idData[3],url[0],"",url[1],jsonString,cName1,"","","",updateID]])
}
