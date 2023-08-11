function getFormData() {
  var rateCardSheet = SpreadsheetApp.openById("1UfAdZyUVDRkEsUwmbmdPFJ08rm0wLv1VwWEp9MfnD7Y")
  var sample = rateCardSheet.getSheetByName("Sample")
  var data = formR.getRange(2, 2, formR.getLastRow() - 1, 19).getValues()
  // .filter(e=>e[0] != "" & e[18] =="")

  var dl = data.length
  for (i = 0; i < dl; i++) {
    try {
      if (data[i][0] != "" && data[i][18] == "") {
        var setData = []

        pdf.getRange(16, 3, 20, 6).clearContent()
        pdf.getRange("C11").clearContent()
        setting.getRange("B3").clearContent()
        setting.getRange("B3").setValue(data[i][3])

        var otherD = [data[i][8], data[i][9], data[i][10], data[i][11], data[i][12], data[i][13], data[i][14], data[i][15], data[i][16]]

        var manuFacoring = data[i][8].split("\n")
        var print = data[i][9].split("\n")
        var social = data[i][11].split("\n")
        var digital = data[i][13].split("\n")
        var digitalTime = data[i][14]
        var eventType = data[i][15]
        var whichEvent = data[i][16].split("\n")

        for (j = 0; j < manuFacoring.length; j++) {
          Logger.log(manuFacoring[j])
          try {
            for (k = 0; k < print.length; k++) {
              var rateData1 = sample.getRange(2, 1, sample.getLastRow() - 1, 6).getValues().filter(e => e[0] == manuFacoring[j] && e[2] == print[k])
              rateData1[k].push(otherD[2])
              setData.push(rateData1[k])
              // Logger.log('print')
              // Logger.log(setData)
            }
          } catch { }

          try {
            for (n = 0; n < social.length; n++) {
              var rateData2 = sample.getRange(2, 1, sample.getLastRow() - 1, 6).getValues().filter(e => e[0] == manuFacoring[j] && e[2] == social[n])
              // Logger.log('Social')              
              // Logger.log(rateData2)              
              rateData2[n].push(otherD[4])
              setData.push(rateData2[n])
            }
          } catch { }

          try {
            for (m = 0; m < digital.length; m++) {
              var rateData3 = sample.getRange(2, 1, sample.getLastRow() - 1, 6).getValues().filter(e => e[0] == manuFacoring[j] && e[2] == digital[m] && e[4] == digitalTime)
              rateData3[m].push(otherD[6])
              // Logger.log(rateData3)
              // Logger.log('Digital')
              setData.push(rateData3[m])
            }
          } catch { }

          try {
            for (o = 0; o < whichEvent.length; o++) {
              var rateData4 = sample.getRange(2, 1, sample.getLastRow() - 1, 6).getValues().filter(e => e[0] == manuFacoring[j] && e[2] == eventType && e[4] == whichEvent[o])
              rateData4[o].push(whichEvent[o])
              // Logger.log(rateData4[o])
              setData.push(rateData4[o])
            }

          } catch { }

        }
        // Logger.log(setData)
        for (z = 0; z < setData.length; z++) {
          setData[z].shift()
        }

        pdf.getRange("C11").setValue(data[i][0])


        setData.sort(function (a, b) {
          return a[0].localeCompare(b[0]);
        });

        var jsonData = setData

        var uniqueElements = Array.from(new Set(setData.map(subarray => subarray[0])));
        Logger.log(uniqueElements);

        Logger.log(setData);

        var uniqlen = [];
        for (let q = 0; q < uniqueElements.length; q++) {
          var f = setData.filter(e => e[0] === uniqueElements[q]);
          // Logger.log(`Filtered for ${uniqueElements[q]}:`, f); 
          uniqlen.push(f.length);
        }

        Logger.log(uniqlen);

        var mergeRow = 16;
        for (var r = 0; r < uniqlen.length; r++) {
          if (uniqlen[r] > 1) {
            pdf.getRange(mergeRow, 3, uniqlen[r], 1).merge();
            mergeRow = mergeRow + uniqlen[r]
          } else {
            mergeRow = mergeRow + 1
          }
          Logger.log(uniqlen[r])
        }
        pdf.getRange(16, 3, setData.length, 6).setValues(setData)
        pdf.hideRows(16 + setData.length, 35 - 16 - setData.length)

        SpreadsheetApp.flush()
        var url = getPDF("PDF", 62)
        const invNo = setting.getRange("E2").getValue();
        const count = setting.getRange("E3").getValue();
        const senderName = setting.getRange("B6").getValue();
        const recipientEmail = setting.getRange("B7").getValue();
        const nameIN = setting.getRange("H2").getValue();
        const ccEmail = setting.getRange("H4").getValue();
        const bccEmail = setting.getRange("H5").getValue();
        const subject = setting.getRange("H6").getValue();
        const body = setting.getRange("H8").getValue()
        const replyTo = setting.getRange("H3").getValue();
        const mergedBody = body.replace("[Sender]", senderName);

        const recipientNo = setting.getRange("B8").getValue();
        const whatsAppMsg = setting.getRange("H9").getValue();
        const newMSG = whatsAppMsg.replace("[Sender]", senderName)
        const api = setting.getRange("B4").getValue();
        var ivn = invNo + count;

        let phoneNumber = data[i][2];
        let numericPart = phoneNumber.replace(/\D/g, '');

        var price = pdf.getRange("I16:I35").getValues().filter(e => e[0] != "").flat()

        for (js = 0; js < jsonData.length; js++) {
          jsonData[js].push(price[js])
        }

        var datatoJSON = jsonData.map(row => {
          return {
            row1: row[0],
            row2: row[1],
            row3: row[2],
            row4: row[3],
            row5: row[4],
            row6: row[5],
            row7: row[6]
          }
        })


        var jsonString = JSON.stringify(datatoJSON, null, 2);
        var lastRow = Master.getRange("A1:A").getValues().filter(e => e[0] != "")
        // Logger.log(lastRow.length);
        Master.getRange(lastRow.length + 1, 1, 1, 9).setValues([[ivn, data[i][1], numericPart, data[i][0], url[0], "", url[1], jsonString, data[i][3]]])
        setting.getRange("E3").setValue(count + 1);

        SpreadsheetApp.flush()
        send([recipientEmail, subject, mergedBody, url[1], nameIN, ccEmail, bccEmail, replyTo, recipientNo, newMSG, url[0], api])

        formR.getRange(i + 2, 20).setValue('Sent')
        pdf.getRange(16, 3, 20, 6).clearContent()
        pdf.getRange("C11").clearContent()
        setting.getRange("B3").clearContent()
        pdf.showRows(16 + setData.length, 35 - 16 - setData.length)
        var sourceCell = pdf.getRange("H36"); // Replace "A1" with the actual source cell you want to copy from
        var targetRange = pdf.getRange("C16:C35");
        sourceCell.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
    }
    catch (e) {
      Logger.log(e)
    }
  }
}

function onEditBy(e) {
  var sheetName = "Master";
  var range = e.range;
  var row = e.range.getRow();
  var col = e.range.getColumn();

  var statusData = Master.getRange(row, 1, 1, 12).getValues()

  var email = Master.getRange(row, 11, 1, 1).getValue()
  var number = Master.getRange(row, 12, 1, 1).getValue()
  var name = Master.getRange(row, 4, 10, 1).getValue()
  var url = Master.getRange(row, 5, 1, 1).getValue()
  var fileid = Master.getRange(row, 7, 1, 1).getValue()
  const nameIN = setting.getRange("H2").getValue();
  const ccEmail = setting.getRange("H4").getValue();
  const bccEmail = setting.getRange("H5").getValue();
  const subject = setting.getRange("E7").getValue();
  const body = setting.getRange("E8").getValue()
  const replyTo = setting.getRange("H3").getValue();
  const mergedBody = body.replace("[Sender]", name);

  var date = new Date()
  const whatsAppMsg = setting.getRange("E9").getValue();
  const newMSG = whatsAppMsg.replace("[Sender]", name)

  const api = setting.getRange("B4").getValue();


  if (sheetName === e.source.getActiveSheet().getName() && col === 6 && range.getValue() == "Approved" && e.source.getActiveSheet().getRange(row,10).getValue() != "") {
    Logger.log('Send')  
    statusData[0].unshift(date)
    approved.getRange(approved.getLastRow() + 1, 1, 1, 13).setValues(statusData)

    send([email, subject, mergedBody, fileid, nameIN, ccEmail, bccEmail, replyTo, number, newMSG, url, api])

  } else if (sheetName === e.source.getActiveSheet().getName() && col === 6 && range.getValue() == "Disapproved"){

    var fileid = Master.getRange(row, 7, 1, 1).getValue()
    var rEmail = setting.getRange("B10").getValue()
    var rNumber = setting.getRange("B11").getValue()
    var rSubject = setting.getRange("B12").getValue()
    var rEmailBody = setting.getRange("B13").getValue()
    var rWhatsappBody = setting.getRange("B14").getValue()
    var invID = Master.getRange(row, 1, 1, 1).getValue()
    var fileURL = Master.getRange(row, 5, 1, 1).getValue()
    var updateEBody = rEmailBody.replace("<<invID>>", invID)
    var updateWBody = rWhatsappBody.replace("<<invID>>", invID)
    var date = new Date()
    statusData[0].unshift(date)
    disapproved.getRange(disapproved.getLastRow() + 1, 1, 1, 13).setValues(statusData)
    send([rEmail, rSubject, updateEBody, fileid, nameIN, ccEmail, bccEmail, replyTo, rNumber, updateWBody, fileURL, api])
  }

if("Approved"  == e.source.getActiveSheet().getName() && col === 16 && range.getValue() == "Send" ){
var date = new Date()
const custReply = setting.getRange("B17").getValue();
const custCC = setting.getRange("B18").getValue();
  const custSubject = setting.getRange("B19").getValue();
  const custEmailBody = setting.getRange("B20").getValue();
 const custWhatsappBody = setting.getRange("B21").getValue();
var custData = approved.getRange(row,1, 1, 13).getValues().flat()


  send([custData[2],custSubject,custEmailBody,custData[7],"",custCC,"",custReply,custData[3],custWhatsappBody,custData[5], api])
  approved.getRange(row,14,1,1).setValue(date);
}











}



