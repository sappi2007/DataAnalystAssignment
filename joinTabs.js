/*  
Google App script 
Name: joinTabs
Version : v1.1
Create on : 04/12/2024
Last updated : 04/12/2024
Author BP Swapna
Functionality : Joins Source Data 1, Source Data 2 and Source Data 3 based on Slotid and hostpartnerid.
Creates separte tabs based on matching records of Hostpartnername. 
Checks for tab exists input validation on second iteration.
Font is default for data. 
Script requires    Joins Source Data 1, Source Data 2 and Source Data 3 tabs with certain data format. Otherwise it raises standard exception.
Execution : Create Appscript and Save from Sheets. Click Externtions-->AppScripts and copy past the following code. Save it and run from same AppScript IDE or run form Sheets-->Externsions-->Macros.
You need to import macro joinTabs before run. */

function joinTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceData1 = ss.getSheetByName('Source Data 1');
  var sourceData2 = ss.getSheetByName('Source Data 2');
  var sourceData3 = ss.getSheetByName('Source Data 3');

  var values1 = sourceData1.getDataRange().getValues();
  var values2 = sourceData2.getDataRange().getValues();
  var values3 = sourceData3.getDataRange().getValues();

  values1.shift();
  values2.shift();
  values3.shift();

  //Joining SD1 and SD2 based on slotid 
  var joinedData = [];
  values1.forEach(function(row1) {
    let slotId1 = row1[0];
    let matchingRow2 = values2.find(function(row2) {
      return row2[0] == slotId1;
    });
    if (matchingRow2) {
      joinedData.push([
        row1[0],  // SlotId
        row1[1],  // HostPartenerName
        row1[2],  // Program
        row1[3],  // Term Start (renamed to Term Begin)
        row1[4],  // Term End
        row1[5],  // Matched with Host Partner
        matchingRow2[4],  // Maryland Leg Dist
        matchingRow2[3],  // County
        matchingRow2[1],  // Gender
        matchingRow2[2]   // Ethnic
      ]);
    }
  });

  
  var dataMap = {};
  joinedData.forEach(function(row) {
    let hostPartenerName = row[1];
    // check if hostPartenerName is present in the dataMap
    if (!dataMap[hostPartenerName]) {
      dataMap[hostPartenerName] = [];
    }
    dataMap[hostPartenerName].push(row);
  });

  Object.keys(dataMap).forEach(function(hostPartenerName) {
    let tabName = hostPartenerName.substring(); 
    let existingSheet = ss.getSheetByName(tabName);
    let sheet;
    if (existingSheet) {
      let response = Browser.msgBox('Tab name "' + tabName + '" already exists. Do you want to overwrite it?', Browser.Buttons.YES_NO);
      if (response == 'yes') {
        ss.deleteSheet(existingSheet);
        sheet = ss.insertSheet(tabName);
      }
    } else {
      sheet = ss.insertSheet(tabName);
    }
    if (sheet) {
      let data = dataMap[hostPartenerName];
      let headers = ['SlotId', 'HostPartenerName', 'Program', 'Term Begin', 'Term End', 'Matched with Host Partner', 'Maryland Leg Dist', 'County', 'Gender', 'Ethnic'];

      //Formatting heading like Hostpartner name,date and other 
      let headerRange = sheet.getRange(9, 1, 1, headers.length);
      headerRange.setValues([headers]).setFontWeight("bold").setBackground("#d3d3d3").setHorizontalAlignment("center");
      headerRange.setWrap(true);
 
      sheet.getRange(11, 1, data.length, headers.length).setValues(data);
      let columnToRemove = headers.indexOf('HostPartenerName') + 1; // Adding 1 to match 1-indexed column numbers in Sheets
      sheet.deleteColumn(columnToRemove);
      sheet.getRange('A1:I1').merge().setFontWeight("bold").setHorizontalAlignment("center");
      let rangeA1toI1 = sheet.getRange('A1:I1');
      rangeA1toI1.setFontFamily("Times New Roman").setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center");
      let rangeA2toI2 = sheet.getRange('A2:I2');
      rangeA2toI2.merge().setFontWeight("bold").setHorizontalAlignment("center");
      let currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
      rangeA2toI2.setValue(currentDate).setFontFamily("Times New Roman").setFontSize(12).setHorizontalAlignment("center");
      let rangeA6toC6 = sheet.getRange('A6:C6');
      rangeA6toC6.setFontWeight("bold").setFontFamily("Times New Roman").setFontSize(12);
      sheet.getRange('A7:I7').merge();
      let rangeA7toI7 = sheet.getRange('A7:I7');
      rangeA7toI7.setFontFamily("Times New Roman");
      let rangeA4toI4 = sheet.getRange('A4:I4');
      rangeA4toI4.setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);

    }
  });

// Find the rows in source3 that match with hostPartenerName in joined data map (SD1 and SD2)
  values3.forEach(function(row) {
    let hostPartenerName = row[1];
    if (dataMap[hostPartenerName]) {
      let tabName = hostPartenerName.substring(0, 25);
      let sheet = ss.getSheetByName(tabName);
      if (sheet) {
        sheet.getRange(1, 1, 1, 1).setValue(row[1]);  // Cell A1
        sheet.getRange(6, 1, 1, 1).setValue(row[0]);  // Cell A6
        //sheet.getRange(6, 2, 1, 1).setValue(row[2]);  // Cell B5
        sheet.getRange(6, 2, 1, 1).setValue(row[3]);  // Cell B6
        sheet.getRange(7, 1, 1, 1).setValue(row[2]);  // Cell A7
        sheet.getRange(6, 3, 1, 1).setValue(row[1]);  // Cell A7
        // Rename tabs with "mission" name
        // sheet.setName(row[2]);
      }
      
    }
  });
}
