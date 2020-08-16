var sheetName = 'Main';
var scriptProp = PropertiesService.getScriptProperties()


function intialSetup () {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  scriptProp.setProperty('key', activeSpreadsheet.getId())
}



function createNewSheet (title){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var create = ss.insertSheet(title);
  var newSheet = SpreadsheetApp.getActive().getSheetByName(title);
  newSheet.getRange('A1').setValue('Hello3');
}

function editSheet(){
  var sheet = SpreadsheetApp.getActive().getSheetByName('Test Sheet');
  sheet.getRange('A1').setValue('Hello');
  

}
//
//function groupSheet(){
//  var sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
//  var data = sheet.getDataRange().getValues();
//  var lastRow = sheet.getLastRow();
//  var lastColumn = sheet.getLastColumn();
//  var newData = [];
//  var match = false;
//  for (var i in data) {
//    var row = data[i];
//    for (var j in newData) {
//      if((row[2] == newData[j][2]) && (row[3] == newData[j][3]) ){
//        match = true;
//        Logger.log("Match found")
//        var title = row[2] + " - " + row[3];
//        var getSheet = SpreadsheetApp.getActive().getSheetByName(title);
//        if(!getSheet){
//          Logger.log("New Sheet") 
////          sheet.getRange(i,95).setValue("Yes");
////          sheet.getRange(j,95).setValue("Yes");
//          newData[j][94] = "yes";    
//          data[i][94] = "yes";
//          sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
//          sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
//
//          var ss = SpreadsheetApp.getActiveSpreadsheet();
//          var create = ss.insertSheet(title);
//          var newSheet = SpreadsheetApp.getActive().getSheetByName(title);
//          newSheet.getRange('A1').setValue('Hello3');
//          var headerVals = sheet.getRange("A1:CQ1").getValues();
//          newSheet.getRange("A1:CQ1").setValues(headerVals);
//          newSheet.appendRow(newData[j]);
//          newSheet.appendRow(row);
//
//
//
//          
//        }
//        else{
//          Logger.log("Already exists");
//          var getSheet2 = SpreadsheetApp.getActive().getSheetByName(title);
//          if(!(newData[j][94]=="yes")){
////            newData[j][94] = "yes";    
////            data[i][94] = "yes";
////            getSheet2.appendRow(row);
//            Logger.log("Not in a group")
//            
//          }
//          else{
//            Logger.log("In a group")
//          }
//
//        }
//      }
//      else{
//       Logger.log("No match")
//       match = false;
//       newData.push(row)
//
//      }
//    }
//    if(!match){
//      newData.push(row)
//
//    }
////    if (!match) {
////      var title = row[2] + " - " + row[3];
////      var ss = SpreadsheetApp.getActiveSpreadsheet();
////      var s = ss.getActiveSheet();
////      var create = ss.insertSheet(title);
////      var newSheet = SpreadsheetApp.getActive().getSheetByName(title);
////      newSheet.getRange('A1').setValue('Hello3');
////      newData.push(row);
////      newSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
////
////    }
//  }
////  sheet.clearContents();
////  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
//}

function groupRowNum(groupID){
  var group_sheet = SpreadsheetApp.getActive().getSheetByName("Groups");
  var startRow = 2; // First row of data to process
  var lastRowGrp = group_sheet.getLastRow();

  var group_id = group_sheet.getRange('B'+startRow+':B'+lastRowGrp);
  var data = group_id.getValues();
//  Logger.log("IDs in Sheet : "+data)
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    if(row == groupID){ //[1] because column B
//      Logger.log("Index of row = "+(i+2));
      return i+2;
    }
    else{
      
    }
//    Logger.log(data)

  }
}

function checkGroupExists(groupID){
  var group_sheet = SpreadsheetApp.getActive().getSheetByName("Groups");
  var startRow = 2; // First row of data to process
  var lastRowGrp = group_sheet.getLastRow();
  var group_id = group_sheet.getRange('B'+startRow+':B'+lastRowGrp);
  var data = group_id.getValues();
  var returnVal = false;
//  Logger.log(data);
  for(var i=0; i<data.length; i++){
    returnVal = false;
//    Logger.log(data[i][0].includes(groupID));
    if(data[i][0].includes(groupID)){
      returnVal = true;
      break;
    }
//    else{
//      returnVal = false;
//    }
  }
//  Logger.log(returnVal)
  return returnVal;
}

function checkBoardingExists(groupID){
  var group_sheet = SpreadsheetApp.getActive().getSheetByName("Boarding House");
  var startRow = 2; // First row of data to process
  var lastRowGrp = group_sheet.getLastRow();
  var group_id = group_sheet.getRange('B'+startRow+':B'+lastRowGrp);
  var data = group_id.getValues();
  var returnVal = false;
//  Logger.log(data);
  for(var i=0; i<data.length; i++){
    returnVal = false;
//    Logger.log(data[i][0].includes(groupID));
    if(data[i][0].includes(groupID)){
      returnVal = true;
      break;
    }
//    else{
//      returnVal = false;
//    }
  }
//  Logger.log(returnVal)
  return returnVal;
}


function groupSheet(){
  var groupLeadName = "";
  var propName = "";
  var main_sheet = SpreadsheetApp.getActive().getSheetByName("Main");
  var group_sheet = SpreadsheetApp.getActive().getSheetByName("Groups");
  var startRow = 2; // First row of data to process
  var lastRow = main_sheet.getLastRow();
  var lastColumn = main_sheet.getLastColumn();
  var lastRowGrp = group_sheet.getLastRow();
  var lastColumnGrp = group_sheet.getLastColumn();
  var currentRow = [];
  var groupRow = [];
  var in_a_group = "yes";
  var dataRange = main_sheet.getRange(startRow, 1, lastRow-1, lastColumn);
  // Fetch values for each row in the Range.
  var group_id = group_sheet.getRange('B'+startRow+':B'+lastRowGrp);
  var data = dataRange.getValues();
  var counter = 0;
  for (var i = 0; i < data.length; ++i) {
    groupRow = [];
    var row = data[i];
    currentRow = row;
    //    Logger.log(currentRow);
    groupLeadName = row[3];
    propName = row[2];
    var added_to_grp = row[94];
    var groupID = row[2]+"-"+row[3];
    Logger.log("Check for "+groupID+"row in main sheet")
//    Logger.log("Group= "+ groupID);
    
    if(checkGroupExists(groupID)){
      groupRow = [];
      var rowNum = groupRowNum(groupID);
      Logger.log("Group Found in Row Num: " +rowNum );
      var member1= group_sheet.getRange('X'+rowNum).getValue();
      var member2= group_sheet.getRange('Z'+rowNum).getValue();
      var member3= group_sheet.getRange('AB'+rowNum).getValue();
      var member4= group_sheet.getRange('AD'+rowNum).getValue();
      var member5= group_sheet.getRange('AF'+rowNum).getValue();
      if(!(added_to_grp==in_a_group)){
        if(!(member1=="")){
          Logger.log("Member1 exists");
          if(!(member2=="")){
            Logger.log("Member2 exists");
            if(!(member3=="")){
              Logger.log("Member3 exists");
              if(!(member4=="")){
                Logger.log("Member4 exists");
                if(!(member5=="")){
                  Logger.log("This Group is complete");
                }
                else{
                  Logger.log("Member5 doesn't exist");
                  group_sheet.getRange('AF'+rowNum).setValue(row[7]+" "+row[8]);//member name
                  group_sheet.getRange('AG'+rowNum).setValue(row[10]);//member email

                  group_sheet.getRange('AX'+rowNum).setValue(row[40]+" "+row[41]);//parent 1 name
                  group_sheet.getRange('AY'+rowNum).setValue(row[48]);//parent 1 email

                  group_sheet.getRange('AZ'+rowNum).setValue(row[51]+" "+row[52]);//parent 2 name
                  group_sheet.getRange('BA'+rowNum).setValue(row[59]);//parent 2 email

                  main_sheet.getRange('CQ'+(i+2)).setValue("yes");

                }
              }
              else{
                Logger.log("Member4 doesn't exist");
                group_sheet.getRange('AD'+rowNum).setValue(row[7]+" "+row[8]);//member name
                group_sheet.getRange('AE'+rowNum).setValue(row[10]);//member email

                group_sheet.getRange('AT'+rowNum).setValue(row[40]+" "+row[41]);//parent 1 name
                group_sheet.getRange('AU'+rowNum).setValue(row[48]);//parent 1 email

                group_sheet.getRange('AV'+rowNum).setValue(row[51]+" "+row[52]);//parent 2 name
                group_sheet.getRange('AW'+rowNum).setValue(row[59]);//parent 2 email

                main_sheet.getRange('CQ'+(i+2)).setValue("yes");
              }
            }
            else{
              Logger.log("Member3 doesn't exist");
              group_sheet.getRange('AB'+rowNum).setValue(row[7]+" "+row[8]);//member name
              group_sheet.getRange('AC'+rowNum).setValue(row[10]);//member email

              group_sheet.getRange('AP'+rowNum).setValue(row[40]+" "+row[41]);//parent 1 name
              group_sheet.getRange('AQ'+rowNum).setValue(row[48]);//parent 1 email

              group_sheet.getRange('AR'+rowNum).setValue(row[51]+" "+row[52]);//parent 2 name
              group_sheet.getRange('AS'+rowNum).setValue(row[59]);//parent 2 email

              main_sheet.getRange('CQ'+(i+2)).setValue("yes");
            }
          } 
          else{
            Logger.log("Member2 doesn't exist");
            group_sheet.getRange('Z'+rowNum).setValue(row[7]+" "+row[8]);//member name
            group_sheet.getRange('AA'+rowNum).setValue(row[10]);//member email

            group_sheet.getRange('AL'+rowNum).setValue(row[40]+" "+row[41]);//parent 1 name
            group_sheet.getRange('AM'+rowNum).setValue(row[48]);//parent 1 email

            group_sheet.getRange('AN'+rowNum).setValue(row[51]+" "+row[52]);//parent 2 name
            group_sheet.getRange('AO'+rowNum).setValue(row[59]);//parent 2 email

            main_sheet.getRange('CQ'+(i+2)).setValue("yes");
          }
        }
        else{
          Logger.log("Member1 doesn't exist");
          group_sheet.getRange('X'+rowNum).setValue(row[7]+" "+row[8]);//member name
          group_sheet.getRange('Y'+rowNum).setValue(row[10]);//member email

          group_sheet.getRange('AH'+rowNum).setValue(row[40]+" "+row[41]);//parent 1 name
          group_sheet.getRange('AI'+rowNum).setValue(row[48]);//parent 1 email

          group_sheet.getRange('AJ'+rowNum).setValue(row[51]+" "+row[52]);//parent 2 name
          group_sheet.getRange('AK'+rowNum).setValue(row[59]);//parent 2 email

          main_sheet.getRange('CQ'+(i+2)).setValue("yes");
        }        
      }
    
    }
    else if(!(checkGroupExists(groupID)) && (row[4]=="Group House") && (row[91]=="I Agree") && (row[92]=="I Agree")){
      main_sheet.getRange('CQ'+(i+2)).setValue("yes");
      groupRow = [];
      Logger.log("Group Not Found");
      Logger.log("Creating New Group Row for group "+ groupID);
      groupRow[1] = row[2]+"-"+row[3];
      Logger.log(groupRow[1]);
      groupRow[2] = row[2];
      Logger.log(groupRow[2]);
      groupRow[3] = row[3];
      Logger.log(groupRow[3]);      
      groupRow[23] = row[7]+" "+row[8];//member name
      groupRow[24] = row[10];//member email
      
      groupRow[33] = row[40]+" "+row[41];//parent1 full name
      groupRow[34] = row[48];//parent1 email

      groupRow[35] = row[51]+" "+row[52];//parent2 full name
      groupRow[36] = row[59];//parent2 email

      group_sheet.appendRow(groupRow);
      
    } 
    
//    else if(!(checkGroupExists(groupID)) && (row[4]=="Group House")){
//      main_sheet.getRange('CQ'+(i+2)).setValue("yes");
//      groupRow = [];
//      Logger.log("Group Not Found");
//      Logger.log("Creating New Group Row for group "+ groupID);
//      groupRow[1] = row[2]+"-"+row[3];
//      Logger.log(groupRow[1]);
//      groupRow[2] = row[2];
//      Logger.log(groupRow[2]);
//      groupRow[3] = row[3];
//      Logger.log(groupRow[3]);      
//      groupRow[12] = row[7]+" "+row[8];
//      groupRow[17] = row[40]+" "+row[41];//parent1 full name
//      groupRow[18] = row[51]+" "+row[52];//parent2 full name
//      group_sheet.appendRow(groupRow);
//      
//    } 
  }
  
  
  
}

function transferToBoardingSheet(){
  var main_sheet = SpreadsheetApp.getActive().getSheetByName("Main");
  var group_sheet = SpreadsheetApp.getActive().getSheetByName("Boarding House ");
  var startRow = 2; // First row of data to process
  var lastRow = main_sheet.getLastRow();
  var lastColumn = main_sheet.getLastColumn();
  var lastRowGrp = group_sheet.getLastRow();
  var lastColumnGrp = group_sheet.getLastColumn();
  var groupRow = [];
  var in_a_group = "Yes";
  var dataRange = main_sheet.getRange(startRow, 1, lastRow-1, lastColumn);
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    groupRow = [];
    var row = data[i];
    var added_to_grp = row[94];
    if((row[4]=="Boarding House") && !(row[94]=="Yes")){
      main_sheet.getRange('CQ'+(i+2)).setValue("Yes");
      group_sheet.appendRow(row);

    
    }
  }


}

function mergeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if(row[0] == newData[j][0]){
        duplicate = true;
        if(!(row[91]=="")){
          newData[j][91]=(row[91]);        
        }
        else if(!(row[92]=="")){
          newData[j][92]=(row[92]);          
        }
      
//        Logger.log(newData[j][87])
//        Logger.log(row[87])

      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function sendEmails() {
  var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
  var sheet = doc.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  Logger.log(lastRow);
  Logger.log(lastColumn);
  var dataRange = sheet.getRange(lastRow, 1, 1 , lastColumn);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
//  var idNum = idRange.getValues();
  Logger.log(data)
  for (var i in data) {
    var row = data[i];
    var sentCheck = row[93];
//    var idData = idRange.setValue(generateID());
    if(!(sentCheck =="Yes")){
      var applicantName = row[7]+" "+row[8];
      Logger.log("Applicant name is: "+applicantName);
      var applicantID = row[0];
      Logger.log("Applicant ID is: "+applicantID);
      var propertyNum = row[2];
      
      var applicantMessage = "We recieved your submission! We will Get back to you soon.";
      var applicantSubject = 'Rental Application';
      var parent1Name = row[40]+" "+row[41];
      Logger.log("Parent 1 Name : " + parent1Name);
      var parent1Email = row[48];
      Logger.log("parent 1 email: " + parent1Email);
      var parent1HtmlMsg = "<p>Dear "+parent1Name+",</p>"+"<p>This email confirms that "+applicantName+" has listed you as a guarantor for their College Housing application for " +propertyNum+" . As a guarantor you are financially responsible for any outstanding payments the applicant may fail to pay under the terms of the lease agreement.</p> <p>To confirm yourself as a guarantor, please use this ID: "+applicantID+" and follow this: <a href='http://parent1.rentalapp.collegehousing.us/'>http://parent1.rentalapp.collegehousing.us/</a>. </p>"+ "<p>If you do not want to be a guarantor for "+applicantName+" or if you believe you have received this email in error. Please let us know at contact@collegehousing.us. </p>"+"<p>Thank you,</p>"+"<div><b>College Housing Management</b></div><div>College Housing LLC </div><div>410.680.2868 | 240.367.9669</div><div>10075-10 Tyler Ct. Ijamsville, MD 21754 </div><div>http://www.collegehousing.us</div>";
      Logger.log(parent1HtmlMsg);
      if(!(parent1Email=="")){
        MailApp.sendEmail(parent1Email, applicantSubject , "", {htmlBody: parent1HtmlMsg,name: "College Housing"});    
        //      row[89]=is_sent;
//        sheet.getRange(lastRow,lastColumn).setValue("Yes");
        sheet.getRange('CP'+lastRow).setValue("Yes");
      }
      var parent2Name = row[51]+" "+row[52];
      Logger.log("Parent 2 Name :" + parent2Name);
      var parent2Email = row[59];
      Logger.log(parent2Email);
      var parent2HtmlMsg = "<p>Dear "+parent2Name+",</p>"+"<p>This email confirms that "+applicantName+" has listed you as a guarantor for their College Housing application for " +propertyNum+" . As a guarantor you are financially responsible for any outstanding payments the applicant may fail to pay under the terms of the lease agreement.</p> <p>To confirm yourself as a guarantor, please use this ID: "+applicantID+" and follow this: <a href='http://parent2.rentalapp.collegehousing.us/'>http://parent2.rentalapp.collegehousing.us/</a> </p>"+ "<p>If you do not want to be a guarantor for "+applicantName+" or if you believe you have received this email in error. Please let us know at contact@collegehousing.us. </p>"+"<p>Thank you,</p>"+"<div><b>College Housing Management</b></div><div>College Housing LLC </div><div>410.680.2868 | 240.367.9669</div><div>10075-10 Tyler Ct. Ijamsville, MD 21754 </div><div>http://www.collegehousing.us</div>";
      if(!(parent2Email=="")){
        MailApp.sendEmail(parent2Email, applicantSubject , "", {htmlBody: parent2HtmlMsg,name: "College Housing"});    
//        sheet.getRange(lastRow,lastColumn).setValue("Yes");
        sheet.getRange('CP'+lastRow).setValue("Yes");
      }
    }
  }
}

//function setEmailToYes(){
//  var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
//  var sheet = doc.getSheetByName(sheetName);
//  var lastRow = sheet.getLastRow();
//  var lastColumn = sheet.getLastColumn();
//  Logger.log(lastRow);
//  Logger.log(lastColumn);
//  var dataRange = sheet.getRange(lastRow, 1, 1 , lastColumn);
//  // Fetch values for each row in the Range.
//  var data = dataRange.getValues();
//  for (var i in data) {
//    
////    sheet.getRange(lastRow,lastColumn).setValue("Yes");
//  }
//
//}

function sendEmailIfNoGuarantors(){
  var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
  var sheet = doc.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  Logger.log(lastRow);
  Logger.log(lastColumn);
  var dataRange = sheet.getRange(lastRow, 1, 1 , lastColumn);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
//  var idNum = idRange.getValues();
  Logger.log(data)
  for (var i in data) {
    var row = data[i];
    var checkIfProblemGuarantors= row[37];
    var applicantID = row[0];
    var applicantName = row[7]+" "+row[8];
    var applicantEmail = row[10];
    Logger.log("Applicant Id : "+ applicantID);
    Logger.log("Applicant Name : "+ applicantName);
    Logger.log("Applicant Email : "+ applicantEmail);
    Logger.log("Applicant Issue : "+ checkIfProblemGuarantors);
    var htmlMsg ="<p>An application submitted have problems with both guarantors signing</p><p>Applicant ID:"+applicantID+"</p><p>Applicant Name: "+applicantName+" </p><p>Applicant Email: "+applicantEmail+"</p><p>Spreadsheet Link: <a href='https://docs.google.com/spreadsheets/d/1JjaYhKijzTtCvmkGN3PWpPacPq_n1X_LsGTL-7BD180/edit#gid=0'>https://docs.google.com/spreadsheets/d/1JjaYhKijzTtCvmkGN3PWpPacPq_n1X_LsGTL-7BD180/edit#gid=0</a> </p>";    
    if(!(checkIfProblemGuarantors =="No")){
       MailApp.sendEmail("contact@collegehousing.us", "Pending Application" , "", {htmlBody: htmlMsg, name:"Rental Application"});    
    }   
  }
}

function doPost (e) {
  var lock = LockService.getScriptLock()
  lock.tryLock(10000)
//  mergeDuplicates();

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'))
    var sheet = doc.getSheetByName(sheetName)

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    var nextRow = sheet.getLastRow() + 1
    Logger.log(nextRow);

    var newRow = headers.map(function(header) {
      return header === 'timestamp' ? new Date() : e.parameter[header]
    })

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow])
    Logger.log(newRow)
    mergeDuplicates();    
    sendEmails();
//    mergeDuplicates();
    sendEmailIfNoGuarantors();

    
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
