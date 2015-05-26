/*
 *  Copyright 2014.9.18  Jing Tang  tangjing725@ccnbg.org
 *  此程序是利用google app script api 和 google drive service 实现的营会分房程序。
 *  在通知原作者的前提下，你可以使用传播以及修改本程序，但禁止用于任何商业用途。
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation; either version 2 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful, but
 *  WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>. 
 *
 *  For more information on using the Spreadsheet API, see
 *  https://developers.google.com/apps-script/service_spreadsheet
 *
 */

//NOTICE
//#Line: 265和411, replace with your email address before you test


function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

function generateGroupOverview() {
  //file: groupOverview ID
  var groupOverviewId = "1hc47xmxLMJ9-TGGLJ6Vsaq651Cg_m0-GqHKghG58z7A";
  
  var groupOverviewSheet = SpreadsheetApp.openById(groupOverviewId).getActiveSheet();
  groupOverviewSheet.clear();
  var registerSheet = SpreadsheetApp.getActiveSheet();
  var regData = registerSheet.getDataRange().getValues();
  
  //build map with key:groupNr, value: arrayOfRows
  var gmap = {};
  for (var i = 1; i <= regData.length - 1; i++) {
    var row = regData[i];
    var name = row[0];
    var group = row[14];
    var isGroupLeader = row[15];
    for(j = 1; j< 32; j++) {
      if(group == j) {
        if(gmap[j] == null){
           gmap[j] = [];
        }
        gmap[j].push(row);
      }
      
    }
  
  }
 
  Logger.log(gmap);
  
  //now write into group overview sheet
  for(n = 1; n< 32; n++) {
    var gmapVal = gmap[n];
    if(!gmapVal){
      Logger.log("end!, no group: "+n);
      return 
    }
    groupOverviewSheet.appendRow(['第'+n+'组','性别','信仰','团契','房间']);
    var rw  =  groupOverviewSheet.getLastRow();
    var header = groupOverviewSheet.getRange(rw, 1,1,5);
    header.setBackground('green');
    Logger.log("j "+rw);
    for(r = 0; r< gmapVal.length; r++){
      var grow = gmapVal[r]
      //Logger.log(grow);
      var gname = grow[0]
      var gender = grow[2]
      var gfaith = grow[3]
      if(gfaith == '否' || gfaith == '') {
        gfaith = '慕道友';
      } else {
        gfaith = '基督徒';
      }
      var gchurch = grow[11]
      var groom = grow[13]
      if(groom == '') {
          groom = '日营'; 
      }
      var gleader = grow[15]
      if(gleader != ''){
          gleader = '组长';
          var gw  =  groupOverviewSheet.getLastRow();
          var glrange = groupOverviewSheet.getRange(gw+1, 1);
          Logger.log("z "+gw)
          glrange.setBackground('yellow');
          
      }
    
       groupOverviewSheet.appendRow([gname,gender,gfaith,gchurch,groom, gleader]);
    }
  
  }
  
}


function copyGroupInfo(){
  clearGroupInfo();
  //file 小龙分组 ID
  var xiaolongDocId = '1qbbQ4PIxAWDrL61N1xx3st9uQvs9Z5089HinlG84plU';
  var xiaolongSheet = SpreadsheetApp.openById(xiaolongDocId).getActiveSheet();
  var groupData = xiaolongSheet.getDataRange().getValues();
  
  var registerSheet = SpreadsheetApp.getActiveSheet();
  var regData = registerSheet.getDataRange().getValues();
  var n = 1;
  for (var i = 1; i <= regData.length - 1; i++) {
    var row = regData[i];

    var name = row[0];
    innerloop: 
    for(var j = n; j <= groupData.length - 1; j++) {
      var grow = groupData[j]
      var gname = grow[0];
      var group = grow[14];
      var isGroupLeader = grow[15];
      Logger.log(i+" - "+j)
      if(name.trim() == gname.trim() && group != '') {
           Logger.log(name+" - "+gname)
           var groupRange = registerSheet.getRange(i+1, 15);
           groupRange.setValue(group);
           if(isGroupLeader != ''){
             var groupLeaderRange = registerSheet.getRange(i+1, 16);
             groupLeaderRange.setValue(isGroupLeader);
           }
           n = j+1;
           break innerloop;
      }
      
    }
  }
}

function appendContentInTableCell(body, cell) {
  for (var i = 0; i < body.getNumChildren(); i++) {
     var element = body.getChild(i).copy();
     var type = element.getType();
     if( type == "PARAGRAPH" ){
         //Logger.log(i+" "+element.getText());
       if(i!=4) {
         cell.appendParagraph(element);
       }
     }
  } 
}

function generateEmptyNamecard() {
 //file newCardTemplate ID 
 var templateDocID = "1jnTdlPyu-ZieOl-3U6Dt_MxVjPR7HV9AUDHxz76_wto";//copyOfTemplate hash
  
 var templateDoc = DocumentApp.openById(templateDocID);
 var templateBody = templateDoc.getBody();
 // generated name card file
  //file: finalCards ID
 var finalCardsId = "1-D-DEgtQRTmzDZ8H8KYQJqAhFh5SQ-34mq-GBbfsP_0";
 var newDoc = DocumentApp.openById(finalCardsId);
   
  var body = newDoc.getBody();
  body.clear();
  body.setMarginTop(10);
  body.setMarginLeft(55);
  body.setMarginRight(20);
  body.setMarginBottom(10);
   
  var temp;
  var copyBody;
  var cell_left;
  var cell_right;
  var ar = [];
  var table = body.appendTable();
  table.setBorderWidth(0.5);

   // parse register form
  for (var i = 1; i <= 120; i++) {
  
   
    //copy template on temp
    Logger.log("copy templateBody to copyBody")
 
    //copyContent(templateBody, copyBody)
    copyBody = templateBody.copy();
    // replace text from data in register form
    copyBody.replaceText('%na%', '姓名:');
    
      
      
      copyBody.replaceText('%ch%', '教会:');
      
     
    copyBody.replaceText('%rmn%', '');
    copyBody.replaceText('%gr%', '');
    copyBody.replaceText('%gl%', '');
     
      //build table , put data in it and insert it in finalCards doc
      if((i+1)%2 == 0){
        Logger.log("left copyBody "+copyBody.getText()+" saved in temp");
        temp = copyBody.copy(); // important move :) body must be first detached then assigned
        Logger.log("hoho: "+temp.getText());

      } else {
        Logger.log("left copyBody: "+temp.getText());
        Logger.log("right copyBody "+copyBody.getText())
        ar.push(i+1);
        var idx = ar.indexOf(i+1);
        Logger.log("build table row "+(i+1)+" and two cells")
        Logger.log("append left and right");
        var tableRow = table.appendTableRow();
        tableRow.appendTableCell('');
        tableRow.appendTableCell('');
        tableRow.setMinimumHeight(150);
        
        cell_left= table.getCell(idx,0);
        cell_right= table.getCell(idx,1);
        
        cell_left.setWidth(249.48);
        cell_right.setWidth(249.48);
        
        
        cell_left.setPaddingBottom(0);
        cell_left.setPaddingLeft(0);
        cell_left.setPaddingRight(0);
        cell_left.setPaddingTop(0);
        
        cell_right.setPaddingBottom(0);
        cell_right.setPaddingLeft(0);
        cell_right.setPaddingRight(0);
        cell_right.setPaddingTop(0);
        
       
       
        appendContentInTableCell(temp, cell_left);
        appendContentInTableCell(copyBody,cell_right);
        
        //copyContent(temp, cell_left);
        //copyContent(copyBody,cell_right);
      }
     
  }
  Logger.log(ar);
  Logger.log("process finished, doc is saving");

  newDoc.saveAndClose();
  Logger.log("new Doc saved");

  var newFile = DocsList.getFileById(finalCardsId);

  var pdf = newFile.getAs("application/pdf");
  Logger.log("pdf created");
  
  //replace with your email address
  var emailTo = 'xxxx@gmail.com';
  var subject = 'test sending pdf using google app script';
  var message = "Please see attached";
  MailApp.sendEmail(emailTo, subject, message, {attachments:pdf});
  Logger.log("mail sent");

}

function generateNamecard() {
  
 var registerSheet = SpreadsheetApp.getActiveSheet();
 var regData = registerSheet.getDataRange().getValues();
 
 //template  
 var templateDocID = "1jnTdlPyu-ZieOl-3U6Dt_MxVjPR7HV9AUDHxz76_wto";//copyOfTemplate hash
 var templateDoc = DocumentApp.openById(templateDocID);
 var templateBody = templateDoc.getBody();
 
  
 // generated name card file
 var finalCardsId = "1-D-DEgtQRTmzDZ8H8KYQJqAhFh5SQ-34mq-GBbfsP_0";
 var newDoc = DocumentApp.openById(finalCardsId);
   
  var body = newDoc.getBody();
  body.clear();
  body.setMarginTop(10);
  body.setMarginLeft(55);
  body.setMarginRight(20);
  body.setMarginBottom(10);
   
  var temp;
  var copyBody;
  var cell_left;
  var cell_right;
  var ar = [];
  var table = body.appendTable();
  table.setBorderWidth(0.5);

   // parse register form
  for (var i = 1; i <= regData.length - 1; i++) {
  
    var row = regData[i];
    
    var username = row[0];
    var age = row[1];
    var gender = row[2];
    var church = row[11];
    var roomNr = row[13];
    var group = row[14];
    var groupLeader = row[15];
    
    Logger.log(i+" - "+username+" - "+roomNr +" - "+church+" - "+group)
    
   
    //copy template on temp
    Logger.log("copy templateBody to copyBody")
 
    //copyContent(templateBody, copyBody)
    copyBody = templateBody.copy();
    // replace text from data in register form
    if(username != '' && age > 2) {
      copyBody.replaceText('%na%', username.trim());
      if(roomNr == '') {
        copyBody.replaceText('%rmn%', '日营');
      } else {
        copyBody.replaceText('%rmn%', roomNr);
      }
      
      if(church == '') {
        copyBody.replaceText('%ch%', '');
      } else {
        copyBody.replaceText('%ch%', '教会: '+church.trim());
      }
      
      if(group == '') {
         copyBody.replaceText('%gr%', '');
      } else {
        copyBody.replaceText('%gr%', '小组: '+group);
      }
      
      
      if(groupLeader == '') {
         copyBody.replaceText('%gl%', '');
      } else {
         copyBody.replaceText('%gl%', '(组长)');
      }

    
     
      //build table , put data in it and insert it in finalCards doc
      if((i+1)%2 == 0){
        Logger.log("left copyBody "+copyBody.getText()+" saved in temp");
        temp = copyBody.copy(); // important move :) body must be first detached then assigned
        Logger.log("hoho: "+temp.getText());

      } else {
        Logger.log("left copyBody: "+temp.getText());
        Logger.log("right copyBody "+copyBody.getText())
        ar.push(i+1);
        var idx = ar.indexOf(i+1);
        Logger.log("build table row "+(i+1)+" and two cells")
        Logger.log("append left and right");
        var tableRow = table.appendTableRow();
        tableRow.appendTableCell('');
        tableRow.appendTableCell('');
        tableRow.setMinimumHeight(150);
        
        cell_left= table.getCell(idx,0);
        cell_right= table.getCell(idx,1);
        
        cell_left.setWidth(249.48);
        cell_right.setWidth(249.48);
        
        
        cell_left.setPaddingBottom(0);
        cell_left.setPaddingLeft(0);
        cell_left.setPaddingRight(0);
        cell_left.setPaddingTop(0);
        
        cell_right.setPaddingBottom(0);
        cell_right.setPaddingLeft(0);
        cell_right.setPaddingRight(0);
        cell_right.setPaddingTop(0);
        
       
       
        appendContentInTableCell(temp, cell_left);
        appendContentInTableCell(copyBody,cell_right);
        
        //copyContent(temp, cell_left);
        //copyContent(copyBody,cell_right);
      }
    } 
  }
  Logger.log(ar);
  Logger.log("process finished, doc is saving");

  newDoc.saveAndClose();
  Logger.log("new Doc saved");

  var newFile = DocsList.getFileById(finalCardsId);

  var pdf = newFile.getAs("application/pdf");
  Logger.log("pdf created");

  // replace your email here
  var emailTo = 'xxxx@gmail.com';
  var subject = 'test sending pdf using google app script';
  var message = "Please see attached";
  MailApp.sendEmail(emailTo, subject, message, {attachments:pdf});
  Logger.log("mail sent");

}

// JUST CLEAR room assginment
function justClear() {
  var registerSheet = SpreadsheetApp.getActiveSheet();

  var clearAssignedRange = registerSheet.getRange("N2:N500");

  clearAssignedRange.clear();
}


// JUST CLEAR group info
function clearGroupInfo() {
  var registerSheet = SpreadsheetApp.getActiveSheet();

  var clearGroupRange = registerSheet.getRange("O2:P500");

  clearGroupRange.clear();
}


// clear the column results from last assignment
function clearRegisterAssigned(registerSheet) {
  var clearAssignedRange = registerSheet.getRange("N2:N500");

  clearAssignedRange.clear();
}

// clear room overview and assginee and then copy new room overview from template
function clearRoomOverview(roomSheet) {
  //file: roomTemplate ID
   var roomTemplateId = '1nRkuJkk7ng-Hi2NsgwXcQestJBre9LqC4DbPPnr5Yfw';
   var roomTemplateSheet = SpreadsheetApp.openById(roomTemplateId).getActiveSheet();
   var rangetoCopy = roomTemplateSheet.getRange("A1:L150").getValues()
   
   roomSheet.clearContents();
   roomSheet.clearFormats() 

   roomSheet.getRange("A1:L150").setValues(rangetoCopy);
   skizzeRoomBorder(roomSheet);

}

function skizzeRoomBorder(sheet) {
  var templateData = sheet.getDataRange().getValues();
  
  for(var i = 1; i<=templateData.length -1; i++) {
        var row = templateData[i];
        var roomType = row[1];
         var range = sheet.getRange(i+1,5,1,roomType);
         range.setBorder(true, true, true, true, null, null);
        
  }
}

//get room statistics before and after assgining
function getStatistics(sheet) {
  var templateData =  sheet.getDataRange().getValues();
  
  var allrooms = 0;
  var allbeds = 0;
  var theRestbeds = 0;
  var allsix = 0;
  var allfive = 0;
  var allfour = 0;
  var allthree = 0;
  var alltwo = 0;
  var allone = 0;
  var six = 0;
  var five = 0;
  var four = 0;
  var three = 0;
  var two = 0;
  var one = 0;
  allrooms = templateData.length -1
  for(var i = 1; i<=templateData.length -1; i++) {
    var row = templateData[i];
    var roomType = row[1];
    allbeds = allbeds +roomType;
    var bedLeft = row[2];
    if(roomType == 6) {
        allsix = allsix + roomType
        six = six + bedLeft 
    }
    if(roomType == 5) {
        allfive = allfive + roomType
        five = five + bedLeft 
    }
    if(roomType == 4) {
        allfour = allfour + roomType
        four = four + bedLeft
    }
    if(roomType == 3) {
        allthree = allthree + roomType
        three = three + bedLeft
    }
    if(roomType == 2) {
        alltwo = alltwo + roomType
        two = two + bedLeft
    }
    if(roomType == 1) {
        allone = allone + roomType
        one = one + bedLeft
    }
  }
  theRestbeds = six + five + four + three + two + one;
  Logger.log("allrooms: "+allrooms);
  Logger.log("allbeds: "+allbeds);
  Logger.log("allsix: "+allsix);
  Logger.log("allfive: "+allfive);
  Logger.log("allfour: "+allfour);
  Logger.log("allthree: "+allthree);
  Logger.log("alltwo: "+alltwo);
  Logger.log("allone: "+allone);
  
  Logger.log("theRestbeds: "+theRestbeds);
  Logger.log("six: "+six);
  Logger.log("five: "+five);
  Logger.log("four: "+four);
  Logger.log("three: "+three);
  Logger.log("two: "+two);
  Logger.log("one: "+one);
  
  sheet.appendRow(['Statistics ', '']);
  sheet.appendRow(['All rooms: ', allrooms]);
  sheet.appendRow(['All beds: ', allbeds]);
  sheet.appendRow(['All 6er beds: ', allsix]);
  sheet.appendRow(['All 5er beds: ', allfive]);
  sheet.appendRow(['All 4er beds: ', allfour]);
  sheet.appendRow(['All 3er beds: ', allthree]);
  sheet.appendRow(['All 2er beds: ', alltwo]);
  sheet.appendRow(['All 1er bed: ', allone]);
  sheet.appendRow(['After assgining:','']);
  sheet.appendRow(['The rest beds: ', theRestbeds]);
  sheet.appendRow(['The rest 6er beds: ', six]);
  sheet.appendRow(['The rest 5er beds: ', five]);
  sheet.appendRow(['The rest 4er beds: ', four]);
  sheet.appendRow(['The rest 3er beds: ', three]);
  sheet.appendRow(['The rest 2er beds: ', two]);
  sheet.appendRow(['The rest 1er beds: ', one]);
  
}

//parse the whole form to find if duplicate user exist, time consuming
function findDuplicated(){
  var registerSheet = SpreadsheetApp.getActiveSheet();
  
  var regData = registerSheet.getDataRange().getValues();
  var ar = [];
  for (var i = 1; i <= regData.length - 1; i++) {
       var row = regData[i];
       var name = row[0];
       ar.push(name);
  }
  //Logger.log(ar);
  var sorted_arr = ar.sort();
  //Logger.log(sorted_arr);
  var results = [];
  for (var i = 0; i < ar.length - 1; i++) {
    if (sorted_arr[i + 1] == sorted_arr[i]) {
         results.push(sorted_arr[i]);
    }
  }
  Logger.log(results);
   Browser.msgBox("duplicated name found: "+results);
  
 
}

//flatten ununified date format and convert it to age number
function flattenBirthDate() {
  var registerSheet = SpreadsheetApp.getActiveSheet();
  
  var regData = registerSheet.getDataRange().getValues();
  for (var i = 1; i <= regData.length - 1; i++) {
    var row = regData[i];

    var age = row[1];
    var ageRange = registerSheet.getRange(i+1, 2)
    
    if(age == '') {
       age = 25;
    }
    Logger.log(typeof age+" "+age); 

    
    if((typeof age == 'number' && age >99) || typeof age == 'object') { // handle comishe date value
      ageRange.setNumberFormat("MM/dd/yyyy")
      age = 2014 - ageRange.getValue().getFullYear();
      
    } else if(typeof age == 'string') {
      if(age.indexOf('.') > -1 ) {
        age = age.split('.')[2];
      } else if(age.indexOf('/') > -1) {
        age = age.split('/')[2];
      }
      age = 2014 - parseInt(age);
    } 
 
    
    //unified number format with age number
    ageRange.setNumberFormat('0.0');
    ageRange.setValue(age);

    Logger.log(age); 
  
  }
  
}


//@deprecated! parse sheet two times, family first then random!
function runAssignRoomWithPriority() {
  var registerSheet = SpreadsheetApp.getActiveSheet();
  
  clearRegisterAssigned(registerSheet)
  
  var regData = registerSheet.getDataRange().getValues();
  
  //file: roomOverview ID
  var roomDocId = '11aQlhlpKaTZR-TL_6c3lQThlLeFa39irXd5DxLqyEak';
  var roomSheet = SpreadsheetApp.openById(roomDocId).getActiveSheet()
  clearRoomOverview(roomSheet)
  
  //file warteList ID
  var theRestDocId = '1s6hO_052g6ZzY0Sv7p4MNOKecpQ_9CO4_jcOI8vbEFU';
  var unassignedSheet = SpreadsheetApp.openById(theRestDocId).getActiveSheet();
  unassignedSheet.clearContents();
  unassignedSheet.appendRow(['总表上的行号','name', 'age','gender','faith','work?','mobile','email','note','','','','Fellowship']);

  var roomData = roomSheet.getDataRange().getValues();
  
  var theRest = [];
  
  
  // assign for special people and family first
  Logger.log("assign for special people and family first");
  for (var i = 1; i <= regData.length - 1; i++) {
    var row = regData[i];
    
    var username = row[0];
    var age = row[1];
    var gender = row[2];
    var roomReq = row[12];
    var noRoomReq = row[8];
    
    
    // skip day camp, no room request for those people,here also skip the children younger than 3
    if(noRoomReq !="" || noRoomReq == "是"){
       Logger.log(username+" go to day camp no room requirement, skip it!");
       continue;
    }
    
    var roomNr="";
    var roomRange = registerSheet.getRange(i+1, 14);
    
    // skip already assgined room for special reason 
    if(roomReq[0] != 'F' && (roomReq[0] == 'B' || roomReq[0] == 'A'|| roomReq >10 || roomReq.length > 3)) {
      Logger.log(username+" already get a room: "+roomReq+". skip it!");
      roomRange.setValue(roomReq);
      assignSpecialRoom(roomSheet, roomData, username, roomReq);
      continue; // skip it if room already assigned 
    }
    
    var rowLog = username+" - "+gender+ " --> "+roomReq
    Logger.log(row);
    
    if(roomReq && roomReq[0] == 'F'){ // check family mark condition if start 'F'
      roomNr = familyAllocating(roomSheet,roomData,username, gender, roomReq)
      Logger.log("assigned family room: "+roomNr+" for "+username);
      if(roomNr == null){
        roomRange.setBackground('red');
        row.unshift(i+1);
        unassignedSheet.appendRow(row);
        theRest.push(row);
        //theRestRowNr.push(i);
      } else {
        // set room nr in register form
        roomRange.setValue(roomNr);
      }
    } 
  }
  
  Logger.log("family parse finished, parse again for random");
  //parse again for random
  for (var i = 1; i <= regData.length - 1; i++) {
    var row = regData[i];
    
    var username = row[0];
    var age = row[1];
    var gender = row[2];
    var faith = row[3];

    var roomReq = row[12];
    var noRoomReq = row[8];
    
    
    // skip day camp, no room request for those people,here also skip the children younger than 3
    if(noRoomReq !="" || noRoomReq == "是"){
       Logger.log(username+" go to day camp no room requirement, skip it!");
       continue;
    }
    
    var roomNr="";
    var roomRange = registerSheet.getRange(i+1, 14);
    
    
    //if no mark, then start assigning 6er room 
    if(roomReq == "") {
       roomReq = 6;
    }
    
    //user older than 45 get 4er room first, family shoube be exclued
    if(age > 45 & roomReq[0] != 'F') {
       roomReq = 4;
    }
    
    var rowLog = username+" - "+gender+ " --> "+roomReq
    Logger.log(row);
    
    if (roomReq == 2 || roomReq == 3 || roomReq == 4 || roomReq == 5 || roomReq == 6) { // check normal marked condition
      roomNr = randomAllocating(roomSheet,roomData,username, gender, age, faith, roomReq) // recursiv function start from 6er until 2er,(e.g. if 6er is full then goto assgin 5er)
      Logger.log("assigned random room: "+roomNr+" for "+username);  
      if(roomNr == null){
        roomRange.setBackground('red');
        row.unshift(i+1);
        unassignedSheet.appendRow(row);
        theRest.push(row);
        //theRestRowNr.push(i);
      } else {
        // set room nr in register form
        roomRange.setValue(roomNr);
      }
    } 
    
  }
  Logger.log("here--> "+theRest);
  
  //statistics
  getStatistics(roomSheet);
}

function runAssignRoom() {
  var registerSheet = SpreadsheetApp.getActiveSheet();
  
  clearRegisterAssigned(registerSheet)
  var regData = registerSheet.getDataRange().getValues();
  
  
  var roomDocId = '11aQlhlpKaTZR-TL_6c3lQThlLeFa39irXd5DxLqyEak';
  var roomSheet = SpreadsheetApp.openById(roomDocId).getActiveSheet()
  clearRoomOverview(roomSheet)
  
  
  var theRestDocId = '1s6hO_052g6ZzY0Sv7p4MNOKecpQ_9CO4_jcOI8vbEFU';
  var unassignedSheet = SpreadsheetApp.openById(theRestDocId).getActiveSheet();
  unassignedSheet.clearContents();
  unassignedSheet.appendRow(['rowNr','name','age','gender','faith','work','mobile','email','note','','','','Fellowship']);
  var roomData = roomSheet.getDataRange().getValues();
  
  var theRest = [];
  
  // parse register form
  for (var i = 1; i <= regData.length - 1; i++) {
    var row = regData[i];
    
    var username = row[0];
    var age = row[1];
    var gender = row[2];
    var faith = row[3];
    var roomReq = row[12];
    var noRoomReq = row[8];
    
    
    // skip day camp, no room request for those people,here also skip the children younger than 3
    if(noRoomReq !="" || noRoomReq == "是"){
       Logger.log(username+" go to day camp no room requirement, skip it!");
       continue;
    }
    
    var roomNr="";
    var roomRange = registerSheet.getRange(i+1, 14);
    
    // skip already assgined room for special reason 
    if(roomReq[0] != 'F' && (roomReq[0] == 'B' || roomReq[0] == 'A'|| roomReq >10 || roomReq.length > 3)) {
      Logger.log(username+" already get a room: "+roomReq+". skip it!");
      roomRange.setValue(roomReq);
      assignSpecialRoom(roomSheet, roomData, username, roomReq);
      continue; // skip it if room already assigned 
    }
    
    //if no mark, then start assigning 6er room 
    if(roomReq == "") {
       roomReq = 6;
    }
    
    //user older than 45 get 4er room first, family shoube be exclued
    if(age > 45 & roomReq[0] != 'F') {
       roomReq = 4;
    }
    
    var rowLog = username+" - "+gender+ " --> "+roomReq
    Logger.log(row);
    
    if(roomReq && roomReq[0] == 'F'){ // check family mark condition if start 'F'
      roomNr = familyAllocating(roomSheet,roomData,username, gender, roomReq)
      Logger.log("assigned family room: "+roomNr+" for "+username);
    } else if (roomReq == 2 || roomReq == 3 || roomReq == 4 || roomReq == 5 || roomReq == 6) { // check normal marked condition
      roomNr = randomAllocating(roomSheet,roomData,username, gender, age, faith, roomReq) // recursiv function start from 6er until 2er,(e.g. if 6er is full then goto assgin 5er)
      Logger.log("assigned random room: "+roomNr+" for "+username);
    }
    
    if(roomNr == null){
        roomRange.setBackground('red');
        row.unshift(i+1);
        unassignedSheet.appendRow(row);
        theRest.push(row);
        //theRestRowNr.push(i);
    } else {
        // set room nr in register form
        roomRange.setValue(roomNr);
    }
  }
  Logger.log("here--> "+theRest);
  
  //statistics
  getStatistics(roomSheet);

  
};

function assignSpecialRoom(roomSheet, roomData, username, roomReq) {
  Logger.log("allocating special room for "+ username+" --> "+roomReq);
  for (var i = 1; i <= roomData.length - 1; i++) {
    var row = roomData[i];
    var roomNr = row[0];
    var roomType = row[1];
    if(roomNr == roomReq) {
      var markDef = roomSheet.getRange(i+1, 4);
      var roomDef = markDef.getValue();
      if(roomDef == ''){
        markDef.setValue('预给');
      }
       var bedLeft = roomSheet.getRange(i+1, 3);
       var bedLeftNumber = bedLeft.getValue();
       if(bedLeftNumber != 0) {
         
         for(j=0; j< roomType; j++) {
           var bed = roomSheet.getRange(i+1, j+5);
           var ocuppiedBed = bed.getValue();
           
           if(ocuppiedBed == '' || ocuppiedBed == 'frei'){
             bed.setValue(username);
             bed.setBackground('#00ff00');
             
             Logger.log('assign No.'+(j+1)+' bed in : '+roomNr+' for '+username);
             bedLeftNumber = bedLeftNumber -1; 
             bedLeft.setValue(bedLeftNumber);
             SpreadsheetApp.flush();
             Logger.log("left number of bed : "+bedLeftNumber+" in room: "+roomNr);
             return roomNr;
           } else {
             Logger.log("the "+(j+1)+"th bed is occupid by "+ocuppiedBed+" in room: "+roomNr);
             Logger.log("try to assign No."+(j+2)+" bed for "+ username+" in room: "+roomNr);
           }
         }
       }
    }
  
  }

}

function randomAllocating(roomSheet,roomData,username, gender, age, faith, roomReq){
  Logger.log("random allocating room: "+ username+" - "+gender+" - "+age+" "+faith+" --> "+roomReq);
  
  var assignedRoom = null;
  for (var i = 1; i <= roomData.length - 1; i++) {
    var row = roomData[i];
    
    //unchanged field
    var roomNr = row[0]; 
    var roomType = row[1];
    
    //room without lift, only for young people
    if((roomNr == 850 || roomNr == 851) && roomReq == 4 && age > 45) {
        Logger.log('no lift, not suit for you');
        continue;
    }
    
    // AO,BB only for age < 50
    
    if(((roomNr[0] == 'A' || roomNr[0] == 'B') && age > 50)) {
        Logger.log('AO&BB only for christian');
        continue;
    }
    
    //changeable field, it's acctually assigning bed
    var bedLeft = roomSheet.getRange(i+1, 3);
    var bedLeftNumber = bedLeft.getValue();
    var rowCell = roomNr +" - "+roomType+ "er - bed: "+bedLeftNumber;
    Logger.log(rowCell);
    
    //define for gender
    var markDef = roomSheet.getRange(i+1, 4);
    
    // assign room as required but based on the room type and left beds in a room
    if((roomReq == roomType) && bedLeftNumber != 0 ) {
      
      var roomDef = markDef.getValue();
      
      if(roomDef == ''){ //who is the first in this room, then define this user gender for this room
        markDef.setValue(gender);
        SpreadsheetApp.flush(); // save in real time
        
        roomDef = gender;
        Logger.log("set this room as "+gender);
        Logger.log(username+' is first person to get this room');
      } 
      
      Logger.log('this room： '+roomNr+' is only for '+roomDef);
      
      // only assign bed to user with same gender in a room
      if(roomDef == gender) {
        Logger.log(username+' is '+gender+'-> get room: '+roomNr);
        
        for(j=0; j< roomType; j++) {
          var bed = roomSheet.getRange(i+1, j+5);
          var ocuppiedBed = bed.getValue();
          
          //empty bed, assign to user
          if(ocuppiedBed == '' || ocuppiedBed == 'frei'){
            bed.setValue(username);
            bed.setBackground('#00ff00');
            Logger.log('assign No.'+(j+1)+' bed in : '+roomNr+' for '+username);
            
            bedLeftNumber = bedLeftNumber -1; 
            bedLeft.setValue(bedLeftNumber);
            SpreadsheetApp.flush();
            Logger.log("left number of room: "+bedLeftNumber);
            assignedRoom = roomNr;
            return assignedRoom;
          } else {
            Logger.log("the No."+(j+1)+" bed is occupid by "+ocuppiedBed+" in room: "+roomNr);
            if(roomType == (j+1)) {
              Logger.log(roomNr+ " is full, get no bed from this room, try find another room for "+username);
            } else {
              Logger.log("try to assign No."+(j+2)+" bed for "+ username+" in room: "+roomNr);
            }
          }
        }
      } else {
          Logger.log(username+' is '+gender+'. get no bed from this room');
          Logger.log('find next possible room for '+username);
        
      }
    } 
  }
  
  if(assignedRoom==null && roomReq > 2) {
      Logger.log("--> "+username+ " request: "+roomReq+"er not statisfied. no more "+roomReq+"er room. try to assign: "+(roomReq-1)+ "er room.");
      return randomAllocating(roomSheet, roomData, username, gender, age, faith, roomReq-1); // call random function recursive 
  } else {
      Logger.log("**room list parsed, can not find a suit room for "+username+" **");
  }
}

function familyAllocating(roomSheet,roomData,username, gender, roomReq){
  Logger.log("family allocating room: "+ username+" - "+gender+" --> "+roomReq);
 
  var roomTypeReq = roomReq[1];
  
  for (var i = 1; i <= roomData.length - 1; i++) {
    var row = roomData[i];
    
    //unchangeable field
    var roomNr = row[0];
    var roomType = row[1];
    
    // room without lift, not suit for family
    if(roomNr == 850 || roomNr == 851) {
        continue;
    }
     //changeable field, it's acctually assigning bed
    var bedLeft = roomSheet.getRange(i+1, 3);
    var bedLeftNumber = bedLeft.getValue();
    
    var rowCell = roomNr +" - "+roomType+ "er - bed: "+bedLeftNumber;
    Logger.log(rowCell);
    
    var markDef = roomSheet.getRange(i+1, 4);
    if((roomTypeReq == roomType) && bedLeftNumber != 0 ) {
      
      var roomDef = markDef.getValue();
      if(roomDef == ''){
        markDef.setValue(roomReq);
        SpreadsheetApp.flush();
        roomDef = roomReq;
        
        Logger.log("set this room as "+roomReq);
        Logger.log(username+' is first person to get this room');
      } 
      
      Logger.log('this room： '+roomNr+' is only for '+roomDef);
      if(roomDef == roomReq) {
        Logger.log(username+' is family: '+roomReq);
        Logger.log('so allocating room: '+roomNr);
        for(j=0; j< roomType; j++) {
          var bed = roomSheet.getRange(i+1, j+5);
          var ocuppiedBed = bed.getValue();

          if(ocuppiedBed == '' || ocuppiedBed == 'frei'){
            bed.setValue(username);
            bed.setBackground('#00ff00');

            Logger.log('assign No.'+(j+1)+' bed in : '+roomNr+' for '+username);
            bedLeftNumber = bedLeftNumber -1; 
            bedLeft.setValue(bedLeftNumber);
            SpreadsheetApp.flush();
            Logger.log("left number of bed : "+bedLeftNumber+" in room: "+roomNr);
            return roomNr;
          } else {
            Logger.log("the "+(j+1)+"th bed is occupid by "+ocuppiedBed+" in room: "+roomNr);
            Logger.log("try to assign No."+(j+2)+" bed for "+ username+" in room: "+roomNr);
          }
        }
      } else {
          Logger.log(username+' is family '+roomReq+'. get no bed from this room');
          Logger.log('find next possible room for '+username);
        
      }
    } 
  }
  
}




/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 */

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "Just Clear",
      functionName : "justClear"
    },
    null, 
     {
      name : "Clear group info",
      functionName : "clearGroupInfo"
    },
    null, 
    {
    name : "分房（先家庭后随机）",
    functionName : "runAssignRoomWithPriority"
    },
    null, 
    {
    name : "分房（按报名顺序）",
    functionName : "runAssignRoom"
    },
     null, 
    {
    name : "导入小组信息",
    functionName : "copyGroupInfo"
    },
     null, 
    {
    name : "生成小组预览",
    functionName : "generateGroupOverview"
    },
    null,
      {
    name : "生成名卡",
    functionName : "generateNamecard"
    },
    null, 
    {
    name : "flatten Birthday",
    functionName : "flattenBirthDate"
    },
     null, 
    {
    name : "find Duplicated name",
    functionName : "findDuplicated"
    }
  ];
  spreadsheet.addMenu("分房菜单", entries);
};
