/*****************************************************************************

 DESCRIPTION:
    This collection of scripts is my attempt to creat a useful library
    of common scripts that can be used across different applications. 
      
  PRIMARY FUNCTIONS   
     
     return_value = isValidDate(date);
       Returns a date/time value if a valid date is input, returns "false, if not
       
     DayOfWeek = GetDay(date);
       Returns a string "Sun" through "Sat" if a valid date is input, returns "false, if not
       
     return_value = DateDiff(FirstDate, FutureDate, Period);
       Computes the difference between two dates and returns an value in the units specified by the "Period" scalar
        Allowable Period Values:  DAYS, WEEKS, MONTHS, YEARS
        
     progressMsg(message,title,timeoutSeconds)
       Uses .toast() to displey parameter driven status messages
       
     returnValue = deleteGoogleDoc(DocURL, oCommon) 
    
   REVISION DATE:
    11-25-2017 - First Instance
    12-10-2017 - Mods GenerateReports to enable use for generating multiplt docs
    03-02-2018 - Modiified deleteGoogleDoc() to change ownership and remone non-owned Doogle docs
    08-20-2018 - Modified changeOwner() by adding a trap to prevent "Exception: Action not allowed" errors when the depreciated 
                  "business.manager@sticypress.org" email address is encountered
    12-04-2019 - Add traps to catch and correct the depreciated "business.manager@sticypress.org" email address

   
 UTILITY FUNCTIONS
  
   numToA(num); Converts an input Column Number (Base 0) to an equivalent A1 notation
     var column_letter = numToA(column_number)    
     
   GetParameters(oSourceSheets, TargetSheet, KeyWord, KeyCol, ValueCol, Parameters, ErrorMessage); 
     Parameters = []; must be declared external to this function
     Searches the first column for the Parameter Section "KeyWord" in the Target Sheet
     If "ValueCol" is empty, returns all column values starting with the KeyCol column; 
       if not, just the KeyCol column value and the ValueCol column values are returned   

 AUTHOR:
    Steve Germain
    (208) 870-2727
    
 COPYRIGHT NOTICE:
   Copyright (c) 2009 - 2017 by Legatus, Inc.
   
 NOTES:
    The function assignEditUrls() is invoked whenever a Form Submission is completed.
    The Awesome Table (https://awesome-table.com/) service is utilized for data display.
    

/***************************************************************************************************************
****************************************************************************************************************
****************************************************************************************************************
*
*    Miscellaneous Utility Functions Below
*
****************************************************************************************************************
****************************************************************************************************************
***************************************************************************************************************/

function FormatTagValue(oCommon, TagValue){    
  // Useage:  TagValue = FormatTagValue(oCommon, TagValue);
  var MaxLength = oCommon.RecordTagMaxLength;
  if(ParamCheck(MaxLength) && !isNaN(MaxLength)){
    // Use the defined Tag Value / truncated
    return TagValue.toString().substring(0, MaxLength);
  } else {
    // Use the defined Tag Value As is
    return TagValue.toString();
  } 
}


function GetXrefValue(oCommon, GlobalValue){
  var func = "****GetXrefValue " ;
  var Step = 1000;
  // useage: xref_value = GetXrefValue(oCommon, GlobalValue);
  try {
    if (ParamCheck(oCommon.Replace_Term_Prefix)){
      var Prefix = RegExp(oCommon.Replace_Term_Prefix);
      if(Prefix.test(GlobalValue)){
        Step = 1100; // use the terms xref
        //Logger.log( func + Step + ' Input Value: ' + GlobalValue + ', Return Value: ' + oCommon.TermsXrefAry[GlobalValue]);
        return oCommon.TermsXrefAry[GlobalValue];
      } else {
        Step = 1200; // use the Headings xref
        //Logger.log( func + Step + ' Input Value: ' + GlobalValue + ', Return Value: ' + oCommon.HeadingsXrefAry[GlobalValue]);
        return oCommon.HeadingsXrefAry[GlobalValue];
      }
    } else {
      Step = 1300; // use the Headings xref
      //Logger.log( func + Step + ' Input Value: ' + GlobalValue + ', Return Value: ' + oCommon.HeadingsXrefAry[GlobalValue]);
      return oCommon.HeadingsXrefAry[GlobalValue];
    }
  } catch (err) {
    Step = 1400; 
    var return_value = '';
    Logger.log( func + Step + ' Input Value: ' + GlobalValue + ', Return Value: empty');
    return return_value
  }
}


function CheckStatus(Timestamp, Status, oCommon){    
  var func = "****CheckStatus " ;
  var Step = 100;
  var Trail = ' ';
  //Check for Timestamp
  Logger.log( func + Step + ' - CheckForValidTimestamp: ' + oCommon.CheckForValidTimestamp
            + ', Active_Status_value: ' + oCommon.Active_Status_value
            + ', Inactive_Status_value: ' + oCommon.Inactive_Status_value);
  Step = 1000; // Check Timestamp settings
  var bGoodTimestamp = false;  
  if (oCommon.CheckForValidTimestamp){  
    if (ParamCheck(Timestamp)){
      Step = 1200;// Timestamp value exists
      Trail = Trail + ', ' + Step;
      bGoodTimestamp = true; 
    }
  } else { 
    Step = 1200;// No Timestamp check
    Trail = Trail + ', ' + Step;
    bGoodTimestamp = true; 
  }
  
  Step = 2000; // Check Status
  var bActiveStatus = false;
  if (ParamCheck(Status)){
    Step = 2100; // Check for "Active" Status value match
    Trail = Trail + ', ' + Step;
    if (oCommon.Active_Status_value != ''){
      if (oCommon.Active_Status_value == Trim(Status)) { 
        Step = 2110;// ActiveStatus is Active
        Trail = Trail + ', ' + Step;
        bActiveStatus = true; 
      }
    }
    Step = 2200; // Check for "Inactive" Status value match
    if (oCommon.Inactive_Status_value != ''){
      Trail = Trail + ', ' + Step;
      if (oCommon.Inactive_Status_value != Trim(Status)) { 
        Step = 2210;// InActiveStatus is Active
        Trail = Trail + ', ' + Step;
        bActiveStatus = true; 
      }
    }
  } else { 
    Step = 2300;// No Status check
    Trail = Trail + ', ' + Step;
    bActiveStatus = true; 
  }

  if (bGoodTimestamp && bActiveStatus) {
    Step = 3100;// Return True
    Trail = Trail + ', ' + Step;
    Logger.log( func + Step + Trail);
    return true;
  } else {
    Step = 3200;// Return False
    Trail = Trail + ', ' + Step;
    Logger.log( func + Step + Trail);
    return false;
  }
}


function FormatForDisplay(arySelectedRows){
  // Useage:  FormatForDisplay(oCommon.SelectedSheetRows)
  // Build the display message
  var RowsSelected = [];
  for (var r = 0; r < arySelectedRows.length; r++){
    RowsSelected[r] = [];
    RowsSelected[r].push('\\n' + '  Row: ' + arySelectedRows[r][0] + ' (' 
                         + arySelectedRows[r][1] + ')');
  }
  return RowsSelected;
}

function GetSheetParam(SheetName, ParamKey, oCommon){
  // Sheet Details Headings: Title, Headings Row, 1st Data Row, Terms Row, AWT Row, Param 1
  // var HeadingsRow = GetSheetParam(SheetName, 'Headings Row', oCommon);
  // var 1stDataRow = GetSheetParam(SheetName, '1st Data Row', oCommon);
  // var TermsRow = GetSheetParam(SheetName, 'Terms Row', oCommon);
  // var AWTRow = GetSheetParam(SheetName, 'AWT Row', oCommon); // Awesome Tables parameters
  // var Param1 = GetSheetParam(SheetName, 'Param 1', oCommon); // Misc parameter value
  if (ParamCheck(oCommon.sheetParams[SheetName])){
    return oCommon.sheetParams[SheetName][ParamKey];
  } else {
    return oCommon.sheetParams['Default'][ParamKey];
  }
}
  
function ParamCheckTest(){
  var func = "***ParamCheckTest " ;
  var Step = 100;
  var param = '';
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
 
  var Step = 200;
  var param = 'abc';
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
  
  var Step = 300;
  var param = true;
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
  
  var Step = 400;
  var param = null;
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
  
  var Step = 500;
  var param = false;
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
  
  var Step = 600;
  var param = 0;
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
  
  var Step = 700;
  var param = 123;
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
  
  var Step = 800;
  var param = -123;
  Logger.log(func + Step + ' param: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
  
  var Step = 900;
  var test = [];
  var param = test['a123'];
  Logger.log(func + Step + ' aryparam: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
   
  var Step = 1000;
  var param = [];
  Logger.log(func + Step + ' aryparam: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
    
  var Step = 1100;
  var param = [1,2,3];
  Logger.log(func + Step + ' aryparam: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
    
  var Step = 1200;
  var param = [];
  param.push[1];
  Logger.log(func + Step + ' aryparam: "' + param + '", ParamCheck(param) = ' + ParamCheck(param));
}

function ParamCheck(parameter){
  if (parameter === ''){
    return false;
  } else if (typeof parameter === 'undefined'){
    return false;
  } else if (parameter === null){
    return false;
  } else if(isNaN(parameter)){
    return true;
  } else if(!isNaN(parameter)){
    return true;
  }
  return false;
}

function isValidEmail(email) {
  //Source: https://stackoverflow.com/questions/46155/how-to-validate-an-email-address-in-javascript
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(String(email).toLowerCase());
}

function isValidURL(str) {
  // Source: https://stackoverflow.com/questions/5717093/check-if-a-javascript-string-is-a-url
  // Useage:  bRteurnBool = isValidURL(str);
  var regex = /(http|https):\/\/(\w+:{0,1}\w*)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%!\-\/]))?/;
  if(!regex .test(str)) {
    return false;
  } else {
    return true;
  }
}

function isFormUrl(value) {
  Step = 2210; // Test value to determine if it contains a url reference to a Google Form
  if(ParamCheck(value)){
    if(!isNaN(value)){ return false; }
    value.toString();
  } else {
    return false; 
  }
  if(value.indexOf("https://") > -1){
    Step = 2220; // URL found
    //Logger.log(func + Step + ' Array Row: ' + ArrayRow + ' URL to be tested: ' + test_value);
    var formUrl = RegExp("https://docs.google.com/forms/");
    if(formUrl.test(value)){
      return true;
    }
  }
  return false;
}

function Clean(str) {
  
  if (!ParamCheck(str)){
    return str;
  } else {
    str = str.toString();
    return str.replace(/[^\x20-\x7E]/g, '');
  }
}


function A1CellAddress(row, col){
  // returns a "row,col" cell address formated using A1 notation
  if(!isNaN(row) && !isNaN(col)){
    A1CellAddress = numToA(col) + String(row).split(".")[0]; 
  } else {A1CellAddress = ''}
}


function findInArray(value, data, startrow, startcol, endrow, endcol) {
  // Finds the row and column containing first instance of "value" within a 2-dimensional "data" array
  var func = "***findInArray " ;
  //Logger.log(func + ' BEGIN');
  var iterations = 0;
  if(value == null || value == "") {return null;}
  if(data == null || data == "") {return null;}
  if(endrow == null || endrow == "" || endrow < 0) {endrow = data.length;}
  if(endcol == null || endcol == "" || endcol < 0) {endcol = 99999;}
  if(startrow == null || startrow == "" || startrow < 0) {startrow = 0;}
  if(startcol == null || startcol == "" || startcol < 0) {startcol = 0;}
  if (endrow > data.length) {endrow = data.length;}
  for (var r = startrow; r < data.length; r++) {
    iterations++;
    if (endcol > data[r].length) {endcol = data[r].length;}
    for (var c = startcol; c <= endcol; c++) {
      if (data[r][c] == value) {
        //Logger.log(func + '(' + value +') Returned - ' + r + '/' + c + ' after searching ' + iterations + ' rows');
        //Logger.log(func + ' Returned - ' + r + '/' + c + ' after searching ' + iterations + ' rows');
        return [r,c];
      }
    }
  }
  //Logger.log(func + '(' + value +') Returned - null after searching ' + iterations + ' rows');
  //Logger.log(func + ' Returned - null after searching ' + iterations + ' rows');
  return null;
}

function findRowCol(value, TabName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TabName);
  var data = ss.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] == value) {
        return [i,j];
        
        //return (ss.getSheetByName(TabName).getRange(i + 1, j + 1));
        //This will return range. You can change that line to get the A1Notation:
        // return (ss.getSheetByName(sheet).getRange(i + 1, j + 1).getA1Notation());        
        
      }
    }
  }
  return null;
}


function find(value, range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = range.split("!")[0];
  var A1Ref = range.split("!")[1];
  var activeRange = ss.getSheetByName(sheet).getRange(A1Ref);
  var data = activeRange.getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] == value) {
        return (ss.getSheetByName(sheet).getRange(i + 1, j + 1));
        //This will return range. You can change that line to get the A1Notation:
        // return (ss.getSheetByName(sheet).getRange(i + 1, j + 1).getA1Notation());        
        
      }
    }
  }
  return null;
}



function AuthenticateUser(acceptable_email_addresses) {

  var func = "***AuthenticateUser - " ;
  var Step = 0;
  var user_email = Session.getActiveUser().getEmail();

  if (acceptable_email_addresses.indexOf(user_email) >= 0) {
    Logger.log(func + Step + ' User: ' + user_email + ' Input: ' + acceptable_email_addresses + ' Finding: TRUE');
    return true;
  } else {
    Logger.log(func + Step + ' User: ' + user_email + ' Input: ' + acceptable_email_addresses + ' Finding: FALSE');
    return false;
  }
}
function getIds(array) {
  // Source: https://codereview.stackexchange.com/questions/9241/comparing-two-arrays-to-see-if-they-contain-objects-with-different-ids
    // We have added `Array.prototype.unique`; it does what you'd expect.
    return array.map(function (x) { return x.id; }).unique().sort();
}

function areDifferentByIds(a, b) {
  // Source: https://codereview.stackexchange.com/questions/9241/comparing-two-arrays-to-see-if-they-contain-objects-with-different-ids  
    var idsA = getIds(a);
    var idsB = getIds(b);

    if (idsA.length !== idsB.length) {
        return true;
    }

    for (var i = 0; i < idsA.length; ++i) {
        if (idsA[i] !== idsB[i]) {
            return true;
        }
    }

    return false;
}

function areDifferentByIds(a, b) {
  // Source: https://codereview.stackexchange.com/questions/9241/comparing-two-arrays-to-see-if-they-contain-objects-with-different-ids
    var addArgs = Array.prototype.slice.call(arguments, 2);
    var props = addArgs.length > 0 ? addArgs : ['id'];
    for(var diff=false, i=0; !diff && i<addArgs.length; i++) {
        diff = a[addArgs[i]] !== b[addArgs[i]];
    }
    return diff;
}

  

function changeOwner(DocURL, oCommon){
/*****************************************************************************

  USEAGE:
    NewDocURL = changeOwner(DocURL, oCommon);
  
******************************************************************************/
  
  var func = "***changeOwner - " ;
  Step = 1000;
  Logger.log(func + 'BEGIN - DocURL: ' + DocURL);

  //Set default return value
  var NewDocURL = DocURL;
  
  try{
    
    Step = 2000;
    var oldID = getIdFrom(DocURL);
    var oldFile = DriveApp.getFileById(oldID);
    
    Step = 2100;
    var oldOwnerEmail = oldFile.getOwner().getEmail();
    var newOwnerEmail = oCommon.ScriptManager;
    //Logger.log(func + Step + ' OldOwner: ' + oldOwnerEmail + ' NewOwner: ' + newOwnerEmail);
    
    Step = 2120; // add trap to prevent "Exception: Action not allowed" errors when the depreciated 
                 //   "business.manager@sticypress.org" email address is encountered
    if (newOwnerEmail.indexOf("business.manager")>-1){
      newOwnerEmail = "business.mgr@sticypress.org";
    }
    if (oldOwnerEmail.indexOf("business.manager")>-1){
      oldOwnerEmail = "business.mgr@sticypress.org";
    }
    
    Step = 3000; // Make changes if the owners are different
    if (oldOwnerEmail !== newOwnerEmail){
      
      Logger.log(func + Step + ' OldOwner: ' + oldOwnerEmail + ' NewOwner: ' + newOwnerEmail);

      Step = 3100; //Give the "new owner" edit rights just in case the domains are different
      try {
        var newFile = oldFile.addViewer(newOwnerEmail); 
        var newFile = oldFile.addEditor(newOwnerEmail); 
      } catch(err) {
        var EventMsg = func + Step + ' Change Permissions Failed - OldOwner: ' + oldOwnerEmail
            + ', NewOwner: ' + newOwnerEmail + ' (error msg: ' + err + ')';
        Logger.log(EventMsg);
      }
      
      Step = 3200; //Attempt to change ownership
      try {
        var newFile = oldFile.setOwner(newOwnerEmail); 
        NewDocURL = newFile.getUrl();
        Logger.log(func + Step + ' Ownership Changed - DocURL: ' + DocURL + ' OldOwner: ' + oldOwnerEmail + ' New: ' + newOwnerEmail);
      } catch(err) {
        var EventMsg = func + Step + ' Change Ownership Failed - OldOwner: ' + oldOwnerEmail
            + ', NewOwner: ' + newOwnerEmail + ' (error msg: ' + err + ')';
        Logger.log(EventMsg);
      }
      
      Step = 3300; // Delete (or remove) the original document
      if (!deleteGoogleDoc(DocURL, oCommon)){
        // Capture the failure to Delete the original document file
        var EventMsg = func + Step + ' Warning 500 - Unable to delete the original document file - OldOwner: ' + oldOwnerEmail 
        + ', NewOwner: ' + newOwnerEmail + ' DocURL: ' + DocURL;
        LogEvent(EventMsg, oCommon); 
        Logger.log(EventMsg);
        //Logger.log(func + Step + ' Result - Input DocURL:  ' + DocURL + ' OldOwner: ' + oldOwnerEmail);
        //Logger.log(func + Step + ' Result - Output DocURL: ' + NewDocURL + ' NewOwner: ' + newOwnerEmail);
      }
    } else {
      Step = 3400; // No need to change ownership
      Logger.log(func + Step + ' No need for Ownership Change detected - DocOwner: ' + oldOwnerEmail);
    }
    
  } catch(err) {
    var EventMsg = func + Step + ' Error 502 - Change Owner Failed - OldOwner: "' + oldOwnerEmail 
    + '", NewOwner: "' + newOwnerEmail + '" with error message: ' + err;
    LogEvent(EventMsg, oCommon);    
  }

  Step = 4000; 
  return NewDocURL;
}

function BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol,Prefix,Suffix){
  /*****************************************************************************

   DESCRIPTION:
     This function is used to build and return a key / value object

   USEAGE
     Input Parameter Values:
     // Build Key/Value object array
       var TargetSheet = 'Sheet Name',    // Sheet Containing Keys and Values
           SectionTitle = 'Section Heading Name', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start r,ow is always the following row
           KeyRow = 'Heading or Col #',   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
           ValueRow = 'Heading or Col #', // number or text [Look for text Match in Col 0 (Base 0)]
           KeyCol = 'Heading or Col #',   // number or text [Look for text Match in StartRow (Base 0)]
           ValueCol = 'Heading or Col #', // number or text [Look for text Match in StartRow (Base 0)]
           oKeyValueAry = {}, // if declared before entering function
           Prefix = '<<',
           Suffix = '>>';
       KeyValueAry = BuildKeyValueAry(oSourceSheets, TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol,oKeyValueAry,Prefix,Suffix)


   REVISION DATE:
      12/31/2017 - First instance
      04/10/2019 - Added ability to add "terms' prefix and suffix strung values.


   NOTES:
   
     
  ******************************************************************************/
  
  var func = "***BuildKeyValueAry - " ;
  Step = 100;
  //Logger.log(func + ' BEGIN');
  Logger.log(func + Step + ' TargetSheet: ' + TargetSheet + ', SectionTitle: ' + SectionTitle + ', KeyRow: ' + KeyRow 
    + ', ValueRow: ' + ValueRow + ', KeyCol: ' + KeyCol + ', ValueCol: ' + ValueCol
    + ', Prefix: ' + Prefix + ', Suffix: ' + Suffix);

  Step = 200;
  if (typeof KeyRow === 'undefined' || KeyRow === null || KeyRow === ''){var bKeyRowActive = false;} else {var bKeyRowActive = true;}
  if (typeof ValueRow === 'undefined' || ValueRow === null || ValueRow === ''){var bValueRowActive = false;} else {var bValueRowActive = true;}
  if (typeof KeyCol === 'undefined' || KeyCol === null || KeyCol === ''){var bKeyColActive = false;} else {var bKeyColActive = true;}
  if (typeof ValueCol === 'undefined' || ValueCol === null || ValueCol === ''){var bValueColActive = false;} else {var bValueColActive = true;}
  if (typeof Prefix === 'undefined' || Prefix === null){var Prefix = '';} 
  if (typeof Suffix === 'undefined' || Suffix === null){var Suffix = '';} 
  //Logger.log(func + Step + ' KeyRowActive: ' + bKeyRowActive + ' ValueRowActive: ' + bValueRowActive);
  //Logger.log(func + Step + ' KeyColActive: ' + bKeyColActive + ' ValueColActive: ' + bValueColActive);
  
  /******************************************************************************/
  Step = 1000; // Load the Tab containing the parameter values
  /******************************************************************************/
  try {
    var values = oSourceSheets.getSheetByName(TargetSheet).getDataRange().getValues();
  } catch(err) {
    var EventMsg = 'ERROR - Tab Named "' + TargetSheet + '" could not be found.';
    Browser.msgBox(EventMsg);
    EventMsg = func + Step + EventMsg + ' Error Msg: ' + err;
    Logger.log(func + Step + EventMsg);
    return false;
    
  }

  /******************************************************************************/
  Step = 2000; // Use the SectionTitle to find the Key Start Row, if entered
  /******************************************************************************/
  var start_row = 0;
  var begin_row = 0;
  var bSectionFound = false;
  if (SectionTitle){
    Step = 2100// Find the start row by searching for the KeyWord text
    //Logger.log(func + Step + ' values.length: ' + values.length);
    for (var row = 0; row < values.length; row++) { 
      //Logger.log(func + Step + ' row: ' + row + ', values[row][0]: ' + values[row][0]);
      if (values[row][0].indexOf(SectionTitle) > -1){
        Step = 2150;
        bSectionFound = true;
        var start_row = row + 1;
        var begin_row = start_row;
        //Logger.log(func + Step + ' start_row: ' + start_row);
        break;
      }
    }
    //if (start_row = 0){
    if (!bSectionFound){
      Step = 2200;
      Logger.log(func + Step + ' Error - SectionTitle (' + SectionTitle + ') not found.');
      return false;
    }
  } // End of search for SectionTitle
  
  Step = 2300;
  //Logger.log(func + Step + ' bSectionFound: ' + bSectionFound + ', start_row: ' + start_row + ', begin_row: ' + begin_row);

  /******************************************************************************/
  Step = 3000; // Use the SectionTitle to find the Key Start Row, if entered
  /******************************************************************************/
  //Logger.log(func + Step + ' KeyRow: ' + KeyRow + ' KeyCol: ' + KeyCol);
  if (bKeyRowActive){
    // Using Row-based Assignments
    if (!isNaN(KeyRow) && KeyRow > -1){
      Step = 3010; // The KeyRow identifier is a numeric Value
      var KeyRowNumber = KeyRow;
    } else if (KeyRow != ''){
      Step = 3020; // The KeyRow identifier is a text scalar
      // Find the KeyRowNumber by searching for the KeyRow identifier text in Col 0 (Base 0)
      //for (var row = start_row; row < values.length; row++) { 
      for (var row = begin_row; row < values.length; row++) { 
        if (values[row][0].indexOf(KeyRow) > -1){
          var KeyRowNumber = row;
          break;
        }
      }
      if (!KeyRowNumber){
        Logger.log(func + Step + ' Error - KeyRow identifier text (' + KeyRow + ') not found.');
        return false;
      }
    } 
    //Logger.log(func + Step + ' KeyRowNumber (' + KeyRowNumber + ') found.');

    /******************************************************************************/
    Step = 3100; // If KeyRow is used, find the required ValueRow 
    /******************************************************************************/
    //Logger.log(func + Step + ' KeyRowNumber: ' + KeyRowNumber + ' ValueRowNumber: ' + ValueRow);
    if(KeyRowNumber > -1){
      if (!isNaN(ValueRow) && ValueRow > -1){
        Step = 3110; // The KeyRow identifier is a numeric Value
        var ValueRowNumber = ValueRow;
      } else {
        Step = 3120; // The ValueRow identifier is a text scalar
        // Find the ValueRowNumber by searching for the ValueRow identifier text in Col 0 (Base 0)
        for (var row = start_row; row < values.length; row++) { 
          if (values[row][0].indexOf(ValueRow) > -1){
            var ValueRowNumber = row;
            break;
          }
        }
        if (!ValueRowNumber){
          Logger.log(func + Step + ' Error - ValueRow identifier text (' + ValueRow + ') not found.');
          return false;
        }
      } 
      //Logger.log(func + Step + ' ValueRowNumber: ' + ValueRowNumber);
      
      /******************************************************************************/
      Step = 3200; // Use KeyRow and ValueRow to build the Key / Value array 
      /******************************************************************************/
      //if(!oKeyValueAry){
        var oKeyValueAry = {};
      //}
      //Logger.log(func + Step + ' KeyRowNumber: ' + KeyRowNumber + ', ValueRowNumber: ' + ValueRowNumber);
      var count = 0;
      for (var c = 0; c < values[KeyRowNumber].length; c++){
        // Skip empty or null "keys"
        if ( ParamCheck(values[KeyRowNumber][c])){
          if (ValueRowNumber){
            count++;
            oKeyValueAry[Prefix + values[KeyRowNumber][c] + Suffix] = values[ValueRowNumber][c];
          } else {
            oKeyValueAry[Prefix + values[KeyRowNumber][c] + Suffix] = c;
            count++;
          } 
        }
      }
      // Verify
      //for(var key in oKeyValueAry){
      //  Logger.log(func + Step + ' Key: ' + key + '  Value: ' + oKeyValueAry[key]);
      //}
      
      Logger.log(func + Step + ' oKeyValueAry completed - length: ' + count);

      return oKeyValueAry;
    } 
  }// End of look for KeyRowNumber and ValueRowNumber
  
  /******************************************************************************/
  Step = 4000; // If KeyCol is used, find the KeyColNumber(s)
  /******************************************************************************/
  if(bKeyColActive){
    //Logger.log(func + Step + ' KeyCol: ' + KeyCol + ' ValueCol: ' + ValueCol);
    var KeyColNumbers = [];
    if (!isNaN(KeyCol)){
      Step = 4010; // The KeyCol identifier is a numeric Value
      KeyColNumbers[0] = [];
      KeyColNumbers[0][0] = KeyCol;
      KeyColNumbers[0][1] = 0;
      //Logger.log(func + Step + ' KeyColNumber: ' + KeyColNumbers[0][0]);
    } else {
      Step = 4020; // The KeyCol identifier is a TEXT Value 
                   //  Build Composite Key the array of column headings
      //var HeadingsRow = 0;
      var HeadingsRow = begin_row;
      KeyColNumbers = GetKeyParams(KeyCol, values, HeadingsRow);
      //KeyColNumbers[part][0] = Column Number of the KeyValue part
      //KeyColNumbers[part][1] = Number of Characters to use when forming the Key Value with the current "part"
      
    }
    if (KeyColNumbers.length <= 0){
      Logger.log(func + Step + ' Error - KeyCol identifier text (' + KeyCol + ') not found.');
      return false;
    }
    //Logger.log(func + Step + ' KeyColNumbers: ' + KeyColNumbers);
  
    /******************************************************************************/
    Step = 4100; // If ValueCol is used, find the ValueColNumber
    /******************************************************************************/
    if (bValueColActive){
      //Logger.log(func + Step + ' KeyCol: ' + KeyCol + ' ValueCol: ' + ValueCol);
      if (!isNaN(ValueCol)){
        Step = 4110; // The ValueCol identifier is a numeric Value
        var ValueColNumber = ValueCol;
        //Logger.log(func + Step + ' ValueColNumber: ' + ValueColNumber);
      } else {
        Step = 4120; // The ValueCol identifier is a text scalar
        // Find the ValueColNumber by searching for the ValueCol identifier text in start_row (Base 0)
        if (SectionTitle){start_row = start_row;}
        for (var col = 0; col < values[start_row].length; col++) { 
          //Logger.log(func + Step + ' row: ' + start_row + ', col: ' + col + ', Value: ' + values[start_row][col]);
          if (values[start_row][col].indexOf(ValueCol) > -1){
            var ValueColNumber = col;
            break;
          }
        }
      }
      if (isNaN(ValueColNumber)){
        Logger.log(func + Step + ' Error - ValueCol identifier text (' + ValueCol + ') not found.');
        return false;
      }
    } // End of bValueColActive test
    //Logger.log(func + Step + ' ValueColNumber: ' + ValueColNumber);
  
    /******************************************************************************/
    Step = 4200 // Use KeyCol and ValueCol to build the Key / Value array 
    /******************************************************************************/
    //Logger.log(func + Step + ', KeyColNumbers: ' + KeyColNumbers + ' ValueColNumber: ' + ValueColNumber);
    //if(!oKeyValueAry){var oKeyValueAry = {};}
    var oKeyValueAry = {};
    Step = 4210; // Declare the variabls and examine each row
    var count = 0,
        bSkipEntry,
        AsciiValue,
        InActiveStatusChar = 9744, // Ascii value of an "empty" checkbox symbol
        ActiveStatusChar = 9745;   // Ascii value of a "checked" checkbox symbol
    
    for (var r = begin_row; r < values.length; r++){
      
      Step = 4220; //Check for the presence of an "Active Checkbox" in the first column
      bSkipEntry = false; // default to assign every row entry
      AsciiValue = values[r][0].charCodeAt(0); // ascii value of the first character
      if(AsciiValue == InActiveStatusChar || AsciiValue == ActiveStatusChar){
        Step = 4222; /// A Status checkbox is present
        //Logger.log(func + Step + 'Param row: ' + r + ', Col 0 Value: ' + values[r][0] + ', Ascii Value: ' + AsciiValue);
        if (AsciiValue != ActiveStatusChar){
          bSkipEntry = true;
        }
      }
      
      Step = 4230; //Process an active row
      if(!bSkipEntry){
        var KeyValue = '';
        Step = 4231// assignKeyValue test string value
        if (KeyColNumbers.length == 1){
          var part = 0;
          Step = 4232// Assign the KeyValue string value for the current FormData row (e.g. LastNameFirstName)
          //Logger.log(func + Step + ' r: ' + r + ' KeyColNumbers[' + part + '][0]:' + KeyColNumbers[part][0] + '  Num_Chrs: ' + KeyColNumbers[part][1]);
          KeyValue = Trim(values[r][KeyColNumbers[part][0]]);
          if (KeyColNumbers[part][1] > 0){KeyValue = KeyValue.substr(0,KeyColNumbers[part][1]);}
        } else if (KeyColNumbers.length > 1){
          Step = 4234; // Build the KeyValue string value for the current FormData row (e.g. LastNameFirstName)
          for (var part = 0; part < KeyColNumbers.length; part++){
            var AppendString = Trim(values[r][KeyColNumbers[part][0]]);
            //Logger.log(func + Step + ' r: ' + r + ' KeyColNumbers[' + part + ']:' + KeyColNumbers[part] + ' Num_Chrs: ' + Num_Chrs[part] + ' Value ' + AppendString);
            if (KeyColNumbers[part][1] > 0){AppendString = AppendString.substr(0,KeyColNumbers[part][1]);}
            KeyValue = KeyValue + AppendString;
          }
        }
        Step = 4236; // Test for a valid KeyValue String
        if (bSectionFound && !KeyValue){break;} // Stop if an empty column value is encountered - end of section
        if (isNaN(KeyValue) && ParamCheck(KeyValue)){
          // Build the Key/Value set
          if (ValueColNumber){
            oKeyValueAry[Prefix + KeyValue + Suffix] = values[r][ValueColNumber];
            count++;
            //Logger.log(func + Step + ' r: ' + r + ' KeyValue: ' + KeyValue + ' ColValue: ' + values[r][ValueColNumber]);
          } else {
            oKeyValueAry[Prefix + KeyValue + Suffix] = r;
            count++;
            //Logger.log(func + Step + ' r: ' + r + ' KeyValue: ' + KeyValue + ' ColValue: ' + r);
          } 
        }
      } // End of check for "active" row entry
    } // Get next row
    
    // Verify
    //for(var key in oKeyValueAry){
    //  Logger.log(func + Step + ' Key: ' + key + '  Value: ' + oKeyValueAry[key]);
    //}
    
    Logger.log(func + Step + ' oKeyValueAry completed - length: ' + count);
    
    return oKeyValueAry;
    
  } // End of look for KeyCoNumber and ValueColNumber
  
  Step = 5000;
  Logger.log(func + Step + ' Error - No Valid Input Parameters found.');
  return false;
}


function GetKeyParams(KeyColTitles, SheetData, HeadingsRow){
  var func = "***GetKeyParams - " ;
  var Step = 100;
  //Logger.log(func + Step + ' BEGIN');
  /******************************************************************************/
  Step = 1000; /* Build a Composite Key using column headings
  
               Useage:
               var KeyParams = [];
               KeyParams = KeyParams(KeyColTitles, SheetData, HeadingsRow);
                 KeyParams[part][0] = Column Number of the KeyValue part
                 KeyParams[part][1] = Number of Characters to use when forming the Key Value with the current "part"
               
  *******************************************************************************/
  //oSheetData = Sheet Data object
  var KeyParams = [];
  var heading_row = HeadingsRow;
  var KeyColTitles = KeyColTitles.split('+'); //Used for building a composite key value
  //Logger.log(func + Step + ' KeyColTitles: ' + KeyColTitles);
  for (var part = 0; part < KeyColTitles.length; part++){
    KeyParams[part] = [];
    var KeyColTitle = Trim(KeyColTitles[part]);
    var KeyColNumber = 0;
    var KeyColNumChrs = 0;
    var start_pos = KeyColTitle.indexOf('('); // look for the "(" delimiter
    if (start_pos > -1){
      // only part of the string is to be used
      var end_pos = KeyColTitles[part].indexOf(')'); // look for the ")" delimiter
      KeyColNumChrs = KeyColTitle.substring(start_pos, end_pos).length; // capture the value of n between "(" and ")"
      KeyColTitle = KeyColTitle.substring(0,start_pos); // remove the "(n)" nomenclature
    }
    Step = 1100;// Find the KeyColNumber by searching for the KeyColTitle string in heading_row (Base 0)
    for (var col = 0; col < SheetData[heading_row].length; col++) { 
      //Logger.log(func + Step + ' row: ' + heading_row + ', col: ' + col + ', Value: ' + SheetData[heading_row][col]);
      if (SheetData[heading_row][col].indexOf(KeyColTitle) > -1){
        KeyColNumber = col;
        break;
      }
    } // End of search for matching part heading in the start row
    
    //Logger.log(func + Step + ' part: ' + part + ', ColTitle: ' + KeyColTitle +', col: ' + col + ', NumChrs: ' + KeyColNumChrs);
    KeyParams[part][0] = KeyColNumber;
    KeyParams[part][1] = KeyColNumChrs;
  }
  return KeyParams;
}
    


function GetParameters(oSourceSheets, TargetSheet, StartRow, KeyWord, KeyCol, ValueCol, ErrorMessage){
  /******************************************************************************
  Capture the Parameter values placed in the "Target" Sheet
  *******************************************************************************/
  var func = "***GetParameters - " ;
  Step = 1000;

  /******************************************************************************/
  Step = 1000; // Load the Tab containing the parameter values
  /******************************************************************************/
  var values = oSourceSheets.getSheetByName(TargetSheet).getDataRange().getValues();
  if (!StartRow || StartRow < 0){StartRow = 0;}
  if (!KeyCol || KeyCol < 0){KeyCol = 0;}
  if (!ValueCol || ValueCol < 0){ValueCol = 0;}
Logger.log(func + Step + ' TargetSheet: ' + TargetSheet + ', StartRow: ' + StartRow + ', KeyWord: ' + KeyWord
       + ' KeyCol: ' + KeyCol + ' ValueCol: ' + ValueCol + ', values.length: ' + values.length);

  /******************************************************************************/
  Step = 2000; // Find the Start Row, if necessary
  /******************************************************************************/
  var first_row = null;
  if (!KeyWord){
    // Begin with the StartRow
    first_row = StartRow;
  } else {
    // Find the start row by searching for the KeyWord text
    KeyWord = KeyWord.toUpperCase();
    for (var row = StartRow; row < values.length; row++) { 
      if (values[row][0].toUpperCase().indexOf(KeyWord) >= 0 ){
        first_row = row + 1;
        break;
      }
    }
  }
  
  if (first_row == null){
    // Error - Heading Term Not Found
    ErrorMessage = func + Step + ' Error 410 - Setup Heading Term (' + KeyWord + ') Not Found.';
    Logger.log(ErrorMessage);
    return ;
  }
  //Logger.log(func + Step + ' TargetSheet: ' + TargetSheet + ', StartRow: ' + StartRow + ', KeyWord: ' + KeyWord
  //         + ' KeyCol: ' + KeyCol + ' ValueCol: ' + ValueCol + ' first_row: ' + first_row);
  
  /******************************************************************************/
  Step = 3000; //Evaluate the KeyCol Input value
  /******************************************************************************/
  if (isNaN(KeyCol)){
    // KeyCol value is a string and is assumed to be a heading title string value
    // Search for Key Col Term in the first row
    if(values[0].indexOf(KeyCol) > -1){
      // Matching heading string found
      var KeyColNumber = values[0].indexOf(KeyCol);
    } else {
      // Error - Heading Term Not Found
      ErrorMessage = func + Step + ' Error 412 - Setup Heading Column Term (' + KeyCol + ') Not Found.';
      Logger.log(ErrorMessage);
      return;
    }
  } else { var KeyColNumber = KeyCol; }
  
  /******************************************************************************/
  Step = 4000; //Evaluate the ValueCol Input value
  /******************************************************************************/
  if (isNaN(ValueCol)){
    Step = 4100; // KeyCol value is a string and is assumed to be a heading title string value
    // Search for Key Col Term in the first row
    if(values[0].indexOf(ValueCol) > -1){
      Step = 4200; // Matching heading string found
      var ValueColNumber = values[0].indexOf(ValueCol);
    } else {
      Step = 4300; // Error - ValueCol Heading Term Not Found
      ErrorMessage = func + Step + ' Error 414 - ValueCol Heading Term (' + ValueCol + ') Not Found.';
      Logger.log(ErrorMessage);
      return;
    }
  } else { 
    Step = 4400; 
    var ValueColNumber = ValueCol; 
  }
  //Logger.log(func + Step + ' TargetSheet: ' + TargetSheet + ', StartRow: ' + StartRow + ', KeyWord: ' + KeyWord
  //       + ' KeyCol: ' + KeyCol + ', KeyColNumber: ' + KeyColNumber + ' ValueCol: ' + ValueCol);
  
  /******************************************************************************/
  Step = 5000; // Capture the scalar value assignments
  /******************************************************************************/
  var Parameters = [];
  var bSkipEntry,
      row_skipped = 0,
      AsciiValue,
      InActiveStatusChar = 9744, // Ascii value of an "empty" checkbox symbol
      ActiveStatusChar = 9745;   // Ascii value of a "checked" checkbox symbol
  for (var r = first_row; r < values.length; r++) { 
    if(!values[r][KeyColNumber]){break;}
    Step = 5100; //Check for the presence of an "Active Checkbox" in the first column
    bSkipEntry = false; // default to assign every row entry
    AsciiValue = values[r][0].charCodeAt(0); // ascii value of the first character
    if(AsciiValue == InActiveStatusChar || AsciiValue == ActiveStatusChar){
      Step = 5200; /// A Status checkbox is present
      //Logger.log(func + Step + 'Param row: ' + r + ', Col 0 Value: ' + values[r][0] + ', Ascii Value: ' + AsciiValue);
      if (AsciiValue != ActiveStatusChar){
        bSkipEntry = true;
        row_skipped++;
      }
    }
    Step = 5300; //Capture the values for "active" rows
    if(bSkipEntry == false){   
      Step = 5310; //Parameters[r - first_row] = [];
      var param_row = r - first_row - row_skipped;
      Parameters[param_row] = [];
      // Determine if assignments are to be taken from a single column or all columns
      if(ValueColNumber){
        Step = 5320; 
         Parameters[param_row][KeyColNumber] = values[r][ValueColNumber]; 
         //Logger.log(func + Step + ' Parameters[' + (r - first_row) + '][' + KeyColNumber + ']: ' + Parameters[r - first_row][KeyColNumber]);
       } else {
        Step = 5320; // Assign values from  the range of columns
        for (var col = 0; col < values[r].length; col++){
          Parameters[param_row][col] = values[r][col]; 
        }
      }
    }
  } // End of loop thru each row
  
  // Verify results
  //Logger.log(func + Step + ' Parameters.length: ' + Parameters.length);
  //for (var i = 0; i < Parameters.length; i++){
  //  Logger.log('**** GetParameters(' + i + '): ' + Parameters[i]);
  //}
  
  return Parameters;
}


function Trim(v) {
  // Source: https://gist.github.com/brucemcpherson/3358690
  return LTrim(RTrim(v));
};
/**
 * Removes leading whitespace
 * @param {string|number} s the item to be trimmed
 * @return {string} The trimmed result
 */
function LTrim(s) {
  return CStr(s).replace(/^\s\s*/, "");
};
/**
 * Removes trailing whitespace
 * @param {string|number} s the item to be trimmed
 * @return {string} The trimmed result
 */
function RTrim(s) {
  return CStr(s).replace(/\s\s*$/, "");
};
function CStr(v) {
  return v===null || IsMissing(v) ? ' ' :  v.toString()  ;
}
function IsMissing (x) {
  return isUndefined(x);
}
function isUndefined ( arg) {
  return typeof arg == 'undefined';
}


function moveFile(fileToMoveURL, targetFolderID) {
  // Source: https://stackoverflow.com/questions/38808875/moving-files-in-google-drive-using-google-script
  var func = "***moveFile - " ;
  var Step = 1000;
  Logger.log(func + Step + ' BEGIN - fileToMoveURL: ' + fileToMoveURL);

  if (!ParamCheck(fileToMoveURL)){
    Logger.log(func + Step + ' Error - Invalid fileToMoveURL passed');
    return fileToMoveURL;
  }
  
  if (!ParamCheck(targetFolderID)){
    Logger.log(func + Step + ' Error - Invalid targetFolderID passed, Value: ' + targetFolderID);
    return fileToMoveURL;
  }
  
  try {
  
    Step = 2100;
    var fileToMoveID = getIdFrom(fileToMoveURL);
    var sourceFile = DriveApp.getFileById(fileToMoveID); //Get the file to move
Logger.log(func + Step + ' sourceFile Url: ' + sourceFile.getUrl());
    
    Step = 2200;
    var targetFolder = DriveApp.getFolderById(targetFolderID);
    
    Step = 2300; 
    var sourceFileName = sourceFile.getName();
    
    Step = 2400; 
    var NewFile = sourceFile.makeCopy(sourceFileName, targetFolder); // Create a copy of the newDoc in the shared folder
    
    Step = 2500; 
    var NewFileURL = NewFile.getUrl(); // Create a copy of the newDoc in the shared folder
Logger.log(func + Step + ' NewFileURL: ' + NewFileURL);
    
    Step = 2600; 
    sourceFile.setTrashed(true);  //  sets the file in the trash of the user's Drive    
    
    Step = 2700; 
    Logger.log(func + Step + ' Successful Move (' + sourceFileName + ') New URL: ' + NewFileURL);
    return NewFileURL;
  }
  
  catch (err) {
    Logger.log(func + Step + ' moveFile Error - FolderID: ' + targetFolderID + ', FileURL: ' + fileToMoveURL + ', ErrMsg: ' + err);
    return false;
  }
}

function addMonths (date, count) {
  if (date && count) {
    var m, d = (date = new Date(+date)).getDate()
    date.setMonth(date.getMonth() + count, 1)
    m = date.getMonth()
    date.setDate(d)
    if (date.getMonth() !== m) date.setDate(0)
  }
  return date
}

function getIdFrom(url) {
  // Shamelessly lifted from: https://stackoverflow.com/questions/16840038/easiest-way-to-get-file-id-from-url-on-google-apps-script
  // https://drive.google.com/open?id=0B2q1MltxdeUJVndvczJXcXJvdWM
  var func = "***getIdFrom - " ;
  //Logger.log(func + ' BEGIN - Input URL: ' + url);
  
  if (!url){
    Logger.log(func + ' Error - Empty URL passed');  
    return;
  } else
    if (typeof url !== 'string'){
      Logger.log(func + ' Error - Invalid URL passed  URL Value: ' + url);  
      return;
    }
  
  if(url.toUpperCase().indexOf('HTTP') > -1){
    var id = "";
    var parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
    //Logger.log(func + ' 10  url:' + url + ' parts: ' + parts);
    if (url.indexOf('?id=') >= 0){
      id = (parts[6].split("=")[1]).replace("&usp","");
      return id;
    } else {
      id = parts[5].split("/");
      //Using sort to get the id as it is the longest element. 
      var sortArr = id.sort(function(a,b){return b.length - a.length});
      id = sortArr[0];
      //Logger.log(func + ' END - Output Google ID: ' + id);
      return id;
    }
  } else {
    // Input "url" is probably already a Google ID value, return w/o taking any action
    //Logger.log(func + ' END - Output Google ID: ' + url);
    return Trim(url);
  }
 
}

function numToA(num){
    var a = '',modulo = 0;
    for (var i = 0; i < 6; i++){
        modulo = num % 26;
        if(modulo == 0) {
            a = 'Z' + a;
            num = num / 26 - 1;
        }else{
            a = String.fromCharCode(64 + modulo) + a;
            num = (num - modulo) / 26;
        }
        if (num <= 0) break;
    }
    return a;
}


/**
 * function will return true if the arg is valid date and false if not.
 * @param {*} d - scalar to be tested
 * @return {False or valid date}

function isValidDate(d) {
  // Source: https://stackoverflow.com/questions/25654177/date-validation-with-if-then-function-in-google-apps-script
  if (Object.prototype.toString.call(d) !== "[object Date]"){return false;}
  if (d.getTime()){return false;}
  if (d.getFullYear() < 1899){return false;}
  return true;
}


function isValidDate(d) {
  try{
    var dTest = new Date(d);
    if (!isNaN(dTest.getTime()) && dTest.getFullYear() >= 1899){ return true;}
    return false;
  } catch(e) {
    return false;
  }
}
*/

function isValidDate(d) {
  if (!isNaN(d)){
    if (Object.prototype.toString.call(d) !== "[object Date]"){return false;}
    if (d.getFullYear() < 1899){return false;}
    return true;
  }

  try{
    var dTest = new Date(d);
    if (!isNaN(dTest.getTime()) && dTest.getFullYear() >= 1899){ return true;}
    return false;
  } catch(e) {
    return false;
  }
}

function DateMatch(d1, d2, Period){    
  var func = "***DateMatch - " ;
   var bReturnStatus = false;
 
  if (isValidDate(d1) && isValidDate(d2) && ParamCheck(Period)){
    // Valid input parameters
    
    var Period = Period.toString().toUpperCase();
    
    switch(Period) {
        
      case "YEARS":
        if(DateDiff(FirstDate, FutureDate, "YEARS") == 0){
          bReturnStatus = true;
        }
        break;
        
      case "MONTHS":
        if(DateDiff(FirstDate, FutureDate, "MONTHS") == 0){
          bReturnStatus = true;
        }
        break;
        
      case "WEEKS":
        if(DateDiff(FirstDate, FutureDate, "WEEKS") == 0){
          bReturnStatus = true;
        }
        break;
        
      case "DAYS":
        if(DateDiff(FirstDate, FutureDate, "DAYS") == 0){
          bReturnStatus = true;
        }
        break;
        
      case "HOURS":
        if(DateDiff(d1, d2, "Days") == 0){
          var h1 = d1.getHours();
          var h2 = d2.getHours();
          if(h1 === h2){
            bReturnStatus = true;
          }
        }
        break;
        
      case "MINUTES":
        if(DateDiff(d1, d2, "Days") == 0){
          var h1 = d1.getHours();
          var h2 = d2.getHours();
          var m1 = d1.getMinutes();
          var m2 = d2.getMinutes();
          if(h1 === h2 && m1 === m2){
            bReturnStatus = true;
          }
        }
        break;
        
      case "SECONDS":
        if(DateDiff(d1, d2, "Days") == 0){
          var h1 = d1.getHours();
          var h2 = d2.getHours();
          var m1 = d1.getMinutes();
          var m2 = d2.getMinutes();
          var s1 = d1.getSeconds();
          var s2 = d2.getSeconds();
          if(h1 === h2 && m1 === m2 && s1 === s2){
            bReturnStatus = true;
          }
        }
        
      default:
        
    }
  }
  //Logger.log(func + 'RETURN: ' + bReturnStatus + ', Period: ' + Period + ' d1: ' + d1 + ', d2: ' + d2);
  return bReturnStatus;
}


function DateDiff(FirstDate, FutureDate, Period){    
  //Source: https://stackoverflow.com/questions/11174385/compare-two-dates-google-apps-script
  var d1 = FirstDate;
  var d2 = FutureDate;
  
  if (isValidDate(d1) && isValidDate(d2) && Period){
    // Valid input parameters
    var Period = Period.toString().toUpperCase();
    
    if (Period == 'DAYS'){
      var t2 = d2.getTime();
      var t1 = d1.getTime();
      return parseInt((t2-t1)/(24*3600*1000));
    }
    if (Period == 'WEEKS'){
      var t2 = d2.getTime();
      var t1 = d1.getTime();
      return parseInt((t2-t1)/(24*3600*1000*7));
    }
    if (Period == 'MONTHS'){
      var d1Y = d1.getFullYear();
      var d2Y = d2.getFullYear();
      var d1M = d1.getMonth();
      var d2M = d2.getMonth();
      return (d2M+12*d2Y)-(d1M+12*d1Y);
    }
    if (Period == 'YEARS'){
      return d2.getFullYear()-d1.getFullYear();
    }
  }
  return false;
}

function GetDay(d){
  if (isValidDate(d)) {
    var weekday = new Array(7);
    weekday[0] = "Sunday";
    weekday[1] = "Monday";
    weekday[2] = "Tuesday";
    weekday[3] = "Wednesday";
    weekday[4] = "Thursday";
    weekday[5] = "Friday";
    weekday[6] = "Saturday";
    return weekday[d.getDay()];
  } else { return false; }
}

function AbbrevMonth(m){
  if (!isNaN(m) && m > 0 && m <= 12){
    var month = new Array(12);
    m = m - 1; //Adjust for Base 0
    month[0] = "Jan";
    month[1] = "Feb";
    month[2] = "Mar";
    month[3] = "Apr";
    month[4] = "May";
    month[5] = "Jun";
    month[6] = "Jul";
    month[7] = "Aug";
    month[8] = "Sep";
    month[9] = "Oct";
    month[10] = "Nov";
    month[11] = "Dec";
    return month[m];
  } else { return ""; }
}

function convertDate(inputFormat) {
  // source: https://stackoverflow.com/questions/39083619/how-to-convert-date-in-javascript-to-mm-dd-yyyy-format
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat);
  return [pad(d.getMonth()+1), pad(d.getDate()), d.getFullYear()].join('/');
}

function formatDate(InDate) {
  // Source: https://stackoverflow.com/questions/25275696/javascript-format-date-time isValidDate(d)
  //if(!date){return '';}
  if(!isValidDate(InDate)){return '';}
  date = new Date(InDate);
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var ampm = hours >= 12 ? 'pm' : 'am';
  hours = hours % 12;
  hours = hours ? hours : 12; // the hour '0' should be '12'
  minutes = minutes < 10 ? '0'+minutes : minutes;
  var strTime = hours + ':' + minutes + ' ' + ampm;
  if (date.getFullYear() < 1900){
    //Logger.log(' formatDate - ' + 10 + ' date: ' + date +  ' Time Value: ' + strTime);
    return strTime;
  } else {
    //return date.getMonth()+1 + "/" + date.getDate() + "/" + date.getFullYear() + "  " + strTime;+
    //Logger.log(' formatDate - ' + 20 + ' date: ' + date +  ' Date Value: ' + date.getMonth() + "/" + date.getDate() + "/" + date.getFullYear());
    return date.getMonth()+1 + "/" + date.getDate() + "/" + date.getFullYear();
  }
}

function formatYear(date) {
  // Source: https://stackoverflow.com/questions/25275696/javascript-format-date-time isValidDate(d)
  //if(!date){return '';}
  if(!isValidDate(date)){return '';}
  return date.getFullYear();
}

function formatTime(date) {
  // Source: https://stackoverflow.com/questions/25275696/javascript-format-date-time
  if(!isValidDate(date)){return '';}
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var ampm = hours >= 12 ? 'pm' : 'am';
  hours = hours % 12;
  hours = hours ? hours : 12; // the hour '0' should be '12'
  minutes = minutes < 10 ? '0'+minutes : minutes;
  var strTime = hours + ':' + minutes + ' ' + ampm;
  //return date.getMonth()+1 + "/" + date.getDate() + "/" + date.getFullYear() + "  " + strTime;
  return strTime;
  
}

function formatDateTime(InDate) {
  // Source: https://stackoverflow.com/questions/25275696/javascript-format-date-time isValidDate(d)
  //if(!date){return '';}
  if(!isValidDate(InDate)){return '';}
  date = new Date(InDate);
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var ampm = hours >= 12 ? 'pm' : 'am';
  hours = hours % 12;
  hours = hours ? hours : 12; // the hour '0' should be '12'
  minutes = minutes < 10 ? '0'+minutes : minutes;
  var strTime = hours + ':' + minutes + ' ' + ampm;
  if (date.getFullYear() < 1900){
    //Logger.log(' formatDate - ' + 10 + ' date: ' + date +  ' Time Value: ' + strTime);
    return strTime;
  } else {
    //return date.getMonth()+1 + "/" + date.getDate() + "/" + date.getFullYear() + "  " + strTime;+
    //Logger.log(' formatDate - ' + 20 + ' date: ' + date +  ' Date Value: ' + date.getMonth() + "/" + date.getDate() + "/" + date.getFullYear());
    return date.getMonth()+1 + "/" + date.getDate() + "/" + date.getFullYear() + ' ' + strTime;
  }
}

function formatAMPM(date) { // This is to display 12 hour format like you asked
  // Source: https://stackoverflow.com/questions/25275696/javascript-format-date-time
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var ampm = hours >= 12 ? 'pm' : 'am';
  hours = hours % 12;
  hours = hours ? hours : 12; // the hour '0' should be '12'
  minutes = minutes < 10 ? '0'+minutes : minutes;
  var strTime = hours + ':' + minutes + ' ' + ampm;
  return strTime;
}

function sendTheEmail(oCommon, To, CC, BCC, Subject, Message, error_message, ReplyTo) {
  var func = "***sendTheEmail " + Version + " - ";
  var NoErrors = true;
  Logger.log(func + Step + " Input Params: "
   + " To/ReplyTo/CC/BCC: "  + To + "/" + ReplyTo + "/" + CC + "/" + BCC);

  // Add traps to prevent catch the depreciated "business.manager@sticypress.org" email address
  if (To.indexOf("business.manager")>-1){To = "business.mgr@sticypress.org"};
  if (CC.indexOf("business.manager")>-1){CC = "business.mgr@sticypress.org"};
  if (BCC.indexOf("business.manager")>-1){BCC = "business.mgr@sticypress.org"};
  
  try {
    
    if(!ParamCheck(ReplyTo) || ReplyTo.toUpperCase().indexOf("NOREPLY") > -1){
      // Use "noReply" if ReplyTo missing or contains the string "noReply"
      MailApp.sendEmail(To, Subject, Message,
                      {
                        cc:      CC,
                        bcc:     BCC,
                        noReply: true
                      });
      return true;
       
    } else {
      // incorporate a ReplyTo address
      MailApp.sendEmail(To, Subject, Message,
                      {
                        cc:      CC,
                        bcc:     BCC,
                        replyTo: ReplyTo
                      });
      return true;
    }
    
  } catch (e) {
    var error_message = func + Step + " Error - Unable to Send Email: "
        + " To/ReplyTo/CC/BCC: "  + To + "/" + ReplyTo + "/" + CC + "/" + BCC
        + ", Error Message: " + e.message;
                      
    LogEvent(error_message, oCommon);    
    return false;
  }
}

function ImportFile(SourceSS, TabName){

  var func = "***ImportFile - " ;
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');

  Step = 1000; // Load the Data TabName Source Tab
  var SSName = SourceSS.getName();
  Logger.log(func + Step + ' SSName: ' + SSName + ' TabName: ' + TabName);
  
  var Sheet = SourceSS.getSheetByName(TabName);
  var message = "Do you want to import a new " + TabName + " data file?";
  if (Browser.msgBox(message, Browser.Buttons.YES_NO).toUpperCase() == "YES"){
    Step = 1100; // Yes - Import it!
    // The code below will set the value of "file_name" input by the user, or 'cancel'
    var file_name = Browser.inputBox('Enter File Name', 'Import Data File Name:', Browser.Buttons.OK_CANCEL);
    
    // return if cancelled of filename is empty
    if (file_name.toUpperCase() == "CANCEL" || !file_name) {
      Logger.log(func + Step + ' File name emprty or cancelled');
      return false;
    }
    
    Step = 1200; // Insert a file Source:  https://ctrlq.org/code/20279-import-csv-into-google-spreadsheet
    var file = DriveApp.getFilesByName(file_name).next();
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
    Sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    Logger.log(func + Step + ' File name ' + file_name + ' successfully imported');
    return true;
    
  } else {
  
    Step = 1300; // User answers "NO"
    Logger.log(func + Step + ' Choice to import data file declined by user');
    return false;
  }
 
}

function progressMsg(message,title,timeoutSeconds) {
  if(ParamCheck(message)){
    //title = message; // temp fix because toast is only showing the "title" entry - 12/12/2018
    //replace carrage return 
    message = message.replace("\\n", "");
    SpreadsheetApp.getActiveSpreadsheet().toast(message,title,timeoutSeconds)
  }
}


function deleteGoogleDoc(DocURL, oCommon) {
  // Source: https://stackoverflow.com/questions/28311109/how-to-retry-driveapp-getfilebyid-if-error-occurs-until-successful
  // Also: http://www.mousewhisperer.co.uk/drivebunny/handle-errors-in-apps-script-with-try-and-catch/
  var func = "***deleteGoogleDoc - " ;
  var Step = 1000;
  var NoErrors = true;
  if (!ParamCheck(DocURL)){
    oCommon.DisplayMessage = ' DocURL value empty or invalid.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    return false;
  }
  
  Step = 2000; // First, use the DriveApp to find and access the target file
  try {
    var FileID = getIdFrom(DocURL);
    var DocFile = DriveApp.getFileById(FileID);
  } catch (err1) {
    oCommon.DisplayMessage = ' Unable to GET file.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage + ' err: ' + err1;
    //LogEvent(EventMsg, oCommon);   
    Logger.log(oCommon.ReturnMessage);
    return false;
  }
  
  Step = 3000; //Once found, try and delete, assuming the file is in a Team Drive
  try {
    DocFile.setTrashed(true);
    Logger.log(func + Step + ' Delete Successful: ' + DocURL);
    return true;
  } catch (err2) {
    Logger.log(func + Step + ' Delete Failed. Err2 Msg: ' + err2 + ', DocUrl: ' + DocURL);
    Step = 3100; // if the above fails, next try and remove it
    try {
      DriveApp.removeFile(DocFile);
      Logger.log(func + Step + ' Remove Successful: ' + DocURL);
      return true;
    } catch (err3) {
      oCommon.DisplayMessage = ' Unable to Delete/Remove file.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage + ', Err3 Msg: ' + err3;
      Logger.log(oCommon.ReturnMessage);
      return false;
    }
  }
}

function LogEvent(EventMsg, oCommon) {
  var func = "***LogEvent - " ;
  
  var Step = 1000;
  if (ParamCheck(oCommon.TestMode) && oCommon.TestMode.toUpperCase().indexOf('NO EVENTS') > -1){
    Logger.log(func + Step + ' Event: "' + EventMsg);
    return;
  }
  Step = 1100;
  if (!ParamCheck(EventMsg)){
    Logger.log(func + Step + ' Event: "' + EventMsg);
    return;
  }
  Step = 1200;
  if (EventMsg == ''){
    Logger.log(func + Step + ' Event: "' + EventMsg);
    return;
  }
  
  Step = 1300;// remove unwanted carriage return characters
  EventMsg = EventMsg.replace(/\\n/g, "");
  
  Step = 2000; // Check IgnoreEventsFor
  if (ParamCheck(oCommon.IgnoreEventsFor)){
    if (oCommon.ScriptUser == Trim(oCommon.IgnoreEventsFor)){return;}
  }
  //Logger.log(func + Step + ' Event: "' + EventMsg);
  //Calculate the execution time string value, if applicable
  var strRunTime = RunTime(oCommon.StartTime);
  // Reset time for the next calculation
  oCommon.StartTime = new Date(); 
  
  var message = new Date() + ' (' + oCommon.ScriptUser + ') ' + EventMsg + strRunTime;
  oCommon.EventMessages.push(message);
  
  Logger.log(func + Step + ' Event: "' + EventMsg + '" posted.');
  
  return;  
    
}

function RunTime(StartTime) {
  // Useage:  strRunTime = RunTime(StartTime); // returns string e.g. 200 sec.
  if (isValidDate(StartTime)){
    var strRunTime = ' [' + Math.round((Number(new Date()) 
       - Number(StartTime)) / 1000).toString() + ' sec.]'; // in seconds
  } else { strRunTime = ""; }
  return strRunTime;
}

function WriteEventMessages(EventMsg, oCommon) {
  var func = "***WrtEventMsgs - " ;
  
  /******************************************************************************/
  var Step = 1000;// Writing Status Messages to Event Messages Tab
  /******************************************************************************/
  if (oCommon.TestMode.toUpperCase().indexOf('NO EVENTS') > -1){return;}
  if (ParamCheck(EventMsg)){ LogEvent(EventMsg, oCommon); }
  var EventCount = oCommon.EventMessages.length;
  Logger.log(func + Step + ' EventCount: ' + EventCount + ', EventMsg: ' + EventMsg);
  if (!ParamCheck(EventCount) || EventCount <= 0 ){return;}
  
  Logger.log(func + Step + ' EventCount: ' + EventCount + ', EventMsg: ' + EventMsg);

  /******************************************************************************/
  Step = 2000;// Capture Warning and/or Error Messages to be written and sent to the script.manager
  /******************************************************************************/
  var outputRow = [];
  var ErrorMessages = '';
  for (var k = 0; k < EventCount; k++) {
    outputRow[k] = new Array();
    outputRow[k][0] = oCommon.EventMessages[k];
    Logger.log(func + Step + ' k: ' + k + ', outputRow[k][0]: ' + outputRow[k][0]);
    if(outputRow[k][0].indexOf("ERROR") > -1 && outputRow[k][0].toUpperCase().indexOf("ERRORS AND/OR WARNINGS") <= -1){
      ErrorMessages = ErrorMessages + '\r\n' + outputRow[k][0];
    }
  }
  // Compute the total script run time an append it to the last message
  var strTotalTime = RunTime(oCommon.TotalTime);
  outputRow[k-1][0] = outputRow[k-1][0] + strTotalTime;

  /******************************************************************************/
  Step = 3000; // Send an email noification for any detected errors 
  /******************************************************************************/
  if (ErrorMessages != ''){
    var Recipient = Trim(oCommon.Send_error_report_to);
    if (typeof Recipient != 'undefined' && Recipient != null && Recipient != ''){
      try {
        MailApp.sendEmail(Recipient, "Event Error Report ", 
                          "\r\nApplication: " + oCommon.ApplicationName 
                          + "\r\nScript Version: " + oCommon.ScriptVersion
                          + "\r\nErrors: " + ErrorMessages
                          + "\r\n*********** END ***********");
        Logger.log(func + Step + ' - Email Sent (' + Recipient + ') ErrorMessages: ' + ErrorMessages);
      } catch (err) {
        var message = new Date() + ' (' + oCommon.ScriptUser + ') ' + func + Step 
          + ' Error - Unable to Send Email (' + Recipient + ') ErrorMessage: ' + err.message;   
        Logger.log(func + Step + ' Error - Unable to Send Email (' + Recipient + ') ErrorMessage: ' + err.message);
        outputRow.push([message]);   
      }
    }
  }
  
  /******************************************************************************/
  Step = 4000; // Write the Event messages to the Event Messages Tab, if enabled
  /******************************************************************************/
  if (!ParamCheck(oCommon.EventMesssagesTab)){ return; }
  
  var MessagesSheet = oCommon.Sheets.getSheetByName(oCommon.EventMesssagesTab);
  var MessagesData = MessagesSheet.getDataRange().getValues();
  
  Step = 4100; // wGet a public lock, one that locks for all invocations
  // http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  //var lock = LockService.getPublicLock();
  //lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
    
    Step = 4100; // Delete rows(s) if there are more event messages than the maximum allowed
    var first_write_row = MessagesData.length + 1 ;
    if (ParamCheck(oCommon.MaxEventMessages) && oCommon.MaxEventMessages != ''){
      Step = 4110; // Gather the extents
      var first_message_row = GetSheetParam(oCommon.EventMesssagesTab, '1st Data Row', oCommon)+1;
      var current_last_row = MessagesData.length ;
      var num_2b_added = outputRow.length;
      var MaxEventMessages = +oCommon.MaxEventMessages;
      var num_2b_deleted = current_last_row - MaxEventMessages + num_2b_added;
      Logger.log(func + Step + ' first_write_row: ' + first_write_row + ' first_message_row: ' + first_message_row);
      Logger.log(func + Step + ' cur last: ' + current_last_row + ', num_2b_added: ' + num_2b_added
                 + ', num_2b_deleted: ' + num_2b_deleted );
      
      Step = 4120; // Test to determine if any rows need to be deleted
      if (num_2b_deleted > 0){
        
        Step = 4130; // Delete num_rows_to_be_deleted using sheet.deleteRows(rowPosition, howMany)
        MessagesSheet.deleteRows(first_message_row,num_2b_deleted);
        
        // Adjust the starting row for the new entries
        first_write_row = current_last_row - num_2b_deleted + 1;
      }
    }
    
    /******************************************************************************/
    Step = 5000; // Write new events 
    /******************************************************************************/
    // Use .getRange(row, column, numRows, numColumns) 
    Logger.log(func + Step + ' range values: (' + first_write_row + ', 1, ' + outputRow.length + ', ' + outputRow[0].length +')');
    MessagesSheet.getRange(first_write_row,1,outputRow.length,outputRow[0].length).setValues(outputRow);
    
    SpreadsheetApp.flush();

    // Release the public lock
    //lock.releaseLock();
    
  } catch (err) {
    // Release the public lock
    //lock.releaseLock();
    Logger.log(func + Step + ' Error writing to the EventMessages Tab: ' + err.message);
  }
  
  /******************************************************************************/
  Step = 5000; // Reset oCommon.EventMessages array list and Hide the Messages Tab
  /******************************************************************************/
  oCommon.EventMessages = [];
  MessagesSheet.hideSheet();
  
  return;
 
}

