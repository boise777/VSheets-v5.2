/***********************************************************************************

oCommon object list for easy reference:
  var oCommon = {
    Globals: Globals,
    Sheets: Sheets, 
    ValidDomain: Globals['validDomain'],
    ApplicationName: Sheets.getName(),
    App_Id: Globals['App_Id'],
    ScriptVersion: Version,
    ScriptManager: Globals['ScriptManager'],
    Using_TeamDrive: Globals['Using_TeamDrive'],
    StartTime: new Date(),
    TotalTime: new Date(),
    ScriptUser: '',
    sheetParams: {},
    HeadingsXrefAry: {},
    TermsXrefAry: {},
    RunCheckXrefAry: {},
    XrefSourceSheet: null,
    FormTimestamp: null,
    FormSubmit_Delay:Globals['FormSubmit_Delay'],
    onSubmit_Sensitivity:Globals['onSubmit_Sensitivity'],
    bActiveUserEmpty: true,
      
    SelectedSheetRows: arySelectedRows,
    ChangedSheetRows: aryChangedRows,
    
    TestMode: Globals['Test_Mode'],
    EventMessages: aryEventMessages,
    CallingMenuItem: null,
    bSilentMode: true,
    ReturnMessage: '',
    DisplayMessage: '',
    
    GoogleForm_Url: Globals['GoogleForm_URL'],
    GoogleForm_Name: Globals['GoogleForm_Name'],
    FormSubmit_Delay: Globals['FormSubmit_Delay'],
    
    SetupSheetName: Globals['SetupSheetName'],
    SetupStartRow: +Globals['SetupSheetStartRow']-1,
    
    EventMesssagesTab: Globals['EventMesssagesTab'],
    MaxEventMessages: +Globals['MaxEventMessages'],
    IgnoreEventsFor: Globals['IgnoreEventsFor'],
    Send_error_report_to: Globals['send_error_report_to'],
    
    FormResponse_SheetName: Globals['FormResponse_SheetName'],
    FormResponse_HeadingRow: null,
    FormResponse_TermsRow: null,
    AwesomeTableRow: null,
    FormResponseStartRow: null,
    TimestampCol: +Globals[Globals['Timestamp_heading']],
    CheckForValidTimestamp: Globals['CheckForValidTimestamp'],
    Form_Owner_Col: +Globals[Globals['Form_Owner_Heading']],
    Record_statusCol: +Globals[Globals['Record_status_heading']],
    Active_Status_value: Globals['Active_Status_value'],
    Inactive_Status_value: Globals['Inactive_Status_value'],
    RecordTagValueCol: +Globals[Globals['Record_Tag_Value_heading']],
    RecordTagMaxLength: Globals['Record_Tag_Max_Length'],
    AsOfDateCol: +Globals[Globals['As_Of_Date_Heading']],
    Edit_URLCol: +Globals[Globals['Edit_URLCol_heading']],
    Edit_ButtonCol: +Globals[Globals['Edit_Button_heading']],
    Edit_Button_Text: Globals['Edit_Button_Text'],
    
    ListItemsSheetName: Globals['ListItemsSheetName'],
    
    BuildMenuSection_Heading: Globals['BuildMenuSection_Heading'], 
    Menu_Title: Globals['Menu_Title'], 
    SheetDetails_Heading: Globals['SheetDetails_Heading'], 
    DataConversions_Heading: Globals['DataConversions_Heading'],
    ValidationSection_Heading: Globals['ValidationSection_Heading'],
    AuditSection_Heading: Globals['AuditSection_Heading'],
    DocumentsSection_Heading: Globals['DocumentsSection_Heading'],
    Default_Doc_Folder_URL: Globals['Default_Doc_Folder_URL'],
    Replace_Term_Prefix: Globals['Replace_Term_Prefix'],
    Replace_Term_Suffix: Globals['Replace_Term_Suffix'],
    EmailSection_Heading: Globals['EmailSection_Heading'],
    ParseSymbol: Globals['ParseSymbol'],
    PersonalMessage: "",

    TransactionTypeSection_Heading: Globals['TransactionTypeSection_Heading'],
    BankTransactionSection_Heading: Globals['BankTransactionSection_Heading'],
    SummarySheetSection_Heading: Globals['SummarySheetSection_Heading'],
    CheckSum_Col: Globals['CheckSum_Col']-1
  }


Noteworthy Script Examples:
  
  Validation Check Boxes
      InActiveStatusChar = 9744, // Ascii value of an "empty" checkbox symbol
      ActiveStatusChar = 9745;   // Ascii value of a "checked" checkbox symbol

  jdoc @param syntax: http://usejsdoc.org/tags-param.html (JSDoc 3 is an API documentation generator)
  https://www.w3schools.com/js/default.asp
  
  Simple data field validation:
  if (typeof firstDataRow === 'undefined' || firstDataRow === null || firstDataRow === ''){
    var EventMsg = func + Step + ' ERROR - firstDataRow not Valid';
    LogEvent(EventMsg, oCommon);
    return false;
  }
 
  Test for a valid number: if(!isNaN(editButtonCol)){}...https://www.w3schools.com/jsref/jsref_isnan.asp
  
  // OKAY TO USE THIS EXAMPLE or code based on it.
  var cell = sheet.getRange('a1');
  var colors = new Array(100);
  for (var y = 0; y < 100; y++) {
    xcoord = xmin;
    colors[y] = new Array(100);
    for (var x = 0; x < 100; x++) {
      colors[y][x] = getColorFromCoordinates(xcoord, ycoord);
      xcoord += xincrement;
    }
    ycoord -= yincrement;
  }
  sheet.getRange(1, 1, 100, 100).setBackgroundColors(colors);
  SpreadsheetApp.flush();
  
DATE MANIPULATION

  // Difference between two dates
  // Reference: https://developers.google.com/google-ads/scripts/docs/features/dates
  var dF = new Date(FormTimestamp).getTime();
  var dS = new Date(sheetTimestamp).getTime();
  var mDiff = Math.abs((dF - dS)/1000);  // time difference in seconds
  
  // This formats the date as Greenwich Mean Time in the format
  // year-month-dateThour-minute-second.
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");

      CellAddr = numToA(ActionDateCol+1) + String(ArrayRow+1).split(".")[0];
      var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
      var formattedDate = Utilities.formatDate(new Date(), timeZone, "M/d/yyy HH:mm:ss");
      SourceSheet.getRange(CellAddr).setValue(formattedDate); 
      SourceSheet.getRange(CellAddr).setNumberFormat('M/d/yyyy H:mm:ss');
  
  
STRING MANIPULATION 

  String w/in string: var res = str.replace(/blue/g, "red");
  String w/in string: var n = str.indexOf("welcome"); < -1
  
  String.substr(startPos, endPos)
  str.toUpperCase()
  var n = num.toString()
  var sheets = ss.getSheets();
  for (var sheetNum = 1; sheetNum < sheets.length; sheetNum++)
  var partsOfStr = str.split(',');

  str.indexOf(String target) -- searches left-to-right for target
  Returns index where found, or -1 if not found
  Use to find the first (leftmost) instance of target
  str.lastIndexOf(String target) -- searches right-to-left instead  
  
 // This formats the date as Greenwich Mean Time in the format
 // year-month-dateThour-minute-second.
 var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  .hideSheet();
  .showSheet()
  
  // we want a public lock, one that locks for all invocations
  // http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  // got the lock, we may now proceed

  // Release the public lock and leave
  lock.releaseLock();


  jdoc @param syntax: http://usejsdoc.org/tags-param.html (JSDoc 3 is an API documentation generator)
  https://www.w3schools.com/js/default.asp
  
  // OKAY TO USE THIS EXAMPLE or code based on it.
  var cell = sheet.getRange('a1');
  var colors = new Array(100);
  for (var y = 0; y < 100; y++) {
    xcoord = xmin;
    colors[y] = new Array(100);
    for (var x = 0; x < 100; x++) {
      colors[y][x] = getColorFromCoordinates(xcoord, ycoord);
      xcoord += xincrement;
    }
    ycoord -= yincrement;
  }
  sheet.getRange(1, 1, 100, 100).setBackgroundColors(colors);
  SpreadsheetApp.flush();
  String.substr(startPos, endPos)
  String w/in string: var n = str.indexOf("welcome");
  str.toUpperCase()
  var n = num.toString()
  var sheets = ss.getSheets();
  for (var sheetNum = 1; sheetNum < sheets.length; sheetNum++)
  var partsOfStr = str.split(',');

  str.indexOf(String target) -- searches left-to-right for target
  Returns index where found, or -1 if not found
  Use to find the first (leftmost) instance of target
  str.lastIndexOf(String target) -- searches right-to-left instead  
  
  
  
  SpreadsheetApp.flush();

******************************************************************************/
