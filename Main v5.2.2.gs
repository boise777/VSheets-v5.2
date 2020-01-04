var Version = '(v5.2.2)';
/********************************* Build 12/04/2019 (rev 1) **********************************

 DESCRIPTION:
 
    This collection of scripts is my attempt to creat a useful library
    of common scripts that can be used across different Form Handling applications. 
    
 PRIMARY FUNCTIONS   
     
    onFormSubmit(oCommon)
     This function is executed on every Form Submission event
     
    SelectSheetRows(oCommon,mode,TargetSheetName,firstDataRow){
      This function is invoked whenever another function needs to process ALL or SELECTED 
        contiguous rows in the target "dataTab" (passed by Globals) found in the active Google Sheet.
        
    GenerateReport(oCommon,SelectedSheetRows,ReportTitle)
      This procedure is used to generate a Google Doc of PDF Doc from the submitted 
        Form Responses using the input Google Doc template for the given report title.
        The parameters for the input ReportTitle are normally defined in a Setup Tab
     
    BuildGlobals(oSourceSheets) {
      Reads the Setup parameters and builds the Globals script properties
      
    BuildCommon()
      This function is used to build an object containing the common parameter values
        used by the other functions in this library.
  
    BuildGlobals(oSourceSheets)
      Reads the Setup parameters and builds the Globals script properties
    
    ParamSetBuilder(oCommon, ParameterSet)
      This function is used to build an object containing the Document parameter values
        used by the other functions in this library.
    
    FormatAwesomeTableParameters(oCommon)
      This little utility was developed to verify that the Awesome Table "VIEW" and "Edit" 
        button template string are pointing to the correct sheet columns containing the URLs   
    
    updateFormFields(oCommon)
      Create / update the custom Drop-Down and CheckBox lists from spreadsheet lists

    assignEditUrls(oCommon) 
      This function is invoked whenever a "new" form submission is made:
         1 - Add / Update the Edit URL
         2 - Write the Employment Status text: "Active" if needed, and
         3 - Write the EDIT url and VIEW button text values used by "Awesome Table" used for data display
         4 - Perform "Data Converstions" as defined in the Setup Sheet

    ValidationScan(oCommon) {
      This function is invoked whenever a "validation" of data is needed to test for:
         1 - expiration dates (Expiry)
         2 - required response to a True of False Form response (T/F)
         3 - required text "string" response to a Form response (InStr)
         4 - valid date test

    PerformAudit(SourceSS, Globals, AuditTitle, SelectedSheetRows, AuditFindings)
      This function is used to verify that entries made in the Form Responses Tab
       (FormData Tab) match corresponding entries place in another "AuditSource" Tab
       
    deleteSelectedRows(oCommon0) 
      This function is invoked whenever one or more sheet rows are to be deleted:
         1 - Locate the target sheet row(s)
         2 - Delete the associated Google Form Response (if any)
         3 - Delete the selected sheet row(s)
       
    updateSelectedRows(oCommon) - NOT IMPLEMENTED 
      This function is invoked whenever a Form Response(s) needs to be updated with "manual"
      entries made to one or more sheet rows:
         1 - Locate the target sheet row(s)
         2 - Find the associated Google Form Response (if any)
         3 - Update the associated Google Form Response (or create a new one)

 AUTHOR:
    Steve Germain
    (208) 870-2727
    
 COPYRIGHT NOTICE:
   Copyright (c) 2009 - 2018 by Legatus, Inc.
   Copyright (c) 2019 - Stephen P. Germain
   
 NOTES:
    The function assignEditUrls() is invoked whenever a Form Submission is completed.
    The Awesome Table (https://awesome-table.com/) service is utilized for data display.
    
    
  
*********************************************************************************/

function onOpenProcedures(oCommon) {
  /* ****************************************************************************
   DESCRIPTION:
     Defines the Name of the Setup Sheet
     Create the custom "Reports" menu
     Load the Global scalars found in the Setup sheet
     Format the Awesome Table Edit and View button parameters

   USEAGE:
     ReturnBool = onOpenProcedures(oCommon); 

   REVISION DATE:
     07-06-2017 - Initial design
     01-03-2018 - Modified to use the Script Library 1.1 set of commonly used form handling functions
     01-27-2018 - Modified to use the Script Library 2.3 set of commonly used form handling functions
                - Added the BuildMenu function
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback
   
   NOTES:
   
   
  ******************************************************************************/
  
  var func = "***onOpenProcedures " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  var bNoErrors = true;
  
  Step = 1000; // Log the authorization status (REQUIRED or NOT_REQUIRED).
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  Logger.log(func + Step + ' Authorization status: ' + authInfo.getAuthorizationStatus());  
  
  Logger.log(func + Step + ' BuildMenuSection_Heading: ' + oCommon.BuildMenuSection_Heading
            + ', ListItemsSheetName: ' + oCommon.ListItemsSheetName  
            + ', AwesomeTableRow: ' + oCommon.AwesomeTableRow);  
  
  /******************************************************************************/
  Step = 2000; // Structure and create custom menu
  /*******************************************************************************/
  if (ParamCheck(oCommon.BuildMenuSection_Heading)){
    if (!BuildMenu(oCommon)){
      bNoErrors = false;
      Logger.log(func + Step + ' BuildMenu(): ' + oCommon.ReturnMessage);  
    }
  }

  /******************************************************************************/
  Step = 3000; // Update Item List in the attached Google Form, if required
  /*******************************************************************************/
  if (ParamCheck(oCommon.ListItemsSheetName)){
    if (!updateFormLists(oCommon)){
      //oCommon.ReturnMessage = oCommon.ReturnMessage + 'updateFormLists errors...';
      bNoErrors = false;
      Logger.log(func + Step + ' updateFormLists(): ' + oCommon.ReturnMessage);  
    }
  }
  
  /*******************************************************************************/
  Step = 4000; // Reformat Edit & View Buttons just in case column changes were made (Awesome Table paramteters)
  /*******************************************************************************/
  if (!isNaN(oCommon.AwesomeTableRow)){
    if (!FormatAwesomeTableParameters(oCommon)){
      //oCommon.ReturnMessage = oCommon.ReturnMessage + 'formatAwesomeTableParameters errors...';
      bNoErrors = false;
      Logger.log(func + Step + ' FormatAwesomeTableParameters(): ' + oCommon.ReturnMessage);  
    }
  }
  
  // Update Library counters
  var scriptProperties = PropertiesService.getScriptProperties();
  var opens_count = Number(scriptProperties.getProperty('B_Opens_count')) + 1;
  scriptProperties.setProperty('A_Opens_count', opens_count);
  
  
  if (bNoErrors){
    LogEvent(func + 'Completed successfully.', oCommon);
  } 
  
  return bNoErrors;
  
}  


function onFormSubmit(oCommon, oMenuParams) {

  /* ****************************************************************************
  DESCRIPTION:
     This function is executed on every Form Submission event
     
  USEAGE:
     ReturnBool = onFormSubmit(oCommon);
  
  REVISION DATE:
     07-07-2017 - Implemented structured function approach
     12-21-2017 - Modified to use "oCommon" methods
     03-23-2018 - Modified to work with v3.x
                - Modified to return boolean
     11-04-2018 - Added SubmissionAlertTemplate email in Step 4200
     11-12-2018 - Modified Step 4200 to send a Submission Alert & Confirmation emails
                    only if the "owner" of the form just made this submission
                    
  
  NOTES:
  ******************************************************************************/
  
  var func = "**onFormSubmit " + Version + " - ";
  var Step = 100;
  
  /******************************************************************************/
  Step = 1000; // Initialize variables
  /******************************************************************************/
  var ScriptUser = oCommon.ScriptUser;
  var bActiveUserEmpty = oCommon.bActiveUserEmpty;
  var bNoErrors = true;
  var triggerUid = oCommon.FormTriggerUid;
  // Log the email address of the person running the script.
  Logger.log(func + Step + ' triggerUid: ' + triggerUid + ', Script User: ' + ScriptUser);
  
  /******************************************************************************/
  Step = 2000; // Sleep to allow time for a new entry to be placed in the FormResponse Sheet
  /******************************************************************************/
  //if(oCommon.FormSubmit_Delay > 0){
  //  Utilities.sleep(oCommon.FormSubmit_Delay); 
  //  SpreadsheetApp.flush();
  //}
  
  /******************************************************************************/
  Step = 3000; // Log Successful form submission and execute next task
  /******************************************************************************/
  if(assignEditUrls(oCommon)){
    var EventMsg = func + Step + ' Successful onSubmit(' + triggerUid + ') Submission (' 
        + oCommon.ChangedSheetRows[0][1] + ', row ' + oCommon.ChangedSheetRows[0][0] + ')';
    LogEvent(EventMsg, oCommon);
    Logger.log(EventMsg);
    Logger.log(''); // blank line for readability
    
    /******************************************************************************/
    Step = 4000; // Generate the Desired Emails & Reports
    /******************************************************************************/
    oCommon.SelectedSheetRows = oCommon.ChangedSheetRows;
    oCommon.ChangedSheetRows = [];
    if(GenerateItems(oCommon, oMenuParams)){
      Step = 4100; // Log Successful Docs & Emails(s)
      var EventMsg = func + Step + ' Docs and/or emails completed for Submission (' 
        + oCommon.SelectedSheetRows[0][1] + ', row ' + oCommon.SelectedSheetRows[0][0] + ')';
      //LogEvent(EventMsg, oCommon);
      Logger.log(EventMsg);
    } else {
      Step = 4200; // Problematic End to GenerateItems()
      // Error messages are captured in the preceeding function
      Logger.log(func + Step + ' ERROR - Return Message: ' + oCommon.ReturnMessage);
      bNoErrors = false;
    }
  } else {
    Step = 3200; // Problematic End to assignEditUrls()
    // Error messages are captured in the preceeding function
    Logger.log(func + Step + ' ERROR - Return Message: ' + oCommon.ReturnMessage);
    bNoErrors = false;
  }
  
  /******************************************************************************/
  Step = 5000; // Log Results and leave...
  /******************************************************************************/
  if (bNoErrors){
    Logger.log(func + Step + ' Form Submission (' + triggerUid + ') completed successfully.');
  } else {
    Logger.log(func + Step + ' Form Submission (' + triggerUid + ') encountered errors.');
  }

  return bNoErrors;
  
}


function GenerateItems(oCommon, oMenuParams) {
  /* ****************************************************************************

   DESCRIPTION:
     This function is called to execute the actions specified within the passed oMenuParams
     parameters: Send emails and/or generate documents.
     
   USEAGE
   
     var bReturnStatus = GenerateItems(oCommon, oMenuParams);
     
   INPUT PARAMETERS
   
     var bSilentMode is global scalar declared in TopLevel code
     
     Array: oCommon.SelectedSheetRows[n]
       oCommon.SelectedSheetRows[n][0] = SheetData Target row
       oCommon.SelectedSheetRows[n][1] = SheetData TagValue
       oCommon.SelectedSheetRows[n][2] = message / email body
     
     Array: oMenuParams (Reference the Setup Tab, Menu Build Section)
     
       oMenuParams['Doc 1'] = Doc 1 parameter value
       ...
       oMenuParams['Email 3'] = Email 3 parameter value
     
/******************************************************************************/
  var func = "*** GenerateItems " + Version + " - ";
  var Step = 100;
  
  /******************************************************************************/
  Step = 1000; // Verify presence of oMenuParams
  /******************************************************************************/
  Step = 1010; // Validate oMenuParams
  if (oMenuParams.length <= 0){
    oCommon.DisplayMessage = ' Warning 230 - Expected oMenu Parameters not found ('
      + ParameterSet + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);     
    return true;
  }
  Step = 1020; // Do not execute EMAILs if the active sheet is not the Form Response Sheet
  for (var key in oMenuParams) {
    if (Trim(key).toUpperCase().indexOf('EMAIL') > -1){
      //Email template Parameter Key found
      var ActiveTab = oCommon.Sheets.getActiveSheet().getName();
      if (ActiveTab !== oCommon.FormResponse_SheetName){
        oCommon.DisplayMessage = 'Error 231 - This function can only be executed in the Form Response tab '
          + '(Active Sheet: "' + ActiveTab + '", Expected Sheet: "' +  oCommon.FormResponse_SheetName + '").';
        oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
        Logger.log(oCommon.ReturnMessage);
        if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
        return false;
      }
    }
  }

  Step = 1030; // Determine need for user to select records
  if(oCommon.SelectedSheetRows.length <= 0){
    var bUserSelectsRows = true;
  } else {
    var bUserSelectsRows = false;
  }
  Logger.log(func + Step + ' SelectedSheetRows.length:' + oCommon.SelectedSheetRows.length);
  
  /**********************************************************************************/
  Step = 1100; // Assign Variables and prepare objects
  /**********************************************************************************/
  // Communicate with user
  if (oCommon.CallingMenuItem = null){
    var title = oCommon.CallingMenuItem;
  } else {
      var title = 'Preparing Requested Items';
  }
  var prog_message = 'Determining scope of work ...';
  progressMsg(prog_message,title,-3);
  
    
  Step = 1200; // Open the "active" sheet and get the row info
  var oSourceSheets = oCommon.Sheets;
  //var Globals = oCommon.Globals;
  if (!oCommon.bSilentMode){
    // Get the "active" sheet
    var sheet = oSourceSheets.getActiveSheet();
    //var SheetData = sheet.getDataRange().getValues();
    var dataTab = oSourceSheets.getActiveSheet().getName();
    
  } else {
    //var sheet = oSourceSheets.getSheetByName(dataTab);    
    var dataTab = oCommon.FormResponse_SheetName;
    //var firstDataRow = oCommon.FormResponseStartRow;
  }
  var firstDataRow = GetSheetParam(dataTab, '1st Data Row', oCommon);
  
  var scriptProperties = PropertiesService.getScriptProperties();

  /**************************************************************************/
  Step = 2000; // Select records, if necessary
  /**************************************************************************/
  // Examine input RecordRetrievalMode value (Ref Step 1020, above)
  if (bUserSelectsRows){
    Step = 2100; // User needs to select rows
    //bSilentMode = false; // bSilentMode declared in TopLevel code 
    var mode = 2;
    if(!SelectSheetRows(oCommon, mode, dataTab, firstDataRow)){
      Step = 2110; // ERROR - No Selected Rows found
      oCommon.DisplayMessage = ' Error 232 - No Rows selected. Procedure will be terminated.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      return false;
    }
    Logger.log(func + Step + ' SelectedSheetRows.length:' + oCommon.SelectedSheetRows.length);
    
    Step = 2200; // Take time to confirm that the user wants to proceed
    var MsgBoxMessage = ('The following rows were selected: ' + FormatForDisplay(oCommon.SelectedSheetRows)
       + '\\n' + ' Do you want to continue?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      var display_message = "Your response is NO - procedure will be terminated";
      return_message = func + Step + ' User terminated after rows were selected';
      oCommon.ReturnMessage = return_message;
      oCommon.DisplayMessage = display_message;
      return false;
    }
  } 
  
  /**************************************************************************/
  Step = 3000; //Setup and execute requested items passed in the oMenuParams list
  /**************************************************************************/
  var emails_attempted = 0;
  var emails_completed = 0;
  var docs_attempted = 0;
  var docs_completed = 0;
  for (var key in oMenuParams) {
    if (Trim(key).toUpperCase().indexOf('EMAIL') > -1){
      //Email template Parameter Key found
      if (ParamCheck(oMenuParams[key])){
        //Email template name found
        var EmailTemplateName = oMenuParams[key];
        Step = 3100; // Send the specified email
        emails_attempted++;
       
        // Take time to get the personal message from the user
        if (bUserSelectsRows){
          oCommon.PersonalMessage = Browser.inputBox('ADD PERSONAL MESSAGE?', 'Enter your message', Browser.Buttons.YES_NO);
        }
         
        if(SendEmail(oCommon, EmailTemplateName)){
          emails_completed++; 
          prog_message = ' Completed ' + emails_completed + ' of ' + emails_attempted + ' emails attempted.';
          var EventMsg = func + Step + prog_message;
          //LogEvent(EventMsg, oCommon);
          Logger.log(func + Step + prog_message);
          
  
        } else {
          // Something went wrong
          prog_message = oCommon.ReturnMessage;
          progressMsg(prog_message,title,-3);
        }
      }
    } else if (Trim(key).toUpperCase().indexOf('DOC') > -1){
      //Document template Parameter Key found
      if (ParamCheck(oMenuParams[key])){
        //Document template name found
        var DocTemplateName = oMenuParams[key];
        Step = 3200; // Generate the specified Document
        docs_attempted++;
        
        if(WriteDocument(oCommon, DocTemplateName)){
          docs_completed++;
          prog_message = ' Completed ' + docs_completed + ' of ' + docs_attempted + ' docs attempted.';
          Logger.log(func + Step + prog_message);
          

        } else {
          // Something went wrong
          prog_message = oCommon.ReturnMessage;
          progressMsg(prog_message,title,-3);
        }
      }
    }
  }
  
  /**************************************************************************/
  Step = 4000; //Wrap up and go home...
  /**************************************************************************/
  Logger.log(func + Step + ' END');
  
  return true;
  
}
   
function SendEmail(oCommon, EmailTemplateName){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to format and send email
     
   USEAGE
     var bStatus = SendEmail(oCommon);

   REVISION DATE:
    07-22-2017 - Add messageTemplateSet and adminTemplateSet as input parameters
    07-25-2017 - do not send Admin Summary email if Admin_template_addr == ''
    10/14/2017 - eliminated "admin emails" and simplified code
    11-11-2017 - added provision for "personal_message" (ref: sendEditUrl) 
    12-26-2017 - Modified to use the oCommon object and acquire the email parameters
                  and template from the Setup "EmailTemplates" section.
    01-22-2018 - Modified code to accept and input template ID or URL value
    01-25-2018 - Modified to return an error message / number of emails sent
    11-04-2018 - Modified Step 3000 to default to use the "To" email addresss when defined 
                 with the template, otherwise use the Primary_email address from the input sheet row
    11-11-2018 - Added dynamic Primary and CC email address capabilities by adding Steps 2400 and 2500
                   and modifying Step 3000 to look for an email column heading value 
                   in the "To" and "CC" email addresss fields
    01-23-2019 - Added code to reformat date values (Ref Step 3440)
    02-04-2019 - Consolidated all in/out parameters into oCommon and only return T/F
    02-13-2019 - Using new GetMenuParams() function in the Top Level calling function to pass the
                   email template names(s) in the oMenuParams array
               - Eliminating default use of 
                    Secondary_emailCol = oCommon.SecondaryEmailCol,
                 The email address will now be taken from the entries(s) in the "To" col of 
                    the Template Parameters
    11/08/2019 - Added var SourceTab = oCommon.AlertsScanSource as additional check in Step 1010
                    so that WeeklyScan results will initiate alert emails
               - Reference AlertsScan() Step 1040
    11/14/2019 - Added Step 6414 to fix Email Subject formatting problem causing the subject to remain the same
                    for every record after the first - the Email_subject was not reinitialized for each row. 

   NOTES:
   
     Input Parameter Values:
     
       Array: oCommon.SelectedSheetRows[n]
         oCommon.SelectedSheetRows[n][0] = SheetData Target row for each recipient
         oCommon.SelectedSheetRows[n][1] = TagValue
         oCommon.SelectedSheetRows[n][2] = message / email body
         
       EmailTemplateName - text value of the template to be used for the email(s)
   
     Output Parameter Values:
     
       Array: oCommon.SelectedSheetRows[n] (if not passed by the calling function)
         oCommon.SelectedSheetRows[n][0] = SheetData Target row for each recipient
         oCommon.SelectedSheetRows[n][1] = Record Tag Value
         oCommon.SelectedSheetRows[n][2] = message / email body
         Added:
           oCommon.SelectedSheetRows[n][3] = recipient emailAddress
   
     Check out: https://developers.google.com/apps-script/articles/mail_merge
                https://webapps.stackexchange.com/questions/85017/utilizing-noreply-option-on-google-script-sendemail-function
    
  ******************************************************************************/
  
  var func = "****SendEmails " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  var test_mode = false; //true;
  if (oCommon.CallingMenuItem = null){
    var title = oCommon.CallingMenuItem;
  } else {
      var title = 'Send Emails';
  }
  
  /******************************************************************************/
  Step = 1000; // Examine and vaildate input parameters
  /******************************************************************************/
  Step = 1010; // Do not execute if the active sheet is not the Form Response Sheet
  var SourceTab = oCommon.AlertsScanSource;
  var ActiveTab = oCommon.Sheets.getActiveSheet().getName();
  if (oCommon.bSilentMode) {
    if (ActiveTab !== oCommon.FormResponse_SheetName && SourceTab !== oCommon.FormResponse_SheetName){
      oCommon.DisplayMessage = 'Error 312 - This function can only be executed in the Form Response tab '
      + '(Active Sheet: "' + ActiveTab + '", Expected Sheet: "' +  oCommon.FormResponse_SheetName + '").';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      Logger.log(oCommon.ReturnMessage);
      //if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
      return false;
    }
  }

  Step = 1020; // Validate EmailTemplateName(s)
  if (!ParamCheck(EmailTemplateName)){
    oCommon.DisplayMessage = 'Error 314 - Email Template Name missing or not Valid.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  }
  
  Step = 1030; // Validate oCommon.SelectedSheetRows
  if(oCommon.SelectedSheetRows.length <= 0){
    var bUserSelectsRows = true;
  } else {
    var bUserSelectsRows = false;
  }
  
  Logger.log(func + Step + ' SelectedSheetRows.length:' + oCommon.SelectedSheetRows.length);
  
  /**********************************************************************************/
  Step = 1100; // Assign Variables and prepare objects
  /**********************************************************************************/
  
  Step = 1010; // Set working scalar values
  var bSendEmail = true;
  var NumEmailsAttempted = 0;
  var NumEmailsCompleted = 0;
  var NumEmailsSent = 0;
  var return_message = '';
  var display_message = '';
  var bReturnStatus = false;
  var user_message = '';
  var todaysDate = new Date();
  var scanDate = formatDateTime(todaysDate); //Trim(scanTime.substring(0, 10)); // "6/13/2017 13:27:05" ==> "6/13/2017"
  var scanCount = 0;
  var sentCount = 0;
  var AdminMessage ='';
  var personal_message = '';

  Step = 1120; // Initialize output parameters
  oCommon.ItemsAttempted = 0;
  oCommon.ItemsCompleted = 0;
  oCommon.ReturnMessage = '';
  oCommon.DisplayMessage = '';
  oCommon.AdminTemplateName = '';
  
  Step = 1130; // Assign the oCommon parameters
  var oSourceSheets = oCommon.Sheets, 
      dataTab = oCommon.FormResponse_SheetName,
      firstDataRow = oCommon.FormResponseStartRow,
      timestampCol = oCommon.TimestampCol,
      StatusCol = oCommon.Record_statusCol,
      InactiveValue = oCommon.Inactive_Status_value,
      editUrlCol = oCommon.Edit_URLCol,
      error_email = oCommon.Send_error_report_to,
      ParseSymbol = oCommon.ParseSymbol,
      ParameterSet = oCommon.EmailSection_Heading
    ;
  
  Step = 1140; // Prepare objects for the Form Response data sheet
  var sheet = oSourceSheets.getSheetByName(dataTab);
  var SheetData = sheet.getDataRange().getValues();
 
  /*****************************************************************************/
  Step = 2000; // Load the Email Parameters
  /*****************************************************************************/
  var oEmailParams = {};
  oEmailParams = ParamSetBuilder(oCommon, ParameterSet);
  //Logger.log(func + Step + ' ParameterSet: ' + ParameterSet + ' Parameters: ' + oEmailParams.Parameters);
  // Return if the parameter set is empty or the email template is not found
  
  if(!ParamCheck(oEmailParams)){
    Step = 2010; // ERROR - No Common Paarameters found
    oCommon.DisplayMessage = ' Error 316 - Email Parameters (' + ParameterSet + ') Not Found.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  }
  // Verify results
  //Logger.log(func + Step + ' Verify: ');
  //for (var key in oEmailParams.HeadingKeyAry) {
  //  Logger.log('Key: %s, Value: %s', key, oEmailParams.HeadingKeyAry[key]);
  //} 
  
  Step = 2100; // Test to determine if email template is ACTIVE
  if (!ParamCheck(oEmailParams.TitleKeyAry[EmailTemplateName])){
    oCommon.DisplayMessage = ' Error 318 - Email Template ("' + EmailTemplateName 
      + '") Inactive or Not Found.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  }
  var active_row = oEmailParams.TitleKeyAry[EmailTemplateName];
  Logger.log(func + Step + ' EmailTemplateName: "' + EmailTemplateName + '"'
     + ', active_row: ' + active_row);

  Step = 2200; // Determine EmailTemplateID value
  var EmailTemplateUrl =  oEmailParams.Parameters[active_row][oEmailParams.HeadingKeyAry['Template ID']]
  if(!ParamCheck(EmailTemplateID)){
    Step = 2011; // Use URL value
    var EmailTemplateUrl = getIdFrom(oEmailParams.Parameters[active_row][oEmailParams.HeadingKeyAry['Template URL']]);
  } else {
    Step = 2210; // ERROR - No Email Template found
    oCommon.DisplayMessage = ' Error 320 - Template URL/ID for email (' + EmailTemplateName 
      + ') is Inactive or Not Found.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
    return false;
  }
  Step = 2300; // Extract the Google ID from the URL, if necessary
  var EmailTemplateID = getIdFrom(EmailTemplateUrl);
  Logger.log(func + Step + ' EmailTemplateName: "' + EmailTemplateName + '"'
     + ', EmailTemplateID: ' + EmailTemplateID);
  
  /******************************************************************************/
  Step = 3000; // Assign the input email parameter values
  /*******************************************************************************/
  Step = 3100; // Assign the input email parameter values
  var Subject_Template = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Subject']];
  var Email_to = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['To']];
  var Email_cc = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['CC']];
  var Email_bcc = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['BCC']];
  var DateMailSentColHdg = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Date Mail Sent Col']];
  var DateMailSentCol = GetXrefValue(oCommon, DateMailSentColHdg);
  var DateMailSentColLtr = numToA(DateMailSentCol+1); // adjust to Base 1
  var maxFrequency = +oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Max Frequency (Days)']];
  var Email_subject = '';
  
  Step = 3200; // Assign output values to oCommon
  //oCommon.SuccessMessage = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Success Message']];
  //oCommon.ErrorMessage = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Error Message']];
  oCommon.AdminTemplateName = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Admin Email Title']];

  Step = 3300; // Assign the input email parameter values
  //Add catch for depreciated use of "From" col heading
  var ReplyTo = '';
  if(ParamCheck(oEmailParams.HeadingKeyAry['From'])){
    ReplyTo = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['From']];
  } else {
    ReplyTo = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['ReplyTo']];    
  }

  Step = 3400; // Test for the use of a "Test" email address
  var TestEmailAddress = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Test Email Address']];
  if (ParamCheck(TestEmailAddress)){
    var test_mode = true;
    Logger.log(func + Step + ' TestEmailAddress: ' + TestEmailAddress + ', test_mode: ' + test_mode);
  } 
  
  /******************************************************************************/
  Step = 4000; // Validate the input email parameter values
  /*******************************************************************************/
  Step = 4100; // Determine the Email_to address to be used
  if (!ParamCheck(Email_to)){
    Step = 4110; // ERROR - No Email_to address found
    oCommon.DisplayMessage = ' Error 322 - Template (' + EmailTemplateName 
      + ') parameters specify no "To" address value.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  } else if(isValidEmail(Email_to)){
    Step = 4120; // Email_cc holds a static email address to be used
  } else if(!isNaN(GetXrefValue(oCommon, Email_to))){
    Step = 4130; // Email_to holds a valid column heading value
    var Email_toCol = GetXrefValue(oCommon, Email_to);
  } else {
    Step = 4140; // ERROR - Invalid Email_to address parameter found
    oCommon.DisplayMessage = ' Error 324 - Template (' + EmailTemplateName + ') contains invalid "TO" address value.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  }
  Logger.log(func + Step + ' bUserSelectsRows:' + bUserSelectsRows + ', Email_to: ' + Email_to 
             + ', Email_toCol: ' + Email_toCol);
             
  Step = 4200; // Verify that Email_cc holds a good value
  if (!ParamCheck(Email_cc)){
    Step = 4210; // Email_cc not used
    Email_cc = '';
  } else if(isValidEmail(Email_cc)){
    Step = 4220; // Email_cc holds a static email address to be used
  } else if(!isNaN(GetXrefValue(oCommon, Email_cc))){
    Step = 4230; // Email_to holds a valid column heading value
    var Email_ccCol = GetXrefValue(oCommon, Email_cc);
  } else {
    Step = 4240; // ERROR - Invalid Email_cc address parameter found
    oCommon.DisplayMessage = ' Error 326 - Template (' + EmailTemplateName + ') contains invalid "CC" address value.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  }
  Logger.log(func + Step + ' bUserSelectsRows:' + bUserSelectsRows + ', Email_cc: ' + Email_cc 
             + ', Email_ccCol: ' + Email_ccCol);
  
  Step = 4300; // Verify that Email_cc holds a good value
  if (!ParamCheck(Email_bcc)){
    Step = 4310; // Email_bcc not used
    Email_bcc = '';
  } else if(isValidEmail(Email_bcc)){
    Step = 4320; // Email_bcc holds a static email address to be used
  } else if(!isNaN(GetXrefValue(oCommon, Email_bcc))){
    Step = 4330; // Email_to holds a valid column heading value
    var Email_bccCol = GetXrefValue(oCommon, Email_bcc);
  } else {
    Step = 4340; // ERROR - Invalid Email_bcc address parameter found
    oCommon.DisplayMessage = ' Error 328 - Template (' + EmailTemplateName + ') contains invalid "BCC" address value.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  }
  Logger.log(func + Step + ' bUserSelectsRows:' + bUserSelectsRows + ', Email_bcc: ' + Email_bcc 
             + ', Email_bccCol: ' + Email_bccCol);
  
  Step = 4400; // Determine the "ReplyTo" Email_to address to be used
  if(!ParamCheck(ReplyTo)){
    Step = 4410; // If missing or empty, use the email address of the user sending the email(s)
    ReplyTo = oCommon.ScriptUser;
  } else if(ReplyTo.toUpperCase().indexOf("NOREPLY") > -1){
    Step = 4420; // to force emails to be sent from "noReply"
    ReplyTo = ''; 
  } else if (isValidEmail(ReplyTo)){
    Step = 4430; // Use the supplied value as the address
    ReplyTo = isValidEmail(ReplyTo);
  } else {
    Step = 4440;  // Reply To holds an invalid email address value
    oCommon.DisplayMessage = ' Error 330 - Template (' + EmailTemplateName + ') contains invalid "Reply To" address value.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(return_message, oCommon); }
    return false;
  }
  Logger.log(func + Step + ' bUserSelectsRows:' + bUserSelectsRows + ', ReplyTo: ' + ReplyTo); 

  /**************************************************************************/
  Step = 5000; // Select records, if necessary (Ref Step 1020, above)
  /**************************************************************************/
  if (bUserSelectsRows){
    Step = 5100; // User needs to select rows
    var mode = 2;
    bSilentMode = false;
    if(!SelectSheetRows(oCommon, mode, dataTab, firstDataRow)){
      Step = 3110; // ERROR - No Selected Rows found
      var return_message = func + Step + ' ********** Warning - No Rows Selected';
      return false;
    }
    
    Step = 5200; // Take time to confirm that the user wants to proceed
    var MsgBoxMessage = ('The following rows were selected: ' + FormatForDisplay(oCommon.SelectedSheetRows) 
       + '\\n' + ' Do you want to continue?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      var display_message = "Your response is NO - procedure will be terminated";
      return_message = func + Step + ' ********** WARNING - User terminated after rows were selected';
      oCommon.ReturnMessage = return_message;
      oCommon.DisplayMessage = display_message;
      // Communicate with user
      prog_message = 'User chooses not to continue.';
      progressMsg(prog_message,title,3);
      return false;
    }
     
    Step = 5300; // Take time to get the personal message from the user
    oCommon.PersonalMessage = Browser.inputBox('ADD PERSONAL MESSAGE?', 'Enter your message', Browser.Buttons.YES_NO);
  }   
  
  /**************************************************************************/
  Step = 6000; // Process each row passsed in oCommon.SelectedSheetRows array and prepare the email template for the given row
  /**************************************************************************/
  var total_attempted_messages = oCommon.SelectedSheetRows.length;
  var bypassed_for_freq_violation = 0;
  for (var k = 0; k < oCommon.SelectedSheetRows.length; k++){
    Step = 6100; 
    NumEmailsAttempted++;
    bSendEmail = true;
    scanCount++;
    var ArrayRow = oCommon.SelectedSheetRows[k][0] - 1;
    var attached_message = oCommon.SelectedSheetRows[k][2];
    var TagValue = SheetData[ArrayRow][oCommon.RecordTagValueCol];  
    TagValue = FormatTagValue(oCommon, TagValue);
    
    Logger.log(func + Step + ' *** PROCESSING EMAIL FOR ***** Row: ' + ArrayRow + ' (' + TagValue +')');
    Logger.log(func + Step + ' test_mode: ' + test_mode + ', maxFrequency: ' + maxFrequency + ', DateMailSentCol: ' + DateMailSentCol);
    /**************************************************************************/
    Step = 6200; // Test to determine if a previous email has already been sent within the frequency for this entry
    /**************************************************************************/
    if (ParamCheck(DateMailSentCol) && ParamCheck(maxFrequency)){
      var bFreqViolation = true;
      if(DateMailSentCol <= 0 || maxFrequency <= 0){
        Step = 6210; // No DateMailSentCol or maxFrequency value entered, therefore no need to check any further
        bFreqViolation = false;
      } else {
        // Additional checking is necessary; Get the date that the last email was sent
        var Sent_Date = SheetData[ArrayRow][DateMailSentCol]; 
        if (typeof Sent_Date === 'undefined' || Sent_Date === null){
          Step = 6220; // Sent_Date field is not defined, SendMail Frequency Violation = Unkown
          bFreqViolation = false;
          Logger.log(func + Step + ' Unknown SendMail Frequency Violation for Sheet Row: ' 
            + (ArrayRow+1) + ' Sent_Date value not found.  DateMailSentCol: ' + DateMailSentCol);
        } else {
          // Test for Sent_Date field value
          var num_days_between_emails = DateDiff(Sent_Date, todaysDate, "DAYS");
          if(num_days_between_emails === false){
            Step = 6230; // Sent_Date field is empty - no previously sent emails
            bFreqViolation = false;
          } else if(num_days_between_emails >= maxFrequency){
            Step = 6240; // num_days_between_emails >= maxFrequency
            bFreqViolation = false;
          }
        }
      }
      if(bFreqViolation){
        // SendMail Frequency Violation = True
        bSendEmail = false;
        bypassed_for_freq_violation++;
        Logger.log(func + Step + ' SendMail Frequency Violation for Sheet Row: ' + (ArrayRow+1) 
          + ' Sent_Date: ' + Sent_Date + ' DateDiff: ' + num_days_between_emails 
          + ' maxFrequency: ' + maxFrequency);
        // Alert User to Frequency violation for this row
        user_message = user_message + 'Email not sent for Sheet Row ' + (ArrayRow+1) + ' due to Frequency violation.'
            + '\\n' + '(Check Email Template settings if you believe that this is incorrect.)' + '\\n';
        
      } else {
        // SendMail Frequency Violation = False
        bSendEmail = true;
        Logger.log(func + Step + ' No SendMail Frequency Violation for Sheet Row: ' 
          + (ArrayRow+1) + ' Sent_Date: ' + Sent_Date + ' DateDiff: ' + num_days_between_emails
          + ' maxFrequency: ' + maxFrequency);
      }
    }
    Logger.log(func + Step + ' bSendEmailBool = ' + bSendEmail + ', bFreqViolation: ' + bFreqViolation
      + ' for Sheet Row: ' + (ArrayRow+1));
    
    /**************************************************************************/
    Step = 6300; // Test to determine if a new email is to be sent
    /**************************************************************************/
    if (bSendEmail){
      Step = 6310; // Get and test the Email_to address, if not already specified as a static address
      if(Email_toCol){
        Email_to = SheetData[ArrayRow][Email_toCol];
      }
      if (!isValidEmail(Email_to)){
        Step = 6315; // No valid email_to address found
        bSendEmail = false;
        Logger.log(func + Step + ' k= ' + k + ' Sheet Row: ' + (ArrayRow+1) + ' (' + TagValue 
          + ') Valid Email_to Address Not found');
      }
     
      Step = 6320; // Get and test the Email_cc address, if not already specified as a static address
      if(Email_ccCol){
        Email_cc = SheetData[ArrayRow][Email_ccCol];
      }
      if (!isValidEmail(Email_cc)){
        Step = 6325; // No valid email_cc address found
        Logger.log(func + Step + ' k= ' + k + ' Sheet Row: ' + (ArrayRow+1) + ' (' + TagValue 
          + ') Valid Email_cc Address Not found');
      }
     
      Step = 6330; // Get and test the Email_bcc address, if not already specified as a static address
      if(Email_bccCol){
        Email_bcc = SheetData[ArrayRow][Email_bccCol];
      }
      if (!isValidEmail(Email_bcc)){
        Step = 6335; // No valid Email_bcc address found
        Logger.log(func + Step + ' k= ' + k + ' Sheet Row: ' + (ArrayRow+1) + ' (' + TagValue 
          + ') Valid Email_bcc Address Not found');
      }
    } 
      
    Logger.log(func + Step + ' k= ' + k + ' bSendEmail: ' + bSendEmail + ' TO/CC/BCC: ' 
      + Email_to + '/'+ Email_cc + '/'+ Email_bcc);
    
    /**************************************************************************/
    Step = 6400; // Replace variable search terms
    /**************************************************************************/
    if (bSendEmail){

      try {
        Step = 6410; // Prepare the Text templates 
        Logger.log(func + Step + ' k= ' + k + ' EmailTemplateID: ' + EmailTemplateID);  
        var doc = DocumentApp.openById(EmailTemplateID);
        Step = 6412; // Prepare the Text template 
        var body = doc.getActiveSection();
        var email = body.getText();
        //var email = String(body.editAsText);
        //var email = String(body.getBlob().getDataAsString());
        Step = 6414; // Prepare the Subject template 
        Email_subject = Subject_Template;

      }         
      catch(err) {
        var return_message = return_message + '\n Unable to build Email "' + EmailTemplateName 
          + '" using Google Doc (' + EmailTemplateUrl + ') Error Message: ' + err.message;
        var EventMsg = func + Step + ' ' + return_message;
        //LogEvent(EventMsg, oCommon); 
        Logger.log(EventMsg); 
        var display_message = return_message;
        oCommon.ReturnMessage = return_message;
        oCommon.DisplayMessage = display_message;
        return false;
      }
      
      Step = 6420; // Create the SearchTerm / Value relationships using the data in the SheetData array
      for (var key in oCommon.TermsXrefAry) {
        if (key.indexOf(oCommon.Replace_Term_Prefix) > -1) {
          var value = SheetData[ArrayRow][Number(oCommon.TermsXrefAry[key])]; 
          if (key == '<<Address_2>>' & value != '') {
            /****************************************************************************/
            Step = 6430; // Add a comma and space to a 2nd Address Line value, if it is present
            value = value + ', ';
            //Logger.log(func + Step + ' key: ' + key + ' Value: ' + value);
          } else if (isValidDate(value)){
            /****************************************************************************/
            Step = 6440; //  Reformat time values to hh:mm ampm, if necessary isValidDate(d)
            value = formatDate(value); 
            //Logger.log(func + Step + ' key: ' + key + ' Value: ' + value);
          } else {
          
            /****************************************************************************/
            Step = 6450; // Default action for Keys not handled above
            var value = Trim(SheetData[ArrayRow][Number(oCommon.TermsXrefAry[key])]);
            //Logger.log(func + Step + ' key: ' + key + ' Value: ' + value);
          } 

          email = email.replace(key, value || '');  // this is the code line that does all of the work!
          Email_subject = Email_subject.replace(key, value || '');
        }
      }
      
      Step = 6460; // Add the Attached Message(s)
      var scanTime = SheetData[ArrayRow][timestampCol];          

      email = email.replace('<<Attached_Messages>>', attached_message || '');
      email = email.replace('<<Scan_Time>>', scanTime || '');
      
      Step = 6470; /// Replace email address if mode = test and a TestEmailAddress is provided 
      if (test_mode){
        if (ParamCheck(TestEmailAddress)){
          Email_to = TestEmailAddress;
          Email_cc = '';
          Email_bcc = '';
          Logger.log(func + Step + ' k= ' + k + ' Array Row: ' + ArrayRow + ' Sheet Row: ' + (ArrayRow+1) 
              + ' TestEmailAddress Used: ' + Email_to);
        }  
      }
      
      Step = 6480; // Add the personal_message Message
      var personal_message = '';
      if (oCommon.PersonalMessage){
        if (oCommon.PersonalMessage.toUpperCase() == "NO") {
          personal_message = '';
        } else {
          personal_message = '\n' + Trim(oCommon.PersonalMessage) + '\n';
        }
      }
      email = email.replace('<<Personal Message>>', personal_message);

    } // End of Replace variable search terms
    
    Logger.log(func + Step + ' Email_to: ' + Email_to
                   + ' ReplyTo:  ' + ReplyTo
                   + ' Email_cc: ' + Email_cc
                   + ' Email_bcc: ' + Email_bcc
                   + ' subject: ' + Email_subject
                   //+ ' email ' + email
                   + ' error_message: ' + error_message);
    
    /**************************************************************************/
    Step = 6500; // Send Email
    /**************************************************************************/
    if (bSendEmail){
      sentCount++;
      Step = 6510;
      var error_message = '';
      
      Logger.log(func + Step + ' Email_to: ' + Email_to
                   + ' ReplyTo:  ' + ReplyTo
                   + ' Email_cc: ' + Email_cc
                   + ' Email_bcc: ' + Email_bcc
                   + ' subject: ' + Email_subject
                   + ' error_message: ' + error_message);
      
      if (oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') <= -1){

        if (sendTheEmail(oCommon, Email_to, Email_cc, Email_bcc, Email_subject, email, error_message, ReplyTo)) {
          Step = 6520;
          NumEmailsCompleted++;
          NumEmailsSent++;
          Logger.log(func + Step + ' Email Sent (' + Email_to + ') Count: ' + NumEmailsSent);
          
          Step = 6530; // Record the successful event, if DateMailSentCol is specified
          if (!isNaN(DateMailSentCol)){
            Logger.log(func + Step + ' DateMailSentCol: ' + DateMailSentCol + ', DateMailSentColLtr: ' + DateMailSentColLtr);
            var CellAddress = DateMailSentColLtr + String(ArrayRow + 1).split(".")[0];
            sheet.getRange(CellAddress).setValue(todaysDate);    
            sheet.getRange(CellAddress).setNumberFormat('M/d/yyyy H:mm:ss');
            // Make sure the cell is updated right away in case the script is interrupted
            SpreadsheetApp.flush();
            Logger.log(func + Step + ' maxFrequency: ' + maxFrequency + ', CellAddress: ' 
                + CellAddress + ', Posted Date: ' + todaysDate);
          }
        } else {
          Step = 6540;
          return_message = return_message + "\n Email Sending Error (" + emailAddress + ") Error Message: " + error_message;
          Logger.log(func + Step + ' ' + return_message);
        }
      }
    }  // End of bSendEmail bool test 

    /**************************************************************************/
    Step = 6600; // Update the Sheet row values, if send is successful
    /**************************************************************************/
    if (bSendEmail){
      // Update the oCommon.SelectedSheetRows[k] array
      oCommon.SelectedSheetRows[k][oCommon.SelectedSheetRows[k].length] = Email_to;
    } else {
      // Remove the row that is not being sent
      oCommon.SelectedSheetRows.splice(k);
    }
    
    Logger.log(func + Step + ' Email_to: ' + Email_to
                   + ' ReplyTo:  ' + ReplyTo
                   + ' Email_cc: ' + Email_cc
                   + ' Email_bcc: ' + Email_bcc
                   + ' subject: ' + Email_subject
                   + ' error_message: ' + error_message);
    
  } // iterate k / End of: Process each row passsed in oCommon.SelectedSheetRows array
  
  /**************************************************************************/
  Step = 7000; // Report metrics and go home...
  /**************************************************************************/
  // Update Library counters
  var scriptProperties = PropertiesService.getScriptProperties();
  var emails_count = Number(scriptProperties.getProperty('B_Emails_count'));
  scriptProperties.setProperty('B_Emails_count', emails_count + NumEmailsCompleted);
  oCommon.EmailsGenerated = NumEmailsCompleted;
  oCommon.ItemsAttempted = total_attempted_messages;
  oCommon.ItemsCompleted = NumEmailsCompleted;
  var message = ' SendMail completed ' + NumEmailsCompleted + ' of ' 
    + total_attempted_messages + ' messages attempted. ';
  // Adjust the message for the number bypassed_for_freq_violation
  if (bypassed_for_freq_violation > 0){
    if (bypassed_for_freq_violation > 1){
      message = message + bypassed_for_freq_violation
      + ' were not sent due to Frequency Violations.';
    } else {
      message = message + bypassed_for_freq_violation
      + ' was not sent due to a Frequency Violation.';
    }
  }
  oCommon.DisplayMessage = oCommon.DisplayMessage + message + '\\n';
  
  oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
  Logger.log(func + oCommon.ReturnMessage);

  Logger.log(func + ' END');
  
  return true;

}


function AlertsScan(oCommon, oMenuParams) {
  /* ****************************************************************************
   DESCRIPTION:
     This function is looks for alert conditions when executed by the user

   USEAGE:
     ReturnBool = AlertsScan(oCommon); 
     
   REVISION DATE:
     07-22-2017 - First implementation
     12-25-2017 - Utilize the procedures in Script Library 1.1
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback  
     02-13-2019 - Modified to work with oMenuParams[] passed from te Top Level Calling function
                - Using new GetMenuParams() function in the Top Level calling function to pass the
                   email template names(s) in the aryParams array
                - Eliminating default use of Secondary_emailCol = oCommon.SecondaryEmailCol,
                - The address will now be taken from the entries(s) in the "To" col of the Template Parameters
     11-08-2019 - Modified Step 5400 so that AdminEmails will not be sent if the SendEmails() 
                    function returns False
                - Added oCommon.AlertsScanSource scalar (Step 1040) to pass source sheet name to the 
                    SendElail() function so that Error 312 will not be tripped during weekly scan process

     Input Parameter Values:
     
       Array: oCommon.SelectedSheetRows[n]
         oCommon.SelectedSheetRows[n][0] = SheetData Target row for each recipient
         oCommon.SelectedSheetRows[n][1] = recipient last name
         
       Array: oMenuParams[key] = value
         oMenuParams[Email] = email #1 Template Title
         oMenuParams[Group] = Group #
   
     Output Parameter Values:
     
       Array: oCommon.ChangedSheetRows[n]
         oCommon.ChangedSheetRows[n][0] = SheetData Target row for each recipient
         oCommon.ChangedSheetRows[n][1] = recipient last name
         oCommon.ChangedSheetRows[n][2] = Alert message / email body
                
   NOTES:
  ******************************************************************************/
  var func = "***AlertsScan " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN...');
  
  /******************************************************************************/
  Step = 1000; // Examine and vaildate input parameters
  /******************************************************************************/
  Step = 1020; // Validate EmailTemplateName(s)
  var EmailTemplateName = oMenuParams['Email'];
  var Group = oMenuParams['Group'];
  var AdminTemplateName = oMenuParams['Admin'];
  
  if (oCommon.CallingMenuItem = null){
    var title = oCommon.CallingMenuItem;
  } else {
    var title ='Scan for Alerts';
  }
  
  var return_message = '';
  var display_message = '';
  var prog_message = '';
  var mode = 1;
  
  if (!ParamCheck(EmailTemplateName)){
    oCommon.DisplayMessage = ' Error 240 - EmailTemplateName(s) missing or not Valid';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    return false;
  }
  
  if (!ParamCheck(Group)){
    Group = 0;
  }
  
  Step = 1030; // Validate oCommon.SelectedSheetRows
  if(oCommon.SelectedSheetRows.length <= 0){
    var bUserSelectsRows = true;
  } else {
    var bUserSelectsRows = false;
  }
  
  Step = 1040; // Correction made for triggered executions (i.e. DoWeeklyScan)
  var bSilentMode = oCommon.bSilentMode;
  if (oCommon.bSilentMode){
    bUserSelectsRows = false;
    mode = 1; // Select all rows when executed in silent mode
    Logger.log(func + Step + ' Mode change - bSilentMode: ' + oCommon.bSilentMode + ', mode: ' + mode );
    if(!SelectSheetRows(oCommon, mode, oCommon.FormResponse_SheetName, oCommon.FormResponseStartRow)){
      Step = 1045; // Error 242 - Failed to select any rows in Silent mode
      oCommon.DisplayMessage = ' Error 242 - Failed to select any rows in Silent mode.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon); 
      return false;
    } else {
      oCommon.AlertsScanSource = oCommon.FormResponse_SheetName;
      Logger.log(func + Step + ' NUmber of Selected Rows: ' + SelectSheetRows.length );
    }
  }
  
  /**************************************************************************/
  Step = 2000; // Select records, if necessary (Ref Step 1020, above)
  /**************************************************************************/
  if (bUserSelectsRows){
    Step = 2100; // User needs to select rows
    var MsgBoxMessage = 'Do you want to select individual rows?' ;
    mode = 1;
    if (Browser.msgBox('Select Rows?',MsgBoxMessage, Browser.Buttons.YES_NO) == 'yes'){
      mode = 2;
    } 
    if(!SelectSheetRows(oCommon, mode, oCommon.FormResponse_SheetName, oCommon.FormResponseStartRow)){
      Step = 2110; // ERROR - No Selected Rows found
      return_message = func + Step + ' ********** ERROR - No Rows Selected';
      display_message = 'No Rows selected, procedure will be terminated';
      //LogEvent(return_message, oCommon);
      oCommon.ReturnMessage = return_message;
      oCommon.DisplayMessage = display_message;
      // Communicate with user
      prog_message = display_message;
      progressMsg(prog_message,title,3);
      return false;
    }
    
    Step = 2200; // Take time to confirm that the user wants to proceed
    var MsgBoxMessage = ('The following rows were selected: ' + FormatForDisplay(oCommon.SelectedSheetRows)
       + '\\n\\n' + ' Do you want to continue?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      display_message = "Your response is NO - procedure will be terminated";
      return_message = func + Step + ' ********** WARNING - User terminated after rows were selected';
      oCommon.ReturnMessage = return_message;
      oCommon.DisplayMessage = display_message;
      // Communicate with user
      prog_message = 'User chooses not to continue.';
      progressMsg(prog_message,title,3);
      return false;
    }
     
    //Step = 2300; // Take time to get the personal message from the user
    //oCommon.PersonalMessage = Browser.inputBox('ADD PERSONAL MESSAGE?', 'Enter your message', Browser.Buttons.YES_NO);
    
  }   
   
  /********************************************************************************/
  Step = 3000; // Scan for Expiration Date Alerts
  /********************************************************************************/
  var rows_scanned = oCommon.SelectedSheetRows.length;
  var ValidationSection = oCommon.ValidationSection_Heading;
  //var ChangedSheetRows = [];
  if (!ValidationScan(oCommon, Group)){
    return false;
  }
  
  //ChangedSheetRows[i][0] = SheetDataRow
  //ChangedSheetRows[i][1] = TagValue
  //ChangedSheetRows[i][2] = AlertMessage

  var NumberOfAlertsFound = oCommon.ChangedSheetRows.length;
  Logger.log(func + Step + ' NumberOfAlertsFound ' + NumberOfAlertsFound + ' Values: ' + oCommon.ChangedSheetRows);

  oCommon.SelectedSheetRows = oCommon.ChangedSheetRows;
  
  /********************************************************************************/
  Step = 4000; // View Alerts and determine if UserAlert and Admin Emails are to be sent
  /********************************************************************************/
  
  if(NumberOfAlertsFound <= 0){
    Step = 4100; // Leave now since no further action is required
    oCommon.ReturnMessage = func + Step + ' No Alerts found.';
    oCommon.DisplayMessage = ' No Alerts found among the rows selected.';
    return true;
  }
    
  /********************************************************************************/
  Step = 5000; // Send Alert Emails?
  /********************************************************************************/
  var bSendAlertEmails = true;
  if(!bSilentMode){
    Step = 5100; // Construct the message body
    var MessageBody = '';
    for (m = 0; m < oCommon.ChangedSheetRows.length; m++){
      MessageBody = MessageBody + '\\n' + '  Row: ' + oCommon.ChangedSheetRows[m][0]
      + ' (' + oCommon.ChangedSheetRows[m][1] + ') Alerts: ' + oCommon.ChangedSheetRows[m][2];
    }  
    Step = 5200; // Get User input
    var MsgBoxMessage = ('The following User Alert Results were found: ' 
       + '\\n' + MessageBody + '\\n' 
       + '\\n' + '  Do you want to send individual User Alert email(s)?');
    if (Browser.msgBox('Send Alert Emails?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      return_message = func + Step + ' User elects not to send Alert Emails.';
      display_message = ' You chose NOT to send individual Alert Emails.';
      Browser.msgBox('Send Alert Emails?', display_message, Browser.Buttons.OK);
      bSendAlertEmails = false;
    } else {
      bSendAlertEmails = true;   
      Step = 5250; // Take time to get the personal message from the user for the Alert emails
      oCommon.PersonalMessage = ''; // clear, just in case
      oCommon.PersonalMessage = Browser.inputBox('ADD PERSONAL MESSAGE?', 'Enter your message', Browser.Buttons.YES_NO);
    }
  }
  
  Step = 5300; // Check the email "TestMode" setting
  if (bSendAlertEmails && oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') > -1){
    Step = 5310; // No Alert Emails sent due to TestMode setting
    return_message = func + Step + ' No Alert Emails sent due to TestMode setting.';
    display_message = ' No Alert Emails will be sent due to Global Email TestMode setting.';
    LogEvent(return_message, oCommon);
    if(!bSilentMode){Browser.msgBox('Send Alert Emails?',display_message, Browser.Buttons.OK);}
    bSendAlertEmails = false;
  }
  
  Step = 5400; // Send Alert Email, if above checks are passed
  var bSendAdminEmails = true;
  if(bSendAlertEmails){
    Step = 5410; // Send Alert Emails
    if(!SendEmail(oCommon, EmailTemplateName)){
      // Problems encountered;
      bSendAdminEmails = false;
      oCommon.DisplayMessage = ' Send Alert Email Problems encountered: ' + oCommon.DisplayMessage;
    } else {
      oCommon.DisplayMessage = ' Send Alert Email Results: ' + oCommon.DisplayMessage;
    }
    if(!bSilentMode){
      Browser.msgBox(oCommon.DisplayMessage);
    } else {
      LogEvent(oCommon.DisplayMessage, oCommon);
    }
    Logger.log(func + Step + ' DisplayMsg: ' + oCommon.DisplayMessage);
    Logger.log(func + Step + ' DisplayMsg: ' + oCommon.ReturnMessage);
  }
    
  /********************************************************************************/
  Step = 6000; // Send Admin Summary Emails?
  /********************************************************************************/
  if (oCommon.ItemsCompleted <= 0){ 
    Step = 6100; // No Admin Emails need to be sent
    return_message = func + Step + ' No Alert emails were sent, so no Admin Emails will be sent.'; 
    display_message = ' No Alert emails were sent, so no Admin Emails will be sent.';
    if(!bSilentMode){ Browser.msgBox('Send Admin Emails?',display_message, Browser.Buttons.OK); }
    bSendAdminEmails = false; 
  } else if(!bSilentMode){
    Step = 6200; // Get User input
    var MsgBoxMessage = ('  Do you want to send Admin Summary email(s)?');
    if (Browser.msgBox('Send Admin Emails?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      return_message = func + Step + ' User elects not to send Admin Emails.';
      display_message = ' OK. You chose NOT to send individual Admin Summary Emails. Process will now terminate';
      Browser.msgBox('Send Admin Emails?', display_message, Browser.Buttons.OK);
      bSendAdminEmails = false;        
    }
  }
  //LogEvent(return_message, oCommon);
  Logger.log(return_message);
  
  Step = 6200; // Check the email "TestMode" setting
  if (bSendAdminEmails && oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') > -1){
    Step = 6210; // No Admin Emails sent due to TestMode setting
    return_message = func + Step + ' No Admin Emails sent due to TestMode setting.';
    display_message = ' No Admin Emails will be sent due to Global Email TestMode setting.';
    LogEvent(return_message, oCommon);
    if(!bSilentMode){Browser.msgBox('Send Admin Emails?',display_message, Browser.Buttons.OK);}
    bSendAdminEmails = false;
  }
  
  Step = 6400; // Send Admin Email if above checks are passed
  if(bSendAdminEmails && AdminTemplateName){
    Step = 6410; // Send Admin Emails
    if(!SendAdminEmails(oCommon, AdminTemplateName)){
      // Problems encountered;
      oCommon.DisplayMessage = ' Send Admin Email Problems encountered: ' + oCommon.DisplayMessage;
    } else {
      oCommon.DisplayMessage = ' Send Admin Email Results: ' + oCommon.DisplayMessage;
    }
    if(!bSilentMode){
      Browser.msgBox(oCommon.DisplayMessage);
    } else {
      LogEvent(oCommon.DisplayMessage, oCommon);
    }
    Logger.log(func + Step + ' DisplayMsg: ' + oCommon.DisplayMessage);
    Logger.log(func + Step + ' DisplayMsg: ' + oCommon.ReturnMessage);
    oCommon.DisplayMessage = '';
    oCommon.ReturnMessage = '';
  }
  
  /********************************************************************************/
  Step = 7000; // Wrap up and go home...
  /********************************************************************************/
  oCommon.DisplayMessage = ' Alerts Scan completed: ' + rows_scanned
    + ' rows scanned and ' + NumberOfAlertsFound + ' alerts found.';
  oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
  LogEvent(oCommon.ReturnMessage, oCommon);
  
  return true;
  
}


function assignEditUrls(oCommon, bSetDuplicateTrap) {
  /* ****************************************************************************

   DESCRIPTION:
     This function is invoked whenever a "new" form submission is made:
         1 - Add / Update the Edit URL
         2 - Write the Employment Status text: "Active" if needed, and
         3 - Write the EDIT url and VIEW button text values used by "Awesome Table" used for data display

   REVISION DATE:
    07-06-2017 - First Instance
    12-27-2017 - Change all row/col references to Base 0
    01-27-2018 - Modified code to bypass non-active records
    03-28-2018 - Added Step = 1310 to temporarily disable the form from accepting responses
               - Added Step = 3000 to re-nable the form to accept responses and return
    11-03-2018 - Modified Step 2400 to format the sheetTimestamp value as mm/dd/yyy hh:mm:ss
    01-20-2019 - Added data conversion case "Split" (ref Step 2650)
     
   NOTES:
     1. The array oCommon.SelectedSheetRows[SheetRow][TagValue] is passed identifying the selected rows
     2. oCommon.SelectedSheetRows.length == 0 if an invalid range has been selected by the user     
     
    click on the menu Edit -> Current project's triggers and set up an automated trigger onFormSubmission

  ******************************************************************************/
  var func = "***assignEditUrls " + Version + " - ";
  var Step = 1000;
  Logger.log(func + Step + ' BEGIN');
  
  /****************************************************************************/
  Step = 1100; // Assign the oCommon parameters
  /****************************************************************************/
  var oSourceSheets = oCommon.Sheets, 
      formID = oCommon.GoogleForm_Url,
      dataTab = oCommon.FormResponse_SheetName,
      firstDataRow = oCommon.FormResponseStartRow,
      AwesomeTableRow = oCommon.AwesomeTableRow, // Base 0; ref https://support.awesome-table.com/hc/en-us/articles/115001068129-Data-configuration
      timestampCol = oCommon.TimestampCol,
      StatusCol = oCommon.Record_statusCol,
      ActiveStatusValue = oCommon.Active_Status_value,
      error_email = oCommon.Send_error_report_to,
      FormTimestamp = oCommon.FormTimestamp
    ;
  
   var onSubmit_sensitivity = Math.abs(oCommon.onSubmit_Sensitivity/1000); 
  
   var NoErrors = true;
  
  
   Step = 1150; // Verify that the FormTimestamp is acceptable
   if (!ParamCheck(FormTimestamp)){
     oCommon.DisplayMessage = ' ERROR 220 - Expected FormTimestamp is not found.';
     oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
     LogEvent(oCommon.ReturnMessage, oCommon);
     Logger.log(func + Step + oCommon.ReturnMessage); 
     return false;
   }
  
  Step = 1160; // Set the "Duplicate Trap"
  if (!ParamCheck(bSetDuplicateTrap)){
    // Parameter is missing or undefined, assume thar trap is to be set
    bSetDuplicateTrap = true;
  }
  // TestMode trumps parameter value
  if (oCommon.TestMode.toUpperCase().indexOf("NO TRAPS") > -1){
    bSetDuplicateTrap = false;
  } 
  
  /******************************************************************************/
  Step = 1300; // Retreive the GoogleForm response data stored separately from the Google Sheet data
  /******************************************************************************/
  try {
    //Logger.log(func + Step + ' Got here...')
    //var formUrl = sheet.getFormUrl();             // Use form attached to sheet
    //var form = FormApp.openByUrl(formUrl);
    formID = getIdFrom(formID);
    var form = FormApp.openById(formID); // formID assigned from oCommon
    
    Step = 1310; // Temporarily disable the form from accepting responses
    // Source: https://stackoverflow.com/questions/43421593/reject-google-forms-submission-or-remove-response
    var defaultClosedFor = form.getCustomClosedFormMessage();
    form.setCustomClosedFormMessage("The system is busy processing a submission. Please try again in a few minutes.");
    form.setAcceptingResponses(false);
    
    Step = 1320; // Retrieve the form responses and item objects
    var responses = form.getResponses();  
    var items = form.getItems();
    
    /******************************************************************************/
    Step = 1400; // Capture the Form Timestamps and Edit url's from the Form object
    /******************************************************************************/
    //Logger.log(func + Step + ' Got here...')
    var timestamps =[];
    var urls =[];
    var ResponderEmails = [];
    
    for (var i = 0; i < responses.length; i++) {
      timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
      urls.push(responses[i].getEditResponseUrl());
      ResponderEmails.push(responses[i].getRespondentEmail());
      //Logger.log(func + Step + ' i: ' + i + ', Timestamp: ' + timestamps[i]);
    }

  } catch (err) {
    // Error condition encountered retreiving the GoogleForm response data
    form.setCustomClosedFormMessage(defaultClosedFor);
    form.setAcceptingResponses(true);
    //lock.releaseLock();
    oCommon.DisplayMessage = ' Error 222 - Error condition encountered retreiving '
      + 'the GoogleForm response data: ' + err.message;
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    Logger.log(func + Step + oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }

  /******************************************************************************/
  Step = 1500; // Capture the Data Conversion parameters, if they exist
  /******************************************************************************/
  //Logger.log(func + Step + ' Got here...')
  var bDoConversions = false;
  var ParameterSet = oCommon.DataConversions_Heading;
  if (ParamCheck(ParameterSet)){
    Step = 1510; // Get the parameters from the Setup sheet
    var oConversionParams = {};
    var ParameterSet = oCommon.DataConversions_Heading;  
    oConversionParams = ParamSetBuilder(oCommon, ParameterSet);
    
    if (ParamCheck(oConversionParams)){
      Step = 1520; // Conversion Parameters are active
      Logger.log(func + Step + ' ParameterSet: ' + ParameterSet 
         + ', Parameters: ' + oConversionParams.Parameters);
      bDoConversions = true;
    } else {      
      Step = 1530; // Expected Data Conversion Parameters, not found
      form.setCustomClosedFormMessage(defaultClosedFor);
      form.setAcceptingResponses(true);
      oCommon.DisplayMessage = ' Warning 224 - Expected Data Conversion Parameters, not found.';
      oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon);
      Logger.log(func + Step + oCommon.ReturnMessage);
      bDoConversions = false;

    }
  }
  
  try {
  
    /******************************************************************************/
    Step = 1600; // Set a public lock that locks for all invocations
    /******************************************************************************/
    // http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
    //var lock = LockService.getPublicLock();
    //lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
    // got the lock, we may now proceed
    
    /**********************************************************************************/
    Step = 1700; //Loop through the Form Response sheet to find the row with a matching FormTimestamp value
    /**********************************************************************************/
    // Prepare objects
    //Logger.log(func + Step + ' dataTab: ' + dataTab);
    var sheet = oSourceSheets.getSheetByName(dataTab);
    var SheetData = sheet.getDataRange().getValues();
    var MatchingRow = null;
    
    //Loop through the Form Response sheet
    for (var ArrayRow = oCommon.FormResponseStartRow; ArrayRow < SheetData.length; ArrayRow++) {
      Step = 1720; // 
      var sheetTimestamp = SheetData[ArrayRow][timestampCol];
      Step = 1725; // Get and assign the record's TagValue // && TargetSheetName == oCommon.FormResponse_SheetName
      if (ParamCheck(oCommon.RecordTagValueCol)){
        var TagValue = FormatTagValue(oCommon, SheetData[ArrayRow][oCommon.RecordTagValueCol]);
      } else {
        var TagValue = '';
      }

      //Logger.log(func + Step + ' ******* Examining (' + TagValue + ') SheetRow: ' + (ArrayRow+1) 
      //  + ' SheetTimestamp: ' + sheetTimestamp);
      if (ParamCheck(sheetTimestamp)){
        Step = 1730; // Timestamp exists, compute date/time values for comparison
                     // Reference: https://developers.google.com/google-ads/scripts/docs/features/dates
        var dF = new Date(FormTimestamp).getTime();
        var dS = new Date(sheetTimestamp).getTime();
        var mDiff = Math.abs((dF - dS)/1000);  // time difference in seconds
        
        //Logger.log(func + Step + ' dF: ' + dF + ', dS: ' + dS + ', mDiff: ' + mDiff + ', Sensitivity: ' + onSubmit_sensitivity);

        Step = 1740; // Test to determine if difference is within the "sensitivity" range
        if(mDiff <= onSubmit_sensitivity){         
          Step = 1745; // Form Entry found
          MatchingRow = ArrayRow; 
          Logger.log(func + Step + ' >>>>>>> SUBMIT FOUND - (' + TagValue + ') SheetRow: ' + (MatchingRow+1));
          break;
        }
      }
    }
    
    /******************************************************************************/
    Step = 1800; // Check for fatal error
    /******************************************************************************/
    if (MatchingRow == null){
      Step = 1850; // Form submit entry not found, exit now
      oCommon.DisplayMessage = ' Error 226 - No Form Submission entries found after onSubmit(' 
        + oCommon.FormTriggerUid + ') event.';
      oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon);
      Logger.log(func + Step + oCommon.ReturnMessage);
      return false;
    }
    
    /******************************************************************************/
    Step = 1900; // To get to here, the row updated by a Form Submit action was found
    /******************************************************************************/
    // assign the ArrayRow value
    ArrayRow = MatchingRow;
    
    /******************************************************************************/
    Step = 2000; // Get a Lock and Check for a duplicate entry error
    /******************************************************************************/ 
    var scriptProperties = PropertiesService.getScriptProperties();
    if (bSetDuplicateTrap){
      var lock = LockService.getScriptLock();
      var bSetTraps = true;
      var count_trap = Number(scriptProperties.getProperty('B_Trap_count')); 
      var count_sub = Number(scriptProperties.getProperty('B_Submission_count')); 
    
      try {
        lock.waitLock(oCommon.FormSubmit_Delay); // wait n seconds for others' use of the code section and lock to stop and then proceed
      } catch (err) {
        count_trap++;
        scriptProperties.setProperty('B_Trap_count', count_trap); 
        oCommon.ReturnMessage = func + Step + ' - Could not obtain lock after ' + (oCommon.FormSubmit_Delay/1000) + ' seconds: '
        + 'SheetRow: ' + (ArrayRow+1) + ' (' + TagValue + ') SheetTimestamp: ' + sheetTimestamp;
        //LogEvent(oCommon.ReturnMessage, oCommon);
        Logger.log(func + Step + oCommon.ReturnMessage);
        
        return false; // return immediately as this is probably a duplicate submission
      }
    
      Step = 2100; // Lock obtained, check for a duplicate entry error
      Logger.log(func + Step + ' - Lock obtained: '
                 + 'SheetRow: ' + (ArrayRow+1) + ' (' + TagValue + ') SheetTimestamp: ' + sheetTimestamp);
      if(!isNaN(oCommon.AsOfDateCol)){
        var LastRptDate = SheetData[ArrayRow][oCommon.AsOfDateCol];
        var dL = new Date(LastRptDate).getTime();
        var dC = new Date(sheetTimestamp).getTime();
        var mDiff = Math.abs((dC - dL)/1000);  // time difference in seconds
        //Logger.log(func + Step + ' dF: ' + dF + ', dS: ' + dS + ', mDiff: ' + mDiff + ', Sensitivity: ' + onSubmit_sensitivity);
        Step = 2110; // Test to determine if difference is within the "sensitivity" range
        if(mDiff <= onSubmit_sensitivity){         
          Step = 2120; // stop immediately
          count_trap++;
          scriptProperties.setProperty('B_Trap_count', count_trap); 
          oCommon.ReturnMessage = func + Step + ' Error 227 - Duplicate Action: mDiff: ' + mDiff + ', LastRptDate: ' + LastRptDate
          + ', CurrentTimestamp: ' + sheetTimestamp;
          //LogEvent(oCommon.ReturnMessage, oCommon);
          Logger.log(func + Step + oCommon.ReturnMessage);
          lock.releaseLock();
          return false;
        }
      }
      
      Step = 2150; // Release lock since it is no longer needed for the Trap
      lock.releaseLock();

    }
      
    Step = 2200; // Format and write the rptDate value
    var rptDateColAddr = numToA(oCommon.AsOfDateCol+1) + String(ArrayRow+1).split(".")[0];
    Logger.log(func + Step + ' URL As of Date - SheetRow: ' + (ArrayRow+1) +  ' CellAddress: ' + rptDateColAddr + ' Value: ' + sheetTimestamp);
    //sheet.getRange(rptDateColAddr).setValue(sheetTimestamp);  
    // set format: m/d/yyy h:m:s
    var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(sheetTimestamp, timeZone, "M/d/yyy HH:mm:ss");
    sheet.getRange(rptDateColAddr).setValue(formattedDate); 
    sheet.getRange(rptDateColAddr).setNumberFormat('M/d/yyyy H:mm:ss');
    SpreadsheetApp.flush();
    Logger.log(func + Step + ' Report Date Written: ' + formattedDate + ' to cell: ' + rptDateColAddr);
        
    Step = 2250; // Record the Submission Event for posterity
    count_sub++; 
    scriptProperties.setProperty('B_Submission_count', count_sub);
    Logger.log(func + Step + ' Submission Recorded: ' + count_sub);
     
    /******************************************************************************/
    Step = 2300; // Find and Format the "EditURLS" value and write to the sheet
    /******************************************************************************/
    //Logger.log(func + Step + ' Edit_URLCol: ' + oCommon.Edit_URLCol);
    if(!isNaN(oCommon.Edit_URLCol)){
      var editURL = [sheetTimestamp?urls[timestamps.indexOf(sheetTimestamp.setMilliseconds(0))]:''];
      var CellAddress = numToA(oCommon.Edit_URLCol+1) + String(ArrayRow+1).split(".")[0]; 
      if (editURL != ''){
        sheet.getRange(CellAddress).setValue(editURL); 
      }
    }
    Logger.log(func + Step + ' Edit_URL: ' + editURL);
    
    /******************************************************************************/
    Step = 2400; // Write the FormOwnerEmail if missing (first submittal?) 
    /******************************************************************************/
    if(!isNaN(oCommon.Form_Owner_Col)){
      Step = 2410; // Get the previously stored FormOwnerEmail value
      var FormOwnerEmail = SheetData[ArrayRow][oCommon.Form_Owner_Col];
      if (!ParamCheck(FormOwnerEmail)){
        Step = 2420; // Construct the FormOwnerEmail cell address and write the value retrieved from the response
        var ResponderEmail = [sheetTimestamp?ResponderEmails[timestamps.indexOf(sheetTimestamp.setMilliseconds(0))]:''];
        var FormOwnerEmailAddr = numToA(oCommon.Form_Owner_Col+1) + String(ArrayRow+1).split(".")[0];
        if (ResponderEmail != ''){
          sheet.getRange(FormOwnerEmailAddr).setValue(ResponderEmail); 
          FormOwnerEmail = ResponderEmail; 
        }
        //Logger.log(func + Step + ' >>>>>>> FormOwnerEmail - SheetRow: ' + (ArrayRow+1) +  ' CellAddress: ' 
        //  + FormOwnerEmailAddr + ' Value: ' + FormOwnerEmail + ' bActiveUserEmpty' + oCommon.bActiveUserEmpty);
      }
    }
            
    /******************************************************************************/
    Step = 2500; // Format the "editButtonCol" value and write to the sheet
    /******************************************************************************/ 
    if(!isNaN(oCommon.Edit_ButtonCol)){
      var editButtonAddr = numToA(oCommon.Edit_ButtonCol+1) + String(ArrayRow+1).split(".")[0];
      //Logger.log(func + Step + ' Writing EDIT Text - SheetRow: ' + (ArrayRow+1) +  ' CellAddress: ' + editButtonAddr + ' Value: ' + editButtonText);
      sheet.getRange(editButtonAddr).setValue(oCommon.Edit_Button_Text);  
    }
        
    /******************************************************************************/
    Step = 2600; // Perform data conversions, if defined
    /******************************************************************************/ 
    if (bDoConversions && oConversionParams.Parameters.length > 1){
      for( var prow = 1; prow < oConversionParams.Parameters.length; prow++){
        Step = 2605; // Conduct a conversion for each row in the oConversionParams.Parameters array
        var ConvType = oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Type of Conversion']];
        var ResultCol = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Result Column Heading']]);
        var ResultAddr = numToA(ResultCol+1) + String(ArrayRow+1).split(".")[0];
        
        var Param1Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 1']]);
        if (!isNaN(Param1Col)) {
          // Get the Form Data value for the current row associated with Param1
          var Param1 = Trim(SheetData[ArrayRow][Param1Col]);
        } else {
          // Get the value of the constant
          var Param1 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 1']]);
        }
        
        var Param2Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']]);
        if (!isNaN(Param2Col)) {
          // Get the Form Data value for the current row associated with Param2
          var Param2 = Trim(SheetData[ArrayRow][Param2Col]);
        } else {
          // Get the value of the constant
          var Param2 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']]);
        }
        
        var Param3Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']]);
        if (!isNaN(Param3Col)) {
          // Get the Form Data value for the current row associated with Param3
          var Param3 = Trim(SheetData[ArrayRow][Param3Col]);
        } else {
          // Get the value of the constant
          var Param3 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']]);
        }
        
        var Param4Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 4']]);
        if (!isNaN(Param4Col)) {
          // Get the Form Data value for the current row associated with Param4
          var Param4 = Trim(SheetData[ArrayRow][Param4Col]);
        } else {
          // Get the value of the constant
          var Param4 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 4']]);
        }
        
        var Param5Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 5']]);
        if (!isNaN(Param5Col)) {
          // Get the Form Data value for the current row associated with Param5
          var Param5 = Trim(SheetData[ArrayRow][Param5Col]);
        } else {
          // Get the value of the constant
          var Param5 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 5']]);
        }
        
        var Param6Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 6']]);
        if (!isNaN(Param6Col)) {
          // Get the Form Data value for the current row associated with Param6
          var Param6 = Trim(SheetData[ArrayRow][Param6Col]);
        } else {
          // Get the value of the constant
          var Param6 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 6']]);
        }
          
        Logger.log(func + Step + ' Values:' 
        + ' ArrayRow: ' + ArrayRow
        + ', prow: ' + prow 
        + ', ConvType: ' + ConvType 
        + ', ResultCol: ' + ResultCol
        + ', ResultAddr: ' + ResultAddr
        + ', Param1: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 1']] + ' Param1Col: ' + Param1Col + ' Value: ' + Param1
        + ', Param2: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']] + ' Param2Col: ' + Param2Col + ' Value: ' + Param2
        + ', Param3: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']] + ' Param3Col: ' + Param3Col + ' Value: ' + Param3
        + ', Param4: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 4']] + ' Param4Col: ' + Param4Col + ' Value: ' + Param4
        + ', Param5: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 5']] + ' Param5Col: ' + Param5Col + ' Value: ' + Param5
        + ', Param6: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 6']] + ' Param6Col: ' + Param6Col + ' Value: ' + Param6);
                       
        switch (ConvType) {
            
          case "Date-To-Day": 
            Step = 3410; // Conduct a "Date-To-Day" conversion
            var dParam1 = new Date(Param1);
            if (isValidDate(dParam1)){
              sheet.getRange(ResultAddr).setValue(GetDay(dParam1));  
            }
            break;
            
          case "DateDiff-Year": 
            Step = 3420; // Conduct a "DateDiff-Year" conversion
            var dParam1 = new Date(Param1);
            var dParam2 = new Date(Param2);
            if (isValidDate(Param1) && isValidDate(Param2)){
              var DateDifference = DateDiff(dParam1,dParam2,'YEARS');
              if (DateDifference){
                if (DateDifference > 0){
                  sheet.getRange(ResultAddr).setValue(DateDiff(dParam1,dParam2,'YEARS'));  
                  sheet.getRange(ResultAddr).setNumberFormat("@");
                }
              }
            }
            break;
            
          case "DateDiff-Age": 
            Step = 3430; // Conduct a "DateDiff-Year" conversion
            var dParam1 = new Date(Param1);
            var dParam2 = new Date(Param2);
            if (isValidDate(dParam1) && isValidDate(dParam2)){
              var DateDifference = DateDiff(dParam1,dParam2,'MONTHS');
              var Age = DateDifference / 12;
              var Age = Math.floor(Age);
              if (Age){
                sheet.getRange(ResultAddr).setValue(Age);
                sheet.getRange(ResultAddr).setNumberFormat("@");
              }
            }
            break;
            
          case "Const": 
            Step = 2640; // Conduct a "Const" conversion
            if (Param1){
              sheet.getRange(ResultAddr).setValue(Param1);  
            }
            break;
            
          case "Split": 
            Step = 2650; // Splits the input string(Param1) into 2 parts using the Param 2 delimiter
            // and places Part 1 into the Result col and Part 2 into the Param 3 column 
            var Part1Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Result Column Heading']]);
            if(isNaN(Part1Col)){
              var EventMsg = func + Step + ' ERROR 225 - Invalid entry in Conversion Parameters row ' + prow + ' for "Result Column Heading"';
              Logger.log(EventMsg);
              LogEvent(EventMsg, oCommon);
              break;
            }
            var Part1Addr = numToA(Part1Col+1) + String(ArrayRow+1).split(".")[0];
            var Part2Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']]);
            if(isNaN(Part2Col)){
              var EventMsg = func + Step + ' ERROR 226 - Invalid entry in Conversion Parameters row ' + prow + ' for Param 3';
              Logger.log(EventMsg);
              LogEvent(EventMsg, oCommon);
              break;
            }
            var Part2Addr = numToA(Part2Col+1) + String(ArrayRow+1).split(".")[0];
            
            // write over the contents of the part2 fields
            sheet.getRange(Part2Addr).setValue('')
            
            var InputString = Param1;
            if(!ParamCheck(InputString)){
              // no string found that needs to be split
              break;
            }
            var delimiter = oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']];
            if(!ParamCheck(delimiter)){
              var EventMsg = func + Step + ' ERROR 227 - Invalid entry in Conversion Parameters row ' + prow + ' for Param 2';
              Logger.log(EventMsg);
              LogEvent(EventMsg, oCommon);
              break;
            }
            Logger.log(func + Step + ' InputString: "' + Param1 + '", Delimiter: "' + delimiter 
                       + '", Part1Addr : ' + Part1Addr  + '", Part2Addr: ' + Part2Addr);
            if(InputString.indexOf(delimiter) > -1){
              Step = 2655; // delimiter is found within the input string
              var StringParts = InputString.split(delimiter);
              var Part1String = Trim(StringParts[0]);
              var Part2String = Trim(StringParts[1]);
              Logger.log(func + Step + ' InputString: ' + InputString + ', Delimiter: "' + delimiter 
                         + '", Part1String: "' + Part1String + '", Part2String: "' + Part2String +'"');
              if (Part1String){sheet.getRange(Part1Addr).setValue(Part1String)}  
              if (Part2String){sheet.getRange(Part2Addr).setValue(Part2String)}  
            }
            break;
            
          case "Trunc": 
            Step = 2660; // Returns the first characters of a string up to the location of the Param 2 delimiter
            var delimiter = Param2;
            if (delimiter){
              var InputString = Param1;
              var StringParts = InputString.split(delimiter);
              var OutputString = StringParts[0];
              Logger.log(func + Step + ' InputString: ' + InputString + ', Delimiter: "' + delimiter + '", Output String: ' + OutputString);
              if (OutputString){
                sheet.getRange(ResultAddr).setValue(OutputString);  
              }
            }
            break;
            
          case "IFP1=P2ThenP3":
            Step = 2670; // Performs IF,Then,Else test and writes result 
            if (Param1 == Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } 
            break;
            
          case "IFP1!=P2ThenP3":
            Step = 2675; // Performs IF,Then,Else test and writes result 
            if (Param1 != Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } 
            break;
            
          case "IFP1=P2ThenP3ElseP4":
            Step = 2680; // Performs IF,Then,Else test and writes result 
            if (Param1 == Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } else {
              sheet.getRange(ResultAddr).setValue(Param4);
            }
            break;
            
          case "IFP1!=P2ThenP3ElseP4":
            Step = 2685; // Performs IF,Then,Else test and writes result 
            if (Param1 != Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } else {
              sheet.getRange(ResultAddr).setValue(Param4);
            }
            break;
            
          case "IFP1=P2andP3!=P4ThenP5ElseP6":
            Step = 2690; // Performs IF,Then,Else test and writes result 
            if (Param1 == Param2 && Param3 != Param4){
              sheet.getRange(ResultAddr).setValue(Param5);  
            } else {
              sheet.getRange(ResultAddr).setValue(Param6);
            }
            break;
            
          case "Map": 
            Step = 2695; // Builds a composite "result" data field using the "Param 2" template value 
            var MustBePresent = Param1;
            var template = Param2;
            Logger.log(func + Step + ' ERROR - ConvType (' + ConvType + ') Not Working yet...');
            //Logger.log(func + Step + ' InputString: ' + InputString + ', Delimiter: "' + delimiter + '", Output String: ' + OutputString);
            break;
            
          default: 
            var EventMsg = ' ERROR 228 - ConvType (' + ConvType + ') Not Found';
            Logger.log(func + Step + EventMsg);
            LogEvent(EventMsg, oCommon);
            break;
            
        } // End of Switch case
        
        //Record the results for posterity...
        Logger.log(func + Step + ' Values:' 
        + ', Param1: ' + Param1
        + ', Param2: ' + Param2
        + ', Param3: ' + Param3
        + ', Param4: ' + Param4
        + ', Param5: ' + Param5
        + ', Param6: ' + Param6
        + ', Result: ' + SheetData[ArrayRow][ResultCol]);
        
        
      } // End of Loop for each parameter row
    } // End of test for data Conversions
    
    /******************************************************************************/
    Step = 3000; // Push winning entry into ChangedSheetRow array
    /******************************************************************************/ 
    oCommon.ChangedSheetRows.push([[ArrayRow+1],[TagValue],[FormOwnerEmail]]);
    
  } catch (err) {
    // Error condition encountered
    form.setCustomClosedFormMessage(defaultClosedFor);
    form.setAcceptingResponses(true);
    //lock.releaseLock();
    oCommon.DisplayMessage = ' Error 229 - Error condition encountered processing '
      + 'the GoogleForm response data: ' + err.message;
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    Logger.log(func + Step + oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
  
  /******************************************************************************/
  Step = 3000; // Enable the Form and release the public lock
  /******************************************************************************/
  // Source: https://stackoverflow.com/questions/43421593/reject-google-forms-submission-or-remove-response
  form.setCustomClosedFormMessage(defaultClosedFor);
  form.setAcceptingResponses(true);
  //lock.releaseLock();
  
  Logger.log(func + Step + ' END');

  return true;
} 



function doConversions(oCommon) {
  /* ****************************************************************************

   DESCRIPTION:
     This function is invoked to just run the conversion of input values

   REVISION DATE:
    04-06-2019 - First Instance
     
   NOTES:
     1. The array oCommon.SelectedSheetRows[SheetRow][TagValue] is passed identifying the selected rows
     2. oCommon.SelectedSheetRows.length == 0 if an invalid range has been selected by the user     
     
    click on the menu Edit -> Current project's triggers and set up an automated trigger onFormSubmission

  ******************************************************************************/
  var func = "**doConversions " + Version + " - ";
  var Step = 1000;
  Logger.log(func + Step + ' BEGIN');
  
  /****************************************************************************/
  Step = 1000; // Assign the oCommon parameters
  /****************************************************************************/
  var oSourceSheets = oCommon.Sheets;
  var dataTab = oCommon.FormResponse_SheetName;
  var source_TermsRow = GetSheetParam(dataTab, 'Terms Row', oCommon);
  var firstDataRow = GetSheetParam(dataTab, '1st Data Row', oCommon);
  var NoErrors = true;
  
  /******************************************************************************/
  Step = 1000; // Capture the Data Conversion parameters, if they exist
  /******************************************************************************/
  var ParameterSet = oCommon.DataConversions_Heading;
  if (!ParamCheck(ParameterSet)){
    Step = 1100; // Expected Data Conversion Parameters, not found
    oCommon.DisplayMessage = ' Error 430 - Expected Data Conversion Parameters, not found.';
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + oCommon.ReturnMessage);
    return false; 
  }
    
  Step = 1200; // Get the parameters from the Setup sheet
  var oConversionParams = {};
  var ParameterSet = oCommon.DataConversions_Heading;  
  oConversionParams = ParamSetBuilder(oCommon, ParameterSet);
  
  if (ParamCheck(oConversionParams)){
    Step = 1210; // Conversion Parameters are active
    Logger.log(func + Step + ' ParameterSet: ' + ParameterSet 
               + ', Parameters: ' + oConversionParams.Parameters);
    bDoConversions = true;
  } else {      
    Step = 1220; // Unable to find and load Conversion Parameters.
    oCommon.DisplayMessage = ' Error 432 - Unable to find and load Conversion Parameters..';
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + oCommon.ReturnMessage);
    bDoConversions = false;
    
  }
  
  /**************************************************************************/
  Step = 2000; // Select records
  /**************************************************************************/
  var mode = 2;
  if(!SelectSheetRows(oCommon, mode, dataTab, firstDataRow)){
    Step = 2110; // ERROR - No Selected Rows found
    oCommon.DisplayMessage = ' Error 434 - No Rows selected. Procedure will be terminated.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    return false;
  }
  Logger.log(func + Step + ' SelectedSheetRows.length:' + oCommon.SelectedSheetRows.length);
  
  Step = 2200; // Take time to confirm that the user wants to proceed
  var MsgBoxMessage = ('The following rows were selected: ' + FormatForDisplay(oCommon.SelectedSheetRows)
  + '\\n' + ' Do you want to continue?');
  if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
    var display_message = "Your response is NO - procedure will be terminated";
    return_message = func + Step + ' User terminated after rows were selected';
    oCommon.ReturnMessage = return_message;
    oCommon.DisplayMessage = display_message;
    return false;
  }
  
  /**********************************************************************************/
  Step = 3000; //Loop through the Selected Sheet rows
  /**********************************************************************************/
  
  Step = 3100; // Prepare a Key/Value array for the Source Sheet terms/heading labels
  var TargetSheetName = dataTab, // Sheet Containing Keys and Values
    SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start row is always the following row
    KeyRow = source_TermsRow,   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
    ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
    KeyCol = '',   // number or text [Look for text Match in StartRow (Base 0)]
    ValueCol = '', // number or text [Look for text Match in StartRow (Base 0)]
    oSourceHeadingsXrefAry = {}; // if declared before entering function BuildKeyValueAry()
  //oSourceHeadingsXrefAry = BuildKeyValueAry(oSourceSheets,TargetSheetName,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol,oSourceHeadingsXrefAry)
  oSourceHeadingsXrefAry = BuildKeyValueAry(oSourceSheets,TargetSheetName,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol);
  Step = 3110; // Verify
  for(var key in oSourceHeadingsXrefAry){
    Logger.log(func + Step + ' Key: ' + key + '  Value: ' + oSourceHeadingsXrefAry[key]);
  }
  
  try {
  
    /******************************************************************************/
    Step = 3200; // Set a public lock that locks for all invocations
    /******************************************************************************/
    // http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
    var lock = LockService.getPublicLock();
    lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
    // got the lock, we may now proceed
    
    /**********************************************************************************/
    Step = 3300; //Loop through the Selected Sheet rows
    /**********************************************************************************/
    var sheet = oSourceSheets.getSheetByName(dataTab);
    var SheetData = sheet.getDataRange().getValues();
  
    for (var k = 0; k < oCommon.SelectedSheetRows.length; k++){
      Step = 3310; 
      var ArrayRow = oCommon.SelectedSheetRows[k][0] - 1; // convert back to Base 0
      Logger.log(func + Step + ' k: ' + k + ', source_row: ' + ArrayRow + ', Length: ' + SheetData[ArrayRow].length);
    
      Step = 3320; // Perform data conversions, if defined
      for( var prow = 1; prow < oConversionParams.Parameters.length; prow++){
        Step = 3325; // Conduct a conversion for each row in the oConversionParams.Parameters array
        var ConvType = oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Type of Conversion']];
        var ResultCol = oSourceHeadingsXrefAry[oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Result Column Heading']]];
        var ResultAddr = numToA(ResultCol+1) + String(ArrayRow+1).split(".")[0];
        
        var Param1Col = oSourceHeadingsXrefAry[oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 1']]];
        if (!isNaN(Param1Col)) {
          // Get the Form Data value for the current row associated with Param1
          var Param1 = Trim(SheetData[ArrayRow][Param1Col]);
        } else {
          // Get the value of the constant
          var Param1 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 1']]);
        }
        
        var Param2Col = oSourceHeadingsXrefAry[oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']]];
        if (!isNaN(Param2Col)) {
          // Get the Form Data value for the current row associated with Param2
          var Param2 = Trim(SheetData[ArrayRow][Param2Col]);
        } else {
          // Get the value of the constant
          var Param2 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']]);
        }
        
        var Param3Col = oSourceHeadingsXrefAry[oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']]];
        if (!isNaN(Param3Col)) {
          // Get the Form Data value for the current row associated with Param3
          var Param3 = Trim(SheetData[ArrayRow][Param3Col]);
        } else {
          // Get the value of the constant
          var Param3 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']]);
        }
        
        var Param4Col = oSourceHeadingsXrefAry[oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 4']]];
        if (!isNaN(Param4Col)) {
          // Get the Form Data value for the current row associated with Param4
          var Param4 = Trim(SheetData[ArrayRow][Param4Col]);
        } else {
          // Get the value of the constant
          var Param4 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 4']]);
        }
        
        var Param5Col = oSourceHeadingsXrefAry[oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 5']]];
        if (!isNaN(Param5Col)) {
          // Get the Form Data value for the current row associated with Param5
          var Param5 = Trim(SheetData[ArrayRow][Param5Col]);
        } else {
          // Get the value of the constant
          var Param5 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 5']]);
        }
        
        var Param6Col = oSourceHeadingsXrefAry[oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 6']]];
        if (!isNaN(Param6Col)) {
          // Get the Form Data value for the current row associated with Param6
          var Param6 = Trim(SheetData[ArrayRow][Param6Col]);
        } else {
          // Get the value of the constant
          var Param6 = Trim(oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 6']]);
        }
          
        Logger.log(func + Step + ' Values:' 
        + ' ArrayRow: ' + ArrayRow
        + ', prow: ' + prow 
        + ', ConvType: ' + ConvType 
        + ', ResultCol: ' + ResultCol
        + ', ResultAddr: ' + ResultAddr
        + ', Param1: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 1']] + ' Param1Col: ' + Param1Col + ' Value: ' + Param1
        + ', Param2: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']] + ' Param2Col: ' + Param2Col + ' Value: ' + Param2
        + ', Param3: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']] + ' Param3Col: ' + Param3Col + ' Value: ' + Param3
        + ', Param4: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 4']] + ' Param4Col: ' + Param4Col + ' Value: ' + Param4
        + ', Param5: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 5']] + ' Param5Col: ' + Param5Col + ' Value: ' + Param5
        + ', Param6: ' + oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 6']] + ' Param6Col: ' + Param6Col + ' Value: ' + Param6);
                       
        switch (ConvType) {
            
          case "Date-To-Day": 
            Step = 3410; // Conduct a "Date-To-Day" conversion
            var dParam1 = new Date(Param1);
            if (isValidDate(dParam1)){
              sheet.getRange(ResultAddr).setValue(GetDay(dParam1));  
            }
            break;
            
          case "DateDiff-Year": 
            Step = 3420; // Conduct a "DateDiff-Year" conversion
            var dParam1 = new Date(Param1);
            var dParam2 = new Date(Param2);
            if (isValidDate(Param1) && isValidDate(Param2)){
              var DateDifference = DateDiff(dParam1,dParam2,'YEARS');
              if (DateDifference){
                if (DateDifference > 0){
                  sheet.getRange(ResultAddr).setValue(DateDiff(dParam1,dParam2,'YEARS'));  
                  sheet.getRange(ResultAddr).setNumberFormat("@");
                }
              }
            }
            break;
            
          case "DateDiff-Age": 
            Step = 3430; // Conduct a "DateDiff-Year" conversion
            var dParam1 = new Date(Param1);
            var dParam2 = new Date(Param2);
            if (isValidDate(dParam1) && isValidDate(dParam2)){
              var DateDifference = DateDiff(dParam1,dParam2,'MONTHS');
              var Age = DateDifference / 12;
              var Age = Math.floor(Age);
              if (Age){
                sheet.getRange(ResultAddr).setValue(Age);  
                sheet.getRange(ResultAddr).setNumberFormat("@");
              }
            }
            break;
            
          case "Const": 
            Step = 3440; // Conduct a "Const" conversion
            if (Param1){
              sheet.getRange(ResultAddr).setValue(Param1);  
            }
            break;
            
          case "Split": 
            Step = 3450; // Splits the input string(Param1) into 2 parts using the Param 2 delimiter
            // and places Part 1 into the Result col and Part 2 into the Param 3 column 
            var Part1Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Result Column Heading']]);
            if(isNaN(Part1Col)){
              var EventMsg = func + Step + ' ERROR - Invalid entry in Conversion Parameters row ' + prow + ' for "Result Column Heading"';
              Logger.log(EventMsg);
              LogEvent(EventMsg, oCommon);
              break;
            }
            var Part1Addr = numToA(Part1Col+1) + String(ArrayRow+1).split(".")[0];
            var Part2Col = GetXrefValue(oCommon, oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 3']]);
            if(isNaN(Part2Col)){
              var EventMsg = func + Step + ' ERROR - Invalid entry in Conversion Parameters row ' + prow + ' for Param 3';
              Logger.log(EventMsg);
              LogEvent(EventMsg, oCommon);
              break;
            }
            var Part2Addr = numToA(Part2Col+1) + String(ArrayRow+1).split(".")[0];
            
            // write over the contents of the part2 fields
            sheet.getRange(Part2Addr).setValue('')
            
            var InputString = Param1;
            if(!ParamCheck(InputString)){
              // no string found that needs to be split
              break;
            }
            var delimiter = oConversionParams.Parameters[prow][oConversionParams.HeadingKeyAry['Param 2']];
            if(!ParamCheck(delimiter)){
              var EventMsg = func + Step + ' ERROR - Invalid entry in Conversion Parameters row ' + prow + ' for Param 2';
              Logger.log(EventMsg);
              LogEvent(EventMsg, oCommon);
              break;
            }
            Logger.log(func + Step + ' InputString: "' + Param1 + '", Delimiter: "' + delimiter 
                       + '", Part1Addr : ' + Part1Addr  + '", Part2Addr: ' + Part2Addr);
            if(InputString.indexOf(delimiter) > -1){
              Step = 3455; // delimiter is found within the input string
              var StringParts = InputString.split(delimiter);
              var Part1String = Trim(StringParts[0]);
              var Part2String = Trim(StringParts[1]);
              Logger.log(func + Step + ' InputString: ' + InputString + ', Delimiter: "' + delimiter 
                         + '", Part1String: "' + Part1String + '", Part2String: "' + Part2String +'"');
              if (Part1String){sheet.getRange(Part1Addr).setValue(Part1String)}  
              if (Part2String){sheet.getRange(Part2Addr).setValue(Part2String)}  
            }
            break;
            
          case "Trunc": 
            Step = 3460; // Returns the first characters of a string up to the location of the Param 2 delimiter
            var delimiter = Param2;
            if (delimiter){
              var InputString = Param1;
              var StringParts = InputString.split(delimiter);
              var OutputString = StringParts[0];
              Logger.log(func + Step + ' InputString: ' + InputString + ', Delimiter: "' + delimiter + '", Output String: ' + OutputString);
              if (OutputString){
                sheet.getRange(ResultAddr).setValue(OutputString);  
              }
            }
            break;
            
          case "IFP1=P2ThenP3":
            Step = 3470; // Performs IF,Then,Else test and writes result 
            if (Param1 == Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } 
            break;
            
          case "IFP1!=P2ThenP3":
            Step = 3475; // Performs IF,Then,Else test and writes result 
            if (Param1 != Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } 
            break;
            
          case "IFP1=P2ThenP3ElseP4":
            Step = 3480; // Performs IF,Then,Else test and writes result 
            if (Param1 == Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } else {
              sheet.getRange(ResultAddr).setValue(Param4);
            }
            break;
            
          case "IFP1!=P2ThenP3ElseP4":
            Step = 3485; // Performs IF,Then,Else test and writes result 
            if (Param1 != Param2){
              sheet.getRange(ResultAddr).setValue(Param3);  
            } else {
              sheet.getRange(ResultAddr).setValue(Param4);
            }
            break;
            
          case "IFP1=P2andP3!=P4ThenP5ElseP6":
            Step = 3490; // Performs IF,Then,Else test and writes result 
            if (Param1 == Param2 && Param3 != Param4){
              sheet.getRange(ResultAddr).setValue(Param5);  
            } else {
              sheet.getRange(ResultAddr).setValue(Param6);
            }
            break;
            
          default: 
            var EventMsg = ' Error 436 - Invalid Conversion Type Name (' + ConvType + ').';
            Logger.log(func + Step + EventMsg);
            LogEvent(EventMsg, oCommon);
            break;
            
        } // End of Switch case
        
        //Record the results for posterity...
        Logger.log(func + Step + ' Values:' 
        + ', Param1: ' + Param1
        + ', Param2: ' + Param2
        + ', Param3: ' + Param3
        + ', Param4: ' + Param4
        + ', Param5: ' + Param5
        + ', Param6: ' + Param6
        + ', Result: ' + SheetData[ArrayRow][ResultCol]);
        
      } // End of Loop for each parameter row
    } // Next Selected Sheet row
    
  } catch (err) {
    // Error condition encountered
    lock.releaseLock();
    oCommon.DisplayMessage = ' Error 438 - Error condition encountered processing. '
      + 'the GoogleForm response data: ' + err.message;
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    Logger.log(func + Step + oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
    
  /******************************************************************************/
  Step = 4000; // Enable the Form and release the public lock
  /******************************************************************************/
  lock.releaseLock();
  
  Logger.log(func + Step + ' END');

  return true;
} 


function moveSelectedRows(oCommon, oMenuParams){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to move selected rows from the Response tab to the 
       saved history tab (as defined in the Setup Globals tab)
     
   USEAGE
     return_bool = moveSelectedRows(oCommon);

   REVISION DATE:
    01-16-2019 - First instance
    02-10-2019 - Modified to avoid deletion of associated documents and files
                   if they are not to be deleted
    02-17-2019 - Modified to get Target Tab from oMenuParams object               
    

  ******************************************************************************/
  
  var func = "***MoveSelectedRows " + Version + " - ";
  var Step = 100;
  Logger.log(func + ' BEGIN');
  
  var test_mode = false; //true;
  var NumRowsMoved = 0;
  var return_message = '';
  
  /******************************************************************************/
  Step = 1000; // Examine and validate input parameters
  /******************************************************************************/
  Step = 1010; // Validate oMenuParams
  if (oMenuParams.length <= 0){
    // Leave since there is nothing to do
    oCommon.DisplayMessage = ' Warning 260 - Expected oMenu Parameters not found.';
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + oCommon.ReturnMessage);
    return false;
  }
  
  /****************************************************************************/
  Step = 1100; // Assign the oCommon parameters
  /****************************************************************************/
  var source_TabName = oCommon.XrefSourceSheet;
  var target_TabName = oMenuParams['DestTab'];
  var TimestampCol = oCommon.TimestampCol;
  var Edit_URLCol = oCommon.Edit_URLCol;
  
  Step = 1200; //Get the Source and Target Sheet details from the oCommon.sheetParams object
  var source_HeadingRow = GetSheetParam(source_TabName, 'Headings Row', oCommon);
  var source_FirstDataRow = GetSheetParam(source_TabName, '1st Data Row', oCommon);
  var target_HeadingRow = GetSheetParam(target_TabName, 'Headings Row', oCommon);
  var target_FirstDataRow = GetSheetParam(target_TabName, '1st Data Row', oCommon);

  Step = 1300; //Get the Source and Target Sheet details from the oCommon.sheetParams object
  if (!ParamCheck(source_TabName) || !ParamCheck(source_HeadingRow) || !ParamCheck(source_FirstDataRow)
        || !ParamCheck(TimestampCol) || !ParamCheck(Edit_URLCol) || Edit_URLCol == TimestampCol
        || !ParamCheck(target_TabName) || !ParamCheck(target_HeadingRow) || !ParamCheck(target_FirstDataRow) ){
    var user_message = 'Sorry, we have to close the program due to an internal software error.' 
      + ' (MoveSelectedRows)';
    oCommon.DisplayMessage = ' Error 262 - One or more input values empty or invalid.';
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + ' ' + return_message); 
          Logger.log(func + Step + ' PARAMETERS:' 
      + ' SOURCE: Tab: ' + source_TabName
      + ', HdgRow: ' + source_HeadingRow
      + ', 1stDataRow: ' + source_FirstDataRow
      + ', TimestampCol: ' + TimestampCol
      + ', Edit_URLCol: ' + Edit_URLCol
      + ', TARGET: Tab: ' + target_TabName
      + ', HdgRow: ' + target_HeadingRow
      + ', 1stDataRow: ' + target_FirstDataRow
     );
    return false;
  }
  
  Step = 1400; // Get the selected rows to be moved
  var mode = 2; // get user-selected rows
  if(!SelectSheetRows(oCommon,mode,source_TabName,source_FirstDataRow)){
    // WARNING - No Selected Rows found
    return_message = func + Step + ' ***** WARNING - No Rows Selected';
    Logger.log(return_message);
    return true;
  }
  var row_count = oCommon.SelectedSheetRows.length;
  Logger.log(func + Step + ' SelectedSheetRows.length: ' + row_count);

  Step = 1500; // Take time to confirm that the user wants to proceed
  if (!oCommon.bSilentMode){
    var MsgBoxMessage = ('Google Form Responses for the following row(s) will be moved to the "' 
       + target_TabName + '" tab: ' + FormatForDisplay(oCommon.SelectedSheetRows)
       + '\\n' + ' Do you want to continue?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      oCommon.DisplayMessage = 'OK, the Move process will be terminated.  Thank you!';
      //Browser.msgBox(oCommon.DisplayMessage);
      oCommon.ReturnMessage = func + Step + ' ********** WARNING - User terminated after rows were selected';
      return false;
    }
  }
  
  Step = 1600; // Build the FromTo Key Value Array of headings Column Associations
  var oSourceSheets = oCommon.Sheets;
  var SourceSheetData = oSourceSheets.getSheetByName(source_TabName).getDataRange().getValues();
  var TargetSheet = oSourceSheets.getSheetByName(target_TabName);
  
  Step = 1610; // Prepare a Key/Value array for the Source Sheet heading labels
  var oSourceHeadingsXrefAry = {}; 
  oSourceHeadingsXrefAry = oCommon.HeadingsXrefAry;
  //Step = 1615; // Verify
  //for(var key in oSourceHeadingsXrefAry){
  //  Logger.log(func + Step + ' Key: ' + key + '  Value: ' + oSourceHeadingsXrefAry[key]);
  //}

  Step = 1620; // Prepare a Key/Value array for the Target Sheet heading labels
  var TargetSheetName = target_TabName, // Sheet Containing Keys and Values
    SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start row is always the following row
    KeyRow = target_HeadingRow,   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
    ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
    KeyCol = '',   // number or text [Look for text Match in StartRow (Base 0)]
    ValueCol = '', // number or text [Look for text Match in StartRow (Base 0)]
    oTargetHeadingsXrefAry = {}; // if declared before entering function BuildKeyValueAry()
  oTargetHeadingsXrefAry = BuildKeyValueAry(oSourceSheets,TargetSheetName,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol);
  Step = 1635; // Verify
  //for(var key in oTargetHeadingsXrefAry){
  //  Logger.log(func + Step + ' Key: ' + key + '  Value: ' + oTargetHeadingsXrefAry[key]);
  //}
  
  if(!oTargetHeadingsXrefAry){ return false; }
  
  Step = 1640; // Build the FromTo a Key/Value array
  var oFromToXrefAry = {}; // if declared before entering function BuildKeyValueAry()
  for(var key in oSourceHeadingsXrefAry){
    if(!isNaN(oTargetHeadingsXrefAry[key])){oFromToXrefAry[oSourceHeadingsXrefAry[key]] = oTargetHeadingsXrefAry[key]}
  }
  Step = 1645; // Verify
  //for(var key in oFromToXrefAry){
  //  Logger.log(func + Step + ' Key: ' + key + '  Value: ' + oFromToXrefAry[key]);
  //}
  
  /******************************************************************************/
  Step = 2000; // loop through the SelectedSheetRows array and move rows
  /******************************************************************************/
  var added_rows = 0;
  var outputRows = []; // holder for the range of values to be moved
  //if (!ParamCheck(TimestampCol) || TimestampCol < 0){ TimestampCol = SourceSheetData[0].length + 1; }
  
  for (var k = 0; k < oCommon.SelectedSheetRows.length; k++){
    Step = 2010; 
    var source_row = oCommon.SelectedSheetRows[k][0] - 1; // convert back to Base 0
    Logger.log(func + Step + ' k: ' + k + ', source_row: ' + source_row + ', Length: ' + SourceSheetData[source_row].length);
    Step = 2020; // grab the column values in the selected source row if it is to be moved
    outputRows[k] = [];
    for (var c = 0; c < SourceSheetData[source_row].length; c++){
      var action = '';
      var target_col = oFromToXrefAry[c]; // will contain the column number for the target cell, if the value is to be moved
      if(!isNaN(target_col)){
        Step = 2030; // place the source value into the translated target cell
        outputRows[k][target_col] = SourceSheetData[source_row][c];
        //Logger.log(func + Step + ' k: ' + k + ', c: ' + c + ', contents: ' + SourceSheetData[source_row][c]
        //     + ', Target Col: ' + target_col);
      }
    }
    var event_message = func + Step + ' *** "' + source_TabName + '" Row ' + oCommon.SelectedSheetRows[k][0] 
      + ' (' + oCommon.SelectedSheetRows[k][1] + ') copied to "' + target_TabName + '" tab.';
    LogEvent(event_message, oCommon);
    Logger.log(event_message);
    
    added_rows++;
  }
  if(outputRows.length <=0){
    Step = 2050; // ERROR - No Selected Rows found
    return_message = func + Step + ' ********** ERROR - At least one row to move should have been found';
    Logger.log(return_message);
    oCommon.DisplayMessage = ' Error 266 - No rows to be moved were successfully located.';
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
  
  Step = 2100; // Write the outputRow[] values to the target sheet
  var TargetSheetData = TargetSheet.getDataRange().getValues();
  var target_last_row = TargetSheet.getLastRow();  // last row that has content
  //Logger.log(func + Step + ' range values: (' + (target_last_row + 1) + ', 1, ' + outputRows.length + ', ' + outputRows[0].length +')');
  TargetSheet.getRange(target_last_row + 1,1,outputRows.length,outputRows[0].length).setValues(outputRows);
  SpreadsheetApp.flush();
  Logger.log(func + Step + ' Total Rows Moved: ' + added_rows);
  
  Step = 2200; // Take time to confirm that the selected rows were moved
  if (!oCommon.bSilentMode){
    var MsgBoxMessage = ('Google Form Responses for the following row(s) were successfully copied to the "' 
       + target_TabName + '" tab: ' + FormatForDisplay(oCommon.SelectedSheetRows)
       + '\\n' + ' Do you want to proceed to the "Delete" process?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      oCommon.DisplayMessage = 'OK, the Move process will now be terminated without deleting any rows.  Thank you!';
      //Browser.msgBox(oCommon.DisplayMessage);
      oCommon.ReturnMessage = func + Step + ' User terminated the "Move" process before the moved rows were deleted.';
      return true;
    }
  }

  /******************************************************************************/
  Step = 3000; // Delete the SelectedSheetRows from the SourceSheet
  /******************************************************************************/
  if(deleteSelectedRows(oCommon, oFromToXrefAry)){return true;} else {return false;}

}    
 

function deleteSelectedRows(oCommon, oFromToXrefAry) {
  /* ****************************************************************************

   DESCRIPTION:
    This function is invoked whenever one or more sheet rows are to be deleted:
         1 - Locate the target sheet row(s)
         2 - Delete the associated Google Form Response (if any)
         3 - Delete the selected sheet row(s)
         
   USEAGE:
   
     bReturnBool = deleteSelectedRows(oCommon, oFromToXrefAry)

   REVISION DATE:
    11-03-2018 - First Instance
    02-08-2019 - Added code to delete any non-Google Form file urls found in the row
    04-17-2019 - Added oFromToXrefAry as a passed parameter from the MoveSelectedRows() function
                  so that "moved" cols containing URLs would not have teh associated files deleted
     
   NOTES:
     1. The array oCommon.SelectedSheetRows[SheetRow][TagValue] is passed identifying the selected rows
     2. oCommon.SelectedSheetRows.length == 0 if an invalid range has been selected by the user     
     
  ******************************************************************************/
  var func = "***deleteSelectedRows " + Version + " - ";
  var Step = 1000;
  Logger.log(func + Step + ' BEGIN');
  
  bTestMode = false; //true
  
  if(ParamCheck(oFromToXrefAry)){
    var bCheckForMovedUrl = true;
  } else {
    var bCheckForMovedUrl = false;
  }
  
  /****************************************************************************/
  Step = 1100; // Assign the oCommon parameters
  /****************************************************************************/
  var formID = oCommon.GoogleForm_Url;
  var dataTab = oCommon.XrefSourceSheet; //dataTab = oCommon.FormResponse_SheetName;
  var firstDataRow = GetSheetParam(dataTab, '1st Data Row', oCommon); //firstDataRow = oCommon.FormResponseStartRow;
  var AwesomeTableRow = GetSheetParam(dataTab, 'AWT Row', oCommon); //AwesomeTableRow = oCommon.AwesomeTableRow; // Base 0; ref https://support.awesome-table.com/hc/en-us/articles/115001068129-Data-configuration
  var timestampCol = oCommon.TimestampCol;
  var StatusCol = oCommon.Record_statusCol;
  var ActiveStatusValue = oCommon.ActiveStatusValue;
  var Inactive_Status_value = oCommon.Inactive_Status_value;
  var rptDateCol = oCommon.AsOfDateCol;
  var editUrlCol = oCommon.Edit_URLCol;
  var editButtonCol = oCommon.Edit_ButtonCol;
  var editButtonText = oCommon.Edit_Button_Text;
  var error_email = oCommon.Send_error_report_to;
  
  var Errors = 0;
  var NoErrors = true;
  
  /******************************************************************************/
  Step = 2000; // Get the Selected Rows, if necessary
  /******************************************************************************/
  if (oCommon.SelectedSheetRows.length <= 0){
    
    Step = 2100; // Get the Selected Rows
    
    var mode = 2; // Retrieve user-selected rows / var mode = 1 Retrieve ALL rows
    Logger.log(func + Step + ' mode = ' + mode + ', firstDataRow: ' + firstDataRow);
      
    if(!SelectSheetRows(oCommon, mode, dataTab, firstDataRow+1)){
      // WARNING - No Selected Rows found
      oCommon.DisplayMessage = ' WARNING - No Selected Rows found';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      Logger.log(oCommon.ReturnMessage);
      return true;
    }
    
    Logger.log(func + Step + ' SelectedSheetRows.length: ' + oCommon.SelectedSheetRows.length);
  
  }
  
  Step = 2200; // Take time to confirm that the user wants to proceed
  if (!oCommon.bSilentMode){
    var MsgBoxMessage = ('Google Form Responses for the following row(s) will be deleted from the "' 
       + dataTab + '" tab: ' + FormatForDisplay(oCommon.SelectedSheetRows)
       + '\\n' + ' Do you want to continue?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      oCommon.DisplayMessage = 'OK, the Delete process will be terminated.  Thank you!';
      //Browser.msgBox(oCommon.DisplayMessage);
      oCommon.ReturnMessage = func + Step + ' ********** WARNING - User terminated after rows were selected';
      return false;
    }
  }

  //Determine scope
  if(!isNaN(editUrlCol)){
    var bDeleteGoogleForms = true;
  } else { 
    var bDeleteGoogleForms = false;
  }
  
  /****************************************************************************/
  Step = 3000; // Warn User that Deleting Google Forms cannot be undone...
  /****************************************************************************/
  if(!oCommon.bSilentMode){
    var part_a = '';
    var part_b = '';
    if (bDeleteGoogleForms) {
      Step = 3100; // Confirm that User also wants to delete Google Form Responses
      var part1 = 'The Google FORM ENTRIES for the following row(s) WILL BE DELETED:'
      + FormatForDisplay(oCommon.SelectedSheetRows)
      var part_a = ' also';  
      var user_message = part1
      + '\\n\\nThis action CANNOT BE UNDONE. \\n\\nDo you want to continue?';
      var return_answer = Browser.msgBox('WARNING', user_message, Browser.Buttons.YES_NO);
      if (return_answer !== 'yes') {
        Step = 3150; // User elects to terminate Delete process
        Browser.msgBox('OK, the Delete process will be terminated.  Thank you!');
        oCommon.DisplayMessage = ' User Choice - User elects to terminate Delete process';
        oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
        Logger.log(func + Step + oCommon.ReturnMessage);
        return true;
      } 
    }
    Step = 3200; // Confirm that User also wants to delete associated files
    if (bCheckForMovedUrl) {
      var part_b = ' referenced in columns that were not moved';
    }
    var user_message = 'Do you' + part_a + ' want to delete all of the associated documents' + part_b + '?';
    var return_answer = Browser.msgBox('PLEASE CONFIRM', user_message, Browser.Buttons.YES_NO);
    
    if (return_answer == 'no') {
      var part2 = 'The DOCUMENTS AND FILES' + part_b + ' associated with the selected row(s) will NOT BE DELETED.';
      var part3 = 'The SELECTED ROW(s) WILL BE DELETED from the "' + dataTab + '" Tab.'
      var bDeleteFiles = false;
    } else if(return_answer == 'yes'){ 
      var part2 = 'ALL DOCUMENTS AND FILES' + part_b + ' associated with the selected row(s) WILL BE DELETED.';
      var part3 = 'The SELECTED ROW(s) WILL BE DELETED from the "' + dataTab + '" Tab.'
      var bDeleteFiles = true;
    } else {
      Step = 3250; // User elects to terminate Delete process
      Browser.msgBox('OK, the Delete process will be terminated.  Thank you!');
      oCommon.DisplayMessage = ' User Choice - User elects to terminate Delete process';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      Logger.log(func + Step + oCommon.ReturnMessage);
      return true;
    }
    
    Step = 3300; // Confirm
    if (bDeleteGoogleForms){
      var user_message = 'OK:'
      + '\\n  1. ' + part1
      + '\\n  2. ' + part2
      + '\\n  3. ' + part3
      + '\\n Please confirm that this is what you want to happen.'
    } else {
      var user_message = 'OK:'
      + '\\n  1. ' + part2
      + '\\n  2. ' + part3
      + '\\n Please confirm that this is what you want to happen.'
    }
    var return_answer = Browser.msgBox('PLEASE CONFIRM', user_message, Browser.Buttons.YES_NO);
    if(return_answer !== 'yes'){ 
      Step = 3350; // User elects to terminate Delete process
      Browser.msgBox('OK, the Delete process will be terminated.  Thank you!');
      oCommon.DisplayMessage = ' User Choice - User elects to terminate Delete process';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      Logger.log(func + Step + oCommon.ReturnMessage);
      return true;
    }
  }  // End of bSilentMode test
  
  /******************************************************************************/
  Step = 4000; // Retreive the GoogleForm response data stored separately from the Google Sheet data
  /******************************************************************************/
  Logger.log(func + Step + ' dataTab: ' + dataTab);
  
  Step = 4100; // Open the Google Form
  if (bDeleteGoogleForms){
    try {
      formID = getIdFrom(formID);
      var form = FormApp.openById(formID); // formID assigned from oCommon
    } catch(err) {
      Step = 4150; // PERMISSION ERROR - Delete process will be terminated
      oCommon.DisplayMessage = ' Error 266 - Google Form Permission Error: ' + err.message;
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      Logger.log(func + Step + oCommon.ReturnMessage);
      LogEvent(oCommon.ReturnMessage, oCommon);
      return false;
    }
  
    try {
      Step = 4200; // Temporarily disable the form from accepting responses
      // Source: https://stackoverflow.com/questions/43421593/reject-google-forms-submission-or-remove-response
      var defaultClosedFor = form.getCustomClosedFormMessage();
      form.setCustomClosedFormMessage("The form is currently processing another request, please refresh the page.");
      form.setAcceptingResponses(false);
      
      Step = 4210; // Retrieve the form responses and item objects
      var responses = form.getResponses();  
      var items = form.getItems();
      
      Step = 4220; // Capture the Form Timestamps and Form ID's from the Form object
      //Logger.log(func + Step + ' Got here...')
      var timestamps =[];
      var FormIDs = [];
      for (var i = 0; i < responses.length; i++) {
        timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
        FormIDs.push(responses[i].getId());
      }
    } catch(err) {
      Step = 5150; // Form access ERROR - Delete process will be terminated
      oCommon.DisplayMessage = ' Error 268 - Google Form Access Error: ' + err.message;
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      Logger.log(func + Step + oCommon.ReturnMessage);
      LogEvent(oCommon.ReturnMessage, oCommon);
      // Enable the form to accept responses and return
      form.setAcceptingResponses(true);
      form.setCustomClosedFormMessage(defaultClosedFor);
      return false;
    }
  } // End of "bDeleteGoogleForms" test
    
  /******************************************************************************/
  Step = 5000; // Examine each SelectedSheet row 
  /******************************************************************************/
   try {
    var oSourceSheets = oCommon.Sheets; 
    var sheet = oSourceSheets.getSheetByName(dataTab);
    var SheetData = sheet.getDataRange().getValues();
    var ResponsesDeleted = 0;
    var RowsDeleted = 0;
    var RowsExamined = 0;
    Logger.log(func + Step + ' timestampCol: ' + timestampCol 
                + ', Number of rows: ' + oCommon.SelectedSheetRows.length);
    
    Step = 5100; /// Start with the last row in the selected range
    for (var r = oCommon.SelectedSheetRows.length; r > 0; r--) {
      Step = 5110; // 
      var SheetRow = oCommon.SelectedSheetRows[r-1][0];
      var Tag = oCommon.SelectedSheetRows[r-1][1]; 
      var ArrayRow = SheetRow - 1; // SheetData and SelectedSheetRows are base 0
      var ActiveCell = sheet.getRange(SheetRow,1);
      ActiveCell.activate();
      RowsExamined++;
      
      if (bDeleteGoogleForms){
        /******************************************************************************/
        Step = 5200; // Delete Form Response and selected Sheet rows if the sheetTimestamp is valid
        /******************************************************************************/
        var sheetTimestamp = SheetData[ArrayRow][timestampCol];
        var bResponseExists = false;
        
        if (isValidDate(sheetTimestamp)){
          Step = 5210; 
          var responseId = [sheetTimestamp?FormIDs[timestamps.indexOf(sheetTimestamp.setMilliseconds(0))]:''];
          Logger.log(func + Step + ' responseId: "' + responseId + '"');
          
          if (ParamCheck(responseId) && responseId != ''){
            bResponseExists = true;
            
            Step = 5220; // Delete the Form Response
            if(!bTestMode){
              
              try {
                FormApp.openById(formID).deleteResponse(responseId);
                //FormApp.getActiveForm().deleteResponse(responseId);
                var event_message = func + Step + ' *** "' + dataTab + '" Row ' 
                + SheetRow + ' (' + Tag + ') Google Form Response Deleted.';
                LogEvent(event_message, oCommon);
                ResponsesDeleted++;
                Logger.log(event_message + ', responseId: ' + responseId);
              } catch (err) {
                // Error condition encountered
                var event_message = func + Step + ' Error 270 - Google Form Delete Error - Row: ' 
                + SheetRow + ' (' + Tag + ')' + ' Error Msg: ' + err.message;
                LogEvent(event_message, oCommon);
                Logger.log(event_message + ', responseId: ' + responseId);
                Errors++;
              }
              
            } else {
              Step = 5230;
              var event_message = func + Step + ' *** Form Response NOT DELETED due to bTestMode setting -' 
              + ' SheetRow: ' + SheetRow + ' (' + Tag + ')';
              LogEvent(event_message, oCommon);
              Logger.log(event_message + ', responseId: ' + responseId);
            }
          } else {
            Step = 5240; // No Form Response found
            var event_message = func + Step + ' *** Google Form Response Not Found -' + ' SheetRow: ' 
            + SheetRow + ' (' + Tag + ')';
            LogEvent(event_message, oCommon);
            Logger.log(event_message);
          }
        }// End of sheetTimestamp exists test   
      } // End of bDeleteGoogleForms = true
      
      /****************************************************************************/
      Step = 5300; // Delete any previously written Reports and files
      /****************************************************************************/
      if (bDeleteFiles){
        for(var col = 0; col < SheetData[ArrayRow].length; col++){
          Step = 5310; // Test each cell to determine if it contains a url reference to a Google Doc or pdf
          // bypass the EditUrl column
          var bDeleteOK = true;
          if (col !== editUrlCol){
            var test_value = SheetData[ArrayRow][col].toString();
            //Logger.log(func + Step + ' Array Row: ' + ArrayRow + ' col: ' + col + ' test_value: ' + test_value);
            if(test_value.indexOf("https://") > -1){
              Step = 5320; // URL found
              Logger.log(func + Step + ' Array Row: ' + ArrayRow + ' col: ' + col + ' test_value: ' + test_value);
              if (bCheckForMovedUrl){
                if (oFromToXrefAry[col]) {
                  //valid entry found; DO NOT delete
                  bDeleteOK = false;
                }
              }
              if (bDeleteOK){
                var fileURL = SheetData[ArrayRow][col];
                Logger.log(func + Step + ' Array Row: ' + ArrayRow + ' URL to be deleted: ' + fileURL);
                if (deleteGoogleDoc(fileURL, oCommon)) {
                  Step = 5330; // Log Successful delete
                  Logger.log(func + Step + ' Deleted Doc URL: ' + fileURL);
                } else {
                  Step = 5340; // Log Unsuccessful delete
                  var event_message = func + Step + ' Error 272 - Unable to Delete URL: ' + fileURL;
                  LogEvent(event_message, oCommon);
                  Logger.log(func + Step + EventMsg);
                }
              } // End of "bDeleteOK" test
            } // End of Cell contains URL test   
          } // End of Cell does NOT contain Edit URL test
        } // next column
      }
      
      /****************************************************************************/
      Step = 5400; // Delete sheet row
      /****************************************************************************/
      sheet.deleteRow(SheetRow);
      var event_message = func + Step + ' *** "' + dataTab + '" Row ' 
        + SheetRow + ' (' + Tag + ') Deleted.';
      LogEvent(event_message, oCommon);
      Logger.log(event_message);
      RowsDeleted++;
    } // End of Examine next selected Sheet rows 
    
  } catch (err) {
    
    oCommon.DisplayMessage = ' Error 274 - Delete Rows processing error: ' + err.message;
    oCommon.ReturnMessage = func + Step + ' Sheet Row: ' + SheetRow + ', ' + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(oCommon.ReturnMessage);
    // Enable the form to accept responses and return
    form.setAcceptingResponses(true);
    form.setCustomClosedFormMessage(defaultClosedFor);
    return false;
  }
  
  /******************************************************************************/
  Step = 6000; // Enable the form to accept responses and return
  /******************************************************************************/
  // Source: https://stackoverflow.com/questions/43421593/reject-google-forms-submission-or-remove-response
  if (bDeleteGoogleForms){
    form.setAcceptingResponses(true);
    form.setCustomClosedFormMessage(defaultClosedFor);
  }
 
  if (Errors > 0){
    Step = 6100; // Errors encountered
    oCommon.DisplayMessage = ' Delete Rows Process completed with errors: RowsExamined: ' 
      + RowsExamined + ', RowsDeleted: ' + RowsDeleted + ' ResponsesDeleted: ' + ResponsesDeleted
      + ', Errors: ' + Errors;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(oCommon.ReturnMessage);
    return false;
  } else {
    Step = 6100; // No Errors encountered
    oCommon.DisplayMessage = ' Delete Rows Process completed successfully: RowsExamined: ' 
      + RowsExamined + ', RowsDeleted: ' + RowsDeleted + ' ResponsesDeleted: ' + ResponsesDeleted
      + ', Errors: ' + Errors;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(oCommon.ReturnMessage);
    return true;
  }
} 


function SelectSheetRows(oCommon,mode,TargetSheetName,firstDataRow){
  /* ****************************************************************************

   DESCRIPTION:
     This function is invoked whenever another function needs to process ALL or SELECTED 
     contiguous rows in the target "dataTab" (passed by Globals) found in the active Google Sheet.
     
   RETURNS Object oCommon.SelectedSheetRows
   
   USEAGE:
   
     bReturnBool = SelectSheetRows(oCommon,mode,TargetSheetName,firstDataRow)

   REVISION DATE:
    12-14-2017 - First Instance
    12-27-2017 - Change all row/col references to Base 0

   NOTES:
     1. The array oCommon.SelectedSheetRows[SheetRow][TagValue] is returned identifying the selected rows
     2. oCommon.SelectedSheetRows.length == 0 if an invalid range is selected by the user
   
  ******************************************************************************/
  var func = "***SelectSheetRows " + Version + " - ";
  var Step = 1000;
  Logger.log(func + Step + ' BEGIN <mode=' + mode + '>');
  
  if (!ParamCheck(TargetSheetName)){
    // ERROR - No input TargetSheetName parameter found
    oCommon.DisplayMessage = ' Error 180 - No target Sheet name passed to the function.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
  
  if (!ParamCheck(firstDataRow) || firstDataRow < 0){
    // WARNING - No input firstDataRow parameter found
    var EventMsg = func + Step + ' WARNING 181 - No input firstDataRow parameter found';
    Logger.log(EventMsg);
    LogEvent(EventMsg, oCommon);
    // Assign value from globals
    firstDataRow = oCommon.firstDataRow;
    if (!ParamCheck(firstDataRow) || firstDataRow <= 0){
      // WARNING - firstDataRow parameter set to row 1 (Base 0)
      var EventMsg = func + Step + ' WARNING 182 - firstDataRow parameter set to row 1 (Base 0)';
      Logger.log(EventMsg);
      LogEvent(EventMsg, oCommon);
      firstDataRow = 1;
    }
  }
  
  if (!ParamCheck(mode)){
    mode = 1;
  }
  
  if(!oCommon.bSilentMode){
    // Set the bNoTag boolean since it is likly that a TagValue col has not been specified
    var bNoTag = true;
  } else { var bNoTag = false; }
  
  /******************************************************************************/
  Step = 1000; // Get the Form data (in the Google Sheet) as determined by the "mode"
  /******************************************************************************/
  var SheetData = oCommon.Sheets.getSheetByName(TargetSheetName).getDataRange().getValues();
  var LastSheetRow = SheetData.length; // returns Base 1
  var TagValue = '';
  
  if (!firstDataRow || firstDataRow < 1){firstDataRow = 1;} // Base 1

  if (mode == 1){
    Step = 1100; // Retreive the SheetData object array from all rows
    var SheetRow = firstDataRow; //Base 1
    var ArrayStartRow = firstDataRow; //Base 0
    var ArrayEndRow = LastSheetRow - 1; // change to Base 0
    Logger.log(func + Step + ' Get ALL Rows - startRow: ' + SheetRow + ' endRow: ' + LastSheetRow);
    
  } else {
    Step = 1200; // Retreive the SheetData object array from  the selected row(s) to be updated
    var ActiveRange = oCommon.Sheets.getActiveRange();
    var SheetRow = ActiveRange.getRow();
    var NumRows = ActiveRange.getNumRows();
    var SheetEndRow = SheetRow + NumRows - 1;
    var ArrayStartRow = SheetRow - 1; // beacuse the SheetData array is base 0 and is a subset of the entire sheet
    var ArrayEndRow = ArrayStartRow + NumRows - 1; //; // because the SheetData array is base 0 and is a subset of the entire sheet
    
    Logger.log(func + Step + ' Get SELECTED Rows - SheetStartRow: ' + SheetRow + ' SheetEndRow: ' + SheetEndRow);

    if (SheetRow < firstDataRow || SheetEndRow > LastSheetRow ) {
      Browser.msgBox("Oops!, You must select one or more rows between row " + firstDataRow.toString().split(".")[0] + " and row " + LastSheetRow + ".");
      return false; // result will be oCommon.SelectedSheetRows.length == 0
    }
  }
  
  /******************************************************************************/
  Step = 2000; // Select Sheet rows 
  /******************************************************************************/
  Logger.log(func + Step + ' Get Rows - ArrayStartRow: ' + ArrayStartRow + ' ArrayEndRow: ' + ArrayEndRow);
  for (var r = ArrayStartRow; r <= ArrayEndRow; r++) {
    Logger.log(func + Step + ' r: ' + r);
    Logger.log(func + Step + ' TagValueCol: ' + oCommon.RecordTagValueCol + ' MaxLength: ' + oCommon.RecordTagMaxLength
           + ', TagValue: ' + SheetData[r][oCommon.RecordTagValueCol]);
    
    Step = 2100; // Get and assign the record's TagValue // && TargetSheetName == oCommon.FormResponse_SheetName
    if (ParamCheck(oCommon.RecordTagValueCol)){
      var TagValue = FormatTagValue(oCommon, SheetData[r][oCommon.RecordTagValueCol]);
    } else {
      var TagValue = '';
    }
      
    //Step = 2150; // Place selected row into the oCommon.SelectedSheetRows array
    oCommon.SelectedSheetRows.push([[r+1],[TagValue]]); // capture the Sheet row number and report URL address
    Logger.log(func + Step + ' Selected (' + TagValue + ') SheetRow: ' + SheetRow);
    SheetRow++;
    
  } // Get the next row

  return true;
}


function WriteDocument(oCommon, ReportTitle){
  /* ****************************************************************************

   DESCRIPTION:
     This procedure is used to generate a Google Doc or PDF Doc from the submitted 
       Form Responses using the input Google Doc template for the given report title.
       
     The parameters for the input ReportTitle are normally defined in a Setup Tab

   USEAGE:
     ReturnBool = GenerateReport(oCommon, ReportTitle); 

   REVISION DATE:
     12-27-2017 - Change all row/col references to Base 0
     01-22-2018 - Modified to accepet either Template ID or URL as input value
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback
     03-29-2018 - Added Step = 4400 code to delete the "Newly Created doc" copied from the template
     02-20-2019 - Modified to creat the new doc/pdf in the "Target" folder rather than create it 
                    first and then move it (to solve unidentified duplicate files being created)

   NOTES:
     1. The passed oCommon object contains the generally used parameters
     2. The passed SheetData[row][col] object contains the Form Response Google Sheet values from the selected row(s)

  ******************************************************************************/  
  var func = "****WriteDocument " + Version + " - ";
  var Step = 1000;
  var bNoErrors = true;
  var message ='';

  /******************************************************************************/
  Step = 1000; // Validate input parameters and assign working variables
  /******************************************************************************/
  Step = 1100; // Determine need for User to select rows
  if (oCommon.SelectedSheetRows.length <= 0){
    Step = 1110; // SelectedSheetRows is EMPTY or NULL
    var bUserSelectsRows = true;
    Logger.log(func + Step + ' bUserSelectsRows; ' + bUserSelectsRows
               + ' (oCommon.SelectedSheetRows is EMPTY or NULL)');
  } else {
    Step = 1120; // SelectedSheetRows are present
    var bUserSelectsRows = false;
    Logger.log(func + Step + ' bUserSelectsRows; ' + bUserSelectsRows
               + ' (oCommon.SelectedSheetRows are present)');
  }
   
  Step = 1200; // Leave immediately if no ReportTitle is passed
  if (!ParamCheck(ReportTitle)){
    oCommon.DisplayMessage = ' Error 380 - ReportTitle is EMPTY or NULL.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
  
  /******************************************************************************/
  Step = 1300; // Get the Document Parameters from the Setup Tab
  /******************************************************************************/
  var oDocParams = {};
  var ParameterSet = oCommon.DocumentsSection_Heading;  
  oDocParams = ParamSetBuilder(oCommon, ParameterSet);
  if(!ParamCheck(oDocParams)){
    Step = 1310; // Error - No Document Parameters found
    var EventMsg = func + Step + ' (' + oCommon.ReturnMessage + ')';
    Logger.log(EventMsg);
    return false;
  }
  
  /******************************************************************************/
  Step = 1400; // Assign Setup Documents section values
  /******************************************************************************/
  if (!ParamCheck(oDocParams.TitleKeyAry[ReportTitle])){
    Step = 1410; // ReportTitle does not exist
    oCommon.DisplayMessage = ' Error 382 - Report Title (' + ReportTitle 
      + ') not found in Setup Section (' + ParameterSet + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage + ' SilentMode: ' + oCommon.bSilentMode);
    if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
    return false;
  }
  
  if(oDocParams.HeadingKeyAry['Template ID']){
    var templateurl = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Template ID']];
  } else if(oDocParams.HeadingKeyAry['Template URL']){
    var templateurl = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Template URL']];    
  } else {
    Step = 1420; // ERROR - Specified Document Template not found
    oCommon.DisplayMessage = ' Error 383 - No Document Template found for ' 
     + 'Report Title (' + ReportTitle + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
    return false;
  }

  Step = 1430; // Extract the Google ID from the URL, if necessary
  var templateid = getIdFrom(templateurl);
  Logger.log(func + Step + ' DocTemplateName: "' + ReportTitle + '"'
     + ', EmailTemplateID: ' + templateid);
  
  if (oCommon.CallingMenuItem != null){
    var title = oCommon.CallingMenuItem;
  } else {
    var title = "Preparing Documents";
  }
  var timeoutSeconds = 3;
  
  if(oDocParams.HeadingKeyAry['Default Folder ID']){
    var defaultFolderID = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Default Folder ID']];
  } else if(oDocParams.HeadingKeyAry['Default Folder URL']){   
    var defaultFolderID = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Default Folder URL']];
  }
  
  Step = 1440; // Extract additional parameters, as specified
  defaultFolderID = getIdFrom(defaultFolderID);

  var rptPrefix = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Name Prefix']]; 
  var rptSuffix = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Name Suffix']]; 
  var asofDateterm = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['As of Date Heading']]; 
  
  var summaryRptURLTerm = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['URL Col Heading']]; 
  var imgMaxWidth = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Max Width']];
  var imgMaxHeight = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Max Height']];
  var viewButtonHeading = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Button Heading']];
  var viewButtonText = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Button Text']]; // text to display in the "viewButtonCol" column of the Sheet
  var photo_placeholder = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Photo Place Holder']]; 
  var photo_missing_text = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['No PhotoText']]; 
  var photoURLHeading = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Photo URL Heading']];
  //var photoURLcol = +oCommon.Globals[photoURLHeading];
  var bSaveAsPDF = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Save As PDF']]; // Boolean value
  var asofDateterm =  oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['As of Date Heading']];
  
  Logger.log(func + Step + ' templateid: ' + templateid + ', defaultFolderID: ' + defaultFolderID + ' SaveAsPDF: ' + bSaveAsPDF);
  
  /****************************************************************************/
  Step = 1500; // Assign the oCommon parameters
  /****************************************************************************/
  var oSourceSheets = oCommon.Sheets;
  var ScriptManager = oCommon.ScriptManager;
  var error_email = oCommon.Send_error_report_to;
  var validDomain = oCommon.ValidDomain;
  var ActiveStatusValue = oCommon.Active_Status_value;
  var InActiveStatusValue = oCommon.Inactive_Status_value; 

  Step = 1510; // Open the "active" sheet and get the row info
  if (oCommon.bSilentMode){
    Step = 1510; // Use the form response tab
    var dataTab = oCommon.FormResponse_SheetName;
    var sheet = oSourceSheets.getSheetByName(dataTab);
    var SheetData = sheet.getDataRange().getValues();
  } else {
    Step = 1520; // Use the "Active" Tab
    var sheet = oSourceSheets.getActiveSheet();
    var dataTab = oSourceSheets.getActiveSheet().getName();
    var SheetData = oSourceSheets.getActiveSheet().getDataRange().getValues();
  }
  
  Step = 1530; // Verify that the correct XrefArray was built in buildCommon()
  if (dataTab !== oCommon.XrefSourceSheet) {
    // Someting went terribly wrong...
    Step = 1535; // ERROR - Active Sheet Name does not match
    oCommon.DisplayMessage = ' ERROR 384 - Xref Build error (Active Sheet: "' 
      + dataTab + '", Expected Sheet: "' +  oCommon.XrefSourceSheet + '").';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
    return false;
  }
  
  /******************************************************************************/
  Step = 1600; // Prepare the "Terms" Key/Value Array: Key = Term, Value = Col # for the Term value
  /******************************************************************************/
  var ReplaceTermsAry = {};
  ReplaceTermsAry = oCommon.TermsXrefAry; // The work has already been done when buildCommon() was executed
  
  Step = 1610; // Assign additional term-dependent variables
  var TagValueCol = ReplaceTermsAry[oCommon.Globals['Record_Tag_Value_heading']]
  var timestampCol = ReplaceTermsAry[oCommon.Globals['Timestamp_heading']];
  var StatusCol = ReplaceTermsAry[oCommon.Globals['Record_status_heading']];
  var viewButtonCol = ReplaceTermsAry[viewButtonHeading];
  var asofDateCol = ReplaceTermsAry[asofDateterm];
  var photoURLcol = ReplaceTermsAry[photoURLHeading];
  var rptDateCol = asofDateCol;
  // Search for the viewURL column
  for (var key in ReplaceTermsAry){
    if (key.indexOf(summaryRptURLTerm) > -1){
      var summaryRptURLCol = ReplaceTermsAry[key];
      break;
    }
  }
  Logger.log(func + Step + ' TagValueCol: ' + TagValueCol);
  Logger.log(func + Step + ' timestampCol: ' + timestampCol);
  Logger.log(func + Step + ' StatusCol: ' + StatusCol); 
  Logger.log(func + Step + ' viewButtonCol: ' + viewButtonCol);
  Logger.log(func + Step + ' asofDateCol: ' + asofDateCol);
  Logger.log(func + Step + ' photoURLcol: ' + photoURLcol);
  Logger.log(func + Step + ' rptDateCol: ' + rptDateCol);
  Logger.log(func + Step + ' summaryRptURLCol: ' + summaryRptURLCol);

  /**********************************************************************************/
  Step = 1700; /* Capture the Document Storage File Folder Parameters
  
    When using Alternate Document Storage Folders: UseAltFolders = TRUE
  
    SheetContainingFolderIDs - Tab containing the headings similar to the following
      
      Alt_folder_key_Heading - The Column heading for the Key Value in the "SheetContainingFolderIDs" Tab
      Alt_folder_id_Heading  - The Column heading for the column containing the Google FolderID value in the "SheetContainingFolderIDs" Tab
      
      FormData_AltFolder_Key_Heading - The heading in the Form Response Tab that contains the "Key" values
      
  **********************************************************************************/

  var bUseAltFolders = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Use Alternate Folders']];
  Logger.log(func + Step + ' bUseAltFolders: ' + bUseAltFolders + ', TagValueCol: ' + TagValueCol);
  if (bUseAltFolders) {
    var SheetContainingFolderIDs = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Tab w Alt Folders']];
    var Alt_folder_key_Heading = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Alt Folder Key Heading']];
    if(oDocParams.HeadingKeyAry['Alt Folder ID Heading']){
      var Alt_folder_id_Heading = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Alt Folder ID Heading']];
    } else if(oDocParams.HeadingKeyAry['Alt Folder URL Heading']){
      var Alt_folder_id_Heading = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Alt Folder URL Heading']];
    } else {
      Step = 1710; // WARNING 386 - Alternate Folder specified, but found
      oCommon.DisplayMessage = ' Error 386 - Alternate Folder specified, but found.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      Logger.log(oCommon.ReturnMessage);
      if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
      bUseAltFolders = false;
    }
  }
  
  /**********************************************************************************/
  Step = 1800; // Build the Document Storage File Folder Parameters DocFolderKeyValueAry
  /**********************************************************************************/
  Logger.log(func + Step + ' bUseAltFolders: ' + bUseAltFolders);
  if (bUseAltFolders) {

    var FormData_AltFolder_Key_Heading = oDocParams.Parameters[oDocParams.TitleKeyAry[ReportTitle]][oDocParams.HeadingKeyAry['Form_Response_Term']];
    var FormData_AltFolder_Key_Col = GetXrefValue(oCommon, FormData_AltFolder_Key_Heading);
    
    //Logger.log(func + Step + ' bUseAltFolders: ' + bUseAltFolders);
    //Logger.log(func + Step + ' SheetContainingFolderIDs: ' + SheetContainingFolderIDs);
    //Logger.log(func + Step + ' Alt_folder_key_Heading: ' + Alt_folder_key_Heading); 
    //Logger.log(func + Step + ' Alt_folder_id_Heading: ' + Alt_folder_id_Heading);
    //Logger.log(func + Step + ' oDocParams.HeadingKeyAry[Form_Response_Term]: ' + oDocParams.HeadingKeyAry['Form_Response_Term']);
    //Logger.log(func + Step + ' FormData_AltFolder_Key_Heading: ' + FormData_AltFolder_Key_Heading);
    //Logger.log(func + Step + ' FormData_AltFolder_Key_Col: ' + FormData_AltFolder_Key_Col);
  
    var TargetSheet = SheetContainingFolderIDs;    // Sheet Containing Keys and Values
    var SectionTitle = ''; // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start r,ow is always the following row
    var KeyRow = '';   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
    var ValueRow = ''; // number or text [Look for text Match in Col 0 (Base 0)]
    var KeyCol = Alt_folder_key_Heading;   // number or text [Look for text Match in StartRow (Base 0)]
    var ValueCol = Alt_folder_id_Heading; // number or text [Look for text Match in StartRow (Base 0)]
    var DocFolderKeyValueAry = {}; // if declared before entering function
    
    //DocFolderKeyValueAry = BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol,DocFolderKeyValueAry)
    DocFolderKeyValueAry = BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol);
    // Verify
    //for(var key in DocFolderKeyValueAry){
    //  Logger.log(func + Step + ' Key: ' + key + '  Value: ' + DocFolderKeyValueAry[key]);
    //}
  }
  
  /**************************************************************************/
  Step = 1900; // Get User Selected Records, if necessary (Ref Step 1200, above)
  /**************************************************************************/
  Logger.log(func + Step + ' bUserSelectsRows: ' + bUserSelectsRows);
  if (bUserSelectsRows){
    Step = 1910; // User needs to select rows
    var mode = 2;
    SelectSheetRows(oCommon, mode, dataTab, firstDataRow);
    if(!ParamCheck(oCommon.SelectedSheetRows) || oCommon.SelectedSheetRows.length <= 0){
      // Warning - No Selected Rows found
      var return_message = func + Step + ' ********** Warning - No Rows Selected';
      //LogEvent(return_message, oCommon);
      //oCommon.ReturnMessage = return_message;
      return false;
    }
    
    Step = 1920; // Take time to confirm that all documents are to be prepared
    var MsgBoxMessage = ('The following rows were selected: ' + FormatForDisplay(oCommon.SelectedSheetRows)
       + '\\n' + ' Do you want to prepare document(s)?');
    if (Browser.msgBox('Prepare Documents?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      var display_message = "Your response is NO - no documents will be prepared.";
      var return_message = func + Step + ' User terminated after rows were selected.';
      oCommon.ReturnMessage = return_message;
      oCommon.DisplayMessage = display_message;
      return false;
    }
  } 
  
  /****************************************************************************/
  Step = 2000; // Process each row passsed in oCommon.SelectedSheetRows
  /****************************************************************************/
  var firstDataRow = GetSheetParam(dataTab, '1st Data Row', oCommon);
  //var timestampCol = oCommon.TimestampCol;
  //var StatusCol = oCommon.Record_statusCol;
  //var ActiveStatusValue = oCommon.Active_Status_value;
  //var InActiveStatusValue = oCommon.Inactive_Status_value; 
  //var rptDateCol = oCommon.AsOfDateCol;
  //var rptDateCol = asofDateCol;
  var number_of_rows = oCommon.SelectedSheetRows.length;
  var documents_completed = 0;
  Logger.log(func + Step + ' dataTab: ' + dataTab + ', Number of rows to Process: ' + number_of_rows); 
  
  for (var k = 0; k < number_of_rows; k++){
    
    var ArrayRow = oCommon.SelectedSheetRows[k][0] - 1; // Convert to Base 0
    
    Step = 2100; // Test for bypassing non-active selected record    
    var bProcessRow = true;
    if (dataTab == oCommon.FormResponse_SheetName){
      var TimeStamp = SheetData[ArrayRow][timestampCol].toString();
      var StatusValue = SheetData[ArrayRow][StatusCol];
      bProcessRow = CheckStatus(TimeStamp, StatusValue, oCommon);
      Logger.log(func + Step + ' k= ' + k + ' Array Row: ' + ArrayRow + ', TimeStamp: ' + TimeStamp + ', StatusValue: ' + StatusValue);
      Logger.log(func + Step + ' bProcessRow Value: ' + bProcessRow);
    } 
    Logger.log(func + Step); // Blank row for readability
    
    if (bProcessRow) {
      Step = 2200; // Process the selected record    
      var TagValue = FormatTagValue(oCommon,SheetData[ArrayRow][TagValueCol]);
      message = ' Preparing "' + ReportTitle + '" Report for ' + TagValue;
      Logger.log(func + Step + message);
      progressMsg(message,title,-5);
    
      Step = 2210; // Get the previously stored Summary Report URL value
      if (ParamCheck(summaryRptURLCol)){ var reportURL = SheetData[ArrayRow][summaryRptURLCol]; }
        else { var reportURL = ''; }
      Logger.log(func + Step + ' k= ' + k + ' Array Row: ' + ArrayRow + ' Previous Doc URL: ' + reportURL);
        
      Step = 2220; // Test to determine if a previous report has been prepared
      if(ParamCheck(reportURL) && reportURL != ''){
        // Summary Report URL present, delete previously prepared report
        if (deleteGoogleDoc(reportURL, oCommon)) {
          Step = 2221; // Log Successful delete
          Logger.log(func + Step + ' Deleted Doc URL: ' + reportURL);
        } else {
          Step = 2222; // Log Unsuccessful delete
          oCommon.DisplayMessage = ' Error 388 -' + oCommon.DisplayMessage + ', ViewURL: ' + reportURL;
          oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
          Logger.log(oCommon.ReturnMessage);
          if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
        }        
      }   
        
      /****************************************************************************/
      Step = 2300; // Determine the TargetFolderID that is to be the ultimate home for the doc/pdf to be created
      /****************************************************************************/
      Step = 2310; // First, replace any non-word characters that may affect the file name (docName) that is to be used
      var patt1 = /\W/g;
      //var TagValue = SheetData[ArrayRow][TagValueCol].replace(patt1," ");
      TagValue.replace(patt1," ");
      TagValue = FormatTagValue(oCommon, TagValue);
      var docName = rptPrefix + TagValue + rptSuffix;
        
      Step = 2320; // Determine the TargetFolderID value to use
      var TargetFolderID = null;
      if (ParamCheck(defaultFolderID)){ TargetFolderID = defaultFolderID; }
      
      if (bUseAltFolders) {
        Step = 2322; // Lookup and test for a valid Folder ID value based on the value of the Key
        var Key = SheetData[ArrayRow][FormData_AltFolder_Key_Col];
        var FolderID = DocFolderKeyValueAry[Key];
        Step = 2324; // If the returned "FolderID" (see Step,1700, above) assign as the TargetFolderID
        if (ParamCheck(FolderID)){ TargetFolderID = getIdFrom(FolderID); } 
        Logger.log(func + Step + ' Doc Folder Key: ' + Key + ' FolderID: ' 
                   + FolderID + ' Form Data Col: ' + FormData_AltFolder_Key_Col);
      } else {  
        Step = 2326; // Use the defaultFolderID, if it exists. Otherwise leave it "null"
        if (ParamCheck(defaultFolderID)){ TargetFolderID = defaultFolderID; }
      }
      
      Logger.log(func + Step + ' TargetFolderID: ' + TargetFolderID + ', Template Url: ' + templateurl);

      /****************************************************************************/
      Step = 2400; // Prepare the copy of the Template Google Doc
      /****************************************************************************/
      try {
        Step = 2410; // Determine where, and create a copy of the Google Doc template (docID)
        if (bSaveAsPDF){
          Step = 2330; // Do the work in the root drive, then move the pdf file
          var docID = DriveApp.getFileById(templateid).makeCopy(docName).getId();
        } else if (TargetFolderID != null){
          Step = 2412; // Do the work in the TargetFolder
          var TargetFolder = DriveApp.getFolderById(TargetFolderID);
          var docID = DriveApp.getFileById(templateid).makeCopy(docName, TargetFolder).getId();
        } else {
          Step = 2414; // Create a copy in the current (root?) folder
          var docID = DriveApp.getFileById(templateid).makeCopy(docName).getId();
        }
          
        Step = 2420; // Open the newly created Google Doc template and prepare for inserting the Form Data
        var doc = DocumentApp.openById(docID);
        var body = doc.getActiveSection();
        var header = doc.getHeader();
        
        Logger.log(func + Step + ' k= ' + k + ' Doc Name: ' + docName + ' New Template ID: ' + docID);  
        
      } catch(err) {
        // If errors are encountered, display or log the Event and exit
        Logger.log(func + Step + ' k= ' + k + ' New Template ID: ' + docID);  
        oCommon.DisplayMessage = ' Error 390 - Unable to Build Google Doc (' + TagValue + ') Error Message: ' 
           + err.message;
        oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
        Logger.log(oCommon.ReturnMessage);
        if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
        return false;
      }

      /****************************************************************************/
      Step = 2500; // Create the SearchTerm / Value relationships using the data in the SheetData array
      /****************************************************************************/
      Logger.log(func + Step + ' Creating the SearchTerm / Value relationships array...');
      for (var key in ReplaceTermsAry) {
       
        var value = SheetData[ArrayRow][Number(ReplaceTermsAry[key])];
        Logger.log(func + Step + ' key: ' + key + ' Value: ' + value);

        if (key.indexOf('Address_2') > -1) {
          /****************************************************************************/
          Step = 2510; // Add a comma and space to a 2nd Address Line value, if it is present
          value = value + ', ';
          //Logger.log(func + Step + ' key: ' + key + ' Value: ' + AddressValue);
        } else 
          if (isValidDate(value)){
              /****************************************************************************/
              Step = 2520; //  Reformat time values to hh:mm ampm, if necessary isValidDate(d)
              value = formatDate(value); 
              Logger.log(func + Step + ' key: ' + key + ' Value: ' + value);
          } else {
            /****************************************************************************/
            Step = 2530; // Default action for Keys not handled above
            var value = Trim(value);
          } 
          
          body.replaceText(key, value);  // this is the code line thatdoes all of the work!
          
          /****************************************************************************/
          Step = 2540; // Update the document header and set permissions
          if (header){
            header.replaceText(key, value);  // this is the code line thatdoes all of the work!
          }  
      }  // end of for (var key in ReplaceTermsAry) 
  
      /****************************************************************************/
      Step = 2550; // Insert photo, if available
      var bPhotoIsAvailable = false;
      
      if (ParamCheck(photoURLcol)){
        var url = Trim(SheetData[ArrayRow][photoURLcol]);
        if (isValidURL(url)){
          bPhotoIsAvailable = true;
        }
      }
      //Logger.log(func + Step + ' photoURLcol: ' + photoURLcol + ' url: ' + url);
                                                                                   
      if (bPhotoIsAvailable){
        // See: https://stackoverflow.com/questions/18848144/inserting-a-google-drive-image-into-a-document-using-google-apps-script
        //      https://gsuite-developers.googleblog.com/2013/06/introducing-google-docs-cursorselection.html
        //      https://stackoverflow.com/questions/26248384/google-apps-script-how-to-horizontally-align-an-inlineimage
        //      https://stackoverflow.com/questions/24932514/google-script-insert-image-in-a-google-doc-table

        try {        
          var fileID = getIdFrom(url);
          var image = DriveApp.getFileById(fileID).getBlob();
          var tables = body.getTables();
          /****************************************************************************/
          Step = 2600; // Search for the image placeholder within the document tables
          for (var t in tables) {
            var table = tables[t];
            var tablerows=table.getNumRows();
            for ( var row = 0; row < tablerows; ++row ) {
              var tablerow = table.getRow(row);
              for ( var cell=0; cell < tablerow.getNumCells(); ++cell) {
                var celltext = tablerow.getChild(cell).getText();
                if(celltext == photo_placeholder) {
                  /****************************************************************************/
                  Step = 2610; // Placeholder found - insert image file
                  table.replaceText(photo_placeholder, '');
                  var inlineI = table.getCell(row, cell).insertImage(0, image);
                  var width = inlineI.getWidth();
                  var newW = width;
                  var height = inlineI.getHeight();
                  var newH = height;
                  var ratio = width/height;
                  //Logger.log(func + Step + ' w= ' + width + ' h= ' + height + ' ratio= ' + ratio);
                  /****************************************************************************/
                  // re-size the image file, if necessary
                  if(width > imgMaxWidth){ 
                    //max width of image
                    newW = imgMaxWidth; 
                    newH = parseInt(newW/ratio);
                  }
                  if(newH > imgMaxHeight){ 
                    //max height of image
                    newH = imgMaxHeight; 
                    newW = parseInt(newH*ratio);
                  }
                  /****************************************************************************/
                  Step = 2620; // Position image in the center of the table cell
                  //Logger.log(func + Step + ' Nw= ' + newW + ' Nh= ' + newH);
                  inlineI.setWidth(newW).setHeight(newH); //resizes the image but also needs to center it
                  var styles = {};
                  styles[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
                  inlineI.getParent().setAttributes(styles);
                }
              }
            }
          }
        } catch (err_p) {
          oCommon.DisplayMessage = ' Error 391 - Unable to insert photo (' 
          + TagValue + ') Error Message: ' + err_p.message;
          oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
          Logger.log(oCommon.ReturnMessage);
          LogEvent(oCommon.ReturnMessage, oCommon);
          var bNoErrors = false;
          Step = 2630; // Replace image placeholder with Photo Not Available text
          //Logger.log(func + Step + ' Nwphoto_placeholder= ' + photo_placeholder + ' photo_missing_text= ' + photo_missing_text);
          if (photo_placeholder){
            body.replaceText(photo_placeholder, photo_missing_text);
          }
        }
      } else {
        /****************************************************************************/       
        Step = 2700; // Replace image placeholder with Photo Not Available text
        //Logger.log(func + Step + ' Nwphoto_placeholder= ' + photo_placeholder + ' photo_missing_text= ' + photo_missing_text);
        if (photo_placeholder){
          body.replaceText(photo_placeholder, photo_missing_text);
        }
      }  // end of Insert photo, if available
  
      /****************************************************************************/
      Step = 2800; // Update the "as of report date" field and the Spreadsheet row number
      /****************************************************************************/
      //Logger.log(func + Step + ' Got here...formatYear(TimeStamp) : ' + formatYear(TimeStamp));
      if(asofDateterm){body.replaceText(asofDateterm, (formatDate(new Date()) + ' ' + formatTime(new Date()))); } 
      body.replaceText('<<Row>>', ArrayRow);
      //if(photo_placeholder){body.replaceText(photo_placeholder, photo_missing_text);} 
      body.replaceText('<<Year>>',formatYear(new Date())); 
      body.replaceText('<<Report_Date>>',formatDateTime((new Date()))); 
  
      //doc.addViewer(workEmail);
      //var workEmail = Trim(SheetData[ArrayRow][workEmailcol]);
      //Logger.log(func + Step + ' workEmail: ' + workEmail + ' Valid domain: ' + validDomain);
      //if (workEmail.indexOf(validDomain) > 0) {
      //  Step = 2810; 
      //  doc.addViewer(workEmail);
      //  Logger.log(func + Step + ' view permission set for: ' + workEmail);
      //}
      
      /****************************************************************************/
      Step = 3000; // Disposition and Save the Google document
      /****************************************************************************/
      var SuccessfulSave = false;
      try {
        /****************************************************************************/
        Step = 3100; // Save the now populated Google Doc
        /****************************************************************************/
        // Evaluate and assign a TargetFolderID value if it has not already been determined
        doc.saveAndClose();
        var newDocURL = doc.getUrl();          
        //Logger.log(func + Step + ' Doc just saved: ' + newDocURL);
        
        Step = 3200 // Determine need for ownership change
        if (!oCommon.Using_TeamDrive){
          newDocURL = changeOwner(newDocURL, oCommon);
          Logger.log(func + Step + ' After ownership change - Url: ' + newDocURL);
        }
        SuccessfulSave = true;
        
      }
      
      catch(err) {
        oCommon.DisplayMessage = ' Error 392 - Unable to Save Google Doc (' 
        + TagValue + ') Error Message: ' + err.message;
        oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
        Logger.log(oCommon.ReturnMessage);
        if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
        var bNoErrors = false;
        
      }
      
      /****************************************************************************/
      Step = 4000; // Generate the PDF file if required and move it to the TargetFolder
      /****************************************************************************/
      Logger.log(func + Step + ' bSaveAsPDF:' + bSaveAsPDF + ', SuccessfulSave: ' + SuccessfulSave);
      if (bSaveAsPDF) {
        try{
          Step = 4100; // Use the above-created newDocURL as the SourceDocURL
          var SourceDocURL = newDocURL;
          
          Logger.log(func + Step + ' Doc just saved - Url: ' + SourceDocURL);
          
          Step = 4200; // Retrieve a Copy of  the Source Google Doc
          var SourceDocID = getIdFrom(SourceDocURL);
          var copyofDoc = DriveApp.getFileById(SourceDocID);
          
          Step = 4300; // Create a new pdf document file in the root directory
          var pdfFile = DriveApp.createFile(copyofDoc.getAs('application/pdf'));
          var pdfFileURL = pdfFile.getUrl();
          var pdfOwnerEmail = pdfFile.getOwner().getEmail();
          
          Logger.log(func + Step + ' Owner: ' + pdfOwnerEmail + ', pdf created - Url: ' + pdfFileURL);
          
          // Delete the source Google Doc file
          Step = 4400;
          DriveApp.getFileById(SourceDocID).setTrashed(true);
          //Logger.log(func + Step + ' Deleted source Doc - ID: ' + SourceDocID);
          newDocURL = pdfFileURL;
          SuccessfulSave = true;
          
          //Move the newly created file to the TargetFolder
          if (TargetFolderID != null){
            var pdfFileID = getIdFrom(pdfFileURL);
            var TargetFolder = DriveApp.getFolderById(TargetFolderID);
            var newDocURL = DriveApp.getFileById(pdfFileID).makeCopy(docName, TargetFolder).getUrl();
            // Delete the source pdf file
            DriveApp.getFileById(pdfFileID).setTrashed(true);
          }
          Logger.log(func + Step + ' newDocURL: ' + newDocURL);
          
        }
        
        catch(err) {
          oCommon.DisplayMessage = ' Error 394 - Unable to Create/Save PDF file (' 
            + TagValue + ') Error Message: ' + err.message;
          oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
          Logger.log(oCommon.ReturnMessage);
          if (oCommon.bSilentMode){ LogEvent(oCommon.ReturnMessage, oCommon); }
          var bNoErrors = false;
        }
      }

      //**************************************************************************
      Step = 5000; // save the new report URL to the Google sheet
      //**************************************************************************
      //Logger.log(func + Step + ' RptURLTerm: ' + summaryRptURLTerm + ' RptURLCol: ' + summaryRptURLCol);
      if (!isNaN(summaryRptURLCol)){
        //var URLaddr = summaryRptURLCol + String(ArrayRow+1).split(".")[0];
        var URLaddr = numToA(summaryRptURLCol+1) + String(ArrayRow+1).split(".")[0]; 
        Logger.log(func + Step + ' URLaddr: ' + URLaddr + ' View Url: ' + newDocURL);
        sheet.getRange(URLaddr).setValue(newDocURL);
      }
        
      //**************************************************************************
      Step = 5100; // save the report "as of date" to the sheet if previous Timestamp date does not exist
      //**************************************************************************
Logger.log(func + Step + ' asofDateTerm: ' + asofDateterm + ' asofDateCol: ' + asofDateCol 
           + ', timestampCol: ' + timestampCol);
      if (!isNaN(asofDateCol) && !isNaN(timestampCol)){
        var rptDateAddr = numToA(rptDateCol+1) + String(ArrayRow+1).split(".")[0]; // A1 notation is always Base 1
        var timestampValue = SheetData[ArrayRow][timestampCol];
        if (!timestampValue){
          timestampValue = new Date();
          timestampValue = formatDate(timestampValue);
          Logger.log(func + Step + ' rptDate: ' + timestampValue + ' rptDateAddr: ' + rptDateAddr +' Added reportURL: ' + newDocURL);
          sheet.getRange(rptDateAddr).setValue(timestampValue) ;
          Logger.log(func + Step + ' SummaryData Report Written for: ' + TagValue + ' Sheet Row: ' + (ArrayRow+1));     
        }  
      }
      /**************************************************************************/
      Step = 5200; // Format the "viewButtonCol" value and write to the sheet
      /**************************************************************************/ 
      //Logger.log(func + Step + ' ViewURLTerm: ' + viewButtonHeading + ' ViewURCol: ' + viewButtonCol);
      if (!isNaN(viewButtonCol)){
        var ViewButtonAddr = numToA(viewButtonCol+1) + String(ArrayRow+1).split(".")[0];
        Logger.log(func + Step + ' Writing VIEW Text - SheetRow: ' + (ArrayRow+1) +  ' CellAddress: ' + ViewButtonAddr + ' Value: ' + viewButtonText);
        sheet.getRange(ViewButtonAddr).setValue(viewButtonText);  
      }
      var event_message = func + Step + ' Report "' + ReportTitle + '" prepared for: ' + TagValue + ' Sheet Row: ' + (ArrayRow+1);
      LogEvent(event_message, oCommon);
      Logger.log(event_message);     

      documents_completed++;
      SpreadsheetApp.flush();
      
    } // end of bypass if Timestamp is empty
    
  } // end of test each oCommon.SelectedSheetRows row
  
  /**************************************************************************/
  Step = 7000; // Report metrics and go home...
  /**************************************************************************/
  // Update Library counters
  var scriptProperties = PropertiesService.getScriptProperties();
  var docs_count = Number(scriptProperties.getProperty('B_Docs_count'));
  scriptProperties.setProperty('B_Docs_count', docs_count + documents_completed);
  oCommon.DocsGenerated = documents_completed;
  oCommon.ItemsAttempted = number_of_rows;
  oCommon.ItemsCompleted = documents_completed;
  oCommon.DisplayMessage = oCommon.DisplayMessage + ' Document "' + ReportTitle 
    + '" results: Completed ' + documents_completed 
    + ' of ' + number_of_rows + ' attempted.' + '\\n';
  oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
  //LogEvent(oCommon.ReturnMessage, oCommon);
  Logger.log(func + oCommon.ReturnMessage);

  Logger.log(func + Step + ' END...');
  return bNoErrors;
}




function updateFormLists(oCommon){
  /* ****************************************************************************
   DESCRIPTION:
     Create the custom Drop-Down and CheckBox lists from spreadsheet lists
     
   USEAGE:
     ReturnBool = updateFormLists(oCommon); 

   REVISION DATE:
     08-02-2017 - Initial design
     12-27-2017 - Change all row/col references to Base 0
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback
     03-10-2019 - Modified to find the Form's Timestamp column in the sheet that is used by
                    updateUrls() 
     
   NOTES:

      Sources: http://wafflebytes.blogspot.com/2016/10/google-script-create-drop-down-list.html
               https://stackoverflow.com/questions/26393556/how-to-prefill-google-form-checkboxes
               
  ******************************************************************************/
  var func = "***updateFormLists " + Version + " - ";
  var Step = 1000;
  Logger.log(func + Step + ' BEGIN ');

  /******************************************************************************/
  Step = 1000; // Get the form attached to Google Sheet generating the Form Responses
  /******************************************************************************/
  //var formUrl = sheet.getFormUrl();             // Use form attached to sheet
  //var form = FormApp.openByUrl(formUrl);
  var formID = oCommon.GoogleForm_Url;
  var expected_form_name = Trim(oCommon.GoogleForm_Name);
  
  try {
    Step = 1100; 
    formID = getIdFrom(formID);
    Step = 1200; 
    var form = FormApp.openById(formID);
    Step = 1300;  
    var items = form.getItems();
    //Logger.log(func + Step + ' Num Items: ' + items.length);
    var actual_form_name = DriveApp.getFileById(formID).getName();
  } catch (e) {
    oCommon.DisplayMessage = ' Warning 150 - Error encountered trying to open Google Form: ' + e;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + oCommon.ReturnMessage);
    return false;
  }

  Step = 1500; // Test for matching Form names
  if(actual_form_name !== expected_form_name){
    oCommon.DisplayMessage = ' ERROR 152 - Form Title (' + actual_form_name 
      + ') does not match the name in the Setup parameter GoogleForm_Name (' + expected_form_name + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + oCommon.ReturnMessage);
    return false;
  }

  Step = 2000; // Get the the Google Sheet containing the Form Responses & generate the headers array
  try {
    var dataTab = oCommon.FormResponse_SheetName; 
    var itemsTab = oCommon.ListItemsSheetName; 
    var sheet = oCommon.Sheets.getSheetByName(dataTab);
    var SheetData = sheet.getDataRange().getValues();
    var FormHeaders = SheetData[0]; // Form Response Sheet headers == form titles (questions)
    var listsheet = oCommon.Sheets.getSheetByName(itemsTab);
    var ListData = listsheet.getDataRange().getValues();
    var ListHeaders = ListData[0]; // Form Response Sheet headers == form titles (questions)
    //Logger.log(func + Step + ' ListHeaders: ' + ListHeaders);    
    
  } catch(err1) {
    // Expected Listsheet not found
    oCommon.DisplayMessage = ' Error 154 - Error encountered trying to open List Items sheet ("'
      + itemsTab + '"): ' + err1;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + oCommon.ReturnMessage);
    return false;
  }
  
  /******************************************************************************/
  Step = 3000; // For each itemTitle,look for a matching "item" ListHeaders in the Form
  /******************************************************************************/
  var total_items = items.length;
  var items_updated = 0;
  try {
    for (var i = 0; i < items.length; i++) {
      var itemTitle = items[i].getTitle();
      //Logger.log(func + Step + ' Item: ' + i + ', Title: ' + itemTitle);
      for (var col = 0; col < ListHeaders.length; col++) {
        if (itemTitle == Trim(ListHeaders[col])){
          /*****************************************************************************/
          Step = 3010; //  Match Found - fill the listbox
          // grab the row values in the matching column
          var ListEntries = [];  // ListEntries[row]    
          for (var r=1; r < ListData.length; r++) {
            //if(ListData[r][col] != "" && ListData[r][col] != null) {
            if (ParamCheck(ListData[r][col])){
              ListEntries.push(ListData[r][col]); 
            }
          }
          //Logger.log(func + Step + ' Title: ' + itemTitle + ' ListEntries: ' + ListEntries);
          /*****************************************************************************/
          Step = 3020; //  Place the ListEntries into the Form
          var itemID = items[i].getId();
          var type = items[i].getType().toString();
          //Logger.log(func + Step + ' Title: ' + itemTitle + ' ItemType: ' + type);
          
          Step = 3100; // Need to treat every type of answer as its specific type.
          switch (items[i].getType()) {
            case FormApp.ItemType.TEXT:
              Step = 3110; 
              var item = form.getItemById(itemID).asTextItem();
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Title: "' + itemTitle + '" updated.');
              items_updated++;
              break;
            case FormApp.ItemType.PARAGRAPH_TEXT: 
              Step = 3120; 
              //item = items[i].asParagraphTextItem();
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Title: "' + itemTitle + '" updated.');
              items_updated++;
              break;
            case FormApp.ItemType.LIST:
              Step = 3130; 
              var item = form.getItemById(itemID).asListItem();
              item.setChoiceValues(ListEntries); 
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Title: "' + itemTitle + '" updated.');
              items_updated++;
              break;
            case FormApp.ItemType.MULTIPLE_CHOICE:
              Step = 3140; 
              var item = form.getItemById(itemID).asMultipleChoiceItem();
              item.setChoiceValues(ListEntries); 
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Title: "' + itemTitle + '" updated.');
              items_updated++;
              break;
            case FormApp.ItemType.CHECKBOX:
              Step = 3150; 
              var item = form.getItemById(itemID).asCheckboxItem();
              item.setChoiceValues(ListEntries);          
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Title: "' + itemTitle + '" updated.');
              items_updated++;
              break;
            case FormApp.ItemType.DATE:
              Step = 3160; 
              //item = items[i].asDateItem();
              //resp = new Date( resp );
              //resp.setDate(resp.getDate()+1);
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Title: "' + itemTitle + '" updated.');
              items_updated++;
              break;
            case FormApp.ItemType.DATETIME:
              Step = 3170; 
              //item = items[i].asDateTimeItem();
              //resp = new Date( resp );
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Title: "' + itemTitle + '" updated.');
              items_updated++;
              break;
            default:
              Step = 3180; 
              item = null;  // Not handling DURATION, GRID, IMAGE, PAGE_BREAK, SCALE, SECTION_HEADER, TIME
              Logger.log(func + Step + ' Item: ' + i + ', Type: ' + type + ', Valid Form Item Title not found.');
              break;
          }
        }
      }
    }
    
  } catch (e) {
    
    oCommon.DisplayMessage = ' WARNING 156 - Error encountered trying to update Form Items: ' + e;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
    Logger.log(func + Step + oCommon.ReturnMessage);
    return false;
  }
  
  /**************************************************************************/
  Step = 4000; // Report metrics and go home...
  /**************************************************************************/
  oCommon.ItemsAttempted = total_items;
  oCommon.ItemsCompleted = items_updated;
  oCommon.DisplayMessage = oCommon.DisplayMessage + ' Update Items for form "' + actual_form_name 
    + '" results: Completed ' + items_updated 
    + ' of ' + total_items + ' found.' + '\\n';
  oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
  //LogEvent(oCommon.ReturnMessage, oCommon);
  Logger.log(func + Step + oCommon.ReturnMessage);

  Logger.log(func + Step + ' END...');
  
  return true;
  
}


function ParamSetBuilder(oCommon, ParameterSet){
  /* ****************************************************************************

  DESCRIPTION:
     This function is used to build an object containing the Document parameter values
     used by the other functions in this library.
     
  USEAGE:
  
    // For Menu Parameters:
    var oMenuParams = {};
    var ParameterSet = oCommon.BuildMenu_Heading;  
    var oMenuParams = ParamSetBuilder(oCommon, ParameterSet);
  
    // For Document Parameters:
    var oDocParams = {};
    var ParameterSet = oCommon.DocumentsSection_Heading;  
    oDocParams = ParamSetBuilder(oCommon, ParameterSet);
  
    // For Email Parameters:
    var oDocParams = {};
    var ParameterSet = oCommon.Globals['EmailSection_Heading'];  
    oDocParams = ParamSetBuilder(oCommon, ParameterSet);
    
    // For Data Conversions
    var oConversionParams = {};
    var ParameterSet = oCommon.DataConversions_Heading;  
    oConversionParams = ParamSetBuilder(oCommon, ParameterSet);
    
    // For Validation Parameters
    var oValidationParams = {};
    var ParameterSet = oCommon.ValidationSection_1_Heading,;  
    oValidationParams = ParamSetBuilder(oCommon, ParameterSet);   
    
    
   RETURNS Object
   
     var oDocParams = {
           HeadingKeyAry: ColNumKeyAry,
           TitleKeyAry:   RowNumKeyAry,
           Parameters:    AllParameters
         }
     

   REVISION DATE:
    12-14-2017 - First Instance
    12-27-2017 - Change all row/col references to Base 0

   NOTES:
     1. 
   
  ******************************************************************************/
  var func = "****ParamSetBuilder " + Version + " - ";
  var Step = 100;
  
  /******************************************************************************/
  Step = 2000; // Get the "Target" ParameterSet values from the Setup Tab
  /******************************************************************************/
  var ErrorMessage = null;
  var AllParameters = [];
  var ParameterSourceTab = oCommon.SetupSheetName;
  var KeyWord = ParameterSet;
  var StartRow = 0; 
  var KeyCol = 0;
  var ValueCol = 0;
  var AsciiValue,
      TitleCol = 0,
      InActiveStatusChar = 9744,
      ActiveStatusChar = 9745;

  AllParameters = GetParameters(oCommon.Sheets, ParameterSourceTab, StartRow, KeyWord, KeyCol, ValueCol, ErrorMessage);  
  
  if (!ParamCheck(AllParameters)){
    Step = 2010; // throw error...
    oCommon.DisplayMessage = ' Error 310 - Parameter Set (' + KeyWord + ') Not Found: ' + ErrorMessage;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(oCommon.ReturnMessage);
    return;
  }
  Logger.log(func + Step + ' AllParameters.length: ' + AllParameters.length + ', Keyword: ' + KeyWord);
  
  Step = 2100; // Create the Heading/Column Value Key array
  var ColNumKeyAry = {};
  for (var c = 0; c < AllParameters[0].length; c++){
    ColNumKeyAry[AllParameters[0][c]] = +c;
  }
  // Verify results
  //Logger.log(func + Step + ' Verify ColNumKeyAry: ');
  //for (var key in ColNumKeyAry) {
  //  Logger.log('Key: %s, Value: %s', key, ColNumKeyAry[key]);
  //}  
  
  Step = 2200; // Create the Title/Row Value Key array
  var RowNumKeyAry = {};
  for (var r = 0; r < AllParameters.length; r++){
    //Check for the presence of an "Active Checkbox" in the first column
    AsciiValue = AllParameters[r][0].charCodeAt(0); // ascii value of the first character
    if(AsciiValue == InActiveStatusChar || AsciiValue == ActiveStatusChar){TitleCol = 1;}
    RowNumKeyAry[AllParameters[r][TitleCol]] = +r;
  }
  // Verify results
  //Logger.log(func + Step + ' Verify RowNumKeyAry: ');
  //for (var key in RowNumKeyAry) {
  //  Logger.log('Key: %s, Value: %s', key, RowNumKeyAry[key]);
  //}  
  
  var oDocParams = {
      HeadingKeyAry: ColNumKeyAry,
      TitleKeyAry: RowNumKeyAry,
      Parameters: AllParameters
      }
  
  return oDocParams;

}


function BuildGlobals(oSourceSheets, SetupSheetName) {
  /* ****************************************************************************
   DESCRIPTION:
     Reads the Setup parameters and builds the Globals script properties

   USEAGE:
     var Globals = {};
     Globals = BuildGlobals(oCommon); 

   REVISION DATE:
     12-14-2017 - Initial design
     12-27-2017 - Change all row/col references to Base 0
     01-28-2018 - Modified to build Search and Replace terms if NO "TermsRow" is specified
     03-23-2018 - Modified to work with v3.x
                - Modified to return Empty Globals Set to facilitate error reporting and User feedback

   
   NOTES:


******************************************************************************/
  var func = "***BuildGlobals " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');
  var Globals = {};
  
  /******************************************************************************/
  Step = 1000; // Validate input parameters
  /*******************************************************************************/
  Globals['ErrorMessage'] = '';
  if (!ParamCheck(oSourceSheets)){
    Step = 1100; // ERROR 101 - No oSourceSheets Object found
    Globals['ErrorMessage'] = ' Error 100 - No oSourceSheets object found.';
    Logger.log(func + Step + Globals['ErrorMessage']);
    return Globals;
  }
  
  if (!ParamCheck(SetupSheetName)){
    Step = 1200;// ERROR - No SetupSheetName found
    Globals['ErrorMessage'] = ' Error 101 - No SetupSheetName found.';
    Logger.log(func + Step + Globals['ErrorMessage']);
    return Globals;
  }

  /******************************************************************************/
  Step = 2000; //Capture the Persistent Scalars values from the "Globals" Tab
  /*******************************************************************************/
  // Source: https://stackoverflow.com/questions/16861482/google-apps-script-bulk-set-properties
  // Load the Global Scalars
  // Create a 'Globals' Object for bulk-setting to prevent overusing API calls
  var values = oSourceSheets.getSheetByName(SetupSheetName).getDataRange().getValues();
  // Create a 'Globals' Object for bulk-setting to prevent overusing API calls
  var KeyWord = 'GLOBAL VARIABLES'; // heading text the indicates start of Global scalar assignments
  var col = 0; // Column containing the "Key" value (Base 0)
  // Find the start row for the Global Value
  for (var row = 0; row < values.length; row++) { 
    if (values[row][col].toUpperCase().indexOf(KeyWord) > -1){
      var GlobalStartRow = row + 1;
      break;
    }
  }
  if (row >= values.length){
    Step = 2010;// ERROR - No "Global Variables" Section found
    Globals['ErrorMessage'] = ' Error 102 - "Global Variables" Section not found in Setup Sheet.';
    Logger.log(func + Step + Globals['ErrorMessage']);
    return Globals;
  }
  
  Step = 2100; // Capture the Global scalar value assignments
  for (var row = GlobalStartRow; row < values.length; row++) { 
    if(!values[row][col]){break;}
    Globals[values[row][col]] = values[row][col + 1]; 
  }
  Logger.log(func + Step + ' Part 1 - Last Row: ' + row);
  //for(var Key in Globals){
  //  Logger.log(func + Step + ' Key:' + Key + '  Value: ' + Globals[Key]);
  //}
  
  return Globals;
}
  
function BuildCommon(Sheets, Globals, bSilentMode){
  /* ****************************************************************************

  DESCRIPTION:
     This function is used to build an object containing the common parameter values
     used by the other functions in this library.
     
  USEAGE:
  
     var oCommon = {};
     oCommon = BuildCommon(Sheets,Globals,ErrorMessage);
     

   REVISION DATE:
    12-14-2017 - First Instance
    03-23-2018 - modified for use witk ver 3.x 
    02-04-2019 - added SendEmails() parameters

   NOTES:
     1. The Globals are created using the OnOpenProcedures() function and contain the values found in the Setup Sheet
   
  ******************************************************************************/
  var func = "***BuildCommon " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  if (!ParamCheck(bSilentMode)){
    bSilentMode = true;
  }

  var aryEventMessages = [];
  var arySelectedRows = [];
  var aryChangedRows = [];
 
  /* try {
    var user_email = Session.getActiveUser().getEmail();
  } catch (err) {
    var user_email = '';
  } */
  
  // Build the Common Properties object to be passed to all of the functions
  
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
    bActiveUserEmpty: true,
    
    SelectedSheetRows: arySelectedRows,
    ChangedSheetRows: aryChangedRows,
    
    TestMode: Globals['Test_Mode'],
    EventMessages: aryEventMessages,
    CallingMenuItem: null,
    bSilentMode: bSilentMode,
    ReturnMessage: '',
    DisplayMessage: '',
    DocsGenerated: 0,
    EmailsGenerated: 0,
    
    GoogleForm_Url: Globals['GoogleForm_URL'],
    GoogleForm_Name: Globals['GoogleForm_Name'],
    FormSubmit_Delay: Globals['FormSubmit_Delay'],
    onSubmit_Sensitivity:Globals['onSubmit_Sensitivity'],
    FormTimestamp: null,
    FormTriggerUid: null,
    
    SetupSheetName: Globals['SetupSheetName'],
    SetupStartRow: -1,
    EventMesssagesTab: Globals['EventMesssagesTab'],
    MaxEventMessages: +Globals['MaxEventMessages'],
    IgnoreEventsFor: Globals['IgnoreEventsFor'],
    Send_error_report_to: Globals['send_error_report_to'],
    
    FormResponse_SheetName: Globals['FormResponse_SheetName'],
    AlertsScanSource:'',
    TimestampCol: -1,
    CheckForValidTimestamp: Globals['CheckForValidTimestamp'],
    Form_Owner_Col: -1,
    Record_statusCol: -1,
    Active_Status_value: Globals['Active_Status_value'],
    Inactive_Status_value: Globals['Inactive_Status_value'],
    RecordTagValueCol: -1,
    RecordTagMaxLength: Globals['Record_Tag_Max_Length'],
    AsOfDateCol: -1,
    Edit_URLCol: -1,
    Edit_ButtonCol: -1,
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
    
    FatalErrorMessage: null
    
  }
  
  /******************************************************************************/
  Step = 1000; // Get the email address of the person initiating this procedure
  /******************************************************************************/
  try {
    Step = 1100;
    var user_email = Session.getEffectiveUser().getEmail();
    oCommon.bActiveUserEmpty = false;  
  } catch (err1) {
    try{
      Step = 1200;
      var user_email = Session.getActiveUser().getEmail();
      oCommon.bActiveUserEmpty = false; 
     } catch (err2) {
       Step = 1300; 
       oCommon.ReturnMessage = func + Step + ' Error 104 - Unable to get User email address: ' + err2;
       Logger.log(oCommon.ReturnMessage);
       LogEvent(oCommon.ReturnMessage,oCommon);
       var user_email = '';
       oCommon.bActiveUserEmpty = true;
     }
  }
  oCommon.ScriptUser = user_email;
  Logger.log(func + Step + ' bActiveUserEmpty: ' + oCommon.bActiveUserEmpty
             + ' ScriptUser: ' + oCommon.ScriptUser);
  
  /******************************************************************************/
  Step = 2000 // Validate the "general" parameters
  /******************************************************************************/

  Step = 2200; // Using_TeamDrive check
  if (oCommon.Using_TeamDrive == 'true' || oCommon.Using_TeamDrive == 'TRUE'){
    oCommon.Using_TeamDrive = true;
  } else {
    oCommon.Using_TeamDrive = false;
  }
  Logger.log(func + Step + ' Using_TeamDrive: ' + oCommon.Using_TeamDrive);
  
  /******************************************************************************/
  Step = 3000; // Build the Sheet Format Details Key/Values object
               // Sheet Details Headings: Title, Headings Row, 1st Data Row, Terms Row, AWT Row, Param 1
               // var HeadingsRow =  oCommon.sheetParams[TabName]['Headings Row'];
               // var 1stDataRow = oCommon.sheetParams[TabName]['1st Data Row'];
               // var TermsRow = oCommon.sheetParams[TabName]['Terms Row'];
               // var AWTRow = oCommon.sheetParams[TabName]['AWT Row'];
               // var Param1 =oCommon.sheetParams[TabName]['Param 1'];
  /******************************************************************************/
  Logger.log(func + Step + ' SheetDetails_Headings value: ' + Globals['SheetDetails_Heading']);
  try {
    if(ParamCheck(Globals['SheetDetails_Heading'])){
      Step = 3100; // Get the Sheet Details Setup Section values
      var oSheetParams = {};
      var ParameterSet = Globals['SheetDetails_Heading'];  
      oSheetParams = ParamSetBuilder(oCommon, ParameterSet);
      /*RETURNS Object
      oSheetParams = {
      HeadingKeyAry: ColNumKeyAry,
      TitleKeyAry:   RowNumKeyAry,
      Parameters:    AllParameters
      } */
      Logger.log(func + Step + ' oSheetParams.length: ' + oSheetParams.length);

      if (!ParamCheck(oSheetParams)){
        Step = 3200;// ERROR - No Sheet Format Parameters found
        oCommon.DisplayMessage = ' ERROR 108 - Expected "' + oCommon.SheetDetails_Heading 
          + '" Parameters not found (' + ParameterSet + ').';
        oCommon.FatalErrorMessage = oCommon.DisplayMessage;
        oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
        LogEvent(oCommon.ReturnMessage, oCommon); 
        Logger.log(func + Step + ' ReturnMessage: ' + oCommon.ReturnMessage);
        
      } else {
      
        Step = 3300; // Build the sheetParams[Sheet Name]/[Values,,,,] object
        var paramValue = '';
        var patt1 = /row|Row|ROW/i;
        for (var key1 in oSheetParams.TitleKeyAry) {
          if (oSheetParams.TitleKeyAry[key1] > 0 ){
            // Capture the column values
            oCommon.sheetParams[key1] = [];
            for (var key2 in oSheetParams.HeadingKeyAry){
              if (oSheetParams.HeadingKeyAry[key2] > 1 && key2 != ''){ 
                paramValue = oSheetParams.Parameters[oSheetParams.TitleKeyAry[key1]][oSheetParams.HeadingKeyAry[key2]];
                // Test to see if the parameter is for a row value
                if(patt1.test(key2)){
                  paramValue = paramValue - 1 ;// adjust to Base 0
                }
                // Assign the parameter value
               oCommon.sheetParams[key1][key2] = paramValue;
               //Logger.log(func + Step + ' Sheet Details - Sheet": ' + key1 + '", Param: ' + key2
               //  + ' Value: ' + oCommon.sheetParams[key1][key2]);
              }
            }
          }
        }
        
        Step = 3400; //Verify the results
        //Headings Row, 1st Data Row, Terms Row, AWT Row, Param 1
        /*
        if (ParamCheck(oCommon.FormResponse_SheetName)){ var Tab_Name = oCommon.FormResponse_SheetName; }
            else { var Tab_Name = 'Default'; }
        Logger.log(func + Step + ' sheetParams[' + Tab_Name + '][Headings Row]: ' + oCommon.sheetParams[Tab_Name]['Headings Row']);
        Logger.log(func + Step + ' sheetParams[' + Tab_Name + '][1st Data Row]: ' + oCommon.sheetParams[Tab_Name]['1st Data Row']);
        Logger.log(func + Step + ' sheetParams[' + Tab_Name + '][Terms Row]: ' + oCommon.sheetParams[Tab_Name]['Terms Row']);
        Logger.log(func + Step + ' sheetParams[' + Tab_Name + '][AWT Row]: ' + oCommon.sheetParams[Tab_Name]['AWT Row']);
        Logger.log(func + Step + ' sheetParams[' + Tab_Name + '][Param 1]: ' + oCommon.sheetParams[Tab_Name]['Param 1']);
        */
      }
    }
  } catch(err) {
    oCommon.DisplayMessage = ' ERROR 110 - Unexpected error loading Sheet Details parameters: ' + err;
    oCommon.FatalErrorMessage = oCommon.DisplayMessage;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon);
  }

  /******************************************************************************/
  Step = 4000; // Build the Xref Arrays using the Form Response or the "Active" Tab
  /******************************************************************************/
  Logger.log(func + Step + ' bSilentMode: ' + oCommon.bSilentMode 
             + ', FormResponse_SheetName: ' + oCommon.FormResponse_SheetName);
  if (oCommon.bSilentMode && ParamCheck(oCommon.FormResponse_SheetName)){
    Step = 4110; // Use the form response tab, if defined
    var dataTab = oCommon.FormResponse_SheetName;
    var sheet = Sheets.getSheetByName(dataTab);
    var SheetData = sheet.getDataRange().getValues();
  } else {
    Step = 4120; // Use the "Active" Tab
    var sheet = Sheets.getActiveSheet();
    var dataTab = Sheets.getActiveSheet().getName();
    var SheetData = Sheets.getActiveSheet().getDataRange().getValues();
  }
  oCommon.XrefSourceSheet = dataTab;

  var headings_row = GetSheetParam(dataTab, 'Headings Row', oCommon);
  var terms_row = GetSheetParam(dataTab, 'Terms Row', oCommon);
  Logger.log(func + Step + ' dataTab: ' + dataTab + ' Headings Row: ' + headings_row
             + ', Terms Row: ' + terms_row);
  
  Step = 4200;// Test for the presence of a Headings row
  if (headings_row > -1){
    Step = 4210;// Build the Headings Key/Value Array
    var TargetSheetName = dataTab, // Sheet Containing Keys and Values
        SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start row is always the following row
        KeyRow = headings_row,   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
        ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
        KeyCol = '',   // number or text [Look for text Match in StartRow (Base 0)]
        ValueCol = ''; // number or text [Look for text Match in StartRow (Base 0)]
    //oCommon.HeadingsXrefAry = {}; // if declared above, before entering function BuildKeyValueAry()
    oCommon.HeadingsXrefAry = BuildKeyValueAry(Sheets,TargetSheetName,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol)
    Step = 4215; // Verify
    //Logger.log(func + Step + ' Verify Headings Xref for: ' + dataTab);
    //for(var key in oCommon.HeadingsXrefAry){
    //  Logger.log(func + Step + ' Headings Key: ' + key + '  Value: ' + oCommon.HeadingsXrefAry[key]);
    //}
  } else {
    Step = 4220;// Warning - No Headings row identified
    oCommon.DisplayMessage = ' Warning 112 - No Headings row identified for "' + dataTab + '" Sheet.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
  }

  Step = 4300;// Build the Terms Key/Value Array
  if (terms_row > -1){
    Step = 4310; // Build the Terms Key/Value Array using the "terms" row
    var TargetSheetName = dataTab, // Sheet Containing Keys and Values
        SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start row is always the following row
        KeyRow = terms_row,   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
        ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
        KeyCol = '',   // number or text [Look for text Match in StartRow (Base 0)]
        ValueCol = ''; // number or text [Look for text Match in StartRow (Base 0)]
    //oCommon.TermsXrefAry = {}; // if declared above, before entering function BuildKeyValueAry()
    oCommon.TermsXrefAry = BuildKeyValueAry(Sheets,TargetSheetName,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol)
    Step = 4315; // Verify
    //Logger.log(func + Step + ' Verify Terms Xref for: ' + dataTab);
    //for(var key in oCommon.TermsXrefAry){
    //  Logger.log(func + Step + ' Terms Key: ' + key + '  Value: ' + oCommon.TermsXrefAry[key]);
    //}
  } else if (headings_row > -1) {
    Step = 4320; // Build the Terms Key/Value Array using the "headings" row
    var Prefix = oCommon.Replace_Term_Prefix;
    var Suffix = oCommon.Replace_Term_Suffix;
    var TargetSheetName = dataTab, // Sheet Containing Keys and Values
        SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start row is always the following row
        KeyRow = headings_row,   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
        ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
        KeyCol = '',   // number or text [Look for text Match in StartRow (Base 0)]
        ValueCol = ''; // number or text [Look for text Match in StartRow (Base 0)]
    //oCommon.TermsXrefAry = {}; // if declared above, before entering function BuildKeyValueAry()
    oCommon.TermsXrefAry = BuildKeyValueAry(Sheets,TargetSheetName,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol,Prefix,Suffix)
    Step = 4325; // Verify
    //Logger.log(func + Step + ' Verify Alternate Terms Xref for: ' + dataTab);
    //for(var key in oCommon.TermsXrefAry){
    //  Logger.log(func + Step + ' Alt Terms Key: ' + key + '  Value: ' + oCommon.TermsXrefAry[key]);
    //}
  } else {
    Step = 4330;// Warning - Unable to build "replace terms" cross-reference array
    oCommon.DisplayMessage = ' Warning 114 - Unable to build "replace terms" cross-reference array '
      + ' for Form Response Sheet (' + dataTab + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    Logger.log(func + Step + oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon); 
  }

  Logger.log(func + Step + ' oCommon.Xrefs built using: ' + oCommon.XrefSourceSheet);

  /******************************************************************************/
  Step = 4400; // Build the RunCheckXref Arrays using the "Menu Options" Setup Section
  /******************************************************************************/
  Step = 4210;// Build the Headings Key/Value Array
  var TargetSheetName = oCommon.SetupSheetName, // Sheet Containing Keys and Values
      SectionTitle = oCommon.BuildMenuSection_Heading, // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start row is always the following row
      KeyRow = '',   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
      ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
      KeyCol = 'Function Name',   // number or text [Look for text Match in StartRow (Base 0)]
      ValueCol = 'RunCheck'; // number or text [Look for text Match in StartRow (Base 0)]
  //oCommon.RunCheckXrefAry = {}; // if declared above, before entering function BuildKeyValueAry()
  oCommon.RunCheckXrefAry = BuildKeyValueAry(Sheets,TargetSheetName,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol)
  if (ParamCheck(oCommon.RunCheckXrefAry)){
    Step = 4420; // Verify
    Logger.log(func + Step + ' oCommon.RunCheckXref built using: ' + SectionTitle);
    
    //Logger.log(func + Step + ' Verify RunCheck Xref for: ' + SectionTitle);
    //for(var key in oCommon.RunCheckXrefAry){
    //  Logger.log(func + Step + ' Function Name: ' + key + '  Value: ' + oCommon.RunCheckXrefAry[key]);
    //}
    
  } else {
    Step = 4430;// Warning - No RunCheck parameters found
    oCommon.DisplayMessage = ' Warning 115 - No RunCheck parameters found in Setup Sheet (' + dataTab + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
  }
    
  /******************************************************************************/
  Step = 5000; // Get and verify the presence of key parameters
  /******************************************************************************/
  oCommon.TimestampCol = GetXrefValue(oCommon, Globals['Timestamp_heading']);
  oCommon.Form_Owner_Col = GetXrefValue(oCommon, Globals['Form_Owner_Heading']);
  oCommon.Record_statusCol = GetXrefValue(oCommon, Globals['Record_status_heading']);
  oCommon.RecordTagValueCol = GetXrefValue(oCommon, Globals['Record_Tag_Value_heading']);
  oCommon.AsOfDateCol = GetXrefValue(oCommon, Globals['As_Of_Date_Heading']);
  oCommon.Edit_URLCol = GetXrefValue(oCommon, Globals['Edit_URLCol_heading']);
  oCommon.Edit_ButtonCol = GetXrefValue(oCommon, Globals['Edit_Button_heading']);
  oCommon.SetupStartRow = GetSheetParam(oCommon.SetupSheetName, '1st Data Row', oCommon);

  /******************************************************************************/
  Step = 6000; // Get parameters for accepting Google Form responses
  /******************************************************************************/
  if (dataTab == oCommon.FormResponse_SheetName){
    oCommon.FormResponse_HeadingRow = GetSheetParam(oCommon.FormResponse_SheetName, 'Headings Row', oCommon);
    oCommon.FormResponse_TermsRow = GetSheetParam(oCommon.FormResponse_SheetName, 'Terms Row', oCommon);
    oCommon.AwesomeTableRow = GetSheetParam(oCommon.FormResponse_SheetName, 'AWT Row', oCommon);
    oCommon.FormResponseStartRow = GetSheetParam(oCommon.FormResponse_SheetName, '1st Data Row', oCommon);
    
    Step = 6100; // Verify presence of needed entries
    // Verify presence of TagValueCol
    Logger.log(func + Step + ' TagValueCol: ' + oCommon.RecordTagValueCol
               + ' TimestampCol: ' + oCommon.TimestampCol);
    if(isNaN(oCommon.RecordTagValueCol)){
      Step = 6110; // No TagValue, issue warning and check TimestampCol
      oCommon.FatalErrorMessage = ' Error 116 - TagValue heading value not valid.';
      oCommon.ReturnMessage = func + Step + oCommon.FatalErrorMessage;
      Logger.log(oCommon.ReturnMessage);
      LogEvent(oCommon.ReturnMessage,oCommon);
    }
    
    Step = 6200; // Verify presence of TimestampCol entry
    if(isNaN(oCommon.TimestampCol)){
      // No TimestampCol, issue warning
      oCommon.ReturnMessage = func + Step + ' Warning 118 - Timestamp heading value not valid.';
      Logger.log(oCommon.ReturnMessage);
      LogEvent(oCommon.ReturnMessage,oCommon);
    }
    
    Step = 6300; // Validate CheckForValidTimestamp is a boolean, if specified
    if (oCommon.CheckForValidTimestamp == ''){ 
      oCommon.ReturnMessage = func + Step + ' Warning 120 - Global Variable value "CheckForValidTimestamp" '
      + 'not specified as TRUE or FALSE. Assumed to be TRUE.';
      Logger.log(oCommon.ReturnMessage);
      LogEvent(oCommon.ReturnMessage,oCommon);
      oCommon.CheckForValidTimestamp = true;
    } else if (oCommon.CheckForValidTimestamp == 'TRUE' || oCommon.CheckForValidTimestamp == 'true'){
      oCommon.CheckForValidTimestamp = true; 
    } else { oCommon.CheckForValidTimestamp = false; }
    
    Step = 6400; // Validate EditUrl col
    if(isNaN(oCommon.Edit_URLCol)){
      // No EditUrl col, issue warning
      oCommon.ReturnMessage = func + Step + ' Warning 122 - Edit Url Column heading value not valid.';
      Logger.log(oCommon.ReturnMessage);
      LogEvent(oCommon.ReturnMessage,oCommon);
    }
    
    Step = 6500; // Check for record status check consistancy
    Logger.log(func + Step + ' Active Status value: ' + oCommon.Active_Status_value
               + ', Inctive Status value: ' + oCommon.Inctive_Status_value);
    if (ParamCheck(oCommon.Active_Status_value) || ParamCheck(oCommon.Inctive_Status_value)){
      if (isNaN(oCommon.Record_statusCol)){
        //Throw error
        oCommon.ReturnMessage = func + Step + ' Warning 124 - Record Status value detected, but Status Column value not valid.';
        Logger.log(oCommon.ReturnMessage);
        LogEvent(oCommon.ReturnMessage,oCommon);
      }
    }
    
    Step = 6600; // Verify that the input value for FormSubmit_Delay is acceptable
    if(ParamCheck(oCommon.FormSubmit_Delay)){
      var initial_value = oCommon.FormSubmit_Delay;
      if(isNaN(initial_value)){
        Step = 6610; // No numeric value detected
        oCommon.FormSubmit_Delay = 0;
      } else if(initial_value < 0){
        Step = 6620; // Value less than zero detected
        oCommon.FormSubmit_Delay = 0;
      } else if(initial_value > 50000){
        Step = 6630; // Value exceeding maximum value detected
        oCommon.FormSubmit_Delay = 15000; // Max delay allowed 15 second delay
      }
    } else {
      Step = 6640; // No input value detected
      oCommon.FormSubmit_Delay = 0;
    }
    Logger.log(func + Step + ' FormSubmitDelay Parameter Adjustment Check - Initial value: ' + initial_value
               + ', Adjusted value: ' + oCommon.FormSubmit_Delay);
    
    Step = 6700; // Verify that the input value for onSubmit_Sensitivity is acceptable
    if(ParamCheck(oCommon.onSubmit_Sensitivity)){
      var initial_value = oCommon.onSubmit_Sensitivity;
      if(isNaN(initial_value)){
        Step = 6710; // No numeric value detected
        oCommon.onSubmit_Sensitivity = 2000; // Default is +/- 2 seconds
      } else if(initial_value < 0){
        Step = 6720; // Value less than zero detected
        oCommon.onSubmit_Sensitivity = 0;
      } else if(initial_value > 10000){
        Step = 6730; // Value exceeding maximum value detected
        oCommon.FormSubmit_Delay = 10000; // Maximum allowed is is +/- 10 seconds
      }
    } else {
      Step = 6740; // No input value detected
      oCommon.onSubmit_Sensitivity = 2000; // Default is +/- 2 seconds
    }
    Logger.log(func + Step + ' onSubmit_Sensitivity Parameter Adjustment Check - Initial value: ' + initial_value
               + ', Adjusted value: ' + oCommon.onSubmit_Sensitivity);
    
  } // End of Get parameters for accepting Google Form responses

  return oCommon;

}


function FormatAwesomeTableParameters(oCommon){
  /* ****************************************************************************
   DESCRIPTION:
     This little utility was developed to verify that the Awesome Table "VIEW" and "Edit" 
        button template string are pointing to the correct sheet columns
     
   USEAGE:
     ReturnBool = FormatAwesomeTableParameters(oCommon); 

   REVISION DATE:
     12-14-2017 - Initial design
     12-27-2017 - Change all row/col references to Base 0
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback
     
   NOTES:

    *** If this ever failes with a "Range Not Found" error, it is most likely because the 
      <<Edit_Label>> and/or the <<View_Label>> terms in row 2 have been corrupted

******************************************************************************/
  // this little utility was developed to verify that the Awesome Table "VIEW" and "Edit" 
  //    button template string are pointing to the correct sheet columns
  var func = "***formatAwesomeTableParameters " + Version + " - ";
  var Step = 1000;
  Logger.log(func + Step + ' BEGIN...');
  
  var Step = 1100;// Test to determine if the Awseome Table parameters are used
  var AwesomeTableRow = oCommon.AwesomeTableRow;
  if (!ParamCheck(AwesomeTableRow) || AwesomeTableRow <= 0){
    Logger.log(func + Step + ' Awesome Table Parameters not used.');
    return true;
  }
  Logger.log(func + Step + ' Awesome Row: ' + AwesomeTableRow);
  
  /******************************************************************************/
  Step = 2000; /* Get the Document Parameters from the Setup Tab
                    Returns: oDocParams object containing:
                               HeadingKeyAry: ColNumKeyAry,
                               TitleKeyAry: RowNumKeyAry,
                               Parameters: AllParameters
  *******************************************************************************/
  Logger.log(func + Step + ' Got here...');
  var oDocParams = {};
  var ParameterSet = oCommon.Globals['DocumentsSection_Heading'];  
  oDocParams = ParamSetBuilder(oCommon, ParameterSet);
  if(!oDocParams){
    Step = 2100;// ERROR - No Document Parameters found
    oCommon.DisplayMessage = ' ERROR 210 - Expected Document Parameters not found ('
      + ParameterSet + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    oCommon.DisplayMessage = '';
    LogEvent(oCommon.ReturnMessage, oCommon); 
    return false;
  }
  Step = 2200; // Get the Form Response Sheet object
  Logger.log(func + Step + ' oDocParams.length: ' + oDocParams.Parameters.length);
  var oSourceSheets = oCommon.Sheets;
  var dataTab = oCommon.FormResponse_SheetName;
  var sheet = oSourceSheets.getSheetByName(dataTab);
   
  /******************************************************************************/
  Step = 3000; // format the EDIT button parameter
  /******************************************************************************/
  if (ParamCheck(oCommon.Edit_ButtonCol) && ParamCheck(oCommon.Edit_URLCol)){
    Step = 3100; // Test for valid entries
    var editButtonCol = oCommon.Edit_ButtonCol;
    var editUrlCol = oCommon.Edit_URLCol;
    Logger.log(func + Step + ' editButtonCol: ' + editButtonCol + ' editUrlCol: ' + editUrlCol);
    if (editButtonCol > 0 && editUrlCol > 0){
      Step = 3200; 
      var editButtonLabel = 'buttonType(' + numToA(editUrlCol + 1) + ',#ffc107)'; // buttonType(BR,#ffc107)
      var editButtonAddr = numToA(editButtonCol + 1) + (AwesomeTableRow + 1);    
      sheet.getRange(editButtonAddr).setValue(editButtonLabel); 
      Logger.log(func + Step + ' editButtonLabel: ' + editButtonLabel + ' Addr: ' + editButtonAddr);
    } else if (editUrlCol > 0){ 
      Step = 3300;// No EditButton column value detected
      oCommon.DisplayMessage = ' Warning 212 - No EditButton column value detected.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      oCommon.DisplayMessage = '';
      LogEvent(oCommon.ReturnMessage, oCommon); 
    } else if (editButtonCol > 0) {
      Step = 3400;// Invalid EditURL column value detected
      oCommon.DisplayMessage = ' Warning 214 - Invalid EditURL column value (' 
        + editUrlCol + ') detected.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      oCommon.DisplayMessage = '';
      LogEvent(oCommon.ReturnMessage, oCommon); 
    }
  }

  /******************************************************************************/
  Step = 4000; // format the VIEW button for each report
  /******************************************************************************/
  Logger.log(func + Step + ' Got here...');
  for ( var r = 1; r < oDocParams.Parameters.length; r++){
    Step = 4100; // Get row entries
    var DocURLTerm = oDocParams.Parameters[r][oDocParams.HeadingKeyAry['URL Col Heading']]; 
    var DocURLCol = GetXrefValue(oCommon, DocURLTerm);
    Logger.log(func + Step + ' r: ' + r + ' DocURLTerm ' + DocURLTerm + ' DocURLCol: ' + DocURLCol);
    var DocUrlAddrCol = numToA(DocURLCol + 1);
    var viewButtonHeadingTerm = oDocParams.Parameters[r][oDocParams.HeadingKeyAry['Button Heading']];
    var viewButtonHeadingCol = GetXrefValue(oCommon, viewButtonHeadingTerm); // column Address where View Details ButtonCol text is to be placed for the AwesomeTables filter row  
    Logger.log(func + Step + ' r: ' + r + ' DocURLTerm ' + DocURLTerm + ', Addr: ' + DocUrlAddrCol);
    Logger.log(func + Step + ' r: ' + r + ' viewButtonHeadingTerm: ' + viewButtonHeadingTerm + ', viewButtonHeadingCol: ' + viewButtonHeadingCol);
    
    if (ParamCheck(viewButtonHeadingCol) && ParamCheck(DocUrlAddrCol)){
      Step = 4200; // Test for valid row entry
      if (viewButtonHeadingCol > 0 && isNaN(DocUrlAddrCol)){
        Step = 4210; 
        var viewButtonLabel = 'buttonType(' + DocUrlAddrCol + ',#ffc107)'; // buttonType(BR,#ffc107)
        var viewButtonHeadingAddr = numToA(viewButtonHeadingCol + 1) + (AwesomeTableRow + 1); // A1 notation is always Base 1  
        Logger.log(func + Step + ' r: ' + r + ' viewButtonLabel: ' + viewButtonLabel + ', Addr: ' + viewButtonHeadingAddr);
        sheet.getRange(viewButtonHeadingAddr).setValue(viewButtonLabel); 
      } else {
        Step = 4300; // Invalid viewButton and/or DocURL column value detected in current row
        oCommon.DisplayMessage = ' Error 216 - Invalid viewButton and/or DocURL column value detected in row: ' + r 
          + ' in Setup Section (' + ParameterSet + ').';
        oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
        oCommon.DisplayMessage = '';
        LogEvent(oCommon.ReturnMessage, oCommon); 
      } 
    } else {
      Step = 4400; // No viewButton and/or DocURL column value detected in current row
      oCommon.DisplayMessage = ' Error 218 - No viewButton and/or DocURL column value detected in row: ' + r 
        + ' in Setup Section (' + ParameterSet + ').';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      oCommon.DisplayMessage = '';
      LogEvent(oCommon.ReturnMessage, oCommon); 
    }
  } 
  /******************************************************************************/
  Step = 5000; // Hide the row containing the Awesome Tables parameters
  /******************************************************************************/
  try {
    sheet.hideRows(AwesomeTableRow);
  } catch (e) {
    oCommon.DisplayMessage = ' Error 219 - Unable to hide AwesomeTable Row (' + AwesomeTableRow
      + '): ' + e.message;
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    oCommon.DisplayMessage = '';
    LogEvent(oCommon.ReturnMessage, oCommon); 
  }
  
  Logger.log(func + Step + ' END...');
  return true;
}

function ValidationScan(oCommon, Groups) {
  /* ****************************************************************************

   DESCRIPTION:
     This function is invoked whenever a "validation" of data is needed to test for:
         1 - expiration dates (Expiry)
         2 - required response to a True of False Form response (T/F)
         3 - required text "string" response to a Form response (InStr)
         4 - valid "date" test

   USEAGE:
     ReturnBool = ValidationScan(oCommon); 
     
   OUTPUT:
     oCommon.ChangedSheetRows[i][0] = SheetDataRow
     oCommon.ChangedSheetRows[i][1] = TagValue
     oCommon.ChangedSheetRows[i][2] = AlertMessage

   REVISION DATE:
     10-10-2017 - First Instance
     11-27-2017 - Modified to use GetParameters() and place Validation Parameters into the SetUp Tab
     02-03-2018 - Modified to enable the use of "groups" of validation parameters (i.e allow user
                   to "group" parameters for test to be performed at different times (e.g. daily, 
                   weekly, monthly, etc.)
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback

   NOTES:
     1. "SheetData" is the range object containg the Google Form responses
     2. "Globals" is the keyed array containing global scalar values
     3. The array oCommon.SelectedSheetRows[SheetRow][TagValue] is passed identifying the selected rows
           oCommon.SelectedSheetRows.length == 0 if an invalid range has been selected by the user     
     
    click on the menu Edit -> Current project's triggers and set up an automated trigger onFormSubmission

  ******************************************************************************/
  var func = "***ValidationScan " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN...');

  // Leave immediately if no selected rows are passed
  if(!oCommon.SelectedSheetRows || oCommon.SelectedSheetRows.length <= 0){
    oCommon.DisplayMessage = ' Error 450 - Selected Sheet Rows array is EMPTY or NULL.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    return false;
  }
  
  /****************************************************************************/
  Step = 1100; // Assign the oCommon parameters
  /****************************************************************************/
  Logger.log(func + Step + ' SelectedSheetRows.length: ' + oCommon.SelectedSheetRows.length)
  var oSourceSheets = oCommon.Sheets, 
      dataTab = oCommon.FormResponse_SheetName,
      firstDataRow = oCommon.FormResponseStartRow,
      timestampCol = oCommon.TimestampCol,
      StatusCol = oCommon.Record_statusCol,
      ActiveValue = oCommon.Active_Status_value,
      InactiveValue = oCommon.Inctive_Status_value,
      ParseSymbol = oCommon.ParseSymbol,
      ValidationSection = oCommon.ValidationSection_Heading
    ;
  
  /**********************************************************************************/
  Step = 1200; // Prepare objects for writing results back to the Form Response data sheet
  /**********************************************************************************/
  var sheet = oSourceSheets.getSheetByName(dataTab);
  var SheetData = sheet.getDataRange().getValues();
  
  /*****************************************************************************/
  Step = 1300; // Load the Validation Parameters
  /*****************************************************************************/
  var oValParams = {};
  var ParameterSet = ValidationSection; 
  oValParams = ParamSetBuilder(oCommon, ParameterSet);
  //Logger.log(func + Step + ' ParameterSet: ' + ParameterSet + ' Parameters: ' + oValParams.Parameters);
  
  var numparams = oValParams.Parameters.length;
  
  if (numparams <=0){
    oCommon.DisplayMessage = ' Error 452 - Email Parameters NOT found.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    return false;
  }
  
  /*****************************************************************************/
  Step = 1310; // Determine if Validation Parameters are segregated by a "group number"
  /*****************************************************************************/
  var bUseScanGroups = false;
  var ScanGroupCol = oValParams.HeadingKeyAry['Group'];  
  if (ParamCheck(ScanGroupCol)){
    Step = 1320; // Validation Parameters are segregated by a "group number"
    var GroupNumberAry = [];
    if (ParamCheck(Groups)){
      Step = 1330; // Identify the Group Column and decode the input Group Value
      // Decode the input Group parammeter
      GroupNumberAry = Groups.split(',');
      bUseScanGroups = true;
    }
  }
  Logger.log(func + Step + ' bUseScanGroups: ' + bUseScanGroups + ', ScanGroupCol: ' + ScanGroupCol + ', GroupNumberAry: ' + GroupNumberAry);
  
  /******************************************************************************/
  Step = 2000; // Examine each selected Sheet row and validate each target column
  /******************************************************************************/
  var todaysDate = new Date();
  for (var ssrow = 0; ssrow < oCommon.SelectedSheetRows.length; ssrow++) {
    var SheetDataRow = oCommon.SelectedSheetRows[ssrow][0] - 1; // SheetData and SelectedSheetRows are base 0
    Step = 2010; // Chack for valid input data passed from SelectedSheetRows()
    if (!ParamCheck(SheetDataRow) || SheetDataRow > SheetData.length){
      oCommon.DisplayMessage = ' Error 454 - Invalid Sheet Data row number (' + ssrow + ') passsed to Validation Scan function.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon); 
      break;
    }
    //Logger.log(func + Step + ' SheetDataRow: ' + SheetDataRow + ', ssrow: ' + ssrow + ', oCommon.SelectedSheetRows.length: ' + oCommon.SelectedSheetRows.length);
    if (!ParamCheck(SheetData[SheetDataRow])){
      oCommon.DisplayMessage = ' Error 456 - Invalid Sheet Data row (' + ssrow + ') found.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon); 
      break;
    }
    
    /******************************************************************************/
    Step = 2100; // Verifiy that the selected row is for an active record
    /******************************************************************************/
    var AlertMessage = '';
    var Timestamp = SheetData[SheetDataRow][timestampCol];
    var Status = SheetData[SheetDataRow][StatusCol];
    var TagValue = oCommon.SelectedSheetRows[ssrow][1];
    
    Logger.log(func + Step + '**************************** ' + TagValue); // Insert blank row in Log
    Logger.log(func + Step + ' SheetDataRow ' + SheetDataRow + ' Record_Status Col: ' + StatusCol + ' Value: ' + SheetData[SheetDataRow][StatusCol]);
    var CheckStatusReturnValue = CheckStatus(Timestamp, Status, oCommon);
    Logger.log(func + Step + ' CheckStatusReturnValue: ' + CheckStatusReturnValue + ', Status: ' + Status + ', Timestamp: ' + Timestamp);
    
    if (CheckStatus(Timestamp, Status, oCommon)){
      // Record is active
      var TagValue = SheetData[SheetDataRow][oCommon.RecordTagValueCol]; 
      TagValue = FormatTagValue(oCommon, TagValue);
      /*****************************************************************************/
      Step = 2200; // Iterate thru the Validation Parameters and perform the specified test for the SheetDataRow
      /*****************************************************************************/
      // Col Values: 0-Column Heading, 1-Type, 2-Param1, 3-Param2, 4-Param3
      for (var i = 1; i < oValParams.Parameters.length; i++){
        var bUseParameter = true;
        // Use the parameter if it is a member of the "Group" to be used
        if (bUseScanGroups){
          Step = 2205; 
          bUseParameter = false;
          var TargetGroup = oValParams.Parameters[i][oValParams.HeadingKeyAry['Group']];
          //Logger.log(func + Step + ' TargetGroup: ' + TargetGroup + ' GroupCol: ' + oValParams.HeadingKeyAry['Group']);
          for (var g = 0; g < GroupNumberAry.length; g++){
            Step = 2206; 
            //Logger.log(func + Step + ' GroupNumberAry[' + g + ']: ' + GroupNumberAry[g]);
            if(TargetGroup.toString().toUpperCase() == GroupNumberAry[g].toString().toUpperCase()){
              Step = 2207; 
              bUseParameter = true;
              break;
            }
          }
        }
        //Logger.log(func + Step + ' TargetGroup: ' + TargetGroup + ' bUseParameter: ' + bUseParameter);
        /*****************************************************************************/
        Step = 2210; // Use selected Parameter row to perform the specified test for the SheetDataRow
        /*****************************************************************************/
        if(bUseParameter){
          var return_message = '';
          Step = 2220; // Test for a valid Column Heading
          /*        
          var a = i;
          var b = oValParams.Parameters[i][0];
          var c = oValParams.HeadingKeyAry['Column Heading'];
          var d = oValParams.Parameters[i][oValParams.HeadingKeyAry['Column Heading']];
          var e = GetXrefValue(oCommon, d);
          var f = oValParams.HeadingKeyAry['Precedent Col Heading'];
          var g = oValParams.Parameters[i][oValParams.HeadingKeyAry['Precedent Col Heading']];
          var h = GetXrefValue(oCommon, d);
          
          Logger.log(func + Step + ' Param Row: ' + a);
          Logger.log(func + Step + ' Target Column Col: ' + b);
          Logger.log(func + Step + ' Target Column Hdg: ' + c);
          Logger.log(func + Step + ' Sheet Target Hdg: ' + d);
          Logger.log(func + Step + ' Sheet Target Col: ' + e);
          Logger.log(func + Step + ' Prcdnt Column Hdg: ' + f);
          Logger.log(func + Step + ' Prcdnt Sheet Target Hdg: ' + g);
          Logger.log(func + Step + ' Prcdnt Sheet Target Col: ' + h);
          */            
          
          var TargetTerm = oValParams.Parameters[i][oValParams.HeadingKeyAry['Column Heading']];
          var PrecedentTerm = oValParams.Parameters[i][oValParams.HeadingKeyAry['Precedent Col Heading']];
          //Logger.log(func + Step + ' Param: ' + i + ', TargetTerm: ' + TargetTerm + ', PrecedentTerm: ' + PrecedentTerm);
          var TargetCol = GetXrefValue(oCommon, TargetTerm);
          var PrecedentTestCol = GetXrefValue(oCommon, PrecedentTerm);
          //Logger.log(func + Step + ' i: ' + i + ', TargetTerm: ' + TargetTerm + ', TargetCol: ' + TargetCol + ' PrecedentTestCol: ' + PrecedentTestCol);
          //Logger.log(func + Step + ' i: ' + i + ', ************ Starting Validation of ' + TargetTerm);
          
          /*****************************************************************************/
          Step = 2300; //Test to determine if a Precedent Condition is to be evaluated first
          /*****************************************************************************/
          var passedPrecedentTest = true;
          if(!isNaN(PrecedentTestCol)){
            //Column Heading	Type of Test	Desired Answer	Incorrect Answer	Precedent Col Heading	Precedent Type of Test	Precedant Desired Answer
            //Logger.log(func + Step + ' i: ' + i + ', ************ Starting Validation Precedent Term: ' + PrecedentTerm);
            var TestType = oValParams.Parameters[i][oValParams.HeadingKeyAry['Precedent Type of Test']];
            var InputValue = SheetData[SheetDataRow][PrecedentTestCol];
            var DesiredResult = oValParams.Parameters[i][oValParams.HeadingKeyAry['Precedant Desired Answer']];
            var IncorrectAnswer = ''; // Not used for a Precedent test
            var PrecedentTestMsg = PerformTest(TestType,InputValue,DesiredResult,IncorrectAnswer);
            if (PrecedentTestMsg){
              passedPrecedentTest = false;
            }
          }  
          /*****************************************************************************/
          Step = 2400; //Perform Primary Validation
          /*****************************************************************************/
          if (!isNaN(TargetCol) && passedPrecedentTest){
            // Valid / existing heading, determine the type of test to run
            
            var TestType = oValParams.Parameters[i][oValParams.HeadingKeyAry['Type of Test']].toUpperCase();
            var InputValue = SheetData[SheetDataRow][TargetCol];
            var DesiredResult = oValParams.Parameters[i][oValParams.HeadingKeyAry['Desired Answer']];
            var IncorrectAnswer = oValParams.Parameters[i][oValParams.HeadingKeyAry['Incorrect Answer']];
            
            //Logger.log(func + Step + ' i: ' + i + ', ************ Starting Primary Validation of ' + TargetTerm);
            
            var TestMessage = PerformTest(TestType,InputValue,DesiredResult,IncorrectAnswer);
            
            /*****************************************************************************/
            Step = 3000; // Evaluate the results
            /*****************************************************************************/
            //Logger.log(func + Step + ' i: ' + i + ', TargetTerm: ' + TargetTerm + ', Result: ' + TestMessage);
            if (TestMessage){
              AlertMessage = AlertMessage + ParseSymbol + TestMessage + ' \n';
              //Logger.log(func + Step + ' i: ' + i + ', TargetTerm: ' + TargetTerm + ', AlertMessage: ' + AlertMessage);
              //Logger.log(func + Step + ' i: ' + i + ', ************ Validation of ' + TargetTerm + ' Failed: ' + AlertMessage);
            }  
          } //End of Perform Primary Validation
        } // End of bUseParameter test
      } // End of Iterate thru the Validation Parameters
    } // end of test for "ACTIVE" record
    
    /**************************************************************************/
    Step = 4000; // Capture the ChangedSheetRows data
    /**************************************************************************/
    //Logger.log(func + Step + ' SheetRow: ' + SheetDataRow + ', TagValue: ' + TagValue  + ', AlertMessage: ' + AlertMessage);
    if(AlertMessage){
      Step = 4100; //var TagValue = SheetData[ArrayRow][oCommon.RecordTagValueCol - 1]; 
      Logger.log(func + Step + ' Validation Errors Found - SheetRow: ' + (SheetDataRow+1) + ' TagValue: ' + TagValue  + ' AlertMessage: ' + AlertMessage);
      oCommon.ChangedSheetRows.push([[SheetDataRow+1],[TagValue],[AlertMessage]]); // capture the Sheet row number
    } else {
      // No Validation error found
      Logger.log(func + Step + ' No Validation Errors Found - SheetRow: ' + SheetDataRow + ' TagValue: ' + TagValue);
    }
    
  } // iterate to the next SelectedSheet row
  
  return true; 
  
}



function PerformTest(TestType,InputValue,DesiredResult,IncorrectAnswer){
  var func = "***PerformTest " + Version + " - ";
  var Step = 1000;
  //Logger.log(func + Step + ' BEGIN');
  //Logger.log(func + Step + ' TestType: ' + TestType + ' InputValue: ' + InputValue + ' DesiredResult: ' + DesiredResult
  //   + ' IncorrectAnswer: ' + IncorrectAnswer);
  
  var ValidationMessage = '';
  
  if(!TestType){
    Step = 1100;// Leave immediately if no TestType
    Logger.log(func + Step + ' No TestType Specified');
    return true;
  }
  TestType = TestType.toUpperCase();
  /*****************************************************************************/
  Step = 3000; // Perform the specified test for the TargetColumn
  /*****************************************************************************/
  switch(TestType) {
      /*****************************************************************************/
    case "T/F":
      Step = 3100; // Test for True or False Value 
      /*****************************************************************************/
      if (InputValue != DesiredResult){
        ValidationMessage = IncorrectAnswer;
      }
      break;
      /*****************************************************************************/
    case "DATE":
      Step = 3300; // Check for Expiration date within the Expired_Dates_threshold
      /*****************************************************************************/
      //Logger.log(func + Step + ' Date Check - Input value: ' + InputValue);
      var dateRangeValues = DesiredResult;
      if (!InputValue) {
        Step = 3310; // Date Missing
        //Logger.log(func + Step + ' Date Check - Input value: Missing');
        ValidationMessage = IncorrectAnswer + ' date is missing.';
        
      } else if (isValidDate(InputValue) == false){
        Step = 3320; // Date is Invalid
        //Logger.log(func + Step + ' Date Check - Input value: Invalid');
        ValidationMessage = IncorrectAnswer + ' date is not valid.';
        
      } else if (dateRangeValues){
        Step = 3330; // To get to here, the date is valid and the parameters to test for expiration are present
        var ExpireDate = new Date(InputValue);
        var todaysDate = new Date();
        //Logger.log(func + Step + ' Date Computation - Input value: ' + InputValue + ', ExpireDate: ' + ExpireDate);
        var partsOfStr = dateRangeValues.split("/");
        var monthsToAdd = +partsOfStr[0];
        if (monthsToAdd  > 0){ExpireDate = addMonths(ExpireDate, monthsToAdd);}
        var Expired_Dates_threshold = +partsOfStr[1];
        var DaysToDLExpiration = DateDiff(todaysDate, ExpireDate, "DAYS");
        var formattedDate = formatDate(ExpireDate);
        //Logger.log(func + Step + ' Date Check - Input value: ' + formattedDate + ', Parts: ' + partsOfStr);
        //Logger.log(func + Step + ' Date Check - DaysToDLExpiration: ' + DaysToDLExpiration + ', Threshold: ' + Expired_Dates_threshold);
        if (DaysToDLExpiration <= Expired_Dates_threshold){
          if (DaysToDLExpiration > 0) {
            Step = 3340; // Alert for pending Expiration
            ValidationMessage = IncorrectAnswer + ' date is about to expire on ' + formattedDate + '.';
          } else {
            Step = 3350; // Alert Employee of Expiration
            ValidationMessage = IncorrectAnswer + ' date has expired on ' + formattedDate + '.';
          }
        }
      }
      break;
      /*****************************************************************************/
    case "INSTR":
      Step = 3400; // Check for an expected and necessary response    TestType,InputValue,DesiredResult,IncorrectAnswer,AlertMessage
      /*****************************************************************************/
      InputValue = InputValue.toUpperCase();
      DesiredResult = DesiredResult.toUpperCase();
      var startPos = InputValue.indexOf(DesiredResult);
      //Logger.log(func + Step + ' startPos: ' + startPos + ', DesiredResult: ' + DesiredResult + ', InputValue: ' + InputValue);
      if (!InputValue) {
        Step = 3410; // Entry Missing
        ValidationMessage = IncorrectAnswer + ' entry is missing.';
        
      //} else if (DesiredResult && InputValue.indexOf(DesiredResult) < 0){
      } else if (DesiredResult) {
        if (startPos < 0){
          Step = 3420; // Entry not valid
          ValidationMessage = IncorrectAnswer + ' entry is not valid.';
        }
      }
      break;

      /*****************************************************************************/
    default:
      Step = 3500; // Empty or missing TestType
      /*****************************************************************************/
      //Logger.log(func + Step + ' ********** WARNING - Empty or undefined Validation Type encountered');
      //Logger.log(func + Step + ' SheetDataRow ' + SheetDataRow + ' ColTitle: ' + ColTitle + ' TestType Value: ' + TestType);
  }

  //Logger.log(func + Step + ' TestType: ' + TestType + ' InputValue: ' + InputValue + ' ValidationMessage: ' + ValidationMessage);
  return ValidationMessage;
}


function SendAdminEmails(oCommon, EmailTemplateName, oSpecialMsg){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to format and send "Admin Summary" emails

   REVISION DATE:
    07-22-2017 - Add messageTemplateSet and adminTemplateSet as input parameters
    07-25-2017 - Do not send Admin Summary email if Admin_template_addr == ''
    10-14-2017 - Removed non-admin email code
    11-28-2017 - Modified so that email would not be sent to any "Sort Field" w. empty email address
    11-29-2017 - Changed code to use "objects"
    01-02-2018 - Changed code to use oCommmon
    01-22-2018 - Modified code to accept URL or Google ID for template
    04-13-2019 - Modified to enable sending of the App Status Report (i.e. "bSpecialReport = true" )
    

   NOTES:
   
     Input Parameter Values:
     
       Array: oCommon.SelectedSheetRows[n]
         oCommon.SelectedSheetRows[0] = SheetData Target row for each recipient
         oCommon.SelectedSheetRows[1] = recipient last name
         oCommon.SelectedSheetRows[2] = message / email body
   
       EmailTemplateName = string text name of the Google Doc Admin email template to be used 
   
     Check out: https://developers.google.com/apps-script/articles/mail_merge
                https://webapps.stackexchange.com/questions/85017/utilizing-noreply-option-on-google-script-sendemail-function
    
  ******************************************************************************/
  
  var func = "***sendAdminEmails " + Version + " - ";
  Step = 100;
  Logger.log(func + ' BEGIN');
  
  var test_mode = false; //true;
  var return_message = '';
   
  /****************************************************************************/
  Step = 1000; // Validate Input parameters
  /****************************************************************************/
  if(!oCommon.EmailSection_Heading){
    Step = 1010; // Leave immediately if no EmailSection_Heading is passed
    oCommon.DisplayMessage = ' ERROR 350 - Expected EmailSection_Heading is EMPTY or NULL.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    Logger.log(oCommon.ReturnMessage);
    return false;    
  }
 
  if (ParamCheck(oSpecialMsg)){
    var bSpecialReport = true;
  } else {
    var bSpecialReport = false
    var TotalSelectedSheetRows = oCommon.SelectedSheetRows.length;
    if(TotalSelectedSheetRows <= 0 || TotalSelectedSheetRows == null){
      Step = 1020; // Leave immediately if no selected rows are passed
      oCommon.DisplayMessage = ' ERROR 352 - SelectedSheetRows is EMPTY or NULL.';
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon); 
      Logger.log(oCommon.ReturnMessage);
      return false;    
    } 
  }
  
  if(!ParamCheck(EmailTemplateName)){
    Step = 1030; // Leave immediately if no EmailTemplateName is passed
    //LogEvent(EventMsg, oCommon); 
    return return_message;
    oCommon.DisplayMessage = ' ERROR 354 - Admin EmailTemplateName is EMPTY or NULL.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    Logger.log(oCommon.ReturnMessage);
    return false;    
  }
  
  /****************************************************************************/
  Step = 1100; // Assign the oCommon parameters
  /****************************************************************************/
  Logger.log(func + Step + ' Got here...')
  var oSourceSheets = oCommon.Sheets, 
      dataTab = oCommon.FormResponse_SheetName,
      firstDataRow = oCommon.FormResponseStartRow,
      timestampCol = oCommon.TimestampCol,
      StatusCol = oCommon.Record_statusCol,
      InactiveValue = oCommon.Inactive_Status_value,
      editUrlCol = oCommon.Edit_URLCol,
      error_email = oCommon.Send_error_report_to,
      ParseSymbol = oCommon.ParseSymbol,
      ParameterSet = oCommon.EmailSection_Heading
    ;

  /*****************************************************************************/
  Step = 1300; // Load the Email Parameters
  /*****************************************************************************/
  var oEmailParams = {};
  var ParameterSet = oCommon.EmailSection_Heading;  
  oEmailParams = ParamSetBuilder(oCommon, ParameterSet);
  //Logger.log(func + Step + ' ParameterSet: ' + ParameterSet + ' Parameters: ' + oEmailParams.Parameters);
  if(oEmailParams.Parameters.length <= 1){
    Step = 1310; // Return if the parameter set is empty or the email template is not found
    oCommon.DisplayMessage = ' ERROR 356 - Admin Email Parameters (' 
      + ParameterSet + ') Not Found or Empty.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    Logger.log(oCommon.ReturnMessage);
    return false;    
  }
  // Verify results
  //Logger.log(func + Step + ' Verify Headings: ');
  //for (var key in oEmailParams.HeadingKeyAry) {
  //  Logger.log('Key: %s, Value: %s', key, oEmailParams.HeadingKeyAry[key]);
  //}  
  //Logger.log(func + Step + ' Verify Titles: ');
  //for (var key in oEmailParams.TitleKeyAry) {
  //  Logger.log('Key: %s, Value: %s', key, oEmailParams.TitleKeyAry[key]);
  //}  

  /******************************************************************************/
  Step = 1400; // Assign specific email parameters now stored in the EmailParameters[] array
  /******************************************************************************/
  if(!oEmailParams.TitleKeyAry[EmailTemplateName]){
    Step = 1410; // ERROR - Email Template Not Found
    oCommon.DisplayMessage = ' ERROR 358 - Admin Email Template (' 
      + EmailTemplateName + ') NOT found in Setup Section '
      + ParameterSet + ').';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    Logger.log(oCommon.ReturnMessage);
    return false;    
  }  
  var TestEmailAddress = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Test Email Address']];

  Step = 1420; // Determine test_mode status
  if (TestEmailAddress){var test_mode = true;} else {var test_mode = false;}
  
  Step = 1430; // Get the rest...
  var Admin_Statement = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Admin Message']];
  var Admin_subject = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Subject']];
  var Admin_from = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['From']];
  var Admin_to = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['To']];
  var Admin_cc = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['CC']];
  var Admin_bcc = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['BCC']];
  var Admin_Sort_Vaules_Source = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Admin Sort Values Source Sheet']]; 
  
  Step = 1440; // Load the Google id FOR THE SELECTED TEMPLATE
  if(oEmailParams.HeadingKeyAry['Template URL']){
    var Admin_template_Url =  oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Template URL']];
  } else if(oEmailParams.HeadingKeyAry['Template ID']){
    var Admin_template_Url =  oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Template ID']];
  } else {
    Step = 1445;// ERROR - No Email Template found
    oCommon.DisplayMessage = ' ERROR 360 - Admin Email Template ID/URL NOT found.';
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    Logger.log(oCommon.ReturnMessage);
    return false;    
  }
  // Extract the Google ID from a URL, if necessary
  var Admin_template_ID = getIdFrom(Admin_template_Url);
 
/*        
        Logger.log(func + Step + ' a - EmailTemplateName: ' + EmailTemplateName);
        Logger.log(func + Step + ' b - Email Param Row: ' + oEmailParams.TitleKeyAry[EmailTemplateName]);
        Logger.log(func + Step + ' c - Template ID: ' + Admin_template_ID);
        Logger.log(func + Step + ' d - Admin_subject: ' + Admin_subject);
        Logger.log(func + Step + ' e - Admin_from: ' + Admin_from);
        Logger.log(func + Step + ' f - Admin_to: ' + Admin_to);
        Logger.log(func + Step + ' g - Admin_bcc: ' + Admin_bcc);
        Logger.log(func + Step + ' h - Admin_Statement: ' + Admin_Statement);
        Logger.log(func + Step + ' h - Admin_Sort_Vaules_Source: ' + Admin_Sort_Vaules_Source);
        Logger.log(func + Step + ' i - test_mode: ' + test_mode);
*/
  /**************************************************************************/
  Step = 1600; /* Load the Admin_email_sort_field value vs. admin email address keyed array
                This section of the code is used to develop a keyed array of "Admin_sort_values"
                and their corresponding "admin" email address for each entry.
  ***************************************************************************/
  if (Admin_Sort_Vaules_Source) {
    Step = 1610;// Source for Admin sort values exists
    var Key_Source_Sheet = Admin_Sort_Vaules_Source;
    var Key_Heading_Term = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Key Heading Term']];
    var Value_Heading_Term = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Value Heading Term']];
    var Form_Data_Heading = oEmailParams.Parameters[oEmailParams.TitleKeyAry[EmailTemplateName]][oEmailParams.HeadingKeyAry['Form Data Key Heading']];
    var Admin_email_addr_sourceCol = GetXrefValue(oCommon, Form_Data_Heading);
    
    /******************************************************************************/
    Step = 1620; // Build the Key/Value Array from the Admin Sort Values
    /******************************************************************************/
    var TargetSheet = Admin_Sort_Vaules_Source,  // SheetContainingKeysandValue
        SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start r,ow is always the following row
        KeyRow = '',       // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
        ValueRow = '',     // number or text [Look for text Match in Col 0 (Base 0)]
        //KeyCol = Key_Heading_Term,       // number or text [Look for text Match in StartRow (Base 0)]
        //ValueCol = Value_Heading_Term,     // number or text [Look for text Match in StartRow (Base 0)]
        KeyCol = 0,       // number or text [Look for text Match in StartRow (Base 0)]
        ValueCol = 1,     // number or text [Look for text Match in StartRow (Base 0)]
        AdminKeyArray = {}; //;
    AdminKeyArray = BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol);

    // Verify
    //for(var Key in AdminKeyArray){
    //  Logger.log(func + Step + ' Key:' + Key + '  Value: ' + AdminKeyValueAry[Key]);
    //}

  }  // End of Load the Admin_email_sort_field value
  
  Step = 1700; // Set working scalar values
  var todaysDate = new Date();
  //var scanTime = String(todaysDate);
  //var scanDate = Trim(scanTime.substring(0, 10)); // "6/13/2017 13:27:05" ==> "6/13/2017"
  var scanDate = formatDate(todaysDate); // "6/13/2017 13:27:05" ==> "6/13/2017"
  var scanTime = scanDate + ' ' + formatAMPM(todaysDate);
  var oAdminMessages = [];

  /**************************************************************************/
  Step = 2000; // Process each row passsed in SelectedSheetRows array and prepare the AdminMessage array
  /**************************************************************************/
  var admin_body = '';
  var summary_body = '';
  var sentCount = 0;
  var adminCount = 0;
  var oAdminEmails = [];
  var AdminEmailAddress = '';
  if(!bSpecialReport){
    // Prepare objects for retrieving the Form Response data
    var sheet = oSourceSheets.getSheetByName(dataTab);
    var SheetData = sheet.getDataRange().getValues();
    for (var k = 0; k < oCommon.SelectedSheetRows.length; k++){
      var ArrayRow = oCommon.SelectedSheetRows[k][0] - 1;
      var Admin_email_sort_field = '';
      var AdminEmailAddress = ''; // defaults to empty email address
      if (Admin_email_addr_sourceCol) {
        Step = 2100; // Admin_email_sort_field value exists
        Admin_email_sort_field = SheetData[ArrayRow][Admin_email_addr_sourceCol]; 
        if (AdminKeyArray[Admin_email_sort_field]){
          AdminEmailAddress = AdminKeyArray[Admin_email_sort_field];
          //Logger.log(func + Step + ' Admin_email_sort_field: ' + Admin_email_sort_field + ' Assigned Addr: ' + AdminEmailAddress);
        }
      }
      // Build the oAdminMessages array of objects
      if(!oCommon.SelectedSheetRows[k][2]){oCommon.SelectedSheetRows[k][2] = '';}
      oAdminMessages.push(
        {
          SortField:  Admin_email_sort_field,                                  // 0 - Admin Sort Field Value
          Sheetrow:   oCommon.SelectedSheetRows[k][0],                                 // 1 - SheetData row Base 1
          TagValue:   oCommon.SelectedSheetRows[k][1],                                 // 2 - TagValue
          Message:    oCommon.SelectedSheetRows[k][2],                                 // 3 - message
          PriEmail:   oCommon.SelectedSheetRows[k][3],                                 // 4 - respondant email address used
          AdminEmail: AdminEmailAddress                                        // 5 - Admin email address to be used
        }
      );
    }  // iterate k / End of: Process each row passsed in SelectedSheetRows array
    
    /**************************************************************************/
    Step = 2200; // Sort the AdminMessage array
    /**************************************************************************/
    //Ref: https://stackoverflow.com/questions/1129216/sort-array-of-objects-by-string-property-value-in-javascript
    oAdminMessages.sort(function(a,b) {return (a.SortField > b.SortField) ? 1 : ((b.SortField > a.SortField) ? -1 : 0);} ); 
    
    /**************************************************************************/
    Step = 3000; // Build each Admin Message Body
    /**************************************************************************/
    //var last_Admin_value = oAdminMessages[0][0]; // set the initial value assignment 
    var last_Admin_value = oAdminMessages[0].SortField; // set the initial value assignment 
    //var admin_body = '';
    //var summary_body = '';
    //var sentCount = 0;
    //var adminCount = 0;
    //var oAdminEmails = [];
    //var AdminEmailAddress = '';
    for (var j = 0; j < oAdminMessages.length; j++){
      if (oAdminMessages[j].SortField != last_Admin_value) {
        Step = 3100; // Record and Reset the test values
        var admin_line = last_Admin_value + ': ' + Admin_Statement
        // Build the oAdminEmails array of objects
        oAdminEmails.push(
          {
            SendToAddr: oAdminMessages[j-1].AdminEmail, // 0 - Admin Send To email address
            SentCount:  sentCount,                      // 1 - Number of Alert Emails Sent
            AdminMsg:   admin_line,                     // 2 - Admin Message from AdminEmail Parameters
            AdminBody:  admin_body,                     // 3 - Admin Email main message body
            SortField:  oAdminMessages[j-1].SortField   // 4 - Sort Field
          }
        );
        
        // Reset values for the next oAdminEmails array of objects
        admin_body ='';
        admin_line ='';
        sentCount = 0;
        adminCount++;
      }
      /**************************************************************************/
      Step = 3200; /* Build the admin email message body:
      TagValue, (email address)
      Warning messages
      /***************************************************************************/
      if (oAdminMessages[j].PriEmail){
        admin_body = admin_body + '\n' + oAdminMessages[j].TagValue + ', ' 
        + ' (' + oAdminMessages[j].PriEmail + ')' + '\n' + '    ' + oAdminMessages[j].Message;
      } else {
        if(oAdminMessages[j].Message){
          admin_body = admin_body + '\n' + oAdminMessages[j].TagValue
          + '\n' + '    ' + oAdminMessages[j].Message;
        } else {
          admin_body = admin_body + '\n' + oAdminMessages[j].TagValue;
        }
      }
      
      last_Admin_value = oAdminMessages[j].SortField;
      
      sentCount++;
      
    } // End of Build each Admin Message Body
    
    /**************************************************************************/
    Step = 3300; // Capture the Admin values for the last entry
    /**************************************************************************/
    var admin_line = last_Admin_value + ': ' + Admin_Statement;
    oAdminEmails.push(
      {
        SendToAddr: oAdminMessages[j-1].AdminEmail,  // 0 - Admin Send To email address
        SentCount:  sentCount,                       // 1 - Number of Alert Emails Sent
        AdminMsg:   admin_line,                      // 2 - Admin Message from AdminEmail Parameters
        AdminBody:  admin_body,                      // 3 - Admin Email main message body
        SortField:  oAdminMessages[j-1].SortField    // 4 - Sort Field
      }
    );
    
    /**************************************************************************/
    Step = 4000; // Build the Summary Admin Email
    /**************************************************************************/
    if (oAdminEmails.length > 0){
      var summary_body = '';
      for (var n = 0; n < oAdminEmails.length; n++){
        summary_body = summary_body + '\n' + oAdminEmails[n].AdminMsg + oAdminEmails[n].AdminBody + '\n';
      }
      oAdminEmails.push(
        {
          SendToAddr: Admin_to,                // 0 - Admin Send To email address
          SentCount:  TotalSelectedSheetRows,  // 1 - Number of Alert Emails Sent
          AdminMsg:   Admin_Statement,         // 2 - Admin Message from AdminEmail Parameters
          AdminBody:  summary_body,            // 3 - Admin Email main message body
          SortField:  'Admin Summary'          // 4 - Sort Field
        }
      );
    } // End of "Build the Summary Admin Email" section
    
  } else {
    /**************************************************************************/
    Step = 4500; // Build the SpecialMsg Admin Email
    /**************************************************************************/
    // Build the single Special Report Admin Body
    oAdminEmails.push(
      {
        SendToAddr: Admin_to,                 // 0 - Admin Send To email address
        SentCount:  0,                        // 1 - Number of Alert Emails Sent
        AdminMsg:   oSpecialMsg.SpecMsgIntro, // 2 - Admin Message from AdminEmail Parameters
        AdminBody:  oSpecialMsg.SpecMsgBody,  // 3 - Admin Email main message body
        SortField:  ''                        // 4 - Sort Field
      }
    );
    Logger.log(func + Step + ' AdminMsg: ' + oSpecialMsg.SpecMsgIntro);  
  }
 
  /**************************************************************************/
  Step = 5000; // Construct and send individual Admin emails
  /**************************************************************************/
  var sentCount = 0;
  var attemptedCount = 0;
  var successCount = 0;
  var bypassed_for_errors = 0;
  for (var m = 0; m < oAdminEmails.length; m++){
    Step = 5100; // Get the template
    try {
      Step = 5110; // Prepare the Text template 
      Logger.log(func + Step + ' m = ' + m + ', EmailTemplateUrl: ' + Admin_template_Url);  
      var doc = DocumentApp.openById(Admin_template_ID);
      var body = doc.getActiveSection();
      //var email = String(body.editAsText);
      //var email = String(body.getBlob().getDataAsString());
      var AdminEmail = body.getText();
      Logger.log(func + Step + ' m = ' + m + ', EmailTemplateID: ' + Admin_template_ID);  
    }         
    catch(err) {
      oCommon.DisplayMessage = ' ERROR 364 - Unable to Build Google Doc Admin Email - Error Message: ' + err.message;
      oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon); 
      Logger.log(oCommon.ReturnMessage);
      bypassed_for_errors++;
      break;
    }
    
    Step = 5200;
    AdminEmail = AdminEmail.replace('<<Scan_Time>>', scanTime);
    AdminEmail = AdminEmail.replace('<<Sent_Count>>', oAdminEmails[m].SentCount);
    
    var admin_line = oAdminEmails[m].AdminMsg;
    if (!bSpecialReport && oAdminEmails[m].SentCount <= 0){
      admin_line = '';
    }
    
    AdminEmail = AdminEmail.replace('<<Admin_Message>>', admin_line);
    AdminEmail = AdminEmail.replace('<<Admin_Body>>', oAdminEmails[m].AdminBody);
    Admin_subject = Admin_subject.replace('<<Scan_Date>>', scanDate);
    
    //var admin_address = oAdminEmails[m][0];
    var admin_address = oAdminEmails[m].SendToAddr;
    
 Logger.log(func + Step + ' admin_address[' + m + ']: ' + admin_address);
    
    // Test for a non-empty email address
    if (admin_address){
      if (test_mode){
        admin_address = TestEmailAddress;
        Admin_cc = '';
        Admin_bcc = '';
      }  
      if (!admin_address){
        admin_address = oCommon.Globals['send_error_report_to'];
        Admin_cc = '';
        Admin_bcc = '';
      }  
      
      Step = 5300;
      //Logger.log(func + Step + ' Sending Admin Email - Error Addr: ' + oCommon.Globals['send_error_report_to']
      //          + ', Message: ' + oAdminEmails[m].AdminBody);
      var error_message = '';
      if (oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') <= -1){
        attemptedCount++;
        if (sendTheEmail(oCommon, admin_address, Admin_cc, Admin_bcc, Admin_subject, AdminEmail, error_message)) {
          Step = 5310;
          sentCount++;
          successCount++;
          Logger.log(func + Step + ' Email Sent (' + admin_address + ') Message: ' + oAdminEmails[m].AdminBody);
        } else {
          Step = 5320;
          return_message = return_message + "  Email Sending Error (" + admin_address + ") Error Message: " + err.message;
          Logger.log(func + Step + ' ' + return_message);
        }
      }
    } //Send email if admin_address is valid
  } // End of Construct and send individual Admin emails 
  
  /**************************************************************************/
  Step = 6000; // Report results
  /**************************************************************************/
  if (oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') > -1){
    var EventMsg = func + Step + " No Admin Emails sent due to TestMode setting.";
    LogEvent(EventMsg, oCommon);
  }  
  var scriptProperties = PropertiesService.getScriptProperties();
  var emails_count = Number(scriptProperties.getProperty('B_Emails_count'));
  scriptProperties.setProperty('B_Emails_count', emails_count + successCount);
  oCommon.EmailsGenerated = successCount;
  oCommon.ItemsAttempted = attemptedCount;
  oCommon.ItemsCompleted = successCount;
  oCommon.DisplayMessage = ' Send Admin Email completed ' + successCount + ' of ' 
    + attemptedCount + ' messages attempted. ';
  // Adjust the message for the number bypassed_for_freq_violation
  if (bypassed_for_errors > 0){
    if (bypassed_for_errors > 1){
      oCommon.DisplayMessage = oCommon.DisplayMessage + bypassed_for_errors
      + ' were not sent due to Google Doc Template errors.';
    } else {
      oCommon.DisplayMessage = oCommon.DisplayMessage + bypassed_for_errors
      + ' was not sent due to a Google Doc Template error.';
    }
  }
    
  oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
  
  Logger.log(func + ' END');
  
  return true;
  
}


function CMGAudit(AuditTitle, oCommon){
  /* ****************************************************************************
   DESCRIPTION:
     This function is used to import a .csv file
     
   USEAGE:
     ReturnBool = CMGAudit(AuditTitle, oCommon);

   REVISION DATE:
     11-11-2017 - Initial design
     03-23-2018 - Modified to work with v3.x
                - odified to return booleaan to facilitate completion
                    message for User feedback and error reporting.
   
   NOTES:
  ******************************************************************************/

  var func = "***CMGAudit " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  Step = 1100; // Leave immediately if no selected rows are passed
  if (typeof AuditTitle === 'undefined' || AuditTitle === null || AuditTitle === ''){
    var EventMsg = func + Step + ' Internal ERROR - Audit Title is EMPTY or NULL';
    LogEvent(EventMsg, oCommon); 
    Browser.msgBox(EventMsg);
    return false;
  }
  
  AuditTitle = Trim(AuditTitle);
  
  /****************************************************************************/
  Step = 1200; // Assign the oCommon parameters
  /****************************************************************************/
  var oSourceSheets = oCommon.Sheets, 
      dataTab = oCommon.FormResponse_SheetName,
      firstDataRow = oCommon.FormResponseStartRow,
      timestampCol = oCommon.TimestampCol,
      rptDateCol = oCommon.AsOfDateCol,
      error_email = oCommon.Send_error_report_to,
      timestampCol = oCommon.TimestampCol,
      workEmailcol = oCommon.PrimaryEmailCol,
      validDomain = oCommon.ValidDomain,
      AlertTemplateName = oCommon.AuditAlertEmailTemplate
    ;
  
  // Declare scalars and arrays
  var SourceTabName = oCommon.Sheets.getActiveSheet().getName();
  var SheetData = oCommon.Sheets.getActiveSheet().getDataRange().getValues();
  
  // Verify
  //for(var Key in oCommon.Globals){
  //   Logger.log(func + Step + ' Key:' + Key + '  Value: ' + oCommon.Globals[Key]);
  //}

  Step = 1500; // Verify OK to run audit in the active Tab
  if (SourceTabName != dataTab){
    Browser.msgBox('Error: Audit cannot be run from this Tab');
    Logger.log(func + Step + ' ERROR: Audit cannot be run from this Tab - SourceTab:' + SourceTabName 
      + ' Response Tab: ' + dataTab);
    return false;
  }
  
  /******************************************************************************/
  Step = 2000; // Get the Selected Rows
  /******************************************************************************/+
  Logger.log(func + Step + ' Got here...')
  var mode = 1; // Retrieve ALL rows
  var MsgBoxMessage = ('Do you want to Scan specific row(s)?');
  if (Browser.msgBox('Scope of Alert Scan?',MsgBoxMessage, Browser.Buttons.YES_NO) == 'yes'){
    mode = 2; // Retrieve user-selected rows
  }
  
  Step = 2100; // Leave immediately if no selected rows are found
  if(!SelectSheetRows(oCommon,mode,SourceTabName,firstDataRow)){
    Logger.log(func + Step + ' EARLY RETURN - SelectedSheetRows is EMPTY or NULL');
    Browser.msgBox("No records selected for CMG Audit." + '\\n' + oCommon.SelectedSheetRows);  
    return false;
  }
  var NumberofRecordsSelected = oCommon.SelectedSheetRows.length;
  
  /********************************************************************************/
  Step = 3000; // Audit the Form Responses
  /********************************************************************************/
  //var AuditTitle = "Verify Defensive Driving Video Date";
  var AuditFindings = [];
  if (!PerformAudit(oCommon, AuditTitle, oCommon.SelectedSheetRows, AuditFindings)){
    return false;    
  }
      
  Logger.log(func + Step + ' AuditFindings: ' + AuditFindings);
  
  if (AuditFindings.length <= 0){
    Browser.msgBox("No Discrepancies found found among " + NumberofRecordsSelected 
       + " selected records." + FormatForDisplay(oCommon.SelectedSheetRows));  
    return true;
  }

  /********************************************************************************/
  Step = 4000; // View Audit Findings and determine if Admin Emails are to be sent
  /********************************************************************************/
  if (AuditFindings.length > 0 && AuditFindings.length != null){
   
    // Construct the message body
    var MessageBody = '';
    for (m = 0; m < AuditFindings.length; m++){
      MessageBody = MessageBody + '\n' + AuditFindings[m];
    }  
    var MsgBoxMessage = ('The following Audit Results were found: ' + '\\n' + MessageBody + '\\n' 
      + '  Do you want to send individual Alert email(s)?');
    
    /********************************************************************************/
    Step = 4100; // Send Alert Emails?
    /********************************************************************************/

    if (Browser.msgBox('Send Alert Emails?',MsgBoxMessage, Browser.Buttons.YES_NO) == 'yes'){
    
      Step = 4110; //Browser.msgBox("The Response is YES");+
  
      oCommon.SelectedSheetRows = AuditFindings;

      // The code below will set the value of "personal_message" input by the user, or 'no'
      oCommon.PersonalMessage = Browser.inputBox('Add personal message?', 'Enter your message', Browser.Buttons.YES_NO);
      var personal_message = Trim(oCommon.PersonalMessage) + '\n\n';
      
      if (oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') > -1){
        var EventMsg = func + Step + " No Alert Email sent due to TestMode setting.";
        LogEvent(EventMsg, oCommon);
      } else {
        //MsgBoxMessage = SendAlertEmails(oCommon, AlertTemplateName, personal_message);
        MsgBoxMessage = SendEmails(oCommon, AlertTemplateName);
        Logger.log(func + Step + ' ' + MsgBoxMessage);
        Browser.msgBox(MsgBoxMessage);
      }
      
      Step = 4120; // Send Admin emails
      if (oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') > -1){
        var EventMsg = func + Step + " No Admin Email sent due to TestMode setting.";
        LogEvent(EventMsg, oCommon);
      } else {
        var EmailTemplateName = 'CMG Audit Notice';
        MsgBoxMessage = SendAdminEmails(oCommon, AdminTemplateName);
        Logger.log(func + Step + ' ' + MsgBoxMessage);
        Browser.msgBox(MsgBoxMessage);
      }

    } else {
      /********************************************************************************/
      Step = 4200; // Send Admin Summary Emails?
      /********************************************************************************/
      var MsgBoxMessage = 'Do you want to send Admin Summary email(s)?';
      if (Browser.msgBox('Send Admin Summary Emails?',MsgBoxMessage, Browser.Buttons.YES_NO) == 'yes'){
      
        Step = 4210; //Browser.msgBox("The Response is YES");
        
        if (oCommon.TestMode.toUpperCase().indexOf('NO EMAIL') > -1){
          var EventMsg = func + Step + " No Admin Email sent due to TestMode setting.";
          LogEvent(EventMsg, oCommon);
        } else {
          oCommon.SelectedSheetRows = AuditFindings;
          // Send Admin emails
          //SendAdminEmails(oCommon, AdminTemplateName)
          //MsgBoxMessage = 'Admin Summary Emails Sent';
          MsgBoxMessage = SendAdminEmails(oCommon, AdminTemplateName)
          Logger.log(func + Step + ' ' + MsgBoxMessage);
          Browser.msgBox(MsgBoxMessage);
        } 
      
      } else {
        Step = 4300; // No Emails Sent...
        MsgBoxMessage = 'Your response is NO - no Admin messages will be sent';
        Logger.log(func + Step + ' ' + MsgBoxMessage);
        Browser.msgBox(MsgBoxMessage);
      }
    }
  }
  Logger.log(func + ' END');
  
  return true;

}

function PerformAudit(oCommon, AuditTitle, AuditFindings){
  /* ****************************************************************************
  
   DESCRIPTION:
     This function is used to verify that entries made in the Form Responses Tab
       (FormData Tab) match corresponding entries place in another "AuditSource" Tab
       
   USEAGE:
     var AuditFindings = [];
     ReturnBool = PerformAudit((oCommon, AuditTitle,AuditFindings); 

   REVISION DATE:
    11-25-2017 - First Instance
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback
     
   STEPS
       1. Initialize
       2. Import the .csv file and place it in a designated "AuditSource" Sheet Tab (CMG_Records_Tab parameter)
       3. Prepare the CMG date data for processing
       4. Compare the selected date fields and place the "erroneous" user records 
           into the ChangedSheetRows array for subsequent processing.
           
   NOTES:
     1. SourceSS is the SpreadSheet object containing all sheets
     2. "Globals" is the keyed array containing global scalar values
     3. FormDataTab is the Tab containing the data to be used for the audit
     4. AuditTab is the Tab containing the imported data to be audited 
     4. The array oCommon.SelectedSheetRows[SheetRow][TagValue] is passed identifying the selected rows
           oCommon.SelectedSheetRows.length == 0 if an invalid range has been selected by the user     
     5. AuditParameters id the array containg the particulars for the audit scan:
          AuditParameters[DataSet][0] = Audit Title
          AuditParameters[DataSet][1] = Type of Test
          AuditParameters[DataSet][2] = Form Data Tab to be audited
          AuditParameters[DataSet][3] = Form Col(s) Used for Index
          AuditParameters[DataSet][4] = Form Test Column
          AuditParameters[DataSet][5] = Audit Data Source Tab
          AuditParameters[DataSet][6] = Audit Col(s) Used for Index
          AuditParameters[DataSet][7] = Audit Test Column
          AuditParameters[DataSet][8] = Error Message Prefix 1
          AuditParameters[DataSet][9] = Error Message Prefix 1

  ******************************************************************************/
 
  var func = "***PerformAudit " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  /****************************************************************************/
  Step = 1100; // Assign the oCommon parameters
  /****************************************************************************/
  Logger.log(func + Step + ' Got here...')
  var oSourceSheets = oCommon.Sheets, 
      formID = oCommon.GoogleForm_Url,
      ParameterSourceTab = oCommon.SetupSheetName,
      dataTab = oCommon.FormResponse_SheetName,
      HeadingsRow = oCommon.FormResponse_HeadingRow,
      FirstFormRow = oCommon.FormResponseStartRow,
      AwesomeTableRow = oCommon.AwesomeTableRow, 
      timestampCol = oCommon.TimestampCol,
      StatusCol =oCommon.Record_statusCol,
      InactiveValue = oCommon.Inactive_Status_value,
      rptDateCol = oCommon.AsOfDateCol,
      editUrlCol = oCommon.Edit_URLCol,
      editButtonCol = oCommon.Edit_ButtonCol,
      editButtonText = oCommon.Edit_Button_Text,
      error_email = oCommon.Send_error_report_to,
      timestampCol = oCommon.TimestampCol,
      ParseSymbol = oCommon.ParseSymbol
    ;
  
  /******************************************************************************/
  Step = 2000; // Get the Audit Parameters from the Setup Tab
  /******************************************************************************/
  var oAuditParameters = {};
  var ParameterSet = oCommon.AuditSection_Heading;  
  oAuditParameters = ParamSetBuilder(oCommon, ParameterSet);   
  if(oAuditParameters.length <=0){
    Browser.msgBox('Error: Audit parameters set cannot be found');  
    var EventMsg = func + Step + ' Error: Audit parameters cannot be found - Audit Section Heading: "' 
      + ParameterSet + '"';
    LogEvent(EventMsg, oCommon); 
    return false;
  }
  // Returned items:
  //   HeadingKeyAry: ColNumKeyAry,
  //   TitleKeyAry:   RowNumKeyAry,
  //   Parameters:    AllParameters
  
  var TitleCol = oAuditParameters.HeadingKeyAry['Audit Title'];  // AuditParameters[DataSet][0] = Audit Title
  var TestTypeCol = oAuditParameters.HeadingKeyAry['Type of Test'];
  var FormDataTabCol = oAuditParameters.HeadingKeyAry['Form Data Tab'];
  var FormDataIndexCol = oAuditParameters.HeadingKeyAry['Form Col(s) Used for Index'];
  var FormDataValueCol = oAuditParameters.HeadingKeyAry['Form Test Column'];
  var AuditTabNameCol = oAuditParameters.HeadingKeyAry['Audit Data Source Tab'];
  var AuditDataKeyCol = oAuditParameters.HeadingKeyAry['Audit Col(s) Used for Index'];
  var AuditDataTestValueCol = oAuditParameters.HeadingKeyAry['Audit Test Column'];
  var MsgPrefixACol = oAuditParameters.HeadingKeyAry['Error Msg Prefix 1'];
  var MsgPrefixBCol = oAuditParameters.HeadingKeyAry['Error Msg Prefix 2'];
  
  Step = 2100; // Select the Audit Parameters rows matching the input Audit Title
  var AuditParameters = [];
  var MatchingParmeterRows = [];
  for (var p = 0; p < oAuditParameters.Parameters.length; p++){
    Logger.log(func + Step + ' p: ' + p + ' Value: ' + oAuditParameters.Parameters[p][TitleCol]);
    if (Trim(oAuditParameters.Parameters[p][TitleCol]) == AuditTitle){
      // Matching parameter set found
      MatchingParmeterRows.push(p);
    }
  }
  
  Step = 2200; // Test for matching parameters found
  if (MatchingParmeterRows.length <= 0){
    // Throw error message and leave
    Browser.msgBox(' Error: Audit parameters cannot be found');  
    var EventMsg = func + Step + 'Fatal Error: Audit parameters cannot be found - AuditTitle: ' + AuditTitle;
    LogEvent(EventMsg, oCommon); 
    return false;
  }
  // Verify
  //for (var i = 0; i < AuditParameters.length; i++){
    //Logger.log(func + Step + ' AuditParameters[' + i + ']: ' + AuditParameters[i]); 
  //}
  
  /******************************************************************************/
  Step = 3000; // Build the AuditKeySets Key / Value Arrays
  /******************************************************************************/
  Logger.log(func + Step + ' MatchingParmeterRows.length: ' + MatchingParmeterRows.length); 
  var AuditKeySets = [];
  for (var i = 0; i < MatchingParmeterRows.length; i++){
    // Import new data, if desired by user
    var ParamRow = MatchingParmeterRows[i];
    var AuditDataTabName = oAuditParameters.Parameters[ParamRow][AuditTabNameCol];
    //Logger.log(func + Step + ' Tab Containing AuditParameters[' + i + ']: ' + AuditDataTabName); 
    ImportFile(oSourceSheets, AuditDataTabName);
    
    Step = 3100; // Build the AuditKeySet for AuditParameters[i] row  Key = FormData row TagValue / Value = FormDataTestHeading cell contents
    AuditKeySets[i] = [];
    var KeyHeading = oAuditParameters.Parameters[ParamRow][AuditDataKeyCol];
    var ValueHeading = oAuditParameters.Parameters[ParamRow][AuditDataTestValueCol];
    
    Logger.log(func + Step + ' ParamRow: ' + ParamRow + ' KeyCol: ' + AuditDataKeyCol + ' ValueCol: ' + AuditDataTestValueCol); 
    Logger.log(func + Step + ' ParamRow: ' + ParamRow + ' KeyHeading: ' + KeyHeading + ' ValueHeading: ' + ValueHeading); 

    Step = 3200; // Build Audit Data Source Key/Value object array
    var TargetSheet = AuditDataTabName,    // Sheet Containing Keys and Values
        SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start r,ow is always the following row
        KeyRow = '',   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
        ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
        KeyCol = KeyHeading,   // number or text [Look for text Match in StartRow (Base 0)]
        ValueCol = ValueHeading; // number or text [Look for text Match in StartRow (Base 0)]
        //oAuditKeyValueAry = {}; // if declared before entering function
    AuditKeySets[i] = [];
    AuditKeySets[i] = BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol);
   
    Step = 3200; //verify
    //for (var k in  AuditKeySets[i]){
    //  Logger.log(func + Step + ' KeySet[' + ParamRow + '] key: ' + k + ' Value: ' + AuditKeySets[i][k]);
    //}   
  }
  
  /**********************************************************************************/
  Step = 4000; // Prepare the FormData scalars, objects and Heading/Col# Key/Value array
  /**********************************************************************************/
  // Capture Values for the Tab to be audited from the first AuditParameters array entry
  ParamRow = MatchingParmeterRows[0]; // Use values from the first matching Audit parameters row
  var FormDataTabName = oAuditParameters.Parameters[ParamRow][FormDataTabCol]; 
  var FormDataIndexHeadings = oAuditParameters.Parameters[ParamRow][FormDataIndexCol]; // Audit Test Column Heading Values
  var FormDataValueHeading = oAuditParameters.Parameters[ParamRow][FormDataValueCol]; // Audit Test Column Heading Values
  var MsgPrefixA = oAuditParameters.Parameters[ParamRow][MsgPrefixACol];
  var MsgPrefixB = oAuditParameters.Parameters[ParamRow][MsgPrefixBCol];
  
  Logger.log(func + Step + ' FormDataTabName:       ' + FormDataTabName );
  Logger.log(func + Step + ' FormDataIndexHeadings: ' + FormDataIndexHeadings );
  Logger.log(func + Step + ' FormDataValueHeading:  ' + FormDataValueHeading);
  Logger.log(func + Step + ' MsgPrefixA:            ' + MsgPrefixA);
  Logger.log(func + Step + ' MsgPrefixB:            ' + MsgPrefixB);
  
  try {
    var sheet = oSourceSheets.getSheetByName(FormDataTabName);
    var FormData = sheet.getDataRange().getValues();
  } catch (err) {
    // Throw error message and leave
    var EventMsg = 'Fatal Error: Form Data Sheet (' + FormDataTabName + ') not found.';
    Browser.msgBox(EventMsg);  
    EventMsg = func + Step + EventMsg + ' Error Message: ' + err.message;
    LogEvent(EventMsg, oCommon); 
    return false;
  }

  Step = 4100; // Build Form Data Headings Key/Value object array
  var TargetSheet = dataTab,    // Sheet Containing Keys and Values
      SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start r,ow is always the following row
      KeyRow = HeadingsRow,   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
      ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
      KeyCol = '',   // number or text [Look for text Match in StartRow (Base 0)]
      ValueCol = ''; // number or text [Look for text Match in StartRow (Base 0)]
  FormDataHeadingKeySet = BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol);

  var AuditedCol = FormDataHeadingKeySet[Trim(FormDataValueHeading)];
  
  Step = 4200; // Get the oKeyParams for Building the composit Index Value
  var KeyParams = [];
  KeyParams = GetKeyParams(FormDataIndexHeadings, FormData, HeadingsRow);
  //KeyParams[part][0] = Column Number
  //KeyParams[part][1] = Number of Characters to use when forming the Key Value with the current "part"
  // Veriify
  //for (var part = 0; part < KeyParams.length; part++){
  //  Logger.log(func + Step + ' Part: ' + part + ', Col: ' + KeyParams[part][0] + ' NumChars: ' + KeyParams[part][1]);
  //}

  /******************************************************************************/
  Step = 5000; // Examine each SelectedSheetRows array row
  /******************************************************************************/
  Logger.log(func + Step + ' SelectedSheetRows.length: ' + oCommon.SelectedSheetRows.length);
  for (var count = 0; count < oCommon.SelectedSheetRows.length; count++){
    r = oCommon.SelectedSheetRows[count][0] - 1; // convert back to base 0
    var error_message = "";
    
    /******************************************************************************/
    Step = 5100; // Verifiy that the selected row is for an active record
    /******************************************************************************/
    var Record_Status = Trim(FormData[r][StatusCol].toString().toUpperCase());
    Logger.log(func + Step + ' Record_Status: ' + Record_Status + ' : ' 
            + CheckStatus('', Record_Status, oCommon));
    if (CheckStatus('', Record_Status, oCommon)){
      // Record is active
      Step = 5200; // Build the RowIndex Key Value for the current FormData row (e.g. LastNameFirstName)
      var RowIndex = '';
      for (var part = 0; part < KeyParams.length; part++){
        var AppendString = Trim(FormData[r][KeyParams[part][0]]);
        if (KeyParams[part][1] > 0){AppendString = AppendString.substr(0,KeyParams[part][1]);}
        RowIndex = RowIndex + AppendString;
      }
      Logger.log(func + Step + ' r: ' + r + ' RowIndex: ' + RowIndex);
      
      Step = 5300; // Test for a Valid RowIndex Key Value
      if (RowIndex){
        
        /******************************************************************************/
        Step = 5310; // Compare entries for each AuditParameter row
        /******************************************************************************/
        for (var i = 0; i < MatchingParmeterRows.length; i++){
          var ParamRow = MatchingParmeterRows[i];
          var AuditDataTabName = oAuditParameters.Parameters[ParamRow][AuditTabNameCol];
          Step = 5320;// Capture the audit values to be compared to the Form Values
          var KeySet = AuditKeySets[i];
          //verify
          //for (var k in KeySet){
          //  Logger.log(func + (Step+0) + ' KeySet[' + ParamRow + '] k: ' + k + ' Value: ' + KeySet[k]);
          //}
        
          Step = 5330;
          Logger.log(func + Step + ' i: ' + i + ' KeySet[' + RowIndex + ']: ' + KeySet[RowIndex]);
          
          if (KeySet[RowIndex]) { 
            var AuditSourceValue = KeySet[RowIndex];
          } else {
            var AuditSourceValue = '';
          }
          
          Step = 5340;
          var TestType = oAuditParameters.Parameters[ParamRow][TestTypeCol].toUpperCase();  
          Logger.log(func + Step + ' r: ' + r + ' RowIndex: ' + RowIndex + ' AuditSourceValue[' + ParamRow + ']: '
                     + AuditSourceValue + 'Form Value: ' + FormData[r][AuditedCol] + ' TestType: ' + TestType);
          
          /*****************************************************************************/
          Step = 6000; // Perform the specified test for the AuditSourceValue
          /*****************************************************************************/
          
          switch(TestType) {
              /*****************************************************************************/
            case "T/F":
              /*****************************************************************************/
              Step = 6100; // Test for True or False Value
              break;
              
              /*****************************************************************************/
            case "DATE":
              /*****************************************************************************/
              Step = 6200; // Check for matching dates
              if (AuditSourceValue){
                AuditSourceValue = formatDate(AuditSourceValue);
                
                if(AuditSourceValue != formatDate(FormData[r][AuditedCol])){
                  // The date values do not match
                  Step = 6210; // Record audit error
                  error_message = error_message + ' D-MM' + AuditSourceValue + '*';
                } else {
                  Step = 6220; // No ERROR - Dates match
                  error_message = error_message + ' D-OK' + AuditSourceValue + '*';
                }
              } else {
                Step = 6230; // Form entry exists; however no Audit Source date entry is present
                AuditSourceValue = '';
                error_message = error_message + ' D-NF';
              }
              
              break;
              /*****************************************************************************/
            case "INSTR":
              /*****************************************************************************/
              Step = 6300; // Check for an expected and necessary response
              break;
              
              /*****************************************************************************/
            default:
              /*****************************************************************************/
              Step = 6400; // Empty or missing TestType
              Logger.log(func + Step + ' ********** WARNING - Empty or undefined Validation Type encountered');
              //Logger.log(func + Step + ' SheetDataRow ' + SheetDataRow + ' ColTitle: ' + ColTitle + ' TestType Value: ' + TestType);
              
          } // End of Switch
          
          Logger.log(func + Step + ' ** RowIndex ' + RowIndex );
          Logger.log(func + Step + ' ** SheetDataRow ' + r );
          Logger.log(func + Step + ' ** ParamRow: ' + ParamRow);
          Logger.log(func + Step + ' ** AuditSourceValue: ' + AuditSourceValue);
          Logger.log(func + Step + ' ** Form Value: ' + FormData[r][AuditedCol] );
          Logger.log(func + Step + ' ** TestType: ' + TestType);
          Logger.log(func + Step + ' ** error_message: ' + error_message);
          
        } // End of ParamRow loop
        
        /*****************************************************************************/
        Step = 7000; // Decode and Structure Consolidated Error "Message"
        /*****************************************************************************/
        switch(TestType) {
            /*****************************************************************************/
          case "T/F":
            /*****************************************************************************/
            Step = 7100; // Test for True or False Value
            break;
            
            /*****************************************************************************/
          case "DATE":
            /*****************************************************************************/
            Step = 7200; // Check for errors
            if (error_message.indexOf('D-OK') > -1){
              Step = 7210; // All is Good - a Matching date has been found
              error_message = '';
            } else if (error_message.indexOf('D-MM') > -1){
              Step = 7220; // The Dates do not match
              var startPos = error_message.indexOf('D-MM');
              startPos = startPos + 4;
              var endPos = error_message.indexOf('*', startPos) - startPos; // loking for the end of the date demarked by "*"
              var SourceValue = error_message.substr(startPos, endPos);
              if (FormData[r][AuditedCol]){
                var form_date = formatDate(FormData[r][AuditedCol]);
              } else {
                var form_date = 'Missing';
              }
              error_message = 'Date Mismatch: ' + MsgPrefixA + ' ' + SourceValue + ' ' + MsgPrefixB + ' ' + form_date;
            } else if (error_message.indexOf('D-NF') > -1){
              Step = 7230; // The Audit Source Date entry Not Found
              error_message = MsgPrefixA + ' Not Found';       
            } else {
              Step = 7240; // Unknown error
              error_message = MsgPrefixA + ' Unknown Error: row: ' + (r+1) + ', Tag: ' + FormData[r][oCommon.RecordTagValueCol];       
            }
            break;
            
            /*****************************************************************************/
          case "INSTR":
            /*****************************************************************************/
            Step = 7300; // Check for an expected and necessary response
            break;
            
            /*****************************************************************************/
          default:
            /*****************************************************************************/
            Step = 7400; // Empty or missing TestType
            Logger.log(func + Step + ' ********** WARNING - Empty or undefined Validation Type encountered');
            //Logger.log(func + Step + ' SheetDataRow ' + SheetDataRow + ' ColTitle: ' + ColTitle + ' TestType Value: ' + TestType);
            
        } // End of Switch
        
        /*****************************************************************************/
        //Step = 8000; // Write the Consolidated error message
        /*****************************************************************************/
        Logger.log(func + Step + ' r: ' + r + error_message);
        if (error_message){
          /* Place FormData row info into the AuditFindings array using the format:
          Array: oCommon.SelectedSheetRows[n]
          oCommon.SelectedSheetRows[0] = SheetData Target row for each recipient
          oCommon.SelectedSheetRows[1] = recipient last name
          oCommon.SelectedSheetRows[2] = message / email body
          */
          Logger.log(func + Step + ' ***** Tag Value ' + FormData[r][oCommon.RecordTagValueCol]);
          Logger.log(func + Step + ' * SheetDataRow ' + r );
          Logger.log(func + Step + ' * TestType: ' + TestType);
          Logger.log(func + Step + ' * error_message: ' + error_message);
          AuditFindings.push([[r+1],[FormData[r][oCommon.RecordTagValueCol]],[error_message]]);
        }
        
      } // End of Test for a Valid RowIndex Key Value
      
    } // End of test for ACTIVE Record
      
    } // End of row (r) loop 
  
  return true;
}


function BuildMenu(oCommon){
  /* ****************************************************************************

   DESCRIPTION:
     This function is invoked as part of the "OnOpenProcedures()" set of functions and builds the top
     level user menu from parameters entered into the Setup Sheet

   USEAGE:
     ReturnBool = BuildMenu(oCommon); 

   REVISION DATE:
     01-27-2018 - First Instance
     03-23-2018 - Modified to work with v3.x
                - Modified to return bool to facilitate error reporting and User feedback
     11-03-2018 - Added a new Global scalar "CustomMenu_Title" into the Setup tab 
                - Modified Step 2000 to use the Global "CustomMenu_Title" for the Menu Title      
     11-05-2018 - Added a new Global scalar "CustomMenu2_Title" into the Setup tab 
                - Modified Step 2000 to use the Global "CustomMenu_Title" for the Menu Title      
     
   NOTES:

  ******************************************************************************/
  var func = "***BuildMenu " + Version + " - ";
  Step = 1000;
  Logger.log(func + Step + ' BEGIN');
  
  /******************************************************************************/
  Step = 2000; // Capture the BuildMenu parameters, if they exist
  /******************************************************************************/
    
  if (ParamCheck(oCommon.BuildMenuSection_Heading)){

    var menuEntries = [];
    
    // Determine the parameter set for the first menu
    if(ParamCheck(oCommon.Globals['Menu_Title'])){
      var CustomMenuTitle = oCommon.Menu_Title; 
    } else {
      var CustomMenuTitle = 'User Functions'; 
    }
    
    Step = 2100;
    var ParameterSet = oCommon.BuildMenuSection_Heading;
    var oMenuParams = ParamSetBuilder(oCommon, ParameterSet);
    Step = 2100;
    
    if (!ParamCheck(oMenuParams)){
      Step = 2110; // Errors encountered in ParamSetBuilder()
      Logger.log(func + Step + ' (' + oCommon.ReturnMessage + ')');
      return false;
    }
      
    Logger.log(func + Step + ' CustomMenuTitle: ' + CustomMenuTitle 
               + ', ParameterSet: ' + ParameterSet + ' Parameters: ' + oMenuParams.Parameters);
    
    /******************************************************************************/
    Step = 3000; // Build the Menu Object
    /******************************************************************************/
    var bMenuParamsFound = false;
    var bUsedForOnSubmit = false ;
    if (oMenuParams.Parameters.length > 1){
      for( var prow = 1; prow < oMenuParams.Parameters.length; prow++){
        Step = 3100; // Capture the Menu parameters for each row in the oMenuParams.Parameters array
        var SelectionTitle = Trim(oMenuParams.Parameters[prow][oMenuParams.HeadingKeyAry['Title']]);
        var Function = Trim(oMenuParams.Parameters[prow][oMenuParams.HeadingKeyAry['Function Name']]);
        /*Logger.log(func + Step + ' Values:' 
        + ', prow: ' + prow 
        + ', SelectionTitle: ' + SelectionTitle
        + ', Function: ' + Function);
        */
        // Ignore the special case for the parameters used only by the OnFormSubmit() function
        if (SelectionTitle.toUpperCase().indexOf('SYSTEM VALUE') <= -1){
          bUsedForOnSubmit = false;
        } else {
          bUsedForOnSubmit = true;
        }
        
        if(SelectionTitle && Function && !bUsedForOnSubmit){
          Step = 3200; // Capture the Menu parameters for each row in the oMenuParams.Parameters array
          var menuRow = {name: SelectionTitle, functionName: Function};
          menuEntries.push(menuRow);
          Logger.log(func + Step + ' Added Menu Item#: ' + prow + ' For: ' + SelectionTitle + ' / ' + Function);
          bMenuParamsFound = true;
        } else {
          // WARNING - No useable Menu parameters found
          Logger.log(func + Step + ' Bypassed parameters in Menu row (' + prow + ')'
                     + ' for Function: "' + Function + '"');
        }
      }
      
      if(bMenuParamsFound){
        Step = 3300; // Add the Menu to the Sheets object
        oCommon.Sheets.addMenu(CustomMenuTitle, menuEntries);
        Logger.log(func + Step + ' "' + CustomMenuTitle + '" Menu added to the Sheets object.');
        Logger.log(func + Step + ' END');
        return true;
      } else {
        Step = 3400; //WARNING - No Menu parameters found in the Menu Parameter section
        oCommon.DisplayMessage = ' Fatal Error 020 - No Menu parameters found in the Menu Parameter section.';
        oCommon.ReturnMessage =  func + Step + oCommon.DisplayMessage;
        LogEvent(oCommon.ReturnMessage, oCommon); 
        Logger.log(oCommon.ReturnMessage);
        return false; 
      }
    } else {
      Step = 3400; // Expected Menu Parameter section not found, 
      oCommon.DisplayMessage = ' Fatal Error 021 - Expected Menu Parameter section not found.';
      oCommon.ReturnMessage =  func + Step + oCommon.DisplayMessage;
      LogEvent(oCommon.ReturnMessage, oCommon); 
      Logger.log(oCommon.ReturnMessage);
      return false; 
    }
  } else {
    Step = 4000; // No menu section defined
    var EventMsg = func + Step + ' WARNING - No MenuParameter section defined'; 
    //LogEvent(EventMsg, oCommon); 
    Logger.log(EventMsg);
    return true;    
  }
}

function GetMenuParams(oCommon, CalledFunction, oMenuParams){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to retrieve the parameters associated with a specific 
       function call

   USEAGE:
     ReturnBool = GetMenuParams(oCommon, CalledFunction, aryParams); 

   REVISION DATE:
     02-12-2019 - First Instance
     
   NOTES:
   
   oMenuParams = {} must be declared externally

  ******************************************************************************/
  var func = "***GetMenuParams " + Version + " - ";
  var Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
 /******************************************************************************/
  Step = 1000; // Capture the BuildMenu parameters
  /******************************************************************************/
  var oBuildMenuParams = {};
  var ParameterSet = oCommon.BuildMenuSection_Heading;  
  var oBuildMenuParams = ParamSetBuilder(oCommon, ParameterSet);
      /*RETURNS Object
        oMenuParams = {
          HeadingKeyAry: ColNumKeyAry,
          TitleKeyAry:   RowNumKeyAry,
          Parameters:    AllParameters
        } */
  
  if (!ParamCheck(oBuildMenuParams)){
    // Error encountered in ParamSetBuilder()
    Logger.log(func + Step + ' (' + oCommon.ReturnMessage + ')');
    return false;
  }
  
Logger.log(func + Step + ' CalledFunction: ' + CalledFunction 
   + ', Parameters: ' + oBuildMenuParams.TitleKeyAry[CalledFunction]);
  
  /******************************************************************************/
  Step = 2000; // Find the parameters to be returned
  /******************************************************************************/
  if (oBuildMenuParams.Parameters.length > 1){
    Step = 2100; // Find the Menu Section row containing the desired parameters
    var prow = 0;
    var paramVal = '';
    var menuTitle = '';
    var functionName = '';
    var bMatchFound = false;
    for (var key in oBuildMenuParams.TitleKeyAry) {
      //Logger.log('Key: %s, Value: %s', key, oBuildMenuParams.TitleKeyAry[key]);
      prow = oBuildMenuParams.TitleKeyAry[key];
      paramVal = oBuildMenuParams.Parameters[prow][oBuildMenuParams.HeadingKeyAry['Function Name']];
      //Logger.log(func + Step + ' prow: ' + prow + ', paramVal: ' + paramVal);
      if(Trim(CalledFunction.toUpperCase()) == Trim(paramVal.toUpperCase())){
        // Matching Function Name found
        bMatchFound = true;
        menuTitle = oBuildMenuParams.Parameters[prow][oBuildMenuParams.HeadingKeyAry['Title']];
        functionName = paramVal;
        break;
      }
    } 
  
    Step = 2200; // Search complete
    Logger.log(func + Step + ' bMatchFound: '+ bMatchFound + ' , prow: ' + prow 
               + ', menuTitle: ' + menuTitle + ', FunctionName: ' + functionName);

    /******************************************************************************/
    Step = 3000; // Find the parameters to be returned
    /******************************************************************************/
    if (bMatchFound){
      Step = 3100; // Matching function name found
      // Find the column containing the Keys
      // Verify results
      //Logger.log(func + Step + ' Verify oBuildMenuParams.TitleKeyAry: ');
      //for (var key in oBuildMenuParams.HeadingKeyAry) {
      //  Logger.log('Key: %s, Value: %s', key, oBuildMenuParams.HeadingKeyAry[key]);
      //}  
      var KeysCol = oBuildMenuParams.HeadingKeyAry['Parameter Keys'];
      var ValuesCol = oBuildMenuParams.HeadingKeyAry['Parameter Values'];

      //Logger.log(func + Step + ' bMatchFound: '+ bMatchFound + ' , prow: ' + prow 
      //           + ', KeysCol: ' + KeysCol + ', ValuesCol: ' + ValuesCol);
      //Logger.log(func + Step + ' param_keys: "'+ oBuildMenuParams.Parameters[prow][KeysCol] 
      //           + '", param_values: "' + oBuildMenuParams.Parameters[prow][ValuesCol] + '"');
      
      Step = 3200; // Get and split the Keys
      var param_keys = []
      param_keys = oBuildMenuParams.Parameters[prow][KeysCol].split("/");
      
      Step = 3300; // Get and split the Values
      var param_values = []
      param_values = oBuildMenuParams.Parameters[prow][ValuesCol].split("/");
      
      if (!ParamCheck(param_keys[0])){
        Step = 3350// ERROR - Menu Parameters Inactive or Not Found
        oCommon.DisplayMessage = ' Fatal Error 014 - Menu Parameters Not Found executing function "'  
          + CalledFunction + '".';
        oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
        LogEvent(oCommon.ReturnMessage,oCommon);
        Logger.log(oCommon.ReturnMessage);
        return false;
      }

      Step = 3400; // Create the Key / Value array
      
      oMenuParams["Menu Title"] =  menuTitle;
      oMenuParams["Function Name"] = functionName;
      
      for (var i = 0; i < param_keys.length; i++){
        if(!ParamCheck(param_values[i])){ param_values[i] = ''; }
        oMenuParams[Trim(param_keys[i])] = Trim(param_values[i]);
      }
      Step = 3450; // Verify results
      //Logger.log(func + Step + ' Verify oMenuParams: ');
      for (var key in oMenuParams) {
        Logger.log(func + Step + ' Key: %s, Value: %s', key, oMenuParams[key]);
      }  
      
      return true;
      
    } else {
      Step = 4000;// WARNING - Called Function Inactive or Not Found
      var error_message = ' Fatal Error 015 - Called Function Inactive or Not Found ('
        + CalledFunction + ')';
    }
  } else {
    Step = 5000;// WARNING - Menu Parameters Inactive or Not Found
    var error_message = ' Fatal Error 016 - Menu Parameters Inactive or Not Found ('
      + CalledFunction + ')';
  }
  
  // WARNING - Menu Parameter Set Not Found
  oCommon.ReturnMessage = func + Step + error_message;
  oCommon.DisplayMessage = error_message;

  Logger.log(oCommon.ReturnMessage);
  
  return false;
}



function RestoreRows(oCommon) {
  /* ****************************************************************************
  DESCRIPTION:
     This function is used to restore row entries by finding the associated Google Form
     Response and writing the entries back into the Form Response sheet.
     
  USEAGE:
     ReturnBool = RestoreRow(oCommon);
  
  REVISION DATE:
     03-11-2019 - Initial implementation


  NOTES:
  ******************************************************************************/
  
  var func = "**RestoreRow " + Version + " - ";
  var Step = 100;
  
  /******************************************************************************/
  Step = 1000; // Initialize variables
  /******************************************************************************/
  var ScriptUser = oCommon.ScriptUser;
  var bActiveUserEmpty = oCommon.bActiveUserEmpty;
  var bNoErrors = true;
  var title = "Restore Rows Utility";
  var prog_message = 'Determining scope of work...';
  progressMsg(prog_message,title,3);
  
  Step = 1100; // Assign Variable Values
  var source_TabName = oCommon.FormResponse_SheetName;
  var source_HeadingRow = oCommon.FormResponse_HeadingRow;
  var source_FirstDataRow = oCommon.FormResponseStartRow;
  //var source_TimestampCol = oCommon.TimestampCol;
  var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  Step = 1200; // Prepare objects 
  var oSourceSheets = oCommon.Sheets;
  var sheet = oSourceSheets.getSheetByName(source_TabName);
  var SheetData = sheet.getDataRange().getValues();
  var onSubmit_Sensitivity = oCommon.onSubmit_Sensitivity / 1000;
  
  Step = 1300; // Prepare the Heading/Col# Key/Value array
  var TargetSheet = source_TabName, // Sheet Containing Keys and Values
    SectionTitle = '', // Look for Match in Col 0 (Base 0); if '', Do Not look for a start row; Start row is always the following row
    KeyRow = source_HeadingRow,   // number or text [Look for text Match in Col 0 (Base 0)] or the row following the SectionTitle row
    ValueRow = '', // number or text [Look for text Match in Col 0 (Base 0)]
    KeyCol = '',   // number or text [Look for text Match in StartRow (Base 0)]
    ValueCol = '', // number or text [Look for text Match in StartRow (Base 0)]
    oHeadingsXrefAry = {}; // if declared before entering function BuildKeyValueAry()
  //oHeadingsXrefAry = BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol,oHeadingsXrefAry)
  oHeadingsXrefAry = BuildKeyValueAry(oSourceSheets,TargetSheet,SectionTitle,KeyRow,ValueRow,KeyCol,ValueCol);
  Step = 1245; // Verify
  //for(var key in oHeadingsXrefAry){
  // Logger.log(func + Step + ' Key: ' + key + '  Value: ' + oHeadingsXrefAry[key]);
  //}
  
  /******************************************************************************/
  Step = 2000; // Get the User's selected rows
  /******************************************************************************/
  if(SpreadsheetApp.getActiveSheet().getName() !== source_TabName){
    Step = 2100; // Leave if the active sheet is not the Form Response dataTab
    oCommon.DisplayMessage = ' Sorry, this function will only work in the "' + source_TabName 
      + '" tab that contains the GoogleForm responses.';
    if (!oCommon.bSilentMode){
      Browser.msgBox(oCommon.DisplayMessage);
    }
    oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
    LogEvent(oCommon.ReturnMessage, oCommon); 
    Logger.log(oCommon.ReturnMessage);
    return true;
  }
  
  Step = 2200; // Get the selected rows to be fixed
  var mode = 2; // get user-selected rows
  if(!SelectSheetRows(oCommon,mode,source_TabName,source_FirstDataRow)){
    // WARNING - No Selected Rows found
    return_message = func + Step + ' ********** WARNING - No Rows Selected';
    Logger.log(return_message);
    return true;
  }
  
  Step = 2200; // Take time to confirm that the user wants to proceed
  if (!oCommon.bSilentMode){
    var MsgBoxMessage = ('The following rows were selected: ' + FormatForDisplay(oCommon.SelectedSheetRows)
                         + '\\n' + ' Do you want to continue?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      var display_message = "Your response is NO - procedure will be terminated";
      return_message = func + Step + ' ********** WARNING - User terminated after rows were selected';
      oCommon.ReturnMessage = return_message;
      oCommon.DisplayMessage = display_message;
      // Communicate with user
      prog_message = 'User chooses not to continue.';
      progressMsg(prog_message,title,3);
      return false;
    }
  }
  
  var total_completed = 0;
  var total_attempted = oCommon.SelectedSheetRows.length;
  Logger.log(func + Step + ' SelectedSheetRows.length: ' + total_attempted );

  /******************************************************************************/
  Step = 3000; // Get the Form Responses to acquire the Timestamp values
  /******************************************************************************/
  var formID = getIdFrom(oCommon.GoogleForm_Url);
  var responses = {};
  var timestamps = [];
  var urls =[];
  var ResponderEmails = [];
  
  try {
    var form = FormApp.openById(formID); // formID assigned from oCommon
    //var responses = form.getResponses();
    var formResponses = form.getResponses();
    for (var i = 0; i < formResponses.length; i++) {
      timestamps.push(formResponses[i].getTimestamp().setMilliseconds(0));
      urls.push(formResponses[i].getEditResponseUrl());
      ResponderEmails.push(formResponses[i].getRespondentEmail());
      Logger.log(func + Step + ' i: ' + i + ', Timestamp: ' + timestamps[i]);
    }
    
  } catch(err) {
    // do not have permissions
    // Error condition encountered retreiving the GoogleForm response data
    //form.setCustomClosedFormMessage(defaultClosedFor);
    //form.setAcceptingResponses(true);
    //lock.releaseLock();
    oCommon.DisplayMessage = ' Error 424 - Error condition encountered retreiving '
      + 'the GoogleForm response data: ' + err.message;
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    Logger.log(func + Step + oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
  
  /**************************************************************************/
  Step = 4000; // Loop through the selected rows and search for a matching Form Timestamp
  /**************************************************************************/
  for (var k = 0; k < oCommon.SelectedSheetRows.length; k++){
    Step = 4100; // Get the Sheet Timestamp value
    var source_row = oCommon.SelectedSheetRows[k][0] - 1; // convert back to Base 0
    
    var prog_message = 'Processing row ' + (source_row+1) + ' (' + oCommon.SelectedSheetRows[k][1] + ')';
    progressMsg(prog_message,title,-3);
    
    var SheetTimestamp = SheetData[source_row][oCommon.TimestampCol];
    Logger.log(func + Step + ' k: ' + k + ', source_row: ' + source_row 
      + ', SheetTimestampCol: ' + oCommon.TimestampCol
      + ', SheetTimestamp: ' + SheetTimestamp);
    var bResponseTimeStampFound = false;
    Step = 4200; // Look for the closest Form Response Timestamp value
    if (ParamCheck(SheetTimestamp)){
      for (var i = 0; i < formResponses.length; i++) {
        var dF = timestamps[i];
        var dS = new Date(SheetTimestamp).getTime();
        var mDiff = Math.abs((dF - dS)/1000);  // time difference in seconds
        
        Logger.log(func + Step + ' dF: ' + dF + ', dS: ' + dS + ', mDiff: ' + mDiff 
                   + ', Sensitivity: ' + onSubmit_Sensitivity);
        
        Step = 4300; // Test to determine if difference is within the "sensitivity" range
        if(mDiff <= onSubmit_Sensitivity){         
          Step = 4310; // Matching Form Response found, Loop throught the Form Item Response values and place into the spreadsheet
          var formResponse = formResponses[i];
          var itemResponses = formResponse.getItemResponses();
          for (var j = 0; j < itemResponses.length; j++) {
            var itemResponse = itemResponses[j];
            var itemTitle = itemResponse.getItem().getTitle();
            var itemEntry = itemResponse.getResponse();
            var source_col = oHeadingsXrefAry[itemTitle];
            if (!isNaN(source_col)){
              // Update Cell value
              var CellAddr = numToA(source_col+1) + String(source_row+1).split(".")[0]; 
              sheet.getRange(CellAddr).setValue(itemEntry);
            }
          }
          Step = 4320; // Find and Format the "Timestamp" value and write to the sheet
          /******************************************************************************/
          Logger.log(func + Step + ' TimestampCol: ' + oCommon.TimestampCol);
          if(!isNaN(oCommon.TimestampCol)){
            oCommon.FormTimestamp = Utilities.formatDate(new Date(dF), timeZone, "M/d/yyy HH:mm:ss");
            var CellAddr = numToA(oCommon.TimestampCol+1) + String(source_row+1).split(".")[0]; 
            sheet.getRange(CellAddr).setValue(oCommon.FormTimestamp); 
          }
 
          Step = 4330; // Find and Format the "EditURLS" value and write to the sheet
          /******************************************************************************/
          Logger.log(func + Step + ' Edit_URLCol: ' + oCommon.Edit_URLCol);
          if(!isNaN(oCommon.Edit_URLCol)){
            var editURL = urls[i];
            var CellAddr = numToA(oCommon.Edit_URLCol+1) + String(source_row+1).split(".")[0]; 
            if (editURL != ''){ sheet.getRange(CellAddr).setValue(editURL); }
          }
          Logger.log(func + Step + ' Edit_URL: ' + editURL);
        
          Step = 4340; // Write the FormOwnerEmail if missing (first submittal?) 
          if(!isNaN(oCommon.Form_Owner_Col)){
            Step = 4342; // Get the previously stored FormOwnerEmail value
            var FormOwnerEmail = SheetData[source_row][oCommon.Form_Owner_Col];
            if (FormOwnerEmail != ''){
              Step = 4344; // Construct the FormOwnerEmail cell address and write the value retrieved from the response
              var ResponderEmail = ResponderEmails[i];
              var FormOwnerEmailAddr = numToA(oCommon.Form_Owner_Col+1) + String(source_row+1).split(".")[0];
              if (ParamCheck(ResponderEmail)){ sheet.getRange(FormOwnerEmailAddr).setValue(ResponderEmail); }
              //Logger.log(func + Step + ' >>>>>>> FormOwnerEmail - SheetRow: ' + (ArrayRow+1) +  ' CellAddress: ' 
              //  + FormOwnerEmailAddr + ' Value: ' + FormOwnerEmail + ' bActiveUserEmpty' + oCommon.bActiveUserEmpty);
            }
          }             
          bResponseTimeStampFound = true;
          break;
        } // Matching Timestamp NOT Found
      } // Examine the next Form Response Timestamp value
    
      Step = 4400; // Execute the assignEditUrls function that should have been done onSubmit()
      if (bResponseTimeStampFound){
        var bSetDuplicateTrap = false;
        if(assignEditUrls(oCommon, bSetDuplicateTrap)){
          var EventMsg = func + Step + ' Successful re-Submit for row ' + (source_row+1) + ' (' 
            + oCommon.SelectedSheetRows[k][1] + ').';
          LogEvent(EventMsg, oCommon);
          Logger.log(EventMsg);
          SpreadsheetApp.flush();
          total_completed++;
        } else {
          var EventMsg = func + Step + ' Recovery FAILED for row ' + (source_row+1) + ' (' 
          + oCommon.SelectedSheetRows[k][1] + ').';
          LogEvent(EventMsg, oCommon);
          Logger.log(EventMsg);
        }
      } else {
        Step = 4450; // Mark the Timestamp col as "Not Found"
        if(!isNaN(oCommon.TimestampCol)){
          var CellAddr = numToA(oCommon.TimestampCol+1) + String(source_row+1).split(".")[0]; 
          sheet.getRange(CellAddr).setValue('NOT FOUND'); 
        }
        var EventMsg = func + Step + ' Timestamp NOT FOUND for row ' + (source_row+1) + ' (' 
          + oCommon.SelectedSheetRows[k][1] + ').';
        LogEvent(EventMsg, oCommon);
        Logger.log(EventMsg);
      }
      
    } else {
      Step = 4500; // Throw an error message because the row's timestamp value is not valid
      if(!isNaN(oCommon.TimestampCol)){
        var CellAddr = numToA(oCommon.TimestampCol+1) + String(source_row+1).split(".")[0]; 
        sheet.getRange(CellAddr).setValue('NOT VALID'); 
      }
      var EventMsg = func + Step + ' Re-Submit FAILED for row ' + (source_row+1) + ' (' 
      + oCommon.SelectedSheetRows[k][1] + ') due to invalid Timestamp ( ' + SheetTimestamp + '.';
      LogEvent(EventMsg, oCommon);
      Logger.log(EventMsg);
    }
  } // Next Selected Sheet Row

  /**************************************************************************/
  Step = 5000; // Report metrics and go home...
  /**************************************************************************/
  oCommon.ItemsAttempted = total_attempted;
  oCommon.ItemsCompleted = total_completed;
  oCommon.DisplayMessage = ' ' + total_completed + ' Re-submit actions completed'
    + ' out of ' + total_attempted + ' actions attempted. ';
  oCommon.ReturnMessage = func + Step + oCommon.DisplayMessage;
  LogEvent(oCommon.ReturnMessage, oCommon);
  Logger.log(oCommon.ReturnMessage);

  Logger.log(func + ' END');
  
  return true;
  
}



function GetLibraryStats(){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to gather pertinant Script Library data for the Top Level
       App Status Report
     
   USEAGE
     var oLibraryStats = {
       TrapCount: 0,
       OpensCount: 0,
       SubmissionsCount: 0
     };
     oLibraryStats = GetLibraryStats();

   REVISION DATE:
     04-17-2019 - First Instance
     
   NOTES:

  ******************************************************************************/
  var func = "**GetLibraryStatus " + Version + " - ";
  Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  Step = 1000; // Get Script Library Properties and Declare the scalars
  var scriptProperties = PropertiesService.getScriptProperties();
 
  Step = 1100; // Declare the oStats Return object
  var oLibraryStats = {
    TrapCount: Number(scriptProperties.getProperty('B_Trap_count')),
    OpensCount: Number(scriptProperties.getProperty('B_Opens_count')),
    SubmissionsCount: Number(scriptProperties.getProperty('B_Submission_count')),
    DocsCount: Number(scriptProperties.getProperty('B_Docs_count')),
    EmailsCount: Number(scriptProperties.getProperty('B_Emails_count')),
    LastCountersResetDate: scriptProperties.getProperty('Last_Counters_Reset_Date')
  };
                             
  return oLibraryStats;
}
                             

function ResetLibraryCounters(oCommon){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to reset the Script Library data for the Top Level
       App Status Report
     
   USEAGE
     bReturnStatus = SendAppStatusReport(oCommon);

   REVISION DATE:
     04-17-2019 - First Instance
     
   NOTES:

  ******************************************************************************/
  var func = "**ResetLibraryCounters " + Version + " - ";
  Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  if (!oCommon.bSilentMode){
    Step = 1000; // Take time to confirm that the user wants to proceed
    var MsgBoxMessage = ('This utility will reset ALL of the Script Library counters to zero.' + '\\n'
                         + 'Doing so will also prompt any affiliated scrirpts to reset their App Status Report counters, as well.'
                         + '\\n\\n' + ' Do you want to continue?');
    if (Browser.msgBox('CONTINUE?',MsgBoxMessage, Browser.Buttons.YES_NO) !== 'yes'){
      display_message = "Your response is NO - procedure will be terminated";
      return_message = func + Step + ' User terminated the Reset Counters procedure. No counters have bee reset.';
      oCommon.ReturnMessage = return_message;
      oCommon.DisplayMessage = display_message;
      // Communicate with user
      prog_message = 'User chooses not to continue.';
      progressMsg(prog_message,'Reset Script Library Counters',3);
      return false;
    }
  }
  
  Step = 2000; // Get Script Library Properties and reset the counters
  var scriptProperties = PropertiesService.getScriptProperties();
 
  Step = 2100; // Declare the oStats Return object
  scriptProperties.setProperty('B_Trap_count', 0.0);
  scriptProperties.setProperty('B_Opens_count', 0.0);
  scriptProperties.setProperty('B_Submission_count', 0.0);
  scriptProperties.setProperty('B_Docs_count', 0.0);
  scriptProperties.setProperty('B_Emails_count', 0.0);
  scriptProperties.setProperty('Last_Counters_Reset_Date', new Date());

  return;
}
                             

function SetAppMessage(oCommon, App_Id, message){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to set a unique Script Library property that is accessable by
       specific applications using the library.
     
   USEAGE
     bReturnStatus = SetAppMessage(oCommon, message);

   REVISION DATE:
     04-24-2019 - First Instance
     
   NOTES:

  ******************************************************************************/
  var func = "**SetAppMessage " + Version + " - ";
  Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  Step = 1000; // Validate the parameters
  if (!ParamCheck(message)){
    oCommon.DisplayMessage = ' Error 440 - Missing "message" paramater.';
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    Logger.log(func + Step + oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
  
  Step = 1100; // Validate the parameters
  if (!ParamCheck(App_Id)){
    oCommon.DisplayMessage = ' Error 442 - Missing "App_Id" paramater.';
    oCommon.ReturnMessage  = func + Step + oCommon.DisplayMessage;
    Logger.log(func + Step + oCommon.ReturnMessage);
    LogEvent(oCommon.ReturnMessage, oCommon);
    return false;
  }
  
  Step = 2000; // Get Script Library Properties and set the message scalar
  var scriptProperties = PropertiesService.getScriptProperties();
 
  Step = 2100; // Declare the oStats Return object
  scriptProperties.setProperty(App_Id, message);

  return true;
}
                             

function GetAppMessage(oCommon){
  /* ****************************************************************************

   DESCRIPTION:
     This function is used to set a unique Script Library property that is accessable by
       specific applications using the library.
     
   USEAGE
     message = GetAppMessage(oCommon);

   REVISION DATE:
     04-24-2019 - First Instance
     
   NOTES:

  ******************************************************************************/
  var func = "**GetAppMessage " + Version + " - ";
  Step = 100;
  Logger.log(func + Step + ' BEGIN');
  
  Step = 2000; // Get Script Library Properties and set the message scalar
  var scriptProperties = PropertiesService.getScriptProperties();
 
  var message = scriptProperties.getProperty(oCommon.App_Id);
  
  //Reset the message property
  scriptProperties.setProperty(oCommon.App_Id, '');
  
  return message;

}
