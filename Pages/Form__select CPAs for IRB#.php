<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function DoAddAgendaCPA(){
//On //Error //GoTo //err_DoAddCPA;
;
//If //IsNull(Me.Other_1) //Or //Me.Other_1 = " " //Then //Me.Other_1 = " ";
//If //IsNull(Me.condition1) //Then //Me.condition1 = " ";
//If //IsNull(Me.Condition2) //Then //Me.Condition2 = "  ";
//'If Me.Condition2 = "  " Then
 //'   msgbox "Hit on blanks"
//'Else
 //'   msgbox Me.Condition2
//'End If
;
//'CPAs for exempt studies
//If //Me.Study_Status //Like "*exem*" //Then;
    //If //AddAgendaRecord(Me.[IRB#], //Me.tCPAnumber, //_;
        //Me.MeetingDate, "CPA"//, "Exempt"//, //Me.Category, //Me.condition1, //_;
        //Me.Other_1, //_;
        "z "//, " z"//, "z "//, " z"//, "z "//, //0, //Date, //Me.CheckExpedite, //Forms![newmenu]![UserName], //_;
        //Forms![newmenu]![Location], //_;
        //IIf(IsNull(Me.Date_of_Change), "Not Dated; "//, "Date of CPA Change-" //& //Me.Date_of_Change //& "; "//) //& //_;
        "-" //& //Forms![_select //cpas //for //irb#].Condition2, " z"//, "z "//, //0, //Forms![_select //cpas //for //irb#].txtcyberrefnum) //Then;
        //Exit //Sub;
    //End //If;
//End //If;
//If //Me.Expedite //Then   //' expedited
    //'INSERT Expedite INTO [agenda records] ( [IRB#], SourceNumber, MeetingDate, [Record Type], [Agenda Category], [Agenda Condition1], [Agenda Condition2], UserName, IRBCode, AgendaReportString, [Agenda Explanation] )
    //If //AddAgendaRecord(Me.[IRB#], //Me.tCPAnumber, //_;
        //Me.MeetingDate, "CPA"//, "Expedited"//, //Me.Category, //Me.condition1, //_;
        //Me.Other_1, //_;
        "z "//, " z"//, "z "//, " z"//, "z "//, //0, //Date, //Me.CheckExpedite, //Forms![newmenu]![UserName], //_;
        //Forms![newmenu]![Location], //_;
        //IIf(IsNull(Me.Date_of_Change), "Not Dated; "//, "Received " //& //Me.[Date //Received //by //IRB //Co-Ord] //& "; " //& "Expedited Date-" //& //Me.Date_of_Change //& "; "//) //& //_;
        //IIf(IsNull(Me.CPAAction), "Pre Meeting Action-None"//, "Pre-Meeting Action-" //& //Me.CPAAction), " z"//, "z "//, //0, //Forms![_select //cpas //for //irb#].txtcyberrefnum) //Then;
        //Exit //Sub;
    //End //If;
//End //If;
;
//'INSERT REGULAR INTO [agenda records] ( [IRB#], SourceNumber, MeetingDate, [Record Type], [Agenda Category], [Agenda Condition1], [Agenda Condition2], UserName, IRBCode, AgendaReportString, [Agenda Explanation] )
//If //AddAgendaRecord(Me.[IRB#], //Me.tCPAnumber, //_;
        //Me.MeetingDate, "CPA"//, //Me.Category, //Me.condition1, //Me.Condition2, //_;
        //Me.Other_1, //_;
        "z "//, " z"//, "z "//, " z"//, "z "//, //0, //Date, //Me.CheckExpedite, //Forms![newmenu]![UserName], //_;
        //Forms![newmenu]![Location], //_;
        //IIf(IsNull(Me.Date_of_Change), "Not Dated; "//, "Received " //& //Me.[Date //Received //by //IRB //Co-Ord] //& "; " //& "Date of Change " //& //Me.Date_of_Change //& "; "//) //& //_;
        //Forms![_select //cpas //for //irb#].Condition2, " z"//, "z "//, //0, //Forms![_select //cpas //for //irb#].txtcyberrefnum) //Then;
//End //If;
;
//exit_DoAddCPA:;
//Exit //Sub;
//err_DoAddCPA:;
//MsgBox "Error trying to add CPA Agenda Item--" //& //Err.Description;
//Resume //exit_DoAddCPA;
}
function CheckFULetters(){
//On //Error //GoTo //error_fu;
//Me.Visible = //False;
location.href = "Form_OpenFULetters.php";//, //acNormal, //, //, //, //acDialog;
;
//Exit_checkfu:;
//Exit //Sub;
//error_fu:;
//MsgBox //Err.Description //& " " //& //Err.Number;
//Resume //Exit_checkfu;
;
}
function PlaceOnAgenda(){
//On //Error //GoTo //error_Agenda;
var sqlSTR = '';
//could initialize db here
//could initialize rs here
//Set db = CurrentDb;
;
//Me.Category = //Me.CPATypeList.Column(1);
//Me.condition1 = //Me.CPATypeList;
//If //Me.CpaActions = "(No Action)" //Then;
    //Me.Condition2 = "  ";
//Else;
    //Me.Condition2 = //Me.CpaActions.Column(0);
//End //If;
//If //Me.CheckExpedite //Then;
    //Me.Condition2 = "Expedited";
//End //If;
//If //Me.Study_Status //Like "*exem*" //Then;
    //Me.Condition2 = "Exempt";
//End //If;
 ;
;
//HoldIRBNo = //Me.studyNo;
//DoAssignAgain:;
;
location.href = "Form_ReassignAgendaRecord.php";//, //, //, //, //, //acDialog;
;
//If //ReassignSwitch = //True //Then;
        //Me.MeetingDate = //HoldMeetingDate;
        //If //CheckIfOnCPA = //True //Then //GoTo //DoAssignAgain;
        //DoCmd.SetWarnings //False;
        //'If Me.condition1 Like "*report" Then  changed to below 6/20/01 and presupposes that the maintenance form for CPA types
        //'has been changed to prevent user from changing the Final or Progress report
   //If //Me.Category //Like "*report" //Or //Me.Category //Like "*renewal" //Then;
       //Forms![newmenu]![HoldSource] = "Renewal";
        //On //Error //GoTo //error_Deleting  //'CHLA
        sqlSTR = "Select [agenda records].[IRB#], [agenda records].MeetingDate, [agenda records].SourceNumber, [agenda records].[Record Type] FROM [agenda records]" //_;
        //& " WHERE [agenda records].[IRB#]=" //& "'" //& //[Forms]![_select //cpas //for //irb#]![studyNo] //& "' " //_;
        //& " AND  [agenda records].[record type]= 'Renewal'" //_;
        //& " AND [agenda records].[MeetingDate]= " //& "#" //& //[Forms]![_select //cpas //for //irb#]![MeetingDate] //& "#";
        //'MsgBox sqlSTR
        //Set rs = db.OpenRecordset(sqlSTR);
;
        //If //Not //rs.RecordCount //> //0 //Then;
            //'MsgBox "oRIGINAL NOT FOUND"
        //Else;
            //rs.MoveFirst;
            //Do //While //Not //rs.EOF;
                //rs.Edit;
                //rs.Delete;
                //rs.MoveNext;
            //Loop;
        //End //If;
        //rs.Close;
        //db.Close;
    //End //If;
        //Call //DoAddAgendaCPA;
        //'MsgBox ("Study #- " & HoldIRBNo & "  Assigned to the--" & HoldMeetingDate & "  Meeting")
        //'Forms![_irb input form]![AgendaDate] = Me.MeetingDate
//Else;
    //If //YesNo("No //agenda //item //will //be //recorded.  //Do //you //wish //to //Save //the //CPA //?") //Then;
        //Exit //Sub;
    //Else;
        //CancelUpdate = //True;
                //Me.Undo;
    //End //If;
    ;
//End //If;
//exit_Agenda:;
//Exit //Sub;
//error_Agenda:;
 //MsgBox //Err.Number //& "  Error trying to place CPA on Agenda. " //& //Err.Description;
        //Me.Undo;
        //CancelUpdate = //True;
        //Resume //exit_Agenda;
       ;
;
//error_Deleting:;
 //'MsgBox Err.Number & "  Error trying to delete Original Renewal Agenda Record. " & Err.Description
        //Me.Undo;
        //CancelUpdate = //True;
        //Resume //exit_Agenda;
}
function Combo508_BeforeUpdate(Cancel){
//End //Sub;
;
;
;
//Private //Sub //CheckExpedite_Click();
//If //Me.CheckExpedite = //True //Then;
    //If //Me.Study_Status //Like "*exem*" //Then;
        //MsgBox "CPAs for Exempt Protocols will appear in the Exempt section on the Agenda if the 'Place on Agenda box' is Checked";
        //Me.Expedite = //False;
    //Else;
            //Me.on_agenda_label.ForeColor = //vbRed;
            //Me.Place_on_agenda = //True;
    //End //If;
//Else;
    //If //Me.Study_Status //Like "*exp*" //Then;
        //If //YesNo("This //Protocol //has //an //Expedite //Status.  //Do //you //want //to //expedite //this //CPA //?") //Then;
            //Me.CheckExpedite = //True;
            //Me.Place_on_agenda = //True;
            //Me.on_agenda_label.ForeColor = //vbRed;
        //End //If;
    //End //If;
            ;
//End //If;
    ;
}
function CpaActions_BeforeUpdate(Cancel){
//If //Me.CpaActions.Column(0) = "IRB Approval" //Then;
    //MsgBox //("This //Action //is //for //Display //of //CPAs //created //in //previous //versions //of //PRO_IRB.  //Please //choose //another //action.");
//Exit //Sub;
//End //If;
//ChangedData = //True;
//'Put this in to not change the status of a Study if the CPAAction is mapped to Reserved for system
;
//If //Me.CpaActions.Column(1) //Like "*clos*" //And //Not //Me.CpaActions.Column(1) //Like "*enrollment*" //Then;
 //Forms![_irb //input //form]![TExpireDate] = "";
//End //If;
;
;
//If //Me.Study_Status //Like "*exem*" //Then;
;
    //Exit //Sub;
//Else;
    //'If Me.CpaActions.Column(0) <> "(No Action)" Or Me.CpaActions.Column(1) = "(Reserved for System)" Then
    //If //Me.CpaActions.Column(1) //<> "(Reserved for System)" //Then;
        //Forms![_irb //input //form]![StudyStatus] = //Me.CpaActions.Column(1);
        //Forms![_irb //input //form]![Study //Active?] = //Me.CpaActions.Column(2);
        //Me.TBActive = //Me.CpaActions.Column(2);
        //Me.TBStatus = //Me.CpaActions.Column(1);
;
        //Forms![newmenu]![Status] = //Me.CpaActions.Column(2);
        //If //Forms![newmenu]![Status] //<> "****" //Then;
            //Forms![newmenu]![Status] = //Me.CpaActions.Column(2);
        //End //If;
    //Else;
        //Forms![_irb //input //form]![StudyStatus] = //CpaStatus;
        //Forms![_irb //input //form]![Study //Active?] = //CpaActive;
        //Forms![newmenu]![Status] = //CpaActive;
        ;
        //Me.TBActive = //CpaActive;
        //Me.TBStatus = //CpaStatus;
    //End //If;
//End //If;
;
}
function CPALetterButton_Click(){
//'On Error GoTo Err_CPALetterButton_Click
//'MsgBox Me.Study_Active_.Value
//'If Me.NewRecord Then
 //'   MsgBox "You must Save the Data first.", vbInformation
  //'  Exit Sub
//'End If
;
 //'   Dim stDocName As String
  //'  Dim stLinkCriteria As String
;
   //' stDocName = "_LetterManager"
   //' CPALetterSwitch = True
    //'If Not IsNull(Me.MeetingDate) Then
     //'   HoldMeetingDate = Me.MeetingDate
    //'End If
    //'HoldIRBNo = Me.studyNo
    //'LetterName = "CPA acknowledgment"
    //'stDocName = "_LETTERMANAGER"
    //'CancelUpdate = False
    //'DoCmd.RunCommand acCmdSaveRecord
    //'If Me.NewRecord Then
     //'   MsgBox "No data has been entered"
      //'  Exit Sub
   //' End If
    //''''''---------OUT FOR MERIDIAN 7/10/10
    //'If GetTemplateLocation("CPA") Then
        //'msgbox TemplateLocation
        //'msgbox TemplateName
     //'   If TemplateName = "" Then
      //'      Exit Sub
       //' Else
        //'    If IsNull(Me.Condition2) Then Me.Condition2 = "  "
         //'   DoCmd.SetWarnings False
          //'  DoCmd.OpenQuery ("ExportCPAwithCreateTable")
           //' DoCmd.SetWarnings True
            ;
            //'CustomLetter = True
      //'  End If
    //'Else
        //'msgbox "Do PROIRB SAE Letter"
     //'   CustomLetter = False
    //'End If
  //''  -----------------
  //'CPA
//On //Error //GoTo //Err_CPALetterButton_Click;
//'MsgBox Me.Study_Active_.Value
var stDocName = '';
var stLinkCriteria = '';
var HoldDocNAme = '';
//could initialize i here
var MyCriteria = '';
;
;
                //Forms![newmenu]![txtAmountInvoiced] = //Null;
                //Forms![newmenu]![TxtDatePaid] = //Null;
                //Forms![newmenu]![txtInvCPANumber] = //Null;
                //Forms![newmenu]![txtInvoiceType] = //Null;
                //Forms![newmenu]![txtIRBCode] = //Forms![newmenu]![Location];
                //Forms![newmenu]![txtMeetingDate] = "";
                //Forms![newmenu]![txtPayCPANumber] = //Null;
;
//If //Me.NewRecord //Then;
    //MsgBox "You must Save the Data first."//, //vbInformation;
    //Exit //Sub;
//End //If;
;
    ;
    stDocName = "_LetterManager";
    //CPALetterSwitch = //True;
    //If //Not //IsNull(Me.MeetingDate) //Then;
        //HoldMeetingDate = //Me.MeetingDate;
    //End //If;
    //HoldIRBNo = //Me.studyNo;
    //LetterName = "CPA acknowledgment";
    stDocName = "_LETTERMANAGER";
    //CancelUpdate = //False;
    //DoCmd.RunCommand //acCmdSaveRecord;
    //If //Me.NewRecord //Then;
        //MsgBox "No data has been entered";
        //Exit //Sub;
    //End //If;
;
;
//If //GetTemplateLocation("CPA") //Then;
        //If //TemplateName = "" //Then;
            //Exit //Sub;
        //Else;
            //If //IsNull(Me.Condition2) //Then //Me.Condition2 = "  ";
            //DoCmd.SetWarnings //False;
            i = Len(dir(TemplateName));
            HoldDocNAme = Left$(dir(TemplateName), i - 4);
            ;
            //If HoldDocNAme //Like "*Fee*" //Then;
                MyCriteria = "[TemplateName] = " //& //strQuote //& HoldDocNAme //& //strQuote //& " And [IRBCode] = " //& //strQuote //& //Forms![_irb //input //form]![IRB //Code] //& //strQuote;
                //'MsgBox MyCriteria
                //Forms![newmenu]![txtTemplateName] = HoldDocNAme;
                //Forms![newmenu]![TxtDatePaid] = //Null;
                //Forms![newmenu]![txtInvCPANumber] = //Me.CPAnumber;
                  //Forms![newmenu]![txtInvoiceType] = //DMax("TemplateType", "InvoiceAmountTableQuery"//, //MyCriteria);
              ;
                 //'MsgBox DMax("TemplateType", "InvoiceAmountTableQuery", MyCriteria)
                //'MsgBox (DMax("TemplateType", "InvoiceAmountTable", MyCriteria))
                //Forms![newmenu]![txtAmountInvoiced] = //DMax("AmountToInvoice", "InvoiceAmountTableQuery"//, //MyCriteria);
                //Forms![newmenu]![txtIRBCode] = //Forms![newmenu]![Location];
                ;
                //If //IsNull(Forms![_select //cpas //for //irb#]![MeetingDate]) //Then;
                    //Forms![newmenu]![txtMeetingDate] = "";
                //Else;
                    //Forms![newmenu]![txtMeetingDate] = //Forms![_select //cpas //for //irb#]![MeetingDate];
                //End //If;
                //Forms![newmenu]![txtPayCPANumber] = //Null;
             //End //If;
                           ;
             //DoCmd.OpenQuery //("ExportCPAwithCreateTable");
             //DoCmd.SetWarnings //True;
             //CustomLetter = //True;
        //End //If;
    //Else;
         //CustomLetter = //False;
    //End //If;
 ;
    ;
    ;
    //Forms![newmenu]![SourceNumber] = //Me.tCPAnumber;
    stLinkCriteria = "[IRB#]=" //& "'" //& //HoldIRBNo //& "'";
    //DoCmd.Close;
    location.href = "Form_"+stDocName+".php";//, //acNormal, //, stLinkCriteria;
;
        ;
//Exit_CPALetterButton_Click:;
    //Exit //Sub;
;
//Err_CPALetterButton_Click:;
    //MsgBox "Another user has used this CPA Number.  Click OK then resave the data making sure to take note of the new CPA Number.  " //& //Err.Description;
    //Resume //Exit_CPALetterButton_Click;
    ;
;
}
function CPATypeList_AfterUpdate(){
//If //Me.Study_Status //Like "*exem*" //Then //Exit //Sub;
//If //Me.CPATypeList.Column(2) = //True //Then;
    //Me.Place_on_agenda = //True;
    //Me.on_agenda_label.ForeColor = //vbRed;
//Else;
    //Me.Place_on_agenda = //False;
    //Me.on_agenda_label.ForeColor = //vbBlack;
//End //If;
//'If Me.CPAType = "Modifications Received" Then Me.Expedite = True
;
;
//If //Me.CPATypeList = "Modifications Received" //Then;
    //Me.Expedite = //True;
//Else;
    //Me.Expedite = //False;
//End //If;
;
}
function Form_Activate(){
//DoCmd.Restore;
}
function Form_AfterInsert(){
//MsgBox "Please record the following CPA Internal# to your document--" //& //Hold_CPANumber, //vbInformation  //'done for chla
}
function Form_AfterUpdate(){
//Hold_CPANumber = //Forms![_select //cpas //for //irb#].tCPAnumber;
//Forms![newmenu]![HoldSource] = //Me.tCPAnumber;
//'Forms![_select cpas for irb#].FilterOn = True
//'Forms![_select cpas for irb#].Filter = "[IRB#]=" & "'" & Forms![_irb input form]![ID#] & "'"
;
//'Forms![_select cpas for irb#].Requery
//'Forms![_select cpas for irb#].FilterOn = True
//'DoCmd.ApplyFilter , "[IRB#]=" & "'" & Forms![_irb input form]![ID#] & "'"
//'Forms![_select cpas for irb#].RecordSource = "_selectcpaForrequery"   'done for CHLA
//'Forms![_select cpas for irb#].AllowEdits = True
//'Forms![_select cpas for irb#].DataEntry = False
//'Do Until Forms![_select cpas for irb#].tCPAnumber = Hold_CPANumber  'done for chla
   //'DoCmd.FindRecord "Smith", , True, , True
//'DoCmd.RunCommand acCmdRecordsGoToNext 'done for chla
//'Loop  'done for chla
//If //CyberIRB_Flag //Then //'bridge stuff
    //If //isFormLoaded("CyberBridge") //Then //'bridge stuff
        //gclsCyberIRB.ProcessedByProIRB "CPA"//, //Me.txtcyberrefnum //'bridge stuff
        //Form_CyberBridge_CyberNeedToProcess.RefreshGrid //'bridge stuff   Terry 080902
;
    //End //If;
//End //If //'bridge stuff
}
function Form_ApplyFilter(Cancel, ApplyType){
//MyFilter = //Me.Filter;
}
function Form_Close(){
//On //Error //GoTo //Trap_error;
//BackFromCPA = //True;
;
;
//If //isFormLoaded("_IRB //Input //Form") //Then   //'FOR CHLA
    //BackFromCPA = //True;
    //CancelUpdate = //False;
//'    Forms![_irb input form].Visible = True
    //Forms![_irb //input //form]![StudyStatus] = //CpaStatus;
    //Forms![_irb //input //form]![Study //Active?] = //CpaActive;
    //DoCmd.Close //acForm, "_IRB Input Form"//, //acSaveYes;
    location.href = "Form__IRB Input Form.php";;
    //'032407If isFormLoaded("CyberBridge") Then GoTo Exit_Close
    //'032407DoCmd.OpenForm "_IRB Input Form"
    //'032407Forms![_irb input form].Visible = True
    //'032407Forms![_irb input form].StudyStatus.Requery
    //'032407Forms![_irb input form].AgendaDate.Requery
    //'Forms![_irb input form].lstReviewer.Requery
    //'032407If IsNull(Forms![_irb input form].AgendaDate) Then
       //'032407Forms![_irb input form].NoAgenda.Visible = True
      //'032407Forms![_irb input form].AgendaDate.Visible = False
        //'032407Exit Sub
    //'032407Else
        //'032407Forms![_irb input form].AgendaDate.Visible = True
        //'032407Forms![_irb input form].NoAgenda.Visible = False
    //'032407End If
//End //If;
;
//Exit_Close:;
//DoCmd.Close //acReport, "CPAPrintFullDetail2"//, //acSaveNo;
//Exit //Sub;
//Trap_error:;
//MsgBox "error in cpa  " //& //Err.Number //& "  " //& //Err.Description;
//Resume //Exit_Close;
}
function Form_Current(){
//If //Me.NewRecord //Then;
    //'If Me.Study_Status Like "*exp*" Then Me.CheckExpedite = True
    //Me.Label1.Caption = "New Change in Procedure";
    //Me.Label1.ForeColor = //vbCyan;
    //CpaActive = //Me.Study_Active_;
    //CpaStatus = //Me.Study_Status;
;
    //Exit //Sub;
//Else;
    //Me.Label1.Caption = "Change in Procedure";
    //Me.Label1.ForeColor = //8404992;
    //'Me.AllowAdditions = False
//End //If;
;
;
//If //Me.Place_on_agenda //Then;
   ;
    //Me.on_agenda_label.ForeColor = //vbRed;
    //Else;
    //Me.on_agenda_label.ForeColor = //vbBlack;
//End //If;
;
//CpaActive = //Me.Study_Active_;
//CpaStatus = //Me.Study_Status;
//Me.TBActive = //Me.Study_Active_;
//Me.TBStatus = //Me.Study_Status;
;
;
}
function Form_Open(Cancel){
//txtcyberrefnum.Visible = //True;
//HoldOther1 = "";
//Me.OrderByOn = //False;
//If //Forms![newmenu]![HoldReadOnly] //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.New_Application.Enabled = //False;
    //Me.Save_Data.Enabled = //False;
    //Me.CPALetterButton.Enabled = //False;
//End //If;
;
;
;
;
;
;
}
function New_Application_Click(){
//On //Error //GoTo //error_newapp;
//Me.AllowAdditions = //True;
//DoCmd.GoToRecord //, //, //acNewRec;
//Me.Date_of_Change.SetFocus;
//error_newapp:;
//Exit //Sub;
}
function NextButton_Click(){
//On //Error //GoTo //EndofFile;
//CancelUpdate = //True;
//Me.Undo;
//DoCmd.RunCommand //acCmdRecordsGoToNext;
//exit_next:;
//Exit //Sub;
//EndofFile:;
//MsgBox "End of CPAs"//, //vbInformation;
//DoCmd.RunCommand //acCmdRecordsGoToLast;
;
//GoTo //exit_next;
}
function Place_on_agenda_Click(){
//If //Me.Place_on_agenda //Then;
  ;
    //Me.on_agenda_label.ForeColor = //vbRed;
//Else;
    //If //Me.CheckExpedite = //False //Then;
        //Me.on_agenda_label.ForeColor = //vbBlack;
    //Else;
        //Me.Place_on_agenda = //True;
        //Me.on_agenda_label.ForeColor = //vbRed;
    //End //If;
    //Exit //Sub;
//End //If;
;
}
function preview_print_Click(){
//If //Me.NewRecord //Then;
    //MsgBox "You must Save the Data first."//, //vbInformation;
;
    //Exit //Sub;
//Else;
    //MyFilter = //Me.Filter;
    //'MyFilter = "[selectcpaforprintcpa].[CPANumber]= " & Me.tCPAnumber
    //LetterName = "CPAPrintFullDetail2";
    //Me.Visible = //False;
    ;
    //DoCmd.OpenReport //LetterName, //acViewPreview, //, //MyFilter;
//End //If;
}
function PreviousButton_Click(){
//On //Error //GoTo //BegofFile;
//Me.Undo;
//CancelUpdate = //True;
//DoCmd.RunCommand //acCmdRecordsGoToPrevious;
;
//exit_prev:;
//Exit //Sub;
;
//BegofFile:;
//MsgBox "Beginning of CPAs"//, //vbInformation;
//DoCmd.RunCommand //acCmdRecordsGoToFirst;
;
//GoTo //exit_prev;
;
}
function Returnmenu_Click(){
//Me.Undo;
//DoCmd.Close;
//DoCmd.SelectObject //acForm, "_irb input form";
;
;
;
;
}
function Save_Data_Click(){
//If //SaveCPARecord //Then;
    //Exit //Sub;
//Else;
    //MsgBox //Err.Description;
//End //If;
;
;
}
function CheckIfOnCPA(){
//On //Error //GoTo //Error_Checkifoncpa;
//could initialize hope here
//could initialize dbs here
//Set dbs = CurrentDb;
//could initialize GoForIt here
;
//sqlSTR = "SELECT [agenda records].[IRB#], [agenda records].MeetingDate, [agenda records].[IRB Action Category], [agenda records].[Record Type]" //_;
//& " FROM [agenda records]WHERE [agenda records].[IRB#]= " //& "'" //& //[Forms]![_select //cpas //for //irb#]![IRB#] //& "'" //& " AND [agenda records].MeetingDate <> " //& "#" //& //[Forms]![_select //cpas //for //irb#]![MeetingDate] //& "#" //& " AND isnull([agenda records].[IRB Action Category])";
;
//Set GoForIt = dbs.OpenRecordset(sqlSTR);
;
;
;
;
//If //GoForIt.RecordCount //> //0 //Then;
 //GoForIt.MoveFirst;
            //Do //While //Not //GoForIt.EOF;
                //If //Me.Category //Like "*report" //Or //Me.Category //Like "*renewal" //And //GoForIt![Record //Type] = "renewal" //Then;
                    //GoForIt.Edit;
                    //MsgBox //GoForIt![MeetingDate] //& "  " //& "deleted";
                    //GoForIt.Delete;
                    //GoForIt.MoveNext;
                    ;
                //Else;
                    //If //YesNo("The //Study //is //already //on //the //agenda //for- " & GoForIt![MeetingDate] & " //-Do //you //still //want //to //put //on //this //date?") //Then;
                        //CheckIfOnCPA = //False;
                        ;
                    //Else;
                        //CheckIfOnCPA = //True;
                    //End //If;
                    //GoForIt.Close;
                    //Exit //Function;
                //End //If;
            //Loop;
            //GoForIt.Close;
            //Exit //Function;
//Else;
    //GoForIt.Close;
    //CheckIfOnCPA = //False;
//End //If;
//exit_CheckIfOnCpa:;
//Exit //Function;
;
//Error_Checkifoncpa:;
//MsgBox "Error Checking for correct CPA agenda date" //& //Err.Description //& "  " //& //Err.Number;
//Resume //exit_CheckIfOnCpa;
;
;
}
function DeletePreviousCPAAgendaRecord(){
//On //Error //GoTo //Err_DeletePreviousCPAAgendaRecord;
var sqlSTR = '';
//could initialize db here
//could initialize rs here
//Set db = CurrentDb;
;
            //'this deletes the agenda record previously on file for this cpa
            //Forms![newmenu]![HoldSource] = //Me.tCPAnumber;
            ;
        sqlSTR = "Select [agenda records].[IRB#], [agenda records].MeetingDate, [agenda records].SourceNumber, [agenda records].[Record Type] FROM [agenda records]" //_;
        //& " WHERE ((([agenda records].[IRB#])=" //& "'" //& //[Forms]![_select //cpas //for //irb#]![studyNo] //& "'" //_;
        //& ") AND (( [agenda records].MeetingDate)=" //_;
        //& "#" //& //[Forms]![_select //cpas //for //irb#]![MeetingDate] //& "#" //& ") AND (([agenda records].SourceNumber)= " //_;
        //& "'" //& //[Forms]![newmenu]![HoldSource] //& "') AND (([agenda records].[Record Type])='CPA'" //_;
        //& ")) OR ((([agenda records].[IRB#])=" //& "'" //& //[Forms]![_select //cpas //for //irb#]![studyNo] //& "'" //_;
        //& ") AND (([agenda records].MeetingDate)= #" //& //[Forms]![_select //cpas //for //irb#]![MeetingDate] //& "#" //_;
        //& ") AND (([agenda records].SourceNumber)= '" //& //[Forms]![newmenu]![HoldSource] //& "" //_;
        //& "') AND (([agenda records].[Record Type])='Renewal'))";
        //'MsgBox sqlSTR
        //Set rs = db.OpenRecordset(sqlSTR);
     ;
        //If //Not //rs.RecordCount //> //0 //Then;
            //'MsgBox "Could not find Original Agenda Record.  Processing will continue."
        //Else;
        ;
                        ;
            //rs.MoveFirst;
            //Do //While //Not //rs.EOF;
                //rs.Edit;
                //rs.Delete;
                //rs.MoveNext;
            //Loop;
            //MsgBox "The original CPA has been removed from the agenda.-- You can now schedule the changed CPA";
        //End //If;
        //rs.Close;
        //db.Close;
        //Me.MeetingDate = "";
;
//exit_DeletePreviousCPAAgendaRecord:;
//Exit //Sub;
;
//Err_DeletePreviousCPAAgendaRecord:;
//MsgBox //Err.Number //& "  " //& //Err.Description;
//Resume //exit_DeletePreviousCPAAgendaRecord;
}
function SaveCPARecord(){//As //Boolean;
//On //Error //GoTo //err_SaveCPARecord;
//If //EnrollOver() //And //Me.NewRecord //Then;
    //Me.Other_1 = //Me.Other_1 //& "**Current Enrollment of: " //& //Forms![_irb //input //form]![# //Patients //enrolled] //_;
    //& "  exceeds Approved enrollment level of: " //& //val(Forms![_irb //input //form]![InitEnroll]);
    //MsgBox //("Enrollment //exceeds //approval //level. //Comment //has //been //added //to //the //'Other 1' box."), vbCritical
//End //If;
//If //Not //Me.NewRecord //Then;
    //Forms![newmenu]![SourceNumber] = //Me.tCPAnumber;
//End //If;
  ;
  ;
//If //IsNull(Me.[Date //Received //by //IRB //Co-Ord]) //Then;
     //MsgBox "Date Received cannot be blank"//, //vbInformation;
     //Exit //Function;
//End //If;
//If //Me.CPATypeList.ListIndex = //-1 //Then;
        //MsgBox "You must select a Type of Change"//, //vbInformation;
        //Exit //Function;
//End //If;
//If //Not //IsNull(Me.MeetingDate) //Then //Call //DeletePreviousCPAAgendaRecord;
  ;
  ;
//CancelUpdate = //False;
//Me.AllowEdits = //True;
   ;
//If //Me.NewRecord //Then;
    //'Me.tCPAnumber = 1595
     //Me.tCPAnumber = //IIf(IsNull(DMax(" //[CPA]![CPANumber]", "CPA"//)), //1, //DMax(" //[CPA]![CPANumber]", "CPA"//) //+ //1);
        //Hold_CPANumber = //Me.tCPAnumber  //'done for chla
        //If //IsNull(txtcyberrefnum) //Then //txtcyberrefnum = "NR" //& //Me.[IRB#] //& //Me.tCPAnumber;
//End //If;
//Hold_CPANumber = //Me.tCPAnumber  //'done for chla
//'added for 9.00.0036
//If //Me.Expedite //Then;
 //If //Me.CpaActions = "Modifications Approved" //And //Me.TBStatus = "OPEN" //Then;
    ;
    //[Forms]![_irb //input //form]![Approval //Date  //(original)] = //Me.Date_of_Change;
    //[Forms]![_irb //input //form]![Last //Renewal //Date] = //[Forms]![_irb //input //form]![First //IRB //Review];
    //'MsgBox DMax("[date of change]", "CPA", "[cpanumber]=" & Me.SourceNumber)
       ;
    //HoldMeetingDate = //DateAdd("m", //Forms![_irb //input //form]![Renewal //Cycle], //[Forms]![_irb //input //form]![First //IRB //Review] //- //1);
    //[Forms]![_irb //input //form]![TExpireDate] = //HoldMeetingDate;
    ;
 //End //If;
//End //If;
;
    ;
//On //Error //Resume //Next;
//Save_CPA_Record:;
//DoCmd.RunCommand //acCmdSaveRecord;
//If //Err.Number //<> //0 //Then;
    //If //Err.Description //Like "*dupl*" //Then;
        //Me.tCPAnumber = //IIf(IsNull(DMax(" //[CPA]![CPANumber]", "CPA"//)), //1, //DMax(" //[CPA]![CPANumber]", "CPA"//) //+ //1);
        //Hold_CPANumber = //Me.tCPAnumber;
        //MsgBox "The CPA Number has been changed to " //& //Me.tCPAnumber;
        //Err.Clear;
        //GoTo //Save_CPA_Record;
    //Else;
       //MsgBox "Error trying to save CPA Data--" //& //Err.Number //& "--" //& //Err.Description //' done chla
        //SaveCPARecord = //False;
        //Exit //Function;
    //End //If;
//End //If;
//Hold_CPANumber = //Me.tCPAnumber;
    ;
//If //Me.Place_on_agenda //Then;
        //Call //PlaceOnAgenda;
//End //If;
 ;
;
//If //Me.CheckOpenLetters.ListCount //> //0 //Then;
        //Call //CheckFULetters;
//End //If;
//Forms![_select //cpas //for //irb#].AllowEdits = //True;
//Forms![_select //cpas //for //irb#].DataEntry = //False;
//Forms![_select //cpas //for //irb#].RecordSource = "_selectcpaSForRequery"   //'done for CHLA
//SaveCPARecord = //True;
;
;
//'**************************** Cyber stuff *******************************************************************
//'need flag
//' I have absolutely no idea what to do if the CPA is initiated from ProIRB....
//If //CyberIRB_Flag //Then;
    //If //UpdateCyberIRBCPA(Me.txtcyberrefnum) //Then   //' check to see if the bolupdatecyber is true or not
        //UpdateCyberCPA //Me.studyNo, //Me.CPAnumber;
    //End //If;
    //Form_CyberBridge_CyberNeedToProcess.RefreshGrid    //'Bridge stuff Terry 090108
//End //If;
;
//getout:;
    //Exit //Function;
   ;
//err_SaveCPARecord:;
//SaveCPARecord = //False;
    //'this command will raise and print the description of any message #
    //'msgbox ErrorString(2001)
    //'this test for errors needs to be updated to a cleaner way to allow cancelling of a save operation
//If //Err.Number = //2501 //Or //Err.Number = //2001 //Then;
    //Resume //getout;
//Else;
    //MsgBox //Err.Description //& //Err.Number //& "  Another user has used this CPA Number.  Click OK then resave the data making sure to take note of the new CPA Number."   //'Done for CHLA
    //Resume //getout;
//End //If;
;
}
function CmdViewIRBAction_Click(){
//On //Error //GoTo //Err_CmdViewIRBAction_Click;
;
;
;
//If //IsNull(DMax("[sourcenumber]", "_PostIRBMeetingResults1SingleCPA"//)) //Then;
    //MsgBox "There is no Agenda Record for this Item";
    //Exit //Sub;
//End //If;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName = "_PostMeetingFormSingleItemCPA";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
;
//Exit_CmdViewIRBAction_Click:;
    //Exit //Sub;
;
//Err_CmdViewIRBAction_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_CmdViewIRBAction_Click;
    ;
}
function SentUpLetterSentFields(){
//If //TemplateName = "C:\PRO_IRB\Templates\Invoices\Renewal Fee.dot" //Then;
                //DoCmd.OpenQuery //("ExportRenewalInvoiceWithCreateTable");
                //GoTo //getout;
            //End //If;
            //If //TemplateName = "C:\PRO_IRB\Templates\Invoices\Amendment Fee.dot" //Then;
                //DoCmd.OpenQuery //("ExportAmendmentInvoiceWithCreateTable");
                //GoTo //getout;
            //End //If;
            //If //TemplateName = "C:\PRO_IRB\Templates\Invoices\Renewal and Ammendment Fee.dot" //Then;
                //DoCmd.OpenQuery //("ExportRenewalAmendmentInvoiceWithCreateTable");
                //GoTo //getout;
            //Else;
                //DoCmd.OpenQuery //("ExportStudySpecificWithCreateTable");
            //End //If;
;
//getout:;
//Exit //Sub;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Active();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:768; height:64'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <input type='text' id=tCPAnumber value='empty' style='position:absolute; left:18; top:201; width:168; height:30'>

    <label id=SAE Number Label value='CPA Number' style='position:absolute; left:12; top:165; width:178; height:24; visibility:'>CPA Number</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:612; top:144; width:504; height:0'>

    <label id=Protocol Number & Title Label value='Protocol#/Title' style='position:absolute; left:612; top:120; width:138; height:24; visibility:'>Protocol#/Title</label>
    <input type='text' id=Date of Change value='empty' style='position:absolute; left:216; top:201; width:154; height:30'>

    <label id=Date Change Label value='Effective Date of Change' style='position:absolute; left:198; top:165; width:225; height:24; visibility:'>Effective Date of Change</label>
    <input type='checkbox' id=Signed by PI ? value='False' style='position:absolute; left:942; top:456; width:24; height:30'>

    <label id=Signed by PI ? Label value='Signed ' style='position:absolute; left:912; top:414; width:73; height:24; visibility:'>Signed </label>
    <input type='text' id=Date Signed value='empty' style='position:absolute; left:1014; top:450; width:96; height:30'>

    <label id=Date Signed Label value='Date Signed' style='position:absolute; left:990; top:414; width:114; height:24; visibility:'>Date Signed</label>
    <input type='text' id=date recieved by value='empty' style='position:absolute; left:426; top:199; width:144; height:31'>

    <label id=Label56 value='Date Received ' style='position:absolute; left:426; top:165; width:142; height:24; visibility:'>Date Received </label>
    <input type='button' id=preview_print value='Pre&view\015\012CPAs' style='position:absolute; left:598; top:540; width:123; height:49' onclick='preview_print_Click();'>
    <input type='button' id=Returnmenu value='&Return' style='position:absolute; left:1017; top:540; width:111; height:49' onclick='Returnmenu_Click();'>
    <input type='button' id=New Application value='&Add New CPA' style='position:absolute; left:438; top:540; width:163; height:49' onclick='New Application_Click();'>
    <input type='button' id=Save Data value='&Save Data' style='position:absolute; left:880; top:540; width:136; height:49' onclick='Save Data_Click();'>
    <input type='checkbox' id=Place on agenda value='False' style='position:absolute; left:331; top:493; width:0; height:0'>

    <label id=on agenda label value='Place on Meeting Agenda' style='position:absolute; left:354; top:490; width:228; height:24; visibility:'>Place on Meeting Agenda</label>
    <input type='text' id=Other 1 value='empty' style='position:absolute; left:618; top:264; width:498; height:59'>

    <label id=Label108 value='Other 1' style='position:absolute; left:540; top:288; width:72; height:24; visibility:'>Other 1</label>
    <input type='text' id=other2 value='empty' style='position:absolute; left:621; top:336; width:495; height:65'>

    <label id=Label110 value='Other 2' style='position:absolute; left:540; top:354; width:72; height:24; visibility:'>Other 2</label>
    <label id=Label1 value='Change in Procedure' style='position:absolute; left:42; top:96; width:517; height:54; visibility:'>Change in Procedure</label>
    <input type='text' id=studyNo value='empty' style='position:absolute; left:990; top:36; width:132; height:30'>

    <label id=Label112 value='Study #' style='position:absolute; left:1014; top:6; width:75; height:24; visibility:'>Study #</label>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:390; top:36; width:597; height:30'>

    <label id=Label453 value='IRB' style='position:absolute; left:540; top:6; width:82; height:24; visibility:'>IRB</label>
    <select name='CpaActions' Size='3' style='position:absolute; left:276; top:276; width:246; height:163'>

    <label id=CPAAction_Label value='Pre-Meeting Action Taken' style='position:absolute; left:264; top:240; width:231; height:24; visibility:'>Pre-Meeting Action Taken</label>
    <select name='Sponsor ID' Size='11' style='position:absolute; left:624; top:456; width:264; height:0'>

    <label id=Label130 value='Sponsor' style='position:absolute; left:534; top:456; width:78; height:24; visibility:'>Sponsor</label>
    <select name='Investigator ID' Size='22' style='position:absolute; left:624; top:420; width:264; height:0'>

    <label id=Label131 value='P.I.' style='position:absolute; left:576; top:420; width:39; height:24; visibility:'>P.I.</label>
    <input type='text' id=Study Status value='empty' style='position:absolute; left:960; top:84; width:162; height:30'>

    <label id=Label70 value='Study Status:' style='position:absolute; left:828; top:84; width:123; height:24; visibility:'>Study Status:</label>
    <input type='text' id=Study Active? value='empty' style='position:absolute; left:669; top:84; width:102; height:24'>

    <label id=Label69 value='Active?:' style='position:absolute; left:579; top:84; width:79; height:24; visibility:'>Active?:</label>
    <input type='checkbox' id=CheckExpedite value='False' style='position:absolute; left:331; top:463; width:120; height:21'>

    <label id=Label132 value='Expedite' style='position:absolute; left:354; top:460; width:82; height:24; visibility:'>Expedite</label>
    <!--     <input type='text' id=condition1 value='empty' style='position:absolute; left:324; top:24; width:138; height:72'> -->

    <label id=Label134 value='Text133' style='position:absolute; left:669; top:0; width:78; height:24; visibility:'>Text133</label>
    <!--     <input type='text' id=Condition2 value='empty' style='position:absolute; left:354; top:60; width:114; height:66'> -->

    <label id=Label136 value='Text135' style='position:absolute; left:216; top:78; width:78; height:24; visibility:'>Text135</label>
    <!--     <input type='text' id=Category value='empty' style='position:absolute; left:378; top:24; width:186; height:54'> -->

    <label id=Label138 value='Text137' style='position:absolute; left:486; top:30; width:78; height:24; visibility:'>Text137</label>
    <input type='button' id=NextButton value='&Next CPA Record' style='position:absolute; left:241; top:540; width:196; height:49' onclick='NextButton_Click();'>
    <input type='button' id=PreviousButton value='&Previous CPA Record' style='position:absolute; left:0; top:540; width:243; height:49' onclick='PreviousButton_Click();'>
    <select name='CPATypeList' Size='2' style='position:absolute; left:24; top:270; width:240; height:222'>
    <label id=Label143 value='Type of Change' style='position:absolute; left:60; top:240; width:144; height:24; visibility:'>Type of Change</label>
    <input type='text' id=MeetingDate value='empty' style='position:absolute; left:750; top:490; width:186; height:24'>

    <label id=Label145 value='For Meeting Date' style='position:absolute; left:589; top:490; width:156; height:24; visibility:'>For Meeting Date</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:34; top:40; width:163; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <!--     <select name='CheckOpenLetters' Size='28' style='position:absolute; left:156; top:366; width:168; height:60'> -->

    <label id=Label151 value='List150:' style='position:absolute; left:12; top:366; width:76; height:24; visibility:'>List150:</label>
    <input type='button' id=CPALetterButton value='Acknowledgment\015\012&Letter' style='position:absolute; left:720; top:540; width:160; height:49' onclick='CPALetterButton_Click();'>
    <label id=Label116 value='***Fields with Labels of this color transfer to Agenda' style='position:absolute; left:0; top:517; width:453; height:21; visibility:'>***Fields with Labels of this color transfer to Agenda</label>
    <label id=Label155 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:43; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <input type='button' id=CmdViewIRBAction value='View IRB Action' style='position:absolute; left:954; top:486; width:162; height:40' onclick='CmdViewIRBAction_Click();'>
    <!--     <input type='text' id=TBStatus value='empty' style='position:absolute; left:222; top:18; width:36; height:25'> -->

    <!--     <label id=Label158 value='Text157' style='position:absolute; left:567; top:0; width:78; height:24; visibility:hidden'>Text157</label> -->
    <!--     <input type='text' id=TBActive value='empty' style='position:absolute; left:222; top:60; width:54; height:18'> -->

    <!--     <label id=Label160 value='Text159' style='position:absolute; left:174; top:42; width:78; height:24; visibility:hidden'>Text159</label> -->
    <!--     <input type='text' id=txtcyberrefnum value='empty' style='position:absolute; left:210; top:12; width:180; height:24'> -->
  </body>
</html>
