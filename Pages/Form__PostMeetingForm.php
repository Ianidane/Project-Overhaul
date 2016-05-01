<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function NewExpirationDate(){
//If //Not //IsNull(Me.Last_Renewal_Date) //Then;
          //HoldMeetingDate = //DateAdd("m", //Me.Renewal__Cycle, //Me.Last_Renewal_Date) //- //1;
        //Me.TExpireDate = //HoldMeetingDate;
//Else;
    //If //Not //IsNull(Me.Approval_Date___original_) //Then;
         //HoldMeetingDate = //DateAdd("m", //Me.Renewal__Cycle, //Me.Approval_Date___original_) //- //1;
        //Me.TExpireDate = //HoldMeetingDate;
    //End //If;
//End //If;
//Me.ExpirationDays = //(DateDiff("d", //Me.Meeting_date, //Me.TExpireDate));
//'MsgBox Me.ExpirationDays
;
;
;
;
}
function ActionCombo_AfterUpdate(){
//If //IsNull(Me.ActionCombo.Column(1)) //Then;
        //Me.Undo;
        //Me.NewStudyStatus.Requery;
        //Me.NewStudyStatus = //Me.NewStudyStatus.ItemData(SetComboItem(Me.NewStudyStatus.OldValue));
       //Exit //Sub;
//End //If;
;
//If //Me.SourceNumber = "New Appl" //And //IsNull(Me.Last__IRB_date) //Then;
    //Me.DateFirstReviewed = //Me.Meeting_date;
//End //If;
//If //Me.Agenda_Category //Like "*exped*" //Then;
    //Call //DealWithExpedites;
    //Exit //Sub;
//End //If;
//If //Me.Agenda_Category //Like "*exem*" //Then;
       //Call //DealWithExempts;
        //Exit //Sub;
   ;
//End //If;
;
;
//'this routine deals with which actions will trigger an expriation date recalc
//'
//If //Me.ActionCombo //Like "*NOT*" //Or //Me.ActionCombo //Like "*DIS*" //Then //GoTo //CheckAction;
//If //Me.ActionCombo //Like "*Approv*" //Then //GoTo //ContinueAction;
//If //Me.ActionCombo //Like "*Accept*" //Then //GoTo //ContinueAction;
//If //Me.ActionCombo //Like "*Renew*" //Then //GoTo //ContinueAction;
//If //Me.ActionCombo //Like "*exped*" //Then //GoTo //ContinueAction;
;
//If //Me.ActionCombo = "Closed to Enrollment" //Then //GoTo //ContinueAction;
;
//GoTo //CheckAction;
//'If Me.ActionCombo Like "*Approv*" Or Me.ActionCombo Like "*Renewed*" Or Me.ActionCombo Like "*exempt*" Or Me.ActionCombo Like "*exped*" Or Me.ActionCombo = "Closed to Enrollment" Then
//'leaving the exemp and expedite here until I'm sure we can convert the open stuffOr Me.ActionCombo Like "*exemp*" Then   ' gets contingent and regular approval
//ContinueAction:;
    //If //Me.Record_Type = "Initial" //Then;
        //If //IsNull(Me.Approval_Date___original_) //Then;
            //Me.Approval_Date___original_ = //Me.Meeting_date;
            //Me.Last_Renewal_Date = //Me.Meeting_date;
        //Else;
            //Me.Last_Renewal_Date = //Me.Approval_Date___original_;
        //End //If;
        //Call //NewExpirationDate;
        //GoTo //CheckAction;
    //End //If;
    //If //Me.Agenda_Category //Like "*Renewal*" //Then;
        //Me.Last_Renewal_Date = //Me.Meeting_date;
        //Call //NewExpirationDate;
    //End //If;
;
//'changes for 9.00.0036
//If //Me.ActionCombo = "Modifications Approved" //And //Me.NewStudyStatus = "Pending Modifications" //Then;
    //Me.Last_Renewal_Date = //Me.Meeting_date;
    //Me.Approval_Date___original_ = //Me.Meeting_date;
    //Me.Last__IRB_date = //Me.Meeting_date;
    //'MsgBox DMax("[date of change]", "CPA", "[cpanumber]=" & Me.SourceNumber)
     //Call //NewExpirationDate;
//End //If;
;
//CheckAction:;
//'IRB Tabled  restores original values
//If //Me.ActionCombo //Like "*Tabled*" //Then;
    //HoldIRBAction = //Me.ActionCombo;
    //Me.Undo;
    //If //Me.SourceNumber = "New Appl" //And //IsNull(Me.Last__IRB_date) //Then;
        //Me.DateFirstReviewed = //Me.Meeting_date;
    //End //If;
    //Me.ActionCombo = //HoldIRBAction;
    //Me.Last__IRB_date = //Me.Meeting_date;
    //Exit //Sub;
//Else;
    //Me.Last__IRB_date = //Me.Meeting_date;
       ;
   //If //Me.ActionCombo.Column(1) = "(reserved for system)" //Then;
   //'added in 5_10
  //GoTo //CheckClearExpirationDate:;
        //'Exit Sub
    //Else;
        //If //Me.ActionCombo.Column(2) = "Closed" //Then;
            //Me.NewStudyActive = //Me.NewStudyActive.ItemData(1);
        //Else;
            //Me.NewStudyActive = //Me.NewStudyActive.ItemData(0);
        //End //If;
        //Me.NewStudyStatus.Requery;
        //Me.NewStudyStatus = //Me.NewStudyStatus.ItemData(SetComboItem(Me.ActionCombo.Column(1)));
    //End //If;
//End //If;
//CheckClearExpirationDate:;
//If //Me.NewStudyActive = "Closed" //Then;
    //Me.TExpireDate = "";
    //If //YesNo("Any //Agenda //items //for //this //Study //on //future //meetings //will //be //deleted //as //a //result //of //Closing //this //Study. //Are //you //sure //you //want //to //Close //this //Study? //If //you //do //not //want //to //Close //this //Study //Choose //No //then //Select //an //IRB //Action //other //than //Closed") //Then;
        //On //Error //Resume //Next;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "DeleteFutureAgendaItemsForClosedStudies";
        //DoCmd.SetWarnings //True;
    //Else;
        //Exit //Sub;
        ;
    //End //If;
//Else;
//End //If;
}
function ActionCombo_BeforeUpdate(Cancel){
//If //Me.ActionCombo.OldValue = //Me.ActionCombo //And //Me.ActionCombo //Like "*Tabled*" //Then;
    //MsgBox //("You //cannot //'re'  Table a Study.  If necessary you can reassign to another meeting via the Preliminary Agenda Form")
    //Me.Undo;
    //Cancel = //True;
//End //If;
}
function ButtonReturn_Click(){
//If //Forms![newmenu]![HoldReadOnly] //Then //Me.Undo;
;
//DoCmd.Close;
}
function CopyButton_Click(){
//Me.Action_Explanation = //Me.AgendaReportStringText //& "--" //& //Me.Agenda_Explanation;
}
function cmdAddWording_Click(){
//Me.Action_Explanation = //Me.Action_Explanation //& " " //& //Me.txtSpecial_Wording;
;
}
function DiscussionButton_Click(){
location.href = "Form_MeetingDiscussion.php";;
;
;
}
function Form_Activate(){
//DoCmd.Restore;
//Me.ListriskCat.Requery;
}
function Form_ApplyFilter(Cancel, ApplyType){
//End //Sub;
;
//Private //Sub //Form_BeforeUpdate(Cancel //As //Integer);
//On //Error //GoTo //ErrorBeforeUpdate;
//If //AgendaDetail //Then //Exit //Sub;
//'msgbox Me.ActionCombo.OldValue & "   " & Me.ActionCombo
//If //Not //DataSaved //Then;
    //answer = //YesNo("Do //you //wish //to //post //the //changes //you've //made?");
    //If //answer = //False //Then;
        //Me.Undo;
        //Exit //Sub;
    //End //If;
//End //If;
//If //IsNull(Me.ActionCombo) //Then;
 //'   msgbox ("You must record an action before saving.  Changes will not be saved."), vbInformation
  //'  Me.Undo
  //'  Cancel = True
   //' Exit Sub
   //GoTo //Exit_Save;
//End //If;
//If //AgendaDetail //Then;
    //GoTo //Exit_Save;
//Else;
    //If //Me.ActionCombo.OldValue //Like "*Tabled*" //And //Me.ActionCombo //Like "*Tabled*" //Then;
            //GoTo //Exit_Save;
    //End //If;
    ;
    //If //Me.ActionCombo //Like "*Tabled*" //Then;
        //Call //TabledRecord;
        //If //ReassignSwitch = //False //Then  //' means that the user canceled the new meeting
            //If //YesNo("You //cancelled //rescheduling //the //Study.  //Do //you //still //want //to //save //the //data?") //Then;
            //Else;
                //Me.Undo;
                //Exit //Sub;
            //End //If;
        //Else;
            //Me.Action_Explanation = "Assigned to the " //& //Me.NewAgendaDate //& "  Meeting.  " //& //Me.Action_Explanation;
            ;
        //End //If;
    //Else;
        ;
    //End //If;
//End //If;
//Exit_Save:;
//MsgBox "Data Saved";
//'SpecialSaveFlag = True
//If //CheckIfConsentApproval() //Then;
    //Call //UpdateConsentDate;
//End //If;
;
//Exit_Update:;
//Exit //Sub;
//ErrorBeforeUpdate:;
//MsgBox "Before Update" //& "  " //& //Err.Description, //vbInformation;
//GoTo //Exit_Update;
}
function Form_Close(){
//If //AgendaDetail = //True //Then;
    //AgendaDetail = //False;
//End //If;
//If //isFormLoaded("_display //agenda //records") //Then;
    //Forms![_display //agenda //records].Visible = //True;
//End //If;
   ;
}
function Form_Current(){
//'MsgBox "Form_ Currnet" & Me.Recordset("IRB#")
//Me.ListriskCat.Requery;
//[Forms]![newmenu]![study] = //Forms![_postmeetingform]![IRB#];
//[Forms]![newmenu]![Location] = //Forms![_postmeetingform]![IRB //Code];
;
;
//If //IsNull(Me.TExpireDate) //Then;
   //OrigExpireDate = //#1/1/2001#;
//Else;
   //OrigExpireDate = //Me.TExpireDate;
//End //If;
;
//Me.ExpirationDays = //(DateDiff("d", //Me.Meeting_date, //Me.TExpireDate));
;
;
;
;
//CancelUpdate = //True;
//DataSaved = //False;
//CpaStatus = //Me.NewStudyStatus;
//Me.ReasonLabel.Caption = "------Reason Study is on Agenda-------";
    //Me.ReasonLabel.ForeColor = //8388608;
//If //Me.Agenda_Category //Like "*Exped*" //Then;
    //Me.ReasonLabel.Caption = "Expedited";
    //Me.ReasonLabel.ForeColor = //vbRed;
//End //If;
//If //Me.Agenda_Category //Like "*Exem*" //Then;
    //Me.ReasonLabel.Caption = "Exempted";
    //Me.ReasonLabel.ForeColor = //vbRed;
//End //If;
    ;
;
//If //IsNull(Me.HoldVote) //Or //Me.HoldVote = "" //Then;
    //Me.VoteButton.ForeColor = //8404992;
//Else;
    //Me.VoteButton.ForeColor = //vbCyan;
//End //If;
;
//If //IsNull(Me.holdDiscussion) //Or //Me.holdDiscussion = "" //Then;
    //Me.DiscussionButton.ForeColor = //8404992;
//Else;
    //Me.DiscussionButton.ForeColor = //vbCyan;
//End //If;
;
;
;
;
;
;
;
}
function TabledRecord(){
//On //Error //GoTo //TabledError;
//DoItAgain:;
//HoldIRBNo = //Me.ID_;
location.href = "Form_ReassignAgendaRecord.php";//, //, //, //, //, //acDialog;
//If //ReassignSwitch = //True //Then;
    //CompareDate = //Me.Meeting_date;
    //If //HoldMeetingDate //< //CompareDate //Or //HoldMeetingDate = //CompareDate //Then;
        //MsgBox "Choose a Meeting Date beyond the meeting you are posting to";
        //GoTo //DoItAgain;
    //End //If;
    //Me.NewAgendaDate = //HoldMeetingDate;
    //If //CheckAllReadyOn //Then;
        //GoTo //DoItAgain;
    //Else;
        //Forms![newmenu]![SourceNumber] = //Me.SourceNumber;
        //DoCmd.SetWarnings //False;
        //Call //DoAddTabled;
        //If //Me.Record_Type = "CPA" //Then;
            //DoCmd.OpenQuery "ReassignCPAForTabled";
        //End //If;
        //If //Me.Record_Type = "SAE" //Then;
            //DoCmd.OpenQuery "ReassignSAEForTabled";
        //End //If;
       ;
        //MsgBox //("Study //#- " & HoldIRBNo & " //Internal //#-" //& //Me.SourceNumber //& "  Has been assigned to the--" //& //HoldMeetingDate //& "  Meeting." //_;
        //& "                                    PLEASE NOTE: PRO_IRB no longer automatically tables all items for the Protocol.  " //_;
        //& " You must table each item individually."//);
        //DoCmd.SetWarnings //False;
    //End //If;
//End //If;
//exit_tabled:;
//Exit //Sub;
;
//TabledError:;
//If //Err.Number = //0 //Then;
    //GoTo //exit_tabled;
//Else;
    //MsgBox "Error Creating New Agenda Record" //& //Err.Number //& //Err.Description;
    //Resume //exit_tabled;
//End //If;
;
}
function Form_Open(Cancel){
//Me.MeetingKey = //HoldMeetingKey;
//Me.MeetingKey = //Me.MeetingKey;
//Forms![newmenu]![MeetingKey] = //Me.MeetingKey;
;
//'SpecialSaveFlag = False
//If //Forms![_postmeetingform].RecordsetClone.RecordCount = //0 //Then;
    //MsgBox "There are no Studies on the Agenda for this Meeting Date"//, //vbInformation;
   //DoCmd.CancelEvent;
   //Exit //Sub;
//End //If;
;
;
//If //Not //IsNull(DMax("[Record_ID]", "Agenda Records"//, "[meetingdate] =  #" //& //ThisMeetingDate //& "#" //& " and [agendalinenumber] = 0 and [IRBCode] = " //& //strQuote //& //Me.IRB_Code //& //strQuote)) //Then;
    //'means there are records with Line #s of 00
    //If //PostKeepTogether //Then;
        //Call //NewLineNumbers(True);
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "qryUpdateAgendaRecordsWithLineNumbers";
        //'MsgBox Me.MeetingKey
        //DoCmd.OpenQuery "QryUpdateMeetingDateExpansionFromPost";
        //'docmd.OpenQuery "QryUpdateMeetingDateExpansionFromPost"
        //DoCmd.SetWarnings //True;
        //Me.Requery;
        ;
    //Else;
         //MsgBox //("There //are //Agenda //Records //which //have //been //added //since //you //last //viewed //and //saved //the //Item#s //on //the //Agenda. " _
        & "//Click //OK //to //return //to //the //Main //Menu //and //then //Select //Preliminary //Agenda //Items //to //display //the //Agenda //Records //for //this //Meeting //Date.  //You //must //enter //Item#s //for //any //Item //with //an //Item# //of //0.0, //or //Check //the //Keep //All //Items //Together //box. //It //is //not //necessary //to //prepare //the //Agenda.  //It //is //only //necessary //that //the //Preliminary //Agenda //Items //Form //be //displayed.  //Then //Return //to //the //Main //Menu //and //Select //Post //Meeting //Actions //Again.");
        //DoCmd.CancelEvent;
   ;
    //End //If;
    ;
//End //If;
;
;
;
//If //Forms![newmenu]![HoldReadOnly] //Or //AgendaDetail //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.SendLetterButton.Enabled = //False;
     //Me.Save_Data.Enabled = //False;
//End //If;
;
//DoCmd.Close //acForm, "_agendamonthcheck";
;
//Me.AllowFilters = //False;
//If //IsNull(Me.holdDiscussion) //Then;
    //Me.DiscussionButton.ForeColor = //8404992;
//Else;
    //Me.DiscussionButton.ForeColor = //vbCyan;
//End //If;
//If //IsNull(Me.HoldVote) //Or //Me.HoldVote = "" //Then;
    //Me.VoteButton.ForeColor = //8404992;
//Else;
    //Me.VoteButton.ForeColor = //vbCyan;
//End //If;
//If //RenewAnniversary //Then;
    //Me.TExpireDate.Locked = //False;
//End //If;
;
//'code to Update First review date if Agenda CAt = Initial Sub and reason 2 not "previosly
//'if me.Agenda_Category
}
function NewStudyActive_Change(){
//Me.NewStudyStatus.Requery;
//Me.NewStudyStatus = //Me.NewStudyStatus.ItemData(0);
//If //Me.NewStudyActive = "Closed" //Then;
    //If //YesNo("Any //Agenda //items //for //this //Study //on //future //meetings //will //be //deleted //as //a //result //of //Closing //this //Study. //Are //you //sure //you //want //to //Close //this //Study? //If //you //do //not //want //to //Close //this //Study //Choose //No //then //Select //a //different //Study //Status") //Then;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "DeleteFutureAgendaItemsForClosedStudies";
        //DoCmd.SetWarnings //True;
    //Else;
        //Exit //Sub;
        ;
    //End //If;
//End //If;
;
}
function NextButton_Click(){
//could initialize Response here
//CancelUpdate = //True;
//DoCmd.GoToRecord //, //, //acNext;
//Exit //Sub;
;
;
//NextCheckerror:;
//If //Err.Number = //2105 //Then;
    //MsgBox "End of Studies.  Returning to first Study";
    //DoCmd.GoToRecord //, //, //acFirst;
    //Exit //Sub;
//Else;
    //MsgBox "A error has occurred looping through the records.  Please be certain that 2 versions of PRO_IRB are not running.  In any event, exit PRO_IRB normally, reboot and start PRO_IRB again.  If the error persists call Technical Support";
    //Exit //Sub;
//End //If;
}
function PreviousButton_Click(){
//could initialize Response here
//CancelUpdate = //True;
//DoCmd.GoToRecord //, //, //acPrevious;
//Exit //Sub;
;
;
//PrevCheckerror:;
//If //Err.Number = //2105 //Then;
    //MsgBox "     Beginning of Studies.";
    //DoCmd.GoToRecord //, //, //acFirst;
    //Exit //Sub;
//Else;
    //MsgBox "Big time error";
    //Exit //Sub;
//End //If;
;
}
function Renewal__Cycle_AfterUpdate(){
//Call //NewExpirationDate;
//Me.TExpireDate = //Me.TExpireDate.Value;
}
function Save_Data_Click(){
//Call //SaveRecordSub;
 //UpdateCyberIRB;
//''This doesn't work when you just click on the next arrow......
//'      Dim cmd As ADODB.Command
//'30    Set cmd = New ADODB.Command
//'
//''This record set is from query _postIRBMeetingResults. if you need more data, just look in there, then use Me.recordset("x")
//'
//'
//'  If CyberIRB_Flag Then
//'      UpdateCyberProtocol Me.[IRB#]
//'
//'
//'180        cmd.ActiveConnection = gclsCyberIRB.CyberConnection
//'190        cmd.CommandText = "PIRB_" & gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") & "_UpdateMeetingMinutes"
//'200        cmd.CommandType = adCmdStoredProc
//''
//'210        cmd.Parameters("@varCustNum") = gclsCyberIRB.mcolCyberConfigAttributes("CustNum")
//'220        cmd.Parameters("@varIRBNum") = Me.Recordset("IRB#")  ' field number (6)
//'230        cmd.Parameters("@varRefNum") = Me.Recordset("RefNum")  ' field number (49)
//'235        cmd.Parameters("@varStudyStatus") = Me.Recordset("Study Status")  '    This is the Protocol study status, not the agenda items
//'240        cmd.Parameters("@varAgendaGroup") = Me.Recordset("AgendaGroup")  'field number (4)
//'250        cmd.Parameters("@varMeetingDate") = Me.Recordset("MeetingDate") '(11)
//'260        cmd.Parameters("@varRecordType") = Me.Recordset("Record Type") '(3), (44) SAE, CPA, Initial
//'270        cmd.Parameters("@varAgendaCategory") = Me.Recordset("Agenda Category") '(42)
//'280        cmd.Parameters("@varAgendaCondition1") = Me.Recordset("Agenda Condition1") '(13)
//'290        cmd.Parameters("@varAgnedaCondition2") = Me.Recordset("Agenda Condition2") '(14)
//'300        cmd.Parameters("@varActionCategory") = Me.Recordset("IRB Action Category") '(31)
//'310        cmd.Parameters("@varExpediteFlag") = Me.Recordset("ExpediteFlag") '(12)
//'320        cmd.Execute
//'  End If
//'
;
//'
//'dick's
      //' On Error GoTo SaveError
      //'If IsNull(Me.ActionCombo) Then
       //'   msgbox ("You must record an action before saving"), vbInformation
       //'   Exit Sub
      //'End If
;
      //'DataSaved = True
      //'DoCmd.RunCommand acCmdSaveRecord
      //'SpecialSaveFlag = True  'added 1/10/02
      //'DataSaved = False
      //'exit_save_data:
       //'   Exit Sub
;
      //'SaveError:
      //'msgbox "error saving record" & " " & Err.Number & " " & Err.Description
      //'Resume exit_save_data
;
;
//Exit //Sub;
}
function SendLetterButton_Click(){
//If //SaveRecordSub //Then;
    //Call //GetLetterMAnager;
//End //If;
  ;
;
}
function CheckIfConsentApproval(){//As //Boolean;
//If //Me.Agenda_Category = "Initial Submission" //Then;
     //If //Me.ActionCombo //Like "*Approv*" //Or //Me.ActionCombo //Like "*Accept*" //Or //Me.ActionCombo //Like "*Renew*" //Or //Me.ActionCombo //Like "*exped*" //Or //Me.ActionCombo = "Closed to Enrollment" //Then;
        //CheckIfConsentApproval = //True;
        //Exit //Function;
     //End //If;
    ;
//ElseIf //Me.Agenda_Category = "Procedure" //Then;
        //If //Me.Agenda_Condition1 = "Revised Consent" //Or //Me.Agenda_Condition2 = "Revised Consente" //Then;
            //If //Me.ActionCombo //Like "*Approv*" //Or //Me.ActionCombo //Like "*Accept*" //Or //Me.ActionCombo //Like "*Renew*" //Or //Me.ActionCombo //Like "*exped*" //Or //Me.ActionCombo = "Closed to Enrollment" //Then;
                //CheckIfConsentApproval = //True;
            //Exit //Function;
            //End //If;
        //End //If;
//End //If;
//CheckIfConsentApproval = //False;
}
function UpdateConsentDate(){
//On //Error //GoTo //err_UpdateConsentDate;
;
//could initialize hope here
;
//could initialize dbs here
//Set dbs;
//could initialize GoForIt here
;
//sqlSTR = "SELECT InformedConsent.[IRB#] FROM InformedConsent WHERE InformedConsent.[IRB#]= " //& "'" //& //Me.ID_ //& "'";
//Set GoForIt;
//If //GoForIt.RecordCount = //0 //Then;
    //'MsgBox "An Informed Consent Checklist Form has not been submitted." _
    //'& "Return to the Study Form, enter the Checklist, then return here to RE-Post", vbInformation
    //GoTo //exit_UpdateConsentDate;
//Else;
    //DoCmd.SetWarnings //False;
    //DoCmd.OpenQuery "ICApprovedDateUpdate";
    //DoCmd.SetWarnings //True;
//End //If;
;
//exit_UpdateConsentDate:;
//Exit //Function;
;
//err_UpdateConsentDate:;
//MsgBox //Err.Description;
//Resume //exit_UpdateConsentDate;
;
}
function CheckAllReadyOn(){//As //Boolean;
//On //Error //GoTo //ErrorCheck;
//could initialize hope here
//could initialize dbs here
//Set dbs;
//could initialize GoForIt here
//sqlSTR = "SELECT [agenda records].[IRB#], [agenda records].MeetingDate, [agenda records].[IRB Action Category]" //_;
//& " FROM [agenda records]WHERE [agenda records].[IRB#]= " //& "'" //& //[Forms]![_postmeetingform]![id#] //& "'" //& " AND [agenda records].MeetingDate <> " //& "#" //& //[Forms]![_postmeetingform]![NewAgendaDate] //& "#" //& " AND [agenda records].[IRB Action Category] Is Null AND [agenda records].MeetingDate > " //& "#" //& //[Forms]![_postmeetingform]![MeetingDate] //& "#";
;
;
;
//Set GoForIt;
//If //GoForIt.RecordCount //> //0 //Then;
        //If //YesNo("The //Study //is //already //on //the //agenda //for " & GoForIt![MeetingDate] & "//Do //you //still //want //to //put //on //this //date?") //Then;
            //CheckAllReadyOn = //False;
        //Else;
            //CheckAllReadyOn = //True;
        //End //If;
        //GoForIt.Close;
        //Exit //Function;
//Else;
    //GoForIt.Close;
    //CheckAllReadyOn = //False;
//End //If;
//Exit_Check:;
//Exit //Function;
;
//ErrorCheck:;
//MsgBox "Error Checking if Allready On  " //& //Err.Description //& "  " //& //Err.Number;
//Resume //Exit_Check;
;
}
function GetLetterMAnager(){
//On //Error //GoTo //Error_action;
var stDocName = '';
var stLinkCriteria = '';
;
;
//'MsgBox Me.textSpecialWording
//If //GetTemplateLocation("IRBAction") //Then;
        //'msgbox TemplateLocation
        //'msgbox TemplateName
        //If //TemplateName = "" //Then;
            //Exit //Sub;
        //Else;
            //DoCmd.SetWarnings //False;
            //'MsgBox [Forms]![_PostMeetingForm]![record_id]
            //DoCmd.OpenQuery //("ExportIRBActionWithCreateTable");
            //DoCmd.SetWarnings //True;
            ;
            //CustomLetter = //True;
        //End //If;
    //Else;
        //'msgbox "Do PROIRB SAE Letter"
        //CustomLetter = //False;
    //End //If;
    ;
;
;
;
;
;
;
;
;
;
//LetterName = "IRB Action Letter";
;
    //searchirbs = //False;
    //FromPostIRB = //True;
    //HoldIRBNo = //Me.[ID_];
    stLinkCriteria = "[IRB#]=" && "'" && HoldIRBNo && "'";
    stDocName;
    //DoCmd.SetWarnings //False;
    //HoldMeetingDate = //Me.Meeting_date;
    //HoldIRBAction = //Me.ActionCombo;
   //'Me.Visible = False
    location.href = "Form_"+stDocName+".php";//, //acNormal, //, //stLinkCriteria, //, //acWindowNormal;
//Exit_action:;
//Exit //Sub;
//Error_action:;
//MsgBox //Err.Description;
//Resume //Exit_action;
;
}
function Action_Report_Click(){
//On //Error //GoTo //Err_Action_Report_Click;
;
//If //AllArePosted() //Then;
    //DoCmd.RunCommand //acCmdSaveRecord;
    //GoTo //OpenMeetingMinutes;
//Else;
    ;
    //If //YesNo("All //Agenda //Items //have //not //been //posted.  " & "//Do //you //want //to //continue") //Then;
            //DoCmd.RunCommand //acCmdSaveRecord;
            //GoTo //OpenMeetingMinutes;
    //Else;
        //Exit //Sub;
    //End //If;
//End //If;
 ;
    ;
;
//OpenMeetingMinutes:;
   //Call //fCheckNewMembers;
    //DoCmd.SetWarnings //False;
    //DoCmd.OpenQuery "DuplicateAgendaHeadingtoMinutesHeading";
    //DoCmd.SetWarnings //True;
   ;
    ;
    ;
    location.href = "Form_Meeting Minutes Preparation.php";;
;
//Exit_Action_Report_Click:;
    //Exit //Sub;
;
//Err_Action_Report_Click:;
    //MsgBox "A database error has occured in Meeting Preparation-" //& "Write this message down and call ProIRB--" //& //Err.Description //& "  " //& //Err.Number;
    //Application.Quit;
;
}
function Command529_Click(){
//On //Error //GoTo //Err_Command529_Click;
;
    var stDialStr = '';
    //could initialize PrevCtl here
    //Const //ERR_OBJNOTEXIST = //2467;
    //Const //ERR_OBJNOTSET = //91;
;
    //Set PrevCtl;
    ;
    //If //TypeOf PrevCtl //Is //TextBox //Then;
      stDialStr;
    //ElseIf //TypeOf PrevCtl //Is //ListBox //Then;
      stDialStr;
    //ElseIf //TypeOf PrevCtl //Is //ComboBox //Then;
      stDialStr;
    //Else;
      stDialStr;
    //End //If;
    ;
    //Application.Run "utility.wlib_AutoDial"//, stDialStr;
;
//Exit_Command529_Click:;
    //Exit //Sub;
;
//Err_Command529_Click:;
    //If //(Err = //ERR_OBJNOTEXIST) //Or //(Err = //ERR_OBJNOTSET) //Then;
      //Resume //Next;
    //End //If;
    //MsgBox //Err.Description;
    //Resume //Exit_Command529_Click;
    ;
}
function Command530_Click(){
//On //Error //GoTo //Err_Command530_Click;
;
    var stDialStr = '';
    //could initialize PrevCtl here
    //Const //ERR_OBJNOTEXIST = //2467;
    //Const //ERR_OBJNOTSET = //91;
;
    //Set PrevCtl;
    ;
    //If //TypeOf PrevCtl //Is //TextBox //Then;
      stDialStr;
    //ElseIf //TypeOf PrevCtl //Is //ListBox //Then;
      stDialStr;
    //ElseIf //TypeOf PrevCtl //Is //ComboBox //Then;
      stDialStr;
    //Else;
      stDialStr;
    //End //If;
    ;
    //Application.Run "utility.wlib_AutoDial"//, stDialStr;
;
//Exit_Command530_Click:;
    //Exit //Sub;
;
//Err_Command530_Click:;
    //If //(Err = //ERR_OBJNOTEXIST) //Or //(Err = //ERR_OBJNOTSET) //Then;
      //Resume //Next;
    //End //If;
    //MsgBox //Err.Description;
    //Resume //Exit_Command530_Click;
    ;
}
function TExpireDate_AfterUpdate(){
//If //(DateDiff("m", //Me.Meeting_date, //Me.TExpireDate)) //> //13 //Then;
            //MsgBox "Expiration date 1 more than 1 year beyond the meeting date";
            //Me.TExpireDate.Value = //Me.TExpireDate.OldValue;
//End //If;
//Me.ExpirationDays = //(DateDiff("dddd", //Me.Meeting_date, //Me.TExpireDate));
//MsgBox //Me.ExpirationDays;
}
function VoteButton_Click(){
location.href = "Form_MeetingVote.php";;
;
;
;
}
function DoAddTabled(){
//On //Error //GoTo //err_DoAddTabled;
//If //IsNull(Me.Action_Condition1) //Or //Me.Action_Condition1 = " " //Then;
        //Me.Action_Condition1 = "None";
//End //If;
//If //IsNull(Me.Agenda_Explanation) //Or //Me.Agenda_Explanation = " " //Then;
    //Me.Agenda_Explanation = "None";
//End //If;
//If //IsNull(Me.AgendaReportStringText) //Or //Me.AgendaReportStringText = " " //Then;
    //Me.AgendaReportStringText = "None";
//End //If;
;
;
//'INSERT INTO [agenda records] ( [IRB#], SourceNumber, MeetingDate, [Record Type], [Agenda Category], [Agenda Condition1], [Agenda Condition2], UserName, IRBCode, AgendaReportString, [Agenda Explanation] )
//If //AddTabledAgenda(Me.ID_, //Me.SourceNumber, //_;
//Me.NewAgendaDate, //Me.Record_Type, //Me.Agenda_Category, //IIf(IsNull(Me.Agenda_Condition1), " "//, //Me.Agenda_Condition1), //_;
"Previously" //& //Me.ActionCombo, //_;
 //IIf(Me.Action_Condition1 = "NONE"//, " "//, //Me.Action_Condition1 //& ";"//) //& //IIf(Me.Agenda_Explanation = "NONE"//, " "//, //Me.Agenda_Explanation), //_;
"z "//, " z"//, "z "//, " z"//, "z "//, //0, //Date, //Me.cExpediteFlag, //Forms![newmenu]![UserName], //_;
//Forms![newmenu]![Location], //_;
//Me.AgendaReportStringText, " z"//, "z "//, //0) //Then;
//End //If;
//exit_DoAddTabled:;
//Exit //Sub;
//err_DoAddTabled:;
//MsgBox "Error trying to Table Agenda Item--" //& //Err.Description //& " " //& //Err.Source;
;
//Resume //exit_DoAddTabled;
;
}
function SetComboItem(SearchFor){//As //Integer;
//could initialize Maxinlist here
//could initialize Item here
Maxinlist;
Item;
//Do //Until Item //> Maxinlist //- //1;
    //If //Me.NewStudyStatus.ItemData(Item) = //SearchFor //Then;
        //SetComboItem = Item;
        //Exit //Function;
    //Else;
        Item;
    //End //If;
//Loop;
    ;
//MsgBox "Status - Argument not in list--" //& //SearchFor;
}
function AllArePosted(){//As //Boolean;
//On //Error //GoTo //ErrorCheck1;
;
//If //DCount("[IRB#]", "agendarecordsandIRB"//) //> //0 //Then;
;
;
        //AllArePosted = //False;
//Else;
         //AllArePosted = //True;
//End //If;
;
//exit_check1:;
//'GoForIt.Close
//Exit //Function;
;
//ErrorCheck1:;
//MsgBox "Error Checking if all have been posted " //& //Err.Description //& "  " //& //Err.Number;
//Application.Quit;
;
}
function SaveRecordSub(){//As //Boolean;
 //On //Error //GoTo //SaveError;
//If //IsNull(Me.ActionCombo) //Then;
    //MsgBox //("You //must //record //an //action //before //saving"), //vbInformation;
    //SaveRecordSub = //False;
    //Exit //Function;
//End //If;
;
//DataSaved = //True;
//DoCmd.RunCommand //acCmdSaveRecord;
 //SaveRecordSub = //True;
//DataSaved = //False;
;
;
//'***** new cyber 8/16/2010 *************
      //could initialize cmd here
//30    //Set cmd;
;
//'This record set is from query _postIRBMeetingResults. if you need more data, just look in there, then use Me.recordset("x")
  //If //CyberIRB_Flag //Then;
      //UpdateCyberProtocol //Me.[IRB#];
//'      MsgBox "updating"
//180        //cmd.ActiveConnection = //gclsCyberIRB.CyberConnection;
//190        //cmd.CommandText = "PIRB_" //& //gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") //& "_UpdateMeetingMinutes";
//200        //cmd.CommandType = //adCmdStoredProc;
//'
//210        //cmd.Parameters("@varCustNum") = //gclsCyberIRB.mcolCyberConfigAttributes("CustNum");
//220        //cmd.Parameters("@varIRBNum") = //Me.Recordset("IRB#")  //' field number (6)
//230        //cmd.Parameters("@varRefNum") = //Me.Recordset("RefNum")  //' field number (49)
//235        //cmd.Parameters("@varStudyStatus") = //Me.Recordset("Study //Status")  //'    This is the Protocol study status, not the agenda items
//240        //cmd.Parameters("@varAgendaGroup") = //Me.Recordset("AgendaGroup")  //'field number (4)
//250        //cmd.Parameters("@varMeetingDate") = //Me.Recordset("MeetingDate") //'(11)
//260        //cmd.Parameters("@varRecordType") = //Me.Recordset("Record //Type") //'(3), (44) SAE, CPA, Initial
//270        //cmd.Parameters("@varAgendaCategory") = //Me.Recordset("Agenda //Category") //'(42)
//280        //cmd.Parameters("@varAgendaCondition1") = //Me.Recordset("Agenda //Condition1") //'(13)
//290        //cmd.Parameters("@varAgnedaCondition2") = //Me.Recordset("Agenda //Condition2") //'(14)
//300        //cmd.Parameters("@varActionCategory") = //Me.Recordset("IRB //Action //Category") //'(31)
//310        //cmd.Parameters("@varExpediteFlag") = //Me.Recordset("ExpediteFlag") //'(12)
//320        //cmd.Execute;
  //End //If;
//'***************************************
;
;
;
;
;
//exit_save_data:;
    //Exit //Function;
;
//SaveError:;
//MsgBox "error saving record" //& " " //& //Err.Number //& " " //& //Err.Description;
//Resume //exit_save_data;
}
function DealWithExpedites(){
var CPAStr = '';
//'this routine deals with expedited items
//If //Me.ActionCombo //Like "*Approv*" //Or //Me.ActionCombo //Like "*Accept*" //Or //Me.ActionCombo //Like "*Renew*" //Or //Me.ActionCombo //Like "*exped*" //Or //Me.ActionCombo = "Closed to Enrollment" //Then;
//'If Me.ActionCombo Like "*Approv*" Or Me.ActionCombo Like "*Renew*" Or Me.ActionCombo Like "*exped*" Then   ' gets contingent and regular approval
    //If //Me.Record_Type = "Initial" //Then;
        //If //Not //IsNull(Me.Last_Renewal_Date) //Then //GoTo //CheckAction  //'added 5/23/08
        //If //IsNull(Me.Approval_Date___original_) //Then;
            //Me.Approval_Date___original_ = //Me.Meeting_date;
            //Me.Last_Renewal_Date = //Me.Meeting_date;
        //Else;
            //Me.Last_Renewal_Date = //Me.Approval_Date___original_;
        //End //If;
        //Call //NewExpirationDate;
        //GoTo //CheckAction;
    //End //If;
    //If //Me.Agenda_Condition1 = "Renewal" //Then;
        CPAStr = "[CPAnumber] = " && Forms![_postmeetingform]![SourceNumber];
        //'Me.Last_Renewal_Date = DMax("[Date of Change]", "CPA", CPAStr)
        //'If IsNull(Me.Last_Renewal_Date) Then
        //If //IsNull(DMax("[Date //of //Change]", "CPA"//, //CPAStr)) //Then;
            //MsgBox "This CPA must be corrected to add a Date of Change for the Renewal. Edit the Study and view this CPA entering a Date of Change.";
            //Me.ActionCombo = //Me.ActionCombo.OldValue;
            //Exit //Sub;
        //Else;
            //Me.Last_Renewal_Date = //DMax("[Date //of //Change]", "CPA"//, //CPAStr);
        //End //If;
        ;
        //Call //NewExpirationDate;
    //End //If;
    ;
    //If //Me.ActionCombo = "Modifications Approved" //And //Me.NewStudyStatus = "Pending Modifications" //Then;
        //Me.Last_Renewal_Date = //Me.DateFirstReviewed;
        //Me.Approval_Date___original_ = //DMax("[date //of //change]", "CPA"//, "[cpanumber]=" //& //Me.SourceNumber);
        //Me.Last__IRB_date = //Me.Meeting_date;
        //'MsgBox DMax("[date of change]", "CPA", "[cpanumber]=" & Me.SourceNumber)
         //Call //NewExpirationDate;
    //End //If;
    ;
//End //If;
//CheckAction:;
//'IRB Tabled  restores original values
//If //Me.ActionCombo //Like "*Tabled*" //Then;
    //HoldIRBAction = //Me.ActionCombo;
    //Me.Undo;
    //If //Me.SourceNumber = "New Appl" //And //IsNull(Me.Last__IRB_date) //Then;
        //Me.DateFirstReviewed = //Me.Meeting_date;
    //End //If;
    //Me.ActionCombo = //HoldIRBAction;
    //Me.Last__IRB_date = //Me.Meeting_date;
    //Exit //Sub;
//Else;
    //Me.Last__IRB_date = //Me.Meeting_date;
       ;
   //If //IsNull(Me.ActionCombo.Column(1)) //Then;
        //Me.Undo;
        //Exit //Sub;
   //End //If;
   //If //Me.ActionCombo.Column(1) = "(reserved for system)" //Then;
        //Exit //Sub;
    //Else;
        //If //Me.ActionCombo.Column(2) = "Closed" //Then;
            //Me.NewStudyActive = //Me.NewStudyActive.ItemData(1);
        //Else;
            //Me.NewStudyActive = //Me.NewStudyActive.ItemData(0);
        //End //If;
        //Me.NewStudyStatus.Requery;
        //Me.NewStudyStatus = //Me.NewStudyStatus.ItemData(SetComboItem(Me.ActionCombo.Column(1)));
    //End //If;
//End //If;
//If //Me.NewStudyActive = "Closed" //Then //Me.TExpireDate = "";
}
function DealWithExempts(){
//'deals with exempt agenda category only
;
//'IRB Tabled  restores original values
//If //Me.ActionCombo //Like "*Tabled*" //Then;
    //HoldIRBAction = //Me.ActionCombo;
    //Me.Undo;
    //If //Me.SourceNumber = "New Appl" //And //IsNull(Me.Last__IRB_date) //Then;
        //Me.DateFirstReviewed = //Me.Meeting_date;
    //End //If;
    //Me.ActionCombo = //HoldIRBAction;
    //Me.Last__IRB_date = //Me.Meeting_date;
    //Exit //Sub;
//Else;
    //Me.Last__IRB_date = //Me.Meeting_date;
    //If //IsNull(Me.ActionCombo) //Then;
        //Me.Undo;
    //End //If;
    //If //IsNull(Me.ActionCombo.Column(1)) //Then;
        //Me.ActionCombo.Undo;
        //Exit //Sub;
    //End //If;
   ;
//End //If;
;
;
//If //IsNull(Me.ActionCombo.Column(1)) //Then;
        //Me.Undo;
        //Exit //Sub;
   //End //If;
   //If //Me.ActionCombo.Column(1) = "(reserved for system)" //Then;
        //Exit //Sub;
    //Else;
        //If //Me.ActionCombo.Column(2) = "Closed" //Then;
            //Me.NewStudyActive = //Me.NewStudyActive.ItemData(1);
        //Else;
            //Me.NewStudyActive = //Me.NewStudyActive.ItemData(0);
        //End //If;
        //Me.NewStudyStatus.Requery;
        //Me.NewStudyStatus = //Me.NewStudyStatus.ItemData(SetComboItem(Me.ActionCombo.Column(1)));
//End //If;
;
//If //Me.Record_Type = "Initial" //Then;
   //If //Me.ActionCombo //Like "*Approv*" //Or //Me.ActionCombo //Like "*Accept*" //Or //Me.ActionCombo //Like "*exem*" //Then;
      //If //IsNull(Me.Approval_Date___original_) //Then;
        //Me.Approval_Date___original_ = //Me.Meeting_date;
      //End //If;
   //End //If;
//End //If;
;
;
//If //Me.ActionCombo //Like "*to open*" //Then;
    //If //IsNull(Me.Last_Renewal_Date) //Then;
            //Me.Last_Renewal_Date = //Me.Approval_Date___original_;
            //Call //NewExpirationDate;
    //Else;
            //Me.Last_Renewal_Date = //Me.Meeting_date;
            //Call //NewExpirationDate;
    //End //If;
    ;
//End //If;
//'deal with Navy override
//If //Not //ExemptOverrideSwitch //Then;
    //Exit //Sub;
//Else;
  //If //Me.Record_Type = "Initial" //Or //Me.Record_Type = "Renewal" //Then;
    //If //Me.ActionCombo //Like "*Exempted*" //Then;
        //If //IsNull(Me.Last_Renewal_Date) //Then;
            //Me.Last_Renewal_Date = //Me.Approval_Date___original_;
            //Call //NewExpirationDate;
        //Else;
            //Me.Last_Renewal_Date = //Me.Meeting_date;
            //Call //NewExpirationDate;
        //End //If;
    //End //If;
  //End //If;
//End //If;
}
function bGOTO_Click(){
//On //Error //GoTo //Err_bGOTO_Click;
;
//Me.ID_.SetFocus;
    //'Screen.PreviousControl.SetFocus
    //DoCmd.RunCommand //acCmdFind;
    //'DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70
;
//Exit_bGOTO_Click:;
    //Exit //Sub;
;
//Err_bGOTO_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_bGOTO_Click;
    ;
}
function Command554_Click(){
//On //Error //GoTo //Err_Command554_Click;
;
;
    //Screen.PreviousControl.SetFocus;
    //DoCmd.FindNext;
;
//Exit_Command554_Click:;
    //Exit //Sub;
;
//Err_Command554_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command554_Click;
    ;
}
function UpdateCyberIRB(){
//could initialize cmd here
//30    //Set cmd;
;
//'This record set is from query _postIRBMeetingResults. if you need more data, just look in there, then use Me.recordset("x")
;
;
  //If //CyberIRB_Flag //Then;
      //UpdateCyberProtocol //Me.[IRB#];
;
;
//180        //cmd.ActiveConnection = //gclsCyberIRB.CyberConnection;
//190        //cmd.CommandText = "PIRB_" //& //gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") //& "_UpdateMeetingMinutes";
//200        //cmd.CommandType = //adCmdStoredProc;
//'
//210        //cmd.Parameters("@varCustNum") = //gclsCyberIRB.mcolCyberConfigAttributes("CustNum");
//220        //cmd.Parameters("@varIRBNum") = //Me.Recordset("IRB#")  //' field number (6)
//230        //cmd.Parameters("@varRefNum") = //Me.Recordset("RefNum")  //' field number (49)
//235        //cmd.Parameters("@varStudyStatus") = //Me.Recordset("Study //Status")  //'    This is the Protocol study status, not the agenda items
//240        //cmd.Parameters("@varAgendaGroup") = //Me.Recordset("AgendaGroup")  //'field number (4)
//250        //cmd.Parameters("@varMeetingDate") = //Me.Recordset("MeetingDate") //'(11)
//260        //cmd.Parameters("@varRecordType") = //Me.Recordset("Record //Type") //'(3), (44) SAE, CPA, Initial
//270        //cmd.Parameters("@varAgendaCategory") = //Me.Recordset("Agenda //Category") //'(42)
//280        //cmd.Parameters("@varAgendaCondition1") = //Me.Recordset("Agenda //Condition1") //'(13)
//290        //cmd.Parameters("@varAgnedaCondition2") = //Me.Recordset("Agenda //Condition2") //'(14)
//300        //cmd.Parameters("@varActionCategory") = //Me.Recordset("IRB //Action //Category") //'(31)
//310        //cmd.Parameters("@varExpediteFlag") = //Me.Recordset("ExpediteFlag") //'(12)
//320        //cmd.Execute;
  //End //If;
;
;
//'
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Active();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>





    <input type='text' id=ID# value='empty' style='position:absolute; left:12; top:88; width:162; height:36'></input>

    <label id=Text15 value='IRB#:' style='position:absolute; left:60; top:64; width:57; height:24; visibility:'>IRB#:</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:331; top:70; width:700; height:48'></input>

    <label id=Text27 value='Protocol#\015\012  Title:' style='position:absolute; left:228; top:70; width:91; height:43; visibility:'>Protocol#\015\012  Title:</label>
    <input type='text' id=Approval Date  (original) value='empty' style='position:absolute; left:939; top:130; width:126; height:25'></input>

    <label id=Text31 value='Original Approval Date:' style='position:absolute; left:708; top:130; width:216; height:24; visibility:'>Original Approval Date:</label>
    <input type='text' id=Last Renewal Date value='empty' style='position:absolute; left:939; top:163; width:126; height:25'></input>

    <label id=Text33 value='Last IRB Renewal Date:' style='position:absolute; left:712; top:163; width:214; height:24; visibility:'>Last IRB Renewal Date:</label>
    <select name='Investigator ID' Size='12' style='position:absolute; left:330; top:130; width:330; height:0'>
    <?php
        $serverName = $SVNM;
        $connectionInfo = array( "Database"=>$DB, "UID"=>$UID, "PWD"=>$PWD);
        $conn = sqlsrv_connect( $serverName, $connectionInfo);

        if( $conn ) {
        }
        else{
            echo "Connection could not be established.<br />";
            die( print_r( sqlsrv_errors(), true));
        }

        $sql = sqlsrv_query($conn, "SELECT [StudyStatusCodes1].[Study Status] FROM StudyStatusCodes1 WHERE ((([StudyStatusCodes1].[StudyActiveCode])=[Forms]![_PostMeetingForm]![NewStudyActive])) ORDER BY [StudyStatusCodes1].[Seq]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Study Status']."'".">".$row['Study Status']."</Option>";
        }
    ?>
    </select>

    <label id=PI LAbel value='P.I.' style='position:absolute; left:282; top:130; width:39; height:24; visibility:'>P.I.</label>
    <label id=Label1 value='Agenda Detail' style='position:absolute; left:0; top:6; width:205; height:40; visibility:'>Agenda Detail</label>
    <input type='text' id=Agenda Explanation value='empty' style='position:absolute; left:208; top:439; width:282; height:55'></input>

    <label id=Label465 value='Additional/ Other:' style='position:absolute; left:36; top:439; width:160; height:24; visibility:'>Additional/ Other:</label>
    <label id=ReasonLabel value='------Reason Study is on Agenda-------' style='position:absolute; left:37; top:295; width:468; height:34; visibility:'>------Reason Study is on Agenda-------</label>

    <select name='Agenda Condition1' Size='19' style='position:absolute; left:208; top:378; width:282; height:0'>    </select>

    <label id=Label464 value='Reason1:' style='position:absolute; left:111; top:382; width:90; height:24; visibility:'>Reason1:</label>

    <select name='Action Condition1' Size='1' style='position:absolute; left:718; top:378; width:282; height:0'>    </select>

    <label id=Label488 value='Condition 1:' style='position:absolute; left:600; top:378; width:111; height:24; visibility:'>Condition 1:</label>
    <label id=Label489 value='------IRB Meeting Action--------' style='position:absolute; left:544; top:295; width:469; height:34; visibility:'>------IRB Meeting Action--------</label>
    <select name='Renewal  Cycle' Size='17' style='position:absolute; left:555; top:190; width:57; height:25'>
        <Option value='12'>12</option>
        <Option value='11'>11</option>
        <Option value='10'>10</option>
        <Option value='9'>9</option>
        <Option value='8'>8</option>
        <Option value='7'>7</option>
        <Option value='6'>6</option>
        <Option value='5'>5</option>
        <Option value='4'>4</option>
        <Option value='3'>3</option>
        <Option value='2'>2</option>
        <Option value='1'>1</option>
    </select>

    <label id=Label490 value='Renewal Cycle:' style='position:absolute; left:546; top:160; width:141; height:24; visibility:'>Renewal Cycle:</label>
    <select name='Action Condition2' Size='2' style='position:absolute; left:718; top:408; width:282; height:0'>    </select>

    <label id=Label492 value='Condition 2:' style='position:absolute; left:600; top:408; width:111; height:24; visibility:'>Condition 2:</label>
    <input type='text' id=Meeting_date value='empty' style='position:absolute; left:957; top:12; width:0; height:30'></input>
    <input type='text' id=Date Received value='empty' style='position:absolute; left:330; top:163; width:0; height:0'></input>

    <label id=Label494 value='Date Received (Initial Submission):' style='position:absolute; left:12; top:160; width:307; height:24; visibility:'>Date Received (Initial Submission):</label>
    <label id=Label497 value='Meeting Date' style='position:absolute; left:831; top:18; width:123; height:24; visibility:'>Meeting Date</label>
    <input type='text' id=Last  IRB date value='empty' style='position:absolute; left:330; top:231; width:0; height:0'></input>

    <label id=Label498 value='Date Last seen by IRB:' style='position:absolute; left:112; top:228; width:207; height:24; visibility:'>Date Last seen by IRB:</label>
    <select name='Agenda Condition2' Size='20' style='position:absolute; left:208; top:408; width:282; height:0'>    </select>

    <label id=Label501 value='Reason2:' style='position:absolute; left:109; top:412; width:90; height:24; visibility:'>Reason2:</label>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:264; top:12; width:561; height:48'></input>

    <label id=Label461 value='IRB:' style='position:absolute; left:210; top:12; width:46; height:24; visibility:'>IRB:</label>
    <select name='ActionCombo' style='position:absolute; left:718; top:343; width:282; height:30'>
    <?php
        $serverName = $SVNM;
        $connectionInfo = array( "Database"=>$DB, "UID"=>$UID, "PWD"=>$PWD);
        $conn = sqlsrv_connect( $serverName, $connectionInfo);

        if( $conn ) {
        }
        else{
            echo "Connection could not be established.<br />";
            die( print_r( sqlsrv_errors(), true));
        }

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_Select PIs for irb#].[Common Contact ID], [_Select PIs for irb#].[display name] FROM [_Select PIs for irb#]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select>

    <label id=CPAAction_Label value='Action Taken' style='position:absolute; left:586; top:343; width:123; height:24; visibility:'>Action Taken</label>
    <!--     <input type='text' id=longdate value='empty' style='position:absolute; left:408; top:54; width:156; height:54'></input> -->

    <label id=Label507 value='Text506:' style='position:absolute; left:330; top:54; width:70; height:24; visibility:'>Text506:</label>
    <!--     <input type='text' id=NewAgendaDate value='empty' style='position:absolute; left:210; top:66; width:162; height:48'></input> -->
    <input type='text' id=Texpiredate value='empty' style='position:absolute; left:939; top:196; width:126; height:25'></input>

    <label id=Label518 value='Expiration Date:' style='position:absolute; left:780; top:196; width:145; height:24; visibility:'>Expiration Date:</label>
    <input type='text' id=SourceNumber value='empty' style='position:absolute; left:57; top:499; width:0; height:0'></input>

    <label id=Label519 value='Internal Source' style='position:absolute; left:58; top:469; width:139; height:24; visibility:'>Internal Source</label>
    <!--     <input type='text' id=record_id value='empty' style='position:absolute; left:174; top:54; width:0; height:0'></input> -->

    <label id=Label520 value='record_id:' style='position:absolute; left:114; top:54; width:78; height:24; visibility:'>record_id:</label>
    <input type='text' id=Action Explanation value='empty' style='position:absolute; left:718; top:439; width:282; height:79'></input>

    <label id=Label524 value='Action Explanation:' style='position:absolute; left:537; top:439; width:174; height:24; visibility:'>Action Explanation:</label>
    <!--     <input type='checkbox' id=cExpediteFlag value='False' style='position:absolute; left:300; top:39; width:0; height:0'></input> -->

    <!--     <label id=Label525 value='ExpediteFlag' style='position:absolute; left:323; top:36; width:99; height:24; visibility:hidden'>ExpediteFlag</label> -->
    <label id=Label500 value='Months' style='position:absolute; left:624; top:192; width:70; height:24; visibility:'>Months</label>
    <input type='text' id=AgendaReportStringText value='empty' style='position:absolute; left:208; top:499; width:282; height:0'></input>
    <input type='button' id=DiscussionButton value='D\015\012I\015\012S\015\012C\015\012U\015\012S\015\012S\015\012I\015\012O\015\012' style='position:absolute; left:1032; top:288; width:39; height:256'></input>
    <input type='button' id=VoteButton value='V\015\012O\015\012T\015\012E\015\012' style='position:absolute; left:1072; top:288; width:30; height:256' onclick='VoteButton_Click();'></input>
    <!--     <input type='text' id=holdDiscussion value='empty' style='position:absolute; left:426; top:102; width:102; height:0'></input> -->

    <label id=Label542 value='Text541:' style='position:absolute; left:282; top:102; width:70; height:24; visibility:'>Text541:</label>
    <!--     <input type='text' id=HoldVote value='empty' style='position:absolute; left:588; top:54; width:120; height:36'></input> -->

    <label id=Label544 value='Text541:' style='position:absolute; left:444; top:54; width:70; height:24; visibility:'>Text541:</label>
    <input type='button' id=Save Data value='&Save Data' style='position:absolute; left:874; top:555; width:135; height:45' onclick='Save Data_Click();'></input>
    <input type='button' id=ButtonReturn value='&Return' style='position:absolute; left:1009; top:555; width:93; height:45' onclick='ButtonReturn_Click();'></input>
    <input type='button' id=SendLetterButton value='Send &Correspondence' style='position:absolute; left:615; top:555; width:259; height:45' onclick='SendLetterButton_Click();'></input>
    <input type='button' id=NextButton value='Next' style='position:absolute; left:115; top:555; width:87; height:45' onclick='NextButton_Click();'></input>
    <input type='button' id=PreviousButton value='Previous' style='position:absolute; left:0; top:555; width:117; height:45' onclick='PreviousButton_Click();'></input>
    <input type='button' id=Action Report value='&Meeting Minutes' style='position:absolute; left:354; top:555; width:261; height:45' onclick='Action Report_Click();'></input>
    <!--     <input type='text' id=Record Type value='empty' style='position:absolute; left:174; top:120; width:0; height:0'></input> -->

    <label id=Label545 value='Record Type:' style='position:absolute; left:30; top:120; width:105; height:24; visibility:'>Record Type:</label>
    <input type='text' id=Agenda Category value='empty' style='position:absolute; left:208; top:343; width:282; height:0'></input>

    <label id=Label463 value='Agenda Category:' style='position:absolute; left:40; top:346; width:160; height:24; visibility:'>Agenda Category:</label>
    <select name='NewStudyStatus' Size='6' style='position:absolute; left:793; top:253; width:276; height:30'>
    <?php
        $serverName = $SVNM;
        $connectionInfo = array( "Database"=>$DB, "UID"=>$UID, "PWD"=>$PWD);
        $conn = sqlsrv_connect( $serverName, $connectionInfo);

        if( $conn ) {
        }
        else{
            echo "Connection could not be established.<br />";
            die( print_r( sqlsrv_errors(), true));
        }

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [IRB Actions1].[IRBAction], [IRB Actions1].[StudyStatus], [StudyStatusCodes].[StudyActiveCode] FROM StudyStatusCodes RIGHT JOIN [IRB Actions1] ON [StudyStatusCodes].[Study Status]=[IRB Actions1].[StudyStatus] ORDER BY [IRB Actions1].[SortSeq]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['IRBAction']."'".">".$row['IRBAction']."</Option>";
        }
    ?>
    </select>

    <label id=Label441 value='Study Status:' style='position:absolute; left:852; top:228; width:123; height:24; visibility:'>Study Status:</label>
    <select name='NewStudyActive' Size='26' style='position:absolute; left:643; top:253; width:0; height:30'>
        <Option value='Open'>Open</option>
        <Option value='Closed'>Closed</option>
    </select>

    <label id=Label454 value='Active?:' style='position:absolute; left:667; top:228; width:79; height:24; visibility:'>Active?:</label>
    <input type='text' id=DateFirstReviewed value='empty' style='position:absolute; left:330; top:196; width:0; height:0'></input>

    <label id=Label548 value='First IRB Review:' style='position:absolute; left:112; top:193; width:207; height:24; visibility:'>First IRB Review:</label>
    <!--     <input type='text' id=MeetingKey value='empty' style='position:absolute; left:103; top:174; width:31; height:21'></input> -->

    <label id=Label550 value='Text549:' style='position:absolute; left:54; top:174; width:70; height:24; visibility:'>Text549:</label>
    <input type='button' id=bGOTO value='Find Record' style='position:absolute; left:202; top:555; width:151; height:45' onclick='bGOTO_Click();'></input>
    <input type='text' id=AgendaLineNumber value='empty' style='position:absolute; left:436; top:529; width:0; height:25'></input>

    <label id=Label552 value='Item #' style='position:absolute; left:366; top:529; width:63; height:24; visibility:'>Item #</label>
    <input type='text' id=AgendaGroup value='empty' style='position:absolute; left:204; top:528; width:0; height:0'></input>

    <label id=Label553 value='AgendaGroup:' style='position:absolute; left:70; top:528; width:126; height:24; visibility:'>AgendaGroup:</label>
    <input type='button' id=cmdAddWording value='Add wording to explanation' style='position:absolute; left:480; top:238; width:150; height:51' onclick='cmdAddWording_Click();'></input>
    <input type='text' id=ListriskCat value='empty' style='position:absolute; left:169; top:259; width:306; height:0'></input>

    <label id=Label536 value='Risk Category' style='position:absolute; left:19; top:259; width:138; height:24; visibility:'>Risk Category</label>
    <!--     <input type='text' id=txtSpecial Wording value='empty' style='position:absolute; left:408; top:36; width:0; height:76'></input> -->

    <label id=Label555 value='Special Wording:' style='position:absolute; left:264; top:36; width:129; height:24; visibility:'>Special Wording:</label>
    <!--     <input type='text' id=ExpirationDays value='empty' style='position:absolute; left:522; top:372; width:36; height:18'></input> -->
    <input type='text' id=Text560 value='empty' style='position:absolute; left:18; top:132; width:108; height:0'></input>
  </body>
</html>
