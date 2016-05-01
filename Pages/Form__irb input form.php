<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function MoreThan1Dash(str){//As //Boolean;
;
//'Lee Marsh;   8/21/2001
//'Returns TRUE if the supplied string contains
//'more than two dashes (2 or more)
;
;
    //could initialize Count here
    //could initialize i here
    ;
    Count;
    ;
    //For i;
        //If //Mid(str, //i, //1) = "-" //Then Count;
    //Next i;
    ;
    //MoreThan1Dash = Count //> //1;
        ;
}
function AddNewSponsor(){
var stDocName = '';
    var stLinkCriteria = '';
    //HoldSponsor = //Me.[Sponsor //ID];
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.PI.SetFocus;
    //Me.[Sponsor //ID] = //HoldSponsor;
    //Me.Sponsor.RowSourceType = "Table/Query";
;
;
}
function NewExpirationDate(){
//If //IsNull(Me.Approval_Date___original_) //And //RenewAnniversary //Then;
    //'msgbox "Original Approval Date cannot be blank"
    //'Me.Approval_Date___original_.SetFocus
    //Exit //Sub;
//End //If;
//If //Not //IsNull(Me.Last_Renewal_Date) //Then;
        //If //RenewAnniversary //Then  //' CLIENT USES ANNIVERSARY DATE FOR EXPIRATION
            //CompareDate = //Me.Approval_Date___original_;
            //Do //Until //CompareDate //> //Me.Last_Renewal_Date //+ //32;
               //CompareDate = //DateAdd("m", //Me.Renewal_Cycle, //CompareDate);
            //Loop;
            //Me.TExpireDate = //CompareDate;
        //Else    //' CLIENT USES Renewal Meeting date plus 12 months forDATE FOR EXPIRATION
   ;
            //HoldMeetingDate = //DateAdd("m", //Me.Renewal_Cycle, //Me.Last_Renewal_Date) //- //1;
            //'msgbox HoldMeetingDate
            //Me.TExpireDate = //HoldMeetingDate;
        //End //If;
//Else;
    ;
    //If //Not //IsNull(Me.Approval_Date___original_) //Then;
        //If //RenewAnniversary //Then  //' CLIENT USES ANNIVERSARY DATE FOR EXPIRATION
            //CompareDate = //Me.Approval_Date___original_;
            //Do //Until //CompareDate //> //Date;
               //CompareDate = //DateAdd("m", //Me.Renewal_Cycle, //CompareDate);
            //Loop;
            //Me.TExpireDate = //CompareDate;
        //Else //' CLIENT USES Renewal Meeting date plus 12 months forDATE FOR EXPIRATION
         //HoldMeetingDate = //DateAdd("m", //Me.Renewal_Cycle, //Me.Approval_Date___original_) //- //1;
         //Me.TExpireDate = //HoldMeetingDate;
        //End //If;
    //End //If;
//End //If;
;
}
function CheckDataError(){//As //Boolean;
;
;
//10    //CheckDataError = //False;
//20    //If //Me.NewRecord //And //Not //gclsCyberIRB.gbolBridgeProtocolAlreadyThere //Then  //' Cyberbridge 090117
//30        //checkirbno.Requery;
//40        //If //Not //IsNull(Me.checkirbno.Column(0)) //Then;
//50            //MsgBox "Study Number already on file"//, //vbInformation;
//60            //CheckDataError = //True;
//70        //End //If;
//80    //End //If;
//NoProtocolNum:;
//90    //If //Me.[IRB#] = " " //Or //IsNull(Me.[IRB#]) //Or //IsNull(Me.Protocol_Number___Title) //Or //(Me.Protocol_Number___Title = ""//) //Then;
//100      //MsgBox "The IRB#  or the Protocol Title cannot be blank"//, //vbInformation;
//110     //CheckDataError = //True;
        //'***** Cyberbridge 090117 start *****
//120   //ElseIf //CyberIRB_Flag //Then //'new code that now checks if the Bridge has this IRB# waiting to come into "IRB - Research  Database" as IRB# 'probably want to if a flag here if cyber code is enabled ALTHOUGH if this Org doesn't have Cyber gclsCyberIRB.IRBNumWaitingInBridge will return False ' Terry 090110   I had to comment this out for now, it wouldn't let me bring in any protocols from the bridge Terry 090110 added CyberFlag 100725
          var sstrCyberRefNum = '';
//130          //If //IsNull(txtcyberrefnum) //Then;
//'140                 sstrCyberRefNum = ""
//140                 sstrCyberRefNum = "NR" && Me.[IRB#] 'Terry 7/28/2010  change with Nikki's at meridian through gotomeeting;
;
//150          //Else;
//160                 sstrCyberRefNum;
//170          //End //If;
//180          //If //IsNull(sstrCyberRefNum) //Or //Trim(sstrCyberRefNum) = "" //Then //GoTo //NoProtocolNum;
//190          //If //gclsCyberIRB.IRBNumWaitingInBridge(Me.[IRB#], //sstrCyberRefNum) //Then;
//200              //MsgBox "This Study Number is being used by a study waiting in the Cyber Bridge to come into ProIRB. Select OK then Change Study Number or Return without saving Changes"//, //vbInformation;
//210              //CheckDataError = //True;
//220   //End //If;
         //'***** Cyberbridge 090117 end   *****
;
;
;
;
//230   //End //If;
//240   //If //IsNull(Me.Sponsor) //Then;
//250      //MsgBox "You must assign a Sponsor.  Select  '...Unknown'  if you do not know the Sponsor as yet"//, //vbInformation;
//260    //CheckDataError = //True;
//270   //End //If;
//280   //If //IsNull(Me.PI) //Then;
//290      //MsgBox "You must assign a P.I.  Select  '...Unknown'  if you do not know the P.I. as yet"//, //vbInformation;
//300      //CheckDataError = //True;
//310   //End //If;
//320   //If //IsNull(Me.Date_Received) //Then;
//330       //MsgBox "Date Received cannot be blank"//, //vbInformation;
//340          //CheckDataError = //True;
//350   //End //If;
;
      //'checking for exempt or expeditecite and Status not like Exempt or Expedite
//360   //If //IsNull(Me.ExemptCite) //Or //Me.ExemptCite //Like "*not exempt*" //Then;
//370       //If //Me.[Study //Status] //Like "*exempt*" //Then;
//380           //MsgBox "Study Status conflicts with the Exempt Category. Either the Status or Exempt Category is wrong";
//390           //CheckDataError = //True;
//400       //End //If;
//410   //Else;
//420       //If //Not //Me.[Study //Status] //Like "*exempt*" //Then;
//430           //MsgBox "Study Status conflicts with the Exempt Category. Either the Status or Exempt Category is wrong";
//440           //CheckDataError = //True;
//450       //End //If;
//460   //End //If;
//470   //If //IsNull(Me.ExpediteCite) //Or //Me.ExpediteCite //Like "* full*" //Then;
//480       //If //Me.[Study //Status] //Like "*exped*" //Then;
//490           //MsgBox "Study Status conflicts with the Expedite Category. Either the Status or Expedite Category is wrong";
//500           //CheckDataError = //True;
//510       //End //If;
//520   //Else;
//530       //If //Not //Me.[Study //Status] //Like "*exped*" //Then;
//540           //MsgBox "Study Status conflicts with the Expedite Category. Either the Status or Expedite Category is wrong";
//550           //CheckDataError = //True;
//560       //End //If;
//570   //End //If;
}
function AgendaDate_DblClick(Cancel){
var stDocName = '';
var stLinkCriteria = '';
//FromIRB = //True;
  ;
   //StudySelected = //False;
 //CancelUpdate = //False;
    ;
//'Me.Visible = False
 location.href = "Form_CheckAgendaRecordStatus1.php";//, //, //, //, //, //acDialog;
 ;
 ;
;
}
function Cal_Updated(Code){
//End //Sub;
;
//Private //Sub //Approval_Date___original__AfterUpdate();
//If //IsNull(Me.Approval_Date___original_) //Then;
    //Me.TExpireDate = "";
    //Me.Last_Renewal_Date = "";
//Else;
    //If //Me.NewRecord //Then //Me.Last_Renewal_Date = //Me.Approval_Date___original_ //' aded 5/23/08
//End //If;
//If //Me.[Study //Status] //Like "*exem*" //And //Not //ExemptOverrideSwitch //Then //Exit //Sub;
//'If Me.[Study Status] Like "*exem*" Then Exit Sub
//Call //NewExpirationDate;
//Me.TExpireDate = //Me.TExpireDate.Value;
;
;
}
function Approval_Date___original__Change(){
//If //IsNull(Me.Approval_Date___original_) //Then;
    //Me.TExpireDate = "";
    //Me.Last_Renewal_Date = "";
//Else;
    //If //Me.NewRecord //Then //Me.Last_Renewal_Date = //Me.Approval_Date___original_  //' aded 5/23/08
//End //If;
//If //Me.[Study //Status] //Like "*exem*" //And //Not //ExemptOverrideSwitch //Then //Exit //Sub;
//'If Me.[Study Status] Like "*exem*" Then Exit Sub
//Call //NewExpirationDate;
//Me.TExpireDate = //Me.TExpireDate.Value;
;
}
function bSendInvoice_Click(){
var stDocName = '';
var stLinkCriteria = '';
;
//On //Error //GoTo //Error_bSendInvoice;
//If //CheckDataError //Then;
     //Exit //Sub;
//End //If;
//CancelUpdate = //False;
;
;
//If //Not //WasNew //Then //'done for Chla
    //DoCmd.Close //acForm, "_irb input form"//, //acSaveYes   //'done for CHLA
    location.href = "Form__irb input form.php";//, //acNormal //'Done for CHLA
//End //If //'done for Chla
;
;
;
//If //GetTemplateLocation("Invoice") //Then;
        //If //TemplateName = "" //Then;
            //Exit //Sub;
        //Else;
            //DoCmd.SetWarnings //False;
            //DoCmd.OpenQuery //("ExportStudySpecificWithCreateTable");
            //DoCmd.SetWarnings //True;
            //CustomLetter = //True;
        //End //If;
//Else;
    //CustomLetter = //False;
//End //If;
;
;
    ;
    ;
    ;
    ;
//FromIRB = //False;
//If //Not //IsNull(Forms![_irb //input //form]!AgendaDate) //Then  //'Done for CHLA
        //HoldMeetingDate = //Forms![_irb //input //form]!AgendaDate //'Done for CHLA
    //Else //'Done for CHLA
        //HoldMeetingDate = "" //'Done for CHLA
//End //If //'Done for CHLA
//HoldIRBNo = //Forms![_irb //input //form]![id#] //'Done for CHLA
//HoldStudy = //Forms![_irb //input //form]![id#] //'Done for CHLA
//HoldTitle = //Forms![_irb //input //form]![Protocol //Number //& //Title] //'Done for CHLA
//IRBInvoiceSwitch = //True //' StudySpecificSwitch = True
//LetterName = "Invoice for IRB Services" //'"Study Specific (comments only)"
stDocName;
;
stLinkCriteria = "[IRB#]=" && "'" && HoldIRBNo && "'";
//'DoCmd.Close
location.href = "Form_"+stDocName+".php";//, //acNormal, //, //stLinkCriteria, //, //acWindowNormal;
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
    ;
    ;
       ;
//Exit_bSendInvoice:;
//Exit //Sub;
;
//Error_bSendInvoice:;
//If //Err.Number = //2501 //Then;
    //If //RegComDlg32 //Then //MsgBox "Registered Comdlg32.ocx...Restart PRO_IRB";
    //Application.Quit;
//Else;
    //MsgBox "Error Loading Letter Choices Form..Call PRO_IRB Technical Support...." //& //Err.Number //& "..." //& //Err.Description;
    //Resume //Exit_bSendInvoice;
//End //If;
;
}
function CmdCoINV_Click(){
//If //Me.NewRecord //Then;
    //MsgBox "You must first Save Data for this New Study and assign it to a Meeting Date.  Then you can click the Co-Investigator... button to enter additional data.";
    //Exit //Sub;
//End //If;
//Call //OpenForms("CoInvestigatorSelect");
;
}
function CmdReviewers_Click(){
//If //Me.NewRecord //Then;
    //MsgBox "You must first Save Data for this New Study and assign it to a Meeting Date.  Then you can click the Reviewer... button to enter additional data.";
    //Exit //Sub;
//End //If;
//Call //OpenForms("ReviewerSelect");
}
function CoOrd_DblClick(Cancel){
//On //Error //GoTo //Err_CoOrd_DblClick;
//If //IsNull(Me.[Coordinator //ID]) //Then;
    //MsgBox //("No //Coordinator //to //View");
    //Exit //Sub;
//End //If;
;
    var stDocName = '';
    var stLinkCriteria = '';
    stLinkCriteria = "[common contact id]" && " = " && Me![Coordinator ID];
    stDocName;
    //HoldCoord = //Me.[Coordinator //ID];
    //HoldCoordName = " ";
    //Me.Visible = //False;
    //Forms![newmenu].Form.Visible = //False;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    ;
    //Forms![newmenu].Form.Visible = //True;
    //Me.Visible = //True;
    //Me.Date_Received.SetFocus;
    //If //IsNull(HoldCoord) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Coordinator //ID] = //HoldCoord;
        //Me.CoOrd.RowSourceType = "Table/Query";
    //End //If;
;
;
//Exit_CoOrd_DblClick:;
    //Exit //Sub;
;
//Err_CoOrd_DblClick:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_CoOrd_DblClick;
;
}
function CoOrd_NotInList(NewData, Response){
var strMessage = '';
strMessage = "Are you sure you want to add  '" && NewData && _;
   "' as a new Coordinator?";
//If //YesNo(strMessage) //Then;
    ;
    var stDocName = '';
    var stLinkCriteria = '';
    stDocName;
    //HoldCoordName = //NewData;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //If //IsNull(HoldCoord) //Then;
        //Exit //Sub;
     //Else;
        //Me.[Coordinator //ID] = //HoldCoord;
    //End //If;
    //Response = //acDataErrAdded //'
//Else;
    //DisplayMessage //("Select //an //item //or //Press //the //ESC //key //to //exit");
   //Response = //acDataErrContinue;
//End //If;
;
;
}
function coordadd_Click(){
//On //Error //GoTo //coordadd_err_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
    //HoldCoord = //Me.[Coordinator //ID];
    stDocName;
    //HoldCoordName = " ";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.Date_Received.SetFocus;
     //If //IsNull(HoldCoord) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Coordinator //ID] = //HoldCoord;
        //Me.CoOrd.RowSourceType = "Table/Query";
    //End //If;
;
    ;
//coordadd_Click:;
    //Exit //Sub;
;
//coordadd_err_Click:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //coordadd_Click;
    ;
}
function Ctl__Patients_enrolled_BeforeUpdate(Cancel){//'coast whole sub out field change to text
//If //EnrollOver() //Then;
    //If //YesNo("Enrollment //exceeds //approved //level //do //you //wish //to //save //anyway?") //Then;
        //Else;
        //Exit //Sub;
    //End //If;
//End //If;
;
}
function Form_AfterUpdate(){
//could initialize ExPedorExempt here
//If //Me.[Study //Status] //Like "*Ex*" //Then;
    ExPedorExempt;
//Else;
    ExPedorExempt;
//End //If;
;
//If //BackFromCPA //Then;
    //DataSaved = //True;
    //GoTo //Exit_Update;
//End //If;
//MsgBox "Study data saved";
//DataSaved = //True;
//If //CyberIRB_Flag //Then //'bridge stuff
    //If //isFormLoaded("CyberBridge") //Then //'bridge stuff
        //gclsCyberIRB.ProcessedByProIRB "Protocol"//, //Me.txtcyberrefnum //'bridge stuff
        //Form_CyberBridge_CyberNeedToProcess.RefreshGrid //'bridge stuff   Terry 080902
    //End //If //'bridge stuff
//End //If //'bridge stuff
    ;
;
;
//If //CheckAppl //Then;
    //'study has an open initial submission
    //MsgBox "The Original Agenda record for the Initial Submission has been deleted.  You can now place the Study on the Agenda.";
    //GoTo //PutOnAgenda;
//End //If;
;
//'If WasNew And (Not IsNull(Me.Last_Renewal_Date)) Then GoTo Exit_Update ' DONE FOR chla 05/13/08 dick i commented out
//'If WasNew And (Not IsNull(Me.Last_Renewal_Date)) Then Exit Sub   ' DONE FOR chla
    //'renewal date is blank
//If //(WasNew) //Then;
    //'If IsN ull(Me.Last_Renewal_Date) And Not IsNull(Me.Approval_Date___original_) Then 05/13/08 dick i commented out
        //If ExPedorExempt //Then;
            //If //YesNo("The //Study //Data //has //been //Saved.  //Select //'Yes' if you wish to continue and place this Study on the Agenda.  Select 'No' and the Study will not be placed on the Agenda") Then
                //GoTo //PutOnAgenda;
            //Else;
                //GoTo //Exit_Update //'done for chla
                //'Exit Sub  'done for chla
            //End //If;
        //Else;
            //'GoTo Exit_Update 'done for chla
                //'Exit Sub  'done for chla
            ;
        //End //If;
    //'End If 05/13/08 dick i commented out
//End //If;
//If //WasNew //Then;
    //If //IsNull(Me.Last_Renewal_Date) //And //IsNull(Me.Approval_Date___original_) //Then;
            //GoTo //PutOnAgenda;
    //End //If;
//End //If;
//If //WasNew //Then //MsgBox "didn't catch a was new combo";
   ;
   ;
//If //IsNull(Me.Last_Renewal_Date) //And //IsNull(Me.Approval_Date___original_) //Then;
        //If //Not //YesNo("The //Study //Data //has //been //Saved.  //Select //'Yes' if you wish to continue and place this Study on the Agenda.  Select 'No' and the Sudy will not be placed on the Agenda") Then
            //GoTo //Exit_Update //'done for chla
                //'Exit Sub  'done for chla
            ;
        //Else;
            //GoTo //PutOnAgenda;
        //End //If;
//End //If;
;
//If //IsNull(Me.Last_Renewal_Date) //And ExPedorExempt //Then;
        //If //Not //YesNo("The //Study //Data //has //been //Saved.  //Select //'Yes' if you wish to continue and place this Study on the Agenda.  Select 'No' and the Sudy will not be placed on the Agenda") Then
        //GoTo //Exit_Update //'done for chla
                //'Exit Sub  'done for chla
            ;
        //Else;
            //GoTo //PutOnAgenda;
        //End //If;
//Else;
           //GoTo //Exit_Update //'done for chla
                //'Exit Sub  'done for chla
    ;
//End //If;
;
;
//PutOnAgenda:;
//Me.AgendaText1 = "Initial Submission";
//Me.AgendaText2 = "Date Received- " //& //Me.Date_Received;
;
//If //Me.[Study //Status] //Like "*Exemp*" //Or //Me.[Study //Status] //Like "*Exped*" //Then;
        //If //Me.[Study //Status] //Like "*exem*" //Then;
            //Me.AgendaText1 = "Exempt"  //'agenda category
            //If //Not //IsNull(Me.Approval_Date___original_) //Then;
                //Me.AgendaText2 = //Me.AgendaText2 //& "; Date Exempted- " //& //Me.[Approval_Date___original_];
                //Me.AgendaText2 = //Me.AgendaText2 //& "; Reason Exempt- " //& //Me.ExemptCite.Column(0) //& " - " //& //Me.ExemptCite.Column(1);
            //Else;
                //Me.AgendaText2 = //Me.AgendaText2 //& "; Reason Exempt- " //& //Me.ExemptCite.Column(0) //& " - " //& //Me.ExemptCite.Column(1);
            //End //If;
        //End //If;
        //If //Me.[Study //Status] //Like "*exped*" //Then;
            //Me.AgendaText1 = "Expedited"  //'agenda category
            //If //Not //IsNull(Me.Approval_Date___original_) //Then;
                //Me.AgendaText2 = //Me.AgendaText2 //& "; Date Expedited- " //& //Me.[Approval_Date___original_];
                //Me.AgendaText2 = //Me.AgendaText2 //& "; Reason Expedited- " //& //Me.ExpediteCite.Column(0) //& " - " //& //Me.ExpediteCite.Column(1);
            //Else;
                //'msgbox Me.ExpediteCite.Column(0)
                //Me.AgendaText2 = //Me.AgendaText2 //& "; Reason Expedited- " //& //Me.ExpediteCite.Column(0) //& " - " //& //Me.ExpediteCite.Column(1);
            //End //If;
        //End //If;
//End //If;
;
;
;
;
//If //Not //CreateAgendaRecordForNew //Then;
        //'MsgBox "Application will Close" chla
        //GoTo //Exit_Update //'done for chla
                //'Exit Sub  'done for chla
        ;
//End //If;
//Exit_Update: //'done for chla
//BackFromCPA = //False;
//Exit //Sub //'done for chla
;
//Err_Update: //'done for chla
//MsgBox "Error in After Update-- " //& //Err.Number //& "  " //& //Err.Description  //'done for chla
//MsgBox "Exit ProIRB then restart and verify which records where saved." //'done for chla
//Resume //Exit_Update //'done for chla
}
function Form_BeforeUpdate(Cancel){
//could initialize Response here
//If //CancelUpdate //Then;
    //CancelUpdate = //False;
    //If //Forms![newmenu]![HoldReadOnly] //Then;
        //Me.Undo;
        //Exit //Sub;
    //End //If;
    Response;
        //& " " //& "  Do you wish to Save the Changes?"//, //vbYesNo);
    //If Response;
        //Me.Undo;
            //Cancel = //True;
        //MsgBox "The New Changes were not saved!"//, //vbInformation;
        //DoCmd.Close;
        //Exit //Sub;
    //Else;
        //If //CheckDataError //Then;
            //Cancel = //True;
            //Exit //Sub;
        //End //If;
    //End //If;
//End //If;
;
//If //Me.NewRecord //Then //WasNew = //True;
;
;
}
function Investigator_ID_DblClick(Cancel){
//On //Error //GoTo //Err_Investigator_ID_DblClick;
//HoldPI = //Me.[Investigator //ID];
//If //IsNull(Me.[Investigator //ID]) //Then;
//MsgBox "You must click the 'Add' button to add an Investigator"//, //vbInformation;
//Exit //Sub;
//Else;
    var stDocName = '';
    var stLinkCriteria = '';
    stLinkCriteria = "[common contact id]" && " = " && Me![Investigator ID];
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acWindowNormal;
    //Forms![_Contact //PI]![addpi].Visible = //True;
        //Forms![_Contact //PI]![Ignore].Visible = //True;
//End //If;
;
//Exit_Investigator_ID_DblClick:;
    //Exit //Sub;
;
//Err_Investigator_ID_DblClick:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_Investigator_ID_DblClick;
    ;
;
}
function CPA_button_Click(){
//On //Error //GoTo //Err_cpa_BUTTON_Click;
;
//could initialize HoldCPAcount here
HoldCPAcount;
//If //Forms![newmenu]![HoldReadOnly] //And //Me.cpacount = //0 //Then;
    //MsgBox "Read only user--No CPAs to View";
    //Exit //Sub;
//End //If;
//If //CheckDataError //Then;
     //Exit //Sub;
//End //If;
//CancelUpdate = //False;
//DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
;
;
    ;
;
;
;
;
;
;
//If //Not //WasNew //And //Not //Me.Study_Active_ = "Closed" //Then   //'done for Chla
    //DoCmd.Close //acForm, "_irb input form"//, //acSaveYes   //'done for CHLA
    location.href = "Form__irb input form.php";//, //acNormal //'Done for CHLA
    ;
//End //If //'done for Chla
//WasNew = //False;
//GoCPA:;
var stDocName = '';
    var stLinkCriteria = '';
    stDocName;
    stLinkCriteria = "[IRB#]=" && "'" && Forms![_irb input form]![id#] && "'"   'done for chla;
    ;
    ;
    ;
    ;
//If //Forms![newmenu]![HoldReadOnly] //Then;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
    //Exit //Sub;
//End //If;
   ;
    ;
//If HoldCPAcount;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
    //Forms![_select //cpas //for //irb#].DataEntry = //True;
//Else;
    //If //YesNo("Do //you //wish //to //ADD //a //new //CPA? ") Then
      DoCmd.OpenForm stDocName, , , stLinkCriteria
        Forms![_select cpas for irb#].DataEntry = True
    Else
        DoCmd.OpenForm stDocName, , , stLinkCriteria
    End If
End If
Forms![_irb input form].Visible = False


Exit_cpa_BUTTON_Click:
    Exit Sub

Err_cpa_BUTTON_Click:
    MsgBox Err.Number & Err.Description
    Resume Exit_cpa_BUTTON_Click
    
End Sub

Private Sub Form_Activate()
On Error GoTo err_activate
Me.ListriskCat.Requery

GoTo exit_activate  'temp for chla




exit_activate:
Exit Sub
err_activate:
MsgBox "//Error //saving //data //from //the //Activate //event   " & Err.Description
Resume exit_activate
End Sub



Private Sub Form_Close()
CancelUpdate = True
FromIRB = False
DoCmd.Close acReport, "//StudyCPASAEHistory";
;
;
}
function Form_Current(){
//Me.AgendaDate.Requery;
;
//'Me.lstReviewer.Requery
//Me.AgendaDate = //Me.AgendaDate.[ItemData](0);
//DataSaved = //False;
//If //Me.NewRecord //Then;
    //Me.ListActiveOnly = //True;
    //Me.ListAll = //True;
    //Me.PI.Requery;
    //InitialSubmission = //True;
    //Me.ID_.SetFocus;
    //Me.ID_.Locked = //False;
    //Me.IRB_Code.Locked = //True;
    //Me.IRB_Code.Enabled = //False;
    //Me.Renewal_Cycle.Locked = //False;
    //Me.Date_Received.Locked = //False;
    //Me.InitEnroll.Locked = //False;
    //Me.Ctl__Patients_enrolled.Locked = //False;
    //Me.Protocol_Number___Title.Locked = //False;
    //Me.MostRecentIRBMeeting.Locked = //False;
    //Me.Sponsor.Locked = //False;
    //Me.PI.Locked = //False;
    //Me.CoOrd.Locked = //False;
    //Me.SAE_Button.Enabled = //True;
    //Me.CPA_button.Enabled = //True;
;
//End //If;
;
//If //IsNull(Me.AgendaDate) //Then;
    //Me.NoAgenda.Visible = //True;
    //Me.AgendaDate.Visible = //False;
    //Exit //Sub;
//Else;
    //Me.AgendaDate.Visible = //True;
    //Me.NoAgenda.Visible = //False;
//End //If;
//CpaActive = //Forms![_irb //input //form]![Study //Active?];
    //CpaStatus = //Forms![_irb //input //form]![StudyStatus];
    ;
}
function Form_Filter(Cancel, FilterType){
//Me.FilterOn = //True;
//Me.Filter = "";
}
function Form_Open(Cancel){
//If //RenewAnniversary //Then;
    //Me.TExpireDate.Locked = //False;
//Else;
    //Me.TExpireDate.Locked = //True;
//End //If;
;
//WasNew = //False;
;
//DataSaved = //False;
//'Me.Title.Caption = holdcaption & "  at  " & HoldLocation
//Me.Title.Caption = //holdcaption;
//CancelUpdate = //False;
;
//'Put locking code here if user does not have edit privledges.
;
//If //Forms![newmenu]![HoldReadOnly] //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.sponsoradd.Enabled = //False;
    //Me.coordadd.Enabled = //False;
    //Me.addpi.Enabled = //False;
    //Me.SendCorrespondance.Enabled = //False;
    //Me.Save_Data.Enabled = //False;
    //Me.bSendInvoice.Enabled = //False;
//End //If;
//Me.PI.Requery //'Terry 3/29/2009
;
;
//Check_Demo:;
//If //Demo //Then;
    //Me.Study_Active_.Locked = //True;
    //Me.StudyStatus.Locked = //True;
    //Me.TExpireDate.Locked = //True;
    //Me.InitEnroll.Locked = //True;
    //Me.First_IRB_Review.Locked = //True;
    //Me.Approval_Date___original_.Locked = //True;
    //Me.Last_Renewal_Date.Locked = //True;
    //Me.MostRecentIRBMeeting.Locked = //True;
    //Me.InitEnroll.Locked = //True;
    //Me.Date_Received.Locked = //True;
    //Me.Ctl__Patients_enrolled.Locked = //True;
    //Me.Protocol_Number___Title.Locked = //True;
    ;
    ;
//End //If;
//'Me.StudyStatus = Me.StudyStatus.ItemData(0)
//exit_open:;
//Exit //Sub;
}
function ICButton_Click(){
//On //Error //GoTo //error_IC;
//If //Me.NewRecord //Then;
    //If //DataSaved = //True //Then;
        //GoTo //SetupForIC;
    //Else;
        //MsgBox "You must Save the New Application first"//, //vbInformation;
        //Exit //Sub;
    //End //If;
//End //If;
;
//SetupForIC:;
//HoldCoord = //Me.CoOrd.Column(1);
//HoldPI = //Me.PI.Column(1);
//HoldSponsor = //Me.Sponsor.Column(1);
//HoldIRBNo = //Me.ID_;
//could initialize dbs here
//could initialize rst here
//Set dbs;
//sqlSTR = "SELECT InformedConsent.[IRB#] FROM InformedConsent" //_;
//& " WHERE InformedConsent.[IRB#]= " //& //strQuote //& //[Forms]![_irb //input //form]![id#] //& //strQuote;
;
//Set rst;
//If //rst.RecordCount //< //1 //Then;
    //If //Forms![newmenu]![HoldReadOnly] //Then;
        //MsgBox "No Consent Form to View";
        //rst.Close;
        //Exit //Sub;
    //Else;
        //MsgBox "PRO_IRB has created the Original Checklist for this Study";
        //rst.Close;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "AppendToInformedConsent1";
        //DoCmd.OpenQuery "AppendConsentDetail";
        //DoCmd.SetWarnings //True;
    //End //If;
//Else;
    //rst.Close;
//End //If;
;
//Me.Visible = //False;
location.href = "Form_NewInformedConsent.php";;
;
//exit_IC:;
//Exit //Sub;
;
//error_IC:;
//MsgBox //Err.Description;
//Resume //exit_IC;
;
}
function New_Application_Click(){
location.href = "Form__irb input form.php";//, //acFormDS;
}
function ID__Exit(Cancel){
//If //IsNull(Me.ID_) //Then //Exit //Sub;
//Me.ID_ = //Trim$(Me.ID_);
;
//If //Not //GetVersion = "8.0" //Then;
//'1/24/02 Not going to use this until I can check if any users have 2 dashes.
    //If //MoreThan1Dash(Me.ID_) //Then;
        //MsgBox "Study Number can not contain more than 1 dash";
        //If //IsNull(Me.ID_.OldValue) //Then;
            //Me.ID_ = " ";
            //Exit //Sub;
        //Else;
            //Me.ID_ = //Me.ID_.OldValue;
            //Exit //Sub;
        //End //If;
    //End //If;
//End //If;
;
//If //Me.NewRecord //Or //Me.ID_.OldValue //<> //Me.ID_ //Then;
    //checkirbno.Requery;
    //If //Not //IsNull(Me.checkirbno.Column(0)) //Then;
        //MsgBox "Study Number already on file. Change Study Number or Select OK then Return without saving Changes"//, //vbInformation;
        //Me.ID_.SetFocus;
        //'Me.ID_ = " "
        //Me.ID_.SetFocus;
                  //'**** CyberBridge 090117 start
    //Else //'new code that now checks if the Bridge has this IRB# waiting to come into "IRB - Research  Database" as IRB#
               //'probably want to if a flag here if cyber code is enabled ALTHOUGH if this Org doesn't have Cyber gclsCyberIRB.IRBNumWaitingInBridge will return False
               var sstrCyberRefNum = '';
               //If //IsNull(txtcyberrefnum) //Then;
                   sstrCyberRefNum;
               //Else;
                   sstrCyberRefNum;
               //End //If;
               //If //gclsCyberIRB.IRBNumWaitingInBridge(Me.[IRB#], //sstrCyberRefNum) //Then;
                   //MsgBox "This Study Number is being used by a study waiting in the Cyber Bridge to come into ProIRB. Select OK then Change Study Number or Return without saving Changes"//, //vbInformation;
                   //Me.ID_.SetFocus;
                   //'Me.ID_ = " "
                    //Me.ID_.SetFocus;
               //End //If;
              //'**** CyberBridge 090117 end
    ;
    //End //If;
//End //If;
}
function InitEnroll_BeforeUpdate(Cancel){//'coast whole sub out current enroll dbfield changed to text field changed to
//'If EnrollOver() Then
 //'   If YesNo("Enrollment exceeds approved level do you wish to save anyway?") Then
  //'      Else
   //'     Exit Sub
   //' End If
//'End If
;
}
function IRB_Code_BeforeUpdate(Cancel){
//If //IsNull(Me.AgendaDate) //Then;
    //Exit //Sub;
//Else;
    //DisplayMessage //("Study //must //be //removed //from //the //agenda //before //changing //IRB");
    //Me.Undo;
//End //If;
}
function Last_Renewal_Date_AfterUpdate(){
//Call //NewExpirationDate;
//Me.TExpireDate = //Me.TExpireDate.Value;
;
}
function ListActiveOnly_AfterUpdate(){
//Me.PI.Requery;
}
function PI_Change(){
//If //Me.PI.Column(2) = //False //Then;
    //If //YesNo("You //have //selected //an //Inactive //P.I.  //Are //you //sure //this //is //what //you //wish //to //do?") //Then;
        //Exit //Sub;
    //Else;
        //Me.PI = //Me.PI.OldValue;
    //End //If;
//End //If;
;
}
function PI_DblClick(Cancel){
//On //Error //GoTo //Err_PI_DblClick;
//If //IsNull(Me.[Investigator //ID]) //Then;
    //MsgBox //("No //P.I. //to //View");
    //Exit //Sub;
//End //If;
;
    var stDocName = '';
    var stLinkCriteria = '';
    stLinkCriteria = "[common contact id]" && " = " && Me.[Investigator ID];
    //HoldPI = //Me.[Investigator //ID];
    stDocName;
    //FromIRB = //True;
    //Me.Visible = //False;
    //Forms![newmenu].Form.Visible = //False;
    //SearchContacts = //False;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    ;
    //Forms![newmenu].Form.Visible = //True;
    //Me.Visible = //True;
    //Me.CoOrd.SetFocus;
    //If //IsNull(HoldPI) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Investigator //ID] = //HoldPI;
        //Me.PI.RowSourceType = "Table/Query";
    //End //If;
    ;
;
//Exit_PI_DblClick:;
    //Exit //Sub;
;
//Err_PI_DblClick:;
    //'MsgBox Err.Description, vbInformation
    //Resume //Exit_PI_DblClick;
;
}
function PI_NotInList(NewData, Response){
var strMessage = '';
strMessage = "Are you sure you want to add  '" && NewData && _;
   "' as a new PI?";
//If //YesNo(strMessage) //Then;
    ;
    var stDocName = '';
    var stLinkCriteria = '';
    stDocName;
    //HoldPIName = //NewData;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //If //IsNull(HoldPI) //Then;
        //Exit //Sub;
     //Else;
        //Me.[Investigator //ID] = //HoldPI;
    //End //If;
    //Response = //acDataErrAdded //'
//Else;
    //DisplayMessage //("Select //an //entry //or //Press //the //Esc //key //to //Exit");
   //Response = //acDataErrContinue;
//End //If;
;
}
function Preview_Click(){
//On //Error //GoTo //Err_Preview_Study_History_Click;
var stDocName = '';
;
//If //Me.NewRecord //Then;
    //If //DataSaved = //True //Then;
        //GoTo //PreviewStudy;
    //Else;
        //MsgBox "You must Save the New Application first"//, //vbInformation;
  //GoTo //Exit_Preview_Study_History_Click:;
    //End //If;
//End //If;
//DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
;
//PreviewStudy:;
;
        //MyFilter = "[irb#]=" //& "'" //& //Me.[IRB#] //& "'";
          //holdcaption = "";
       //LetterName = "SudyCPASAEHistory";
       stDocName;
       //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
       ;
       //Me.Visible = //False;
;
    ;
//Exit_Preview_Study_History_Click:;
    //Exit //Sub;
;
//Err_Preview_Study_History_Click:;
    //MsgBox "100-Error occured while trying to Preview History" //& " " //& //Err.Description;
    //Resume //Exit_Preview_Study_History_Click;
    ;
}
function Renewal_Cycle_AfterUpdate(){
//Call //NewExpirationDate;
//Me.TExpireDate = //Me.TExpireDate.Value;
}
function Save_Data_Click(){
//On //Error //GoTo //Save_error;
//If //CheckDataError //Then;
     //Exit //Sub;
//End //If;
//CancelUpdate = //False;
 //Forms![newmenu]![study] = //Me.ID_;
//If //IsNull(txtcyberrefnum) //Then //txtcyberrefnum.Value = "NR" //& //Me.ID_ //'Cyberbridge 090117
//DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
;
//' **** Cyberbridge 090117 Start ****
//''I had to move this here because the IRB# must be in the database before you can add the
//'' CoOrdinator and Reviewer
//If //CyberIRB_Flag //Then;
     var tblExists = '';
;
//'    DoCmd.SetWarnings False
//'       in the msysobjects table, type 6 = linked table and type 1 = table
        tblExists;
        //If tblExists //<> "No" //Then;
        //'***Sync to CoInvStudy1****
           //could initialize scolKeyFieldNames here
           //Set scolKeyFieldNames;
           //scolKeyFieldNames.Add "CoInvContactID";
           //scolKeyFieldNames.Add "StudyNo";
           //could initialize scolDoNotSyncFields here
           //Set scolDoNotSyncFields;
           //scolDoNotSyncFields.Add "CoInvStudyID" //'is an autonumber field so we can't set it
           //gclsCyberIRB.SyncTables //CurrentProject.Connection, "CyberTemp_CoInvStudy1"//, //gclsCyberIRB.ProConnection, "CoInvStudy1"//, //scolKeyFieldNames, scolDoNotSyncFields //'061114
        //End //If;
;
//'       in the msysobjects table, type 6 = linked table and type 1 = table
         tblExists;
         //If tblExists //<> "No" //Then;
            //'061121***Sync to ReviewerStudy1****
            //Set scolKeyFieldNames;
            //scolKeyFieldNames.Add "ReviewerContactID";
            //scolKeyFieldNames.Add "StudyNo";
            //Set scolDoNotSyncFields;
            //scolDoNotSyncFields.Add "ReviewerStudyID" //'is an autonumber field so we can't set it
            //gclsCyberIRB.SyncTables //CurrentProject.Connection, "CyberTemp_ReviewerStudy1"//, //gclsCyberIRB.ProConnection, "ReviewerStudy1"//, //scolKeyFieldNames, scolDoNotSyncFields;
         //End //If;
;
    //Set scolKeyFieldNames;
    //Set scolDoNotSyncFields;
;
    //gclsCyberIRB.gclsBridgeTempData.DeleteTEMPTables //'061112
    //'Now send the protocol to Cyber only if bolProUpdateCyber says to.
     //If //UpdateCyberIRB(Me.[IRB#]) //Then;
        //UpdateCyberProtocol //Me.[IRB#];
        //updateUsers //Me.[IRB#];
     //End //If;
   ;
    //'If the protocol was sent down from Cyber, then don't change any of the contacts in the UserByProtocol table.
    //'If the protocol was created in ProIRB, then I do need to populate the UserByProtocol table to give them
    //'access to their protocols.
//''Terry 4/24/2009     If Me.txtcyberrefnum.Value = "NR" & Me.[IRB#] Then updateUsers Me.[IRB#]
    //Form_CyberBridge_CyberNeedToProcess.RefreshGrid    //'Bridge stuff Terry 090108
    //WasNew = //True;
    ;
 //End //If;
;
//' **** Cyberbridge 090117 end ****
        ;
;
//If //WasNew //Then //GoTo //Exit_Save  //' done for CHLA
//If //Me.Study_Active_ = "OPEN" //Then;
    //DoCmd.Close //'done for CHLA
    location.href = "Form__irb input form.php";//, //acNormal //'Done for CHLA
    //Me.Filter = //Forms![newmenu]![study];
    //Me.FilterOn = //True;
//End //If;
;
;
//Exit_Save:;
//Forms![newmenu]![QuickStudyNumber].Requery;
    //Exit //Sub;
;
//Save_error:;
    //'MsgBox "Error Save Data Function--Study Data Not Saved" & " " & Err.Number & " " & Err.Description, vbInformation
    ;
    //Resume //Exit_Save;
}
function Return_to_Menu_Click(){
//could initialize Response here
//On //Error //GoTo //err_return;
//CancelUpdate = //True;
//WasNew = //False;
//'MsgBox Forms![_irb input form]![StudyStatus]
//'MsgBox Forms![_irb input form]![StudyStatus].OldValue
;
;
//If //Forms![_irb //input //form]![StudyStatus].OldValue = //CpaStatus //Or //CpaStatus = "" //Or //IsNull(CpaStatus) //Then;
    //Me.Undo;
    //DoCmd.Close;
    //Exit //Sub;
//End //If;
//DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70 //'done for chla
//DoCmd.Close;
//Exit //Sub;
;
;
;
//Exit_Return:;
//Exit //Sub  //'done for chla else exit sub
;
;
//err_return:;
//Resume //Exit_Return;
}
function SAE_BUTTON_Click(){
//could initialize HoldSAEcount here
//On //Error //GoTo //Err_SAE_BUTTON_Click;
HoldSAEcount;
//If //Forms![newmenu]![HoldReadOnly] //And //Me.saecount = //0 //And //Not //SAEUser //Then;
        //MsgBox "Read only user--No SAEs to View";
        //Exit //Sub;
//End //If;
//If //CheckDataError //Then;
     //Exit //Sub;
//End //If;
//CancelUpdate = //False;
;
//DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
;
//If //Not //WasNew //Then //'done for Chla
    //DoCmd.Close //acForm, "_irb input form"//, //acSaveYes   //'done for CHLA
    location.href = "Form__irb input form.php";//, //acNormal //'Done for CHLA
//End //If //'done for Chla
;
//WasNew = //False;
;
//GoSAE:;
;
//'DoCmd.RunCommand acCmdSaveRecord
 ;
//'Me.Filter = ""
//searchirbs = //False;
var stDocName = '';
var stLinkCriteria = '';
stDocName;
stLinkCriteria = "SAE.[IRB#]=" && "'" && Forms![_irb input form]![id#] && "'";
//'Me.Requery 'done for chla
//Forms![_irb //input //form].Visible = //False;
;
;
//If //Forms![newmenu]![HoldReadOnly] //Then;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
    //Exit //Sub;
//End //If;
;
;
//'MsgBox Forms![_irb input form]![saecount]
//If HoldSAEcount;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
    //Forms![_SAEForm].DataEntry = //True;
//Else;
    //If //YesNo("Do //you //wish //to //ADD //a //new //SAE? ") Then
      DoCmd.OpenForm stDocName, , , stLinkCriteria
        Forms![_SAEForm].DataEntry = True
    Else
        DoCmd.OpenForm stDocName, , , stLinkCriteria
    End If
End If




Exit_SAE_BUTTON_Click:
    Exit Sub

Err_SAE_BUTTON_Click:
    MsgBox Err.Description
    Resume Exit_SAE_BUTTON_Click
    
End Sub
Private Sub Sponsor_ID_NotInList(NewData As String, Response As Integer)
'On Error GoTo Err_Sponsor_ID_NotInList

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "//_Contact //Sponsor";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd;
    //Forms![_Contact //Sponsor]![addsponsor].Visible = //True;
        //Forms![_Contact //Sponsor]![Ignore].Visible = //True;
        ;
;
;
;
//Exit_Sponsor_ID_NotInList:;
    //Exit //Sub;
;
//Err_Sponsor_ID_NotInList:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_Sponsor_ID_NotInList;
    ;
}
function SendCorrespondance_Click(){
var stDocName = '';
var stLinkCriteria = '';
;
//'MsgBox Forms![_irb input form].textSpecialWording.Value
//On //Error //GoTo //Error_SendCorrespondance;
//If //CheckDataError //Then;
     //Exit //Sub;
//End //If;
//CancelUpdate = //False;
//'MsgBox Me.ListChildCat.Column(3)
 //'Forms![_irb input form]![textSpecialWording] = Forms![_irb input form]!Me.ListChildCat
//'MsgBox Forms![_irb input form]![textSpecialWording]
;
;
;
;
//'DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70 'done for CHLA moved below
;
//If //Not //WasNew //Then //'done for Chla
    //DoCmd.Close //acForm, "_irb input form"//, //acSaveYes   //'done for CHLA
    location.href = "Form__irb input form.php";//, //acNormal //'Done for CHLA
    ;
//End //If //'done for Chla
//'MsgBox Forms![_irb input form]!ListChildCat.Column(3, 0)
//'Forms![_irb input form]!textSpecialWording = Forms![_irb input form]!ListChildCat.Column(3, 0)
//'MsgBox Forms![_irb input form]!textSpecialWording
//goletter:;
//If //GetTemplateLocation("StudySpecific") //Then;
        //If //TemplateName = "" //Then;
            //Exit //Sub;
        //Else;
            //DoCmd.SetWarnings //False;
            //DoCmd.OpenQuery //("ExportStudySpecificWithCreateTable");
            //DoCmd.SetWarnings //True;
            //CustomLetter = //True;
        //End //If;
//Else;
        //'msgbox "Do PROIRB SAE Letter"
        //CustomLetter = //False;
//End //If;
;
     ;
//FromIRB = //False;
//If //Not //IsNull(Forms![_irb //input //form]!AgendaDate) //Then  //'Done for CHLA
        //HoldMeetingDate = //Forms![_irb //input //form]!AgendaDate //'Done for CHLA
    //Else //'Done for CHLA
        //HoldMeetingDate = "" //'Done for CHLA
//End //If //'Done for CHLA
//HoldIRBNo = //Forms![_irb //input //form]![id#] //'Done for CHLA
//HoldStudy = //Forms![_irb //input //form]![id#] //'Done for CHLA
//HoldTitle = //Forms![_irb //input //form]![Protocol //Number //& //Title] //'Done for CHLA
//StudySpecificSwitch = //True;
//LetterName = "Study Specific (comments only)";
stDocName;
;
stLinkCriteria = "[IRB#]=" && "'" && HoldIRBNo && "'";
//'DoCmd.Close
location.href = "Form_"+stDocName+".php";//, //acNormal, //, //stLinkCriteria, //, //acWindowNormal;
;
;
    ;
    ;
        ;
//Exit_SendCorrespondance:;
//Exit //Sub;
;
//Error_SendCorrespondance:;
//If //Err.Number = //2501 //Then;
    //If //RegComDlg32 //Then //MsgBox "Registered Comdlg32.ocx...Restart PRO_IRB";
    //Application.Quit;
//Else;
    //MsgBox "Error Call PRO_IRB Technical Support...." //& //Err.Number //& "..." //& //Err.Description;
    //Resume //Exit_SendCorrespondance;
//End //If;
}
function Sponsor_DblClick(Cancel){
//On //Error //GoTo //Sponsor_err_DblClick;
//If //IsNull(Me.[Sponsor //ID]) //Then;
    //MsgBox //("No //Sponsor //to //View");
    //Exit //Sub;
//End //If;
    var stDocName = '';
    var stLinkCriteria = '';
    stLinkCriteria = "[common contact id]" && " = " && Me.[Sponsor ID];
    //HoldSponsor = //Me.[Sponsor //ID];
    stDocName;
    //FromIRB = //True;
    //Me.Visible = //False;
    //Forms![newmenu].Form.Visible = //False;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    ;
    //Forms![newmenu].Form.Visible = //True;
    //Me.Visible = //True;
    //Me.PI.SetFocus;
    //If //IsNull(HoldSponsor) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Sponsor //ID] = //HoldSponsor;
        //Me.Sponsor.RowSourceType = "Table/Query";
    //End //If;
//Sponsor_DblClick:;
    //Exit //Sub;
;
//Sponsor_err_DblClick:;
    //'DisplayMessage ("Select one or click on Add")
    //Resume //Sponsor_DblClick;
    ;
;
}
function Sponsor_NotInList(NewData, Response){
var strMessage = '';
strMessage = "Are you sure you want to add  '" && NewData && _;
   "' as a new Sponsor?.";
//If //YesNo(strMessage) //Then;
    ;
    var stDocName = '';
    var stLinkCriteria = '';
    //HoldSponsorName = //NewData;
    stDocName;
    //HoldSponsorName = //NewData;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //If //IsNull(HoldSponsor) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Sponsor //ID] = //HoldSponsor;
    //End //If;
;
    //Response = //acDataErrAdded;
//Else;
    //DisplayMessage //("Select //an //item //or //Press //the //Esc //key //to //exit");
    //Response = //acDataErrContinue;
   ;
//End //If;
;
}
function sponsoradd_Click(){
//On //Error //GoTo //Err_sponsoradd_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
    //HoldSponsor = //Me.[Sponsor //ID];
    stDocName;
    //HoldSponsorName = " ";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.PI.SetFocus;
    //If //IsNull(HoldSponsor) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Sponsor //ID] = //HoldSponsor;
        //'Me.Sponsor.RowSourceType = "Table/Query"
        //Me.Sponsor.Requery;
    //End //If;
;
;
;
//Exit_sponsoradd_Click:;
    //Exit //Sub;
;
//Err_sponsoradd_Click:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_sponsoradd_Click;
    ;
}
function addPI_Click(){
//On //Error //GoTo //err_addpi_click;
  var stDocName = '';
    var stLinkCriteria = '';
    //HoldPI = //Me.[Investigator //ID];
    stDocName;
    //HoldPIName = " ";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.CoOrd.SetFocus;
    //If //IsNull(HoldPI) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Investigator //ID] = //HoldPI;
        //Me.PI.RowSourceType = "Table/Query";
    //End //If;
;
  ;
   ;
//Exit_addpi_Click:;
    //Exit //Sub;
;
//err_addpi_click:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_addpi_Click;
    ;
}
function Study_Active__Change(){
//Me.StudyStatus.Requery;
//Me.StudyStatus = //Me.StudyStatus.ItemData(0);
//If //Me.Study_Active_ = "Closed" //Then;
    //Me.TExpireDate = "";
    ;
//Else;
//End //If;
}
function cmdAreaDetails_Click(){
//If //Me.NewRecord //Then;
    //MsgBox "You must first Save Data for this New Study and assign it to a Meeting Date.  Then you can click the MORE... button to enter additional data.";
    //Exit //Sub;
//End //If;
//Call //OpenForms("AreaDetails");
}
function GetVersion(){//As //String;
//' Determine the version of Microsoft Access
//' used to create this application.
;
    //GetVersion = //SysCmd(acSysCmdAccessVer);
;
}
function CreateAgendaRecordForNew(){//As //Boolean;
//On //Error //GoTo //err_CreateAgendaRecordForNew;
   ;
;
//HoldIRBNo = //Me.ID_;
;
location.href = "Form_ReassignAgendaRecord.php";//, //, //, //, //, //acDialog;
        //If //Me.[Study //Status] //Like "*exped*" //Then;
            //Me.HoldExpedite = //True;
        //Else;
            //Me.HoldExpedite = //False;
        //End //If;
        //Me.AgendaDate = //HoldMeetingDate;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "DeleteChangedNewAppl";
         //If //Not //ReassignSwitch //Then;
            //MsgBox //("A //New //Agenda //Record //was //not //created");
            //Exit //Function;
        //End //If;
        ;
        //If //IsNull(Me.TxtAgenda) //Or //Me.TxtAgenda = "" //Then;
        //Else;
            //Me.AgendaText2 = //Me.AgendaText2 //& " " //& "Remarks: " //& //Me.TxtAgenda;
        //End //If;
        //If //IsNull(Me.txtcyberrefnum) //Then //Me.txtcyberrefnum = " ";
        //If //AddAgendaFromStudy(Me.ID_, "New Appl"//, //Me.AgendaDate, "Initial"//, //Me.AgendaText1, " "//, " "//, //Me.AgendaText2, //_;
             " "//, " "//, " "//, " "//, " "//, //0, //Date, //Me.HoldExpedite, //([Forms]![newmenu]![UserName]), //Me.IRB_Code, " "//, " "//, " "//, //0, //Forms![_irb //input //form]!txtcyberrefnum) //Then;
        //Else;
           //GoTo //err_CreateAgendaRecordForNew;
        //End //If;
        ;
        //MsgBox //("Study //#- " & HoldIRBNo & "  //Assigned //to //the--" //& //HoldMeetingDate //& "  Meeting"//);
        //DoCmd.SetWarnings //True;
        //Me.AgendaDate.Visible = //True;
        //Me.NoAgenda.Visible = //False;
        //Me.AgendaDate.Requery;
        //Forms![newmenu]![study] = //Me.ID_;
        //Forms![newmenu]![Status] = //Me.Study_Active_;
        //HoldStudy = //Me.ID_;
        //HoldStatus = //Me.Study_Active_;
        //Forms![newmenu]![ReviewMonth] = "***";
        //Forms![newmenu]![RenewalMonth] = "***";
        //Forms![newmenu]![RenewalCycle] = "***";
        //InitialSubmission = //False;
        //CreateAgendaRecordForNew = //True;
        ;
//Exit_CreateAgendaRecordForNew:;
    //Exit //Function;
//err_CreateAgendaRecordForNew:;
    //MsgBox "Error Creating Agenda Record-Agenda Item not Created";
    //CreateAgendaRecordForNew = //False;
    //Resume //Exit_CreateAgendaRecordForNew;
    ;
}
function StudyStatus_BeforeUpdate(Cancel){
//If //Me.[Study //Status] //Like "*exem" //Then;
    //Me.TExpireDate = "";
    //Me.Last_Renewal_Date = "";
//End //If;
}
function CheckAppl(){//As //Boolean;
var CriteriaStr = '';
//On //Error //GoTo //err_checkappl;
//If //Me.AgendaDate.ListCount = //0 //Then;
    //CheckAppl = //False;
    //Exit //Function;
//End //If;
;
;
//'msgbox Forms![_irb input form]![ID#]
CriteriaStr = "[IRB#] = " && "'" && Forms![_irb input form]![id#] && "'" && " And [SourceNumber] =  'new appl' And IsNull([IRB Action Category])";
//'msgbox CriteriaSTR
;
//If //DCount("[SourceNumber]", "agenda records"//, //CriteriaStr) = //0 //Then;
    //CheckAppl = //False;
//Else;
    //CheckAppl = //True;
//'    msgbox CheckNewAppl
    //'Exit Sub
//End //If;
//exit_Checkappl:;
//Exit //Function;
//err_checkappl:;
    //MsgBox "Error trying to determine if on Agenda--" //& //Err.Description;
    //Application.Quit;
}
function TExpireDate_DblClick(Cancel){
//Me.TExpireDate.Locked = //False;
}
function Trans_Click(){
//On //Error //GoTo //err_Trans;
//Forms![newmenu]![study] = //Me.ID_;
//DoCmd.SetWarnings //False;
//DoCmd.OpenQuery "DeletePriorChron";
//DoCmd.OpenQuery "AppendCPAtoChron";
//DoCmd.OpenQuery "AppendSAEtoChron";
//DoCmd.OpenQuery "AppendLetterstoChron";
//DoCmd.OpenQuery "AppendAgendatoChron";
//DoCmd.SetWarnings //True;
//LetterName = "TransHistory";
//DoCmd.OpenReport "TransHistory"//, //acViewPreview;
//DoCmd.Hourglass //False;
//DoCmd.RunCommand //acCmdZoom100;
;
//exit_Trans:;
//Exit //Sub;
;
//err_Trans:;
//MsgBox "Error Loading Transaction Report" //& "  " //& //Err.Description;
//Resume //exit_Trans;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Current();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>




    <input type='text' id=ID# value='empty' style='position:absolute; left:951; top:40; width:133; height:25'></input>

    <label id=Text15 value='IRB #:' style='position:absolute; left:988; top:7; width:64; height:24; visibility:'>IRB #:</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:150; top:127; width:399; height:96'></input>

    <label id=Text27 value='Protocol/Title' style='position:absolute; left:7; top:126; width:132; height:24; visibility:'>Protocol/Title</label>
    <input type='text' id=Approval Date  (original) value='empty' style='position:absolute; left:1020; top:231; width:105; height:0'></input>

    <label id=Text31 value='Original Approval Date:' style='position:absolute; left:786; top:231; width:216; height:24; visibility:'>Original Approval Date:</label>
    <input type='text' id=Last Renewal Date value='empty' style='position:absolute; left:1020; top:264; width:105; height:0'></input>

    <label id=Text33 value='Last IRB Renewal Date:' style='position:absolute; left:787; top:264; width:214; height:24; visibility:'>Last IRB Renewal Date:</label>
    <input type='text' id=# Patients enrolled value='empty' style='position:absolute; left:1020; top:360; width:105; height:25'></input>

    <label id=Text37 value='Currently Enrolled:' style='position:absolute; left:837; top:360; width:165; height:24; visibility:'>Currently Enrolled:</label>
    <input type='button' id=Save Data value='&Save Data' style='position:absolute; left:856; top:589; width:141; height:48' onclick='Save Data_Click();'></input>
    <input type='button' id=Return to Menu value='&Return' style='position:absolute; left:997; top:589; width:135; height:48' onclick='Return to Menu_Click();'></input>
    <input type='button' id=SAE Button value='View or Add &SAEs' style='position:absolute; left:951; top:501; width:178; height:40' onclick='SAE Button_Click();'></input>
    <input type='text' id=saecount value='empty' style='position:absolute; left:900; top:510; width:48; height:25'></input>

    <label id=Label418 value='# of SAEs :' style='position:absolute; left:783; top:510; width:106; height:24; visibility:'># of SAEs :</label>
    <input type='button' id=CPA button value='View or Add &CPAs' style='position:absolute; left:951; top:540; width:178; height:40' onclick='CPA button_Click();'></input>
    <input type='text' id=cpacount value='empty' style='position:absolute; left:900; top:547; width:48; height:0'></input>

    <label id=Label427 value='# of CPAs :' style='position:absolute; left:783; top:547; width:106; height:24; visibility:'># of CPAs :</label>
    <select name='Sponsor' Size='4' style='position:absolute; left:150; top:246; width:399; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [agenda records1].[IRB#], [agenda records1].[MeetingDate] FROM [agenda records1] GROUP BY [agenda records1].[IRB#], [agenda records1].[MeetingDate], [agenda records1].[IRB Action Category] HAVING ((([agenda records1].[IRB#])=Forms![_irb input form]![ID#]) And (([agenda records1].[IRB Action Category]) Is Null)) ORDER BY [agenda records1].[MeetingDate] DESC");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['IRB#']."'".">".$row['IRB#']."</Option>";
        }
    ?>
    </select>

    <label id=Sponsor_Label value='Sponsor' style='position:absolute; left:51; top:246; width:88; height:24; visibility:'>Sponsor</label>
    <input type='button' id=sponsoradd value='Add' style='position:absolute; left:553; top:246; width:58; height:27' onclick='sponsoradd_Click();'></input>
    <input type='button' id=Preview value='Study &History' style='position:absolute; left:685; top:589; width:171; height:48' onclick='Preview_Click();'></input>
    <select name='StudyStatus' Size='11' style='position:absolute; left:355; top:552; width:276; height:30'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_select sponsors for irb#].[Common Contact ID], [_select sponsors for irb#].[Company Name], [_select sponsors for irb#].[Active] FROM [_select sponsors for irb#] WHERE ((([_select sponsors for irb#].[Active])=True)) ORDER BY [_select sponsors for irb#].[Company Name]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select>

    <label id=Label441 value='Study Status:' style='position:absolute; left:229; top:550; width:123; height:24; visibility:'>Study Status:</label>
    <select name='PI' Size='6' style='position:absolute; left:150; top:294; width:399; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [StudyStatusCodes1].[Study Status] FROM StudyStatusCodes1 WHERE ((([StudyStatusCodes1].[StudyActiveCode])=[Forms]![_irb input form]![Study Active?])) ORDER BY [StudyStatusCodes1].[Seq]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Study Status']."'".">".$row['Study Status']."</Option>";
        }
    ?>
    </select>

    <label id=Label443 value='Principal\015\012Investigator' style='position:absolute; left:24; top:285; width:115; height:43; visibility:'>Principal\015\012Investigator</label>
    <input type='button' id=addPI value='Add' style='position:absolute; left:553; top:292; width:57; height:27' onclick='addPI_Click();'></input>
    <select name='CoOrd' Size='8' style='position:absolute; left:150; top:342; width:399; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_Select PIs for irb#].[Common Contact ID], [_Select PIs for irb#].[display name], [_Select PIs for irb#].[Active] FROM [_Select PIs for irb#] WHERE (((([_Select PIs for irb#].[Active])=[Forms]![_irb input form]![ListActiveOnly] Or ([_Select PIs for irb#].[Active])=[Forms]![_irb input form]![ListAll]))) ORDER BY [_Select PIs for irb#].[display name]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select>

    <label id=Label447 value='Study Coordinator' style='position:absolute; left:27; top:333; width:112; height:43; visibility:'>Study Coordinator</label>
    <input type='button' id=coordadd value='Add' style='position:absolute; left:553; top:340; width:57; height:27' onclick='coordadd_Click();'></input>
    <label id=Title value='Study #  2000.008  at  Morton Plant' style='position:absolute; left:9; top:73; width:690; height:45; visibility:'>Study #  2000.008  at  Morton Plant</label>
    <input type='text' id=InitEnroll value='empty' style='position:absolute; left:1020; top:330; width:105; height:0'></input>

    <label id=Label470 value='Approved Patient Enrollment:' style='position:absolute; left:745; top:330; width:256; height:24; visibility:'>Approved Patient Enrollment:</label>
    <select name='Renewal Cycle' Size='16' style='position:absolute; left:978; top:100; width:69; height:25'>
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
        <Option value='24'>24</option>
    </select>

    <label id=Label471 value='Renewal Cycle:' style='position:absolute; left:823; top:100; width:141; height:24; visibility:'>Renewal Cycle:</label>
    <input type='button' id=Trans value='Transaction\015\012History' style='position:absolute; left:552; top:589; width:133; height:48' onclick='Trans_Click();'></input>
    <select name='Study Active?' Size='30' style='position:absolute; left:87; top:552; width:0; height:30'>
        <Option value='Open'>Open</option>
        <Option value='Closed'>Closed</option>
    </select>

    <label id=Label454 value='Active?:' style='position:absolute; left:3; top:549; width:79; height:24; visibility:'>Active?:</label>
    <input type='text' id=Date Received value='empty' style='position:absolute; left:1020; top:165; width:105; height:25'></input>

    <label id=Label475 value='Date Received:' style='position:absolute; left:859; top:165; width:142; height:24; visibility:'>Date Received:</label>
    <input type='text' id=First IRB Review value='empty' style='position:absolute; left:1020; top:198; width:105; height:25'></input>

    <label id=Label476 value='First IRB Review:' style='position:absolute; left:844; top:198; width:157; height:24; visibility:'>First IRB Review:</label>
    <input type='button' id=ICButton value='Consent Checklist' style='position:absolute; left:0; top:589; width:168; height:48' onclick='ICButton_Click();'></input>
    <input type='button' id=SendCorrespondance value='Send Study Specific &Letter' style='position:absolute; left:334; top:589; width:216; height:48' onclick='SendCorrespondance_Click();'></input>
    <!--     <select name='checkirbno' Size='34' style='position:absolute; left:492; top:12; width:84; height:18'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_select coordinators for irb#].[Common Contact ID], [_select coordinators for irb#].[display name], [_select coordinators for irb#].[Active] FROM [_select coordinators for irb#] WHERE ((([_select coordinators for irb#].[Active])=True)) ORDER BY [_select coordinators for irb#].[display name]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select> -->

    <label id=Label484 value='Combo483:' style='position:absolute; left:348; top:12; width:88; height:24; visibility:'>Combo483:</label>
    <label id=NoAgenda value='Not on Agenda' style='position:absolute; left:619; top:489; width:162; height:36; visibility:'>Not on Agenda</label>
    <input type='text' id=TExpireDate value='empty' style='position:absolute; left:1020; top:133; width:105; height:0'></input>

    <label id=Label492 value='Expiration Date:' style='position:absolute; left:856; top:133; width:145; height:24; visibility:'>Expiration Date:</label>
    <input type='text' id=MostRecentIRBMeeting value='empty' style='position:absolute; left:1020; top:297; width:105; height:0'></input>

    <label id=Label497 value='Last seen by IRB:' style='position:absolute; left:786; top:297; width:216; height:24; visibility:'>Last seen by IRB:</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:36; top:37; width:157; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <label id=Label500 value='Months' style='position:absolute; left:1054; top:100; width:70; height:24; visibility:'>Months</label>
    <select name='ExpediteCite' Size='21' style='position:absolute; left:816; top:460; width:310; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRB - Research Database].[IRB#] FROM [IRB - Research Database] WHERE ((([IRB - Research Database].[IRB#])=[Forms]![_irb input form]![ID#]))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['IRB#']."'".">".$row['IRB#']."</Option>";
        }
    ?>
    </select>

    <label id=Label501 value='Expedite Category:' style='position:absolute; left:636; top:460; width:169; height:24; visibility:'>Expedite Category:</label>
    <select name='Exemptcite' Size='20' style='position:absolute; left:816; top:432; width:310; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [Expedited categories1].[ExpediteCite], [Expedited categories1].[ExpediteDesc] FROM [Expedited categories1] ORDER BY [Expedited categories1].[ExpediteCite]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['ExpediteCite']."'".">".$row['ExpediteCite']."</Option>";
        }
    ?>
    </select>

    <label id=Label502 value='Exempt Category:' style='position:absolute; left:648; top:432; width:157; height:24; visibility:'>Exempt Category:</label>
    <!--     <select name='AgendaDate' Size='35' style='position:absolute; left:636; top:519; width:138; height:61'>
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

        $sql = sqlsrv_query($conn, "SELECT [Exempt Categories1].[ExemptCite], [Exempt Categories1].[ExemptDesc] FROM [Exempt Categories1] ORDER BY [Exempt Categories1].[ExemptCite]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['ExemptCite']."'".">".$row['ExemptCite']."</Option>";
        }
    ?>
    </select> -->

    <label id=Label488 value='On Agenda' style='position:absolute; left:637; top:493; width:103; height:24; visibility:'>On Agenda</label>
    <input type='button' id=cmdAreaDetails value='&More Study Details...' style='position:absolute; left:573; top:138; width:211; height:28' onclick='cmdAreaDetails_Click();'></input>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:351; top:40; width:591; height:25'></input>

    <label id=Label453 value='IRB:' style='position:absolute; left:477; top:7; width:85; height:24; visibility:'>IRB:</label>
    <!--     <input type='text' id=AgendaText1 value='empty' style='position:absolute; left:708; top:76; width:267; height:28'></input> -->

    <!--     <label id=Label515 value='AgendaText1' style='position:absolute; left:564; top:76; width:103; height:24; visibility:hidden'>AgendaText1</label> -->
    <!--     <input type='text' id=AgendaText2 value='empty' style='position:absolute; left:241; top:99; width:174; height:25'></input> -->

    <label id=Label517 value='AgendaText2' style='position:absolute; left:97; top:99; width:103; height:24; visibility:'>AgendaText2</label>
    <!--     <input type='checkbox' id=HoldExpedite value='False' style='position:absolute; left:312; top:43; width:88; height:13'></input> -->

    <!--     <label id=Label519 value='Check518' style='position:absolute; left:335; top:40; width:81; height:24; visibility:hidden'>Check518</label> -->
    <input type='text' id=TextRemarks value='empty' style='position:absolute; left:150; top:384; width:397; height:46'></input>

    <label id=Label521 value='Remarks' style='position:absolute; left:60; top:384; width:79; height:24; visibility:'>Remarks</label>
    <input type='button' id=bSendInvoice value='Send &Invoice' style='position:absolute; left:168; top:589; width:168; height:48' onclick='bSendInvoice_Click();'></input>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:42; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <!--     <input type='checkbox' id=ListActiveOnly value='False' style='position:absolute; left:634; top:268; width:10; height:19'></input> -->

    <label id=Label524 value='ListActiveoNLY' style='position:absolute; left:657; top:265; width:117; height:24; visibility:'>ListActiveoNLY</label>
    <!--     <input type='checkbox' id=ListAll value='True' style='position:absolute; left:18; top:228; width:94; height:18'></input> -->

    <label id=Label526 value='ListAll' style='position:absolute; left:41; top:225; width:81; height:24; visibility:'>ListAll</label>
    <input type='button' id=CmdCoINV value='Co-Investigators' style='position:absolute; left:573; top:171; width:211; height:28' onclick='CmdCoINV_Click();'></input>
    <input type='button' id=CmdReviewers value='Reviewer(s)' style='position:absolute; left:573; top:202; width:211; height:28' onclick='CmdReviewers_Click();'></input>
    <input type='text' id=TxtVersion value='empty' style='position:absolute; left:39; top:183; width:94; height:27'></input>

    <label id=Label530 value='Version #' style='position:absolute; left:37; top:159; width:100; height:19; visibility:'>Version #</label>
    <input type='text' id=TxtAgenda value='empty' style='position:absolute; left:150; top:442; width:397; height:46'></input>

    <label id=Label532 value='Data for Agenda' style='position:absolute; left:55; top:442; width:84; height:43; visibility:'>Data for Agenda</label>
    <!--     <input type='text' id=txtcyberrefnum value='empty' style='position:absolute; left:192; top:24; width:66; height:42'></input> -->

    <label id=Label534 value='Text533:' style='position:absolute; left:48; top:24; width:70; height:24; visibility:'>Text533:</label>
    <input type='text' id=ListRiskCat value='empty' style='position:absolute; left:814; top:396; width:306; height:0'></input>

    <label id=Label536 value='Risk Category' style='position:absolute; left:664; top:396; width:138; height:24; visibility:'>Risk Category</label>
  </body>
</html>
