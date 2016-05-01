<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function AddComments(){
//Exit //Sub;
;
    ;
;
}
function DisableSponsor(){
//Me.SponsorCheck = //False;
    //Me.Sponsoremail.Enabled = //False;
    //Me.ChSponsorPostal.Enabled = //False;
    //Me.Sponsoremail = //False;
    //Me.ChSponsorPostal = //False;
}
function DisableCoord(){
//Me.CoordCheck = //False;
    //Me.CoordPostal.Enabled = //False;
    //Me.Coordemail.Enabled = //False;
    //Me.CoordPostal = //False;
    //Me.Coordemail = //False;
    ;
}
function DisablePI(){
//Me.PICheck = //False;
    //Me.pipostal.Enabled = //False;
    //Me.pipostal = //False;
    //Me.PIemail = //False;
    //Me.PIemail.Enabled = //False;
    ;
}
function SendEmail(r, Recipient){
//On //Error //GoTo //MailProblem:;
//could initialize MAilSubject here
var FiletoSend = '';
var MailMessage = '';
;
MailMessage = vbCrLf & vbCrLf & Forms![_LetterManager]![LetterLabel].Caption & vbCrLf & vbCrLf & "If you cannot view the attachment, click this link and install the free SnapShot viewer from Microsoft .....http://www.proirb.com/files/snapshotviewerinstall.exe";
FiletoSend = "C:\CC\Proirb_dev\" //& //Forms![_LetterManager]![LetterLabel].Caption //& ".snp";
;
//If //Demo //Then;
   //Call //DemoMsg;
   //Exit //Sub;
//End //If;
//'recipient = Forms("_LetterManager")(SendToCntrl).Column(15, r)
;
    //DoCmd.OutputTo //acOutputReport, //LetterName, "SnapshotFormat(*.snp)"//, FiletoSend;
    //gsEmailCCs = " ";
   //' If MSMailRoutine(Recipient, gsEmailCCs, "Communication From: " & Me.IRB_Code & "  IRB.", FiletoSend, MailMessage) Then
    //'    Exit Sub
    //'End If
;
    //'to be added for 2007 fix   Sendthem
   //If //dir("C:\CC\PROIRB_Dev\Outlook.ini") = "Outlook.ini" //Then;
      ;
        //If //Outlook2007Email(Recipient, //gsEmailCCs, "Communication From: " //& //Forms![_LetterManager]!IRBText //& "  IRB."//, //FiletoSend, //MailMessage) //Then;
            //Exit //Sub;
        //End //If;
   //Else;
;
        //If //MSMailRoutine(Recipient, //gsEmailCCs, "Communication From: " //& //Me.IRB_Code //& "  IRB."//, //FiletoSend, //MailMessage) //Then;
            //Exit //Sub;
        //End //If;
    //End //If;
;
    ;
    ;
    ;
    ;
    ;
;
    //'MailSubject = "Communication From: " & Me.IRB_Code & "  IRB." & vbCrLf & vbCrLf & _
    //'"If you cannot view the attachment, click this link and install the free SnapShot viewer from Microsoft .....http://support.microsoft.com/default.aspx?scid=http://download.microsoft.com/download/access2000/viewer/1/win98/EN-US/SnpVw90.exe"
    //'DoCmd.SendObject acSendReport, LetterName, "SnapshotFormat(*.snp)", _
    //'recipient, , , "Letter from IRB", MailSubject, True
                                ;
//'End If
;
//Exit_Mail:;
    //Exit //Sub;
//MailProblem:;
//If //Err.Number = //2501 //Then;
    //MsgBox "The Email Message was not sent";
    //Resume //Exit_Mail;
//Else;
    //MsgBox //Err.Description //& "----";
    //Resume //Exit_Mail;
//End //If;
}
function addPI_Click(){
//Call //sInvestigatorAdd;
}
function CCCOINV_Click(){
//could initialize MaxCCCOINV here
;
MaxCCCOINV = Me.Coinvcombo.ListCount;
//If MaxCCCOINV = 0 Then;
        //MsgBox //("No //Co-Investigator //to //Add");
        //Exit //Sub;
//End //If;
;
//could initialize i here
//While i //< MaxCCCOINV;
//If //Not //Me.CCCOINV //Then;
    //Me.HoldCCKey = //Me.Coinvcombo.Column(4, //i);
    //Call //DeleteCC;
//Else;
    //Me.HoldCCKey = //Me.Coinvcombo.Column(4, //i);
    //Call //AppendToCC;
//End //If;
 i = i + 1;
//Wend;
//If //Me.CheckCCEmail //Then;
    //Call //CheckEmailAddresses;
//End //If;
;
;
}
function CCCoord_Click(){
//If //IsNull(Me.[Coordinator //ID]) //Then;
        //MsgBox //("No //Coordinator //to //Add");
        //Exit //Sub;
//End //If;
;
;
//Me.HoldCCKey = //Me.CoordCombo;
//If //Not //Me.CCCoord //Then;
        //Call //DeleteCC;
//Else;
    //Call //AppendToCC;
//End //If;
//If //Me.CheckCCEmail //Then;
    //Call //CheckEmailAddresses;
//End //If;
;
}
function CCIRBMembers_Click(){
//could initialize MaxIRBs here
;
MaxIRBs = Me.IRBCombo.ListCount;
//If MaxIRBs = 0 Then;
        //MsgBox //("No //IRB //Member //to //Add");
        //Exit //Sub;
//End //If;
;
//could initialize i here
//While i //< MaxIRBs;
//If //Not //Me.CCIRBMembers //Then;
    //Me.HoldCCKey = //Me.IRBCombo.Column(0, //i);
    //Call //DeleteCC;
//Else;
    //Me.HoldCCKey = //Me.IRBCombo.Column(0, //i);
    //Call //AppendToCC;
//End //If;
 i = i + 1;
//Wend;
//If //Me.CheckCCEmail //Then;
    //Call //CheckEmailAddresses;
//End //If;
;
;
}
function CCPI_Click(){
//Me.HoldCCKey = //Me.PICombo;
//If //Not //Me.CCPI //Then;
        //Call //DeleteCC;
//Else;
    //Call //AppendToCC;
;
//End //If;
;
//If //Me.CheckCCEmail //Then;
    //Call //CheckEmailAddresses;
//End //If;
;
}
function CCReviewers_Click(){
//could initialize MaxReviewers here
MaxReviewers = Me.ReviewerCombo.ListCount;
//If MaxReviewers = 0 Then;
        //MsgBox //("No //Reviewers //to //Add");
        //Exit //Sub;
//End //If;
//could initialize i here
i = 0;
//While i //< MaxReviewers;
//If //Not //Me.CCReviewers //Then;
    //Me.HoldCCKey = //Me.ReviewerCombo.ItemData(i);
    //Call //DeleteCC;
//Else;
    //Me.HoldCCKey = //Me.ReviewerCombo.ItemData(i);
    //Call //AppendToCC;
//End //If;
 i = i + 1;
//Wend;
//If //Me.CheckCCEmail //Then;
    //Call //CheckEmailAddresses;
//End //If;
;
;
}
function CCSponsor_Click(){
//Me.HoldCCKey = //Me.SponsorCombo.Column(0);
//If //Not //Me.CCSponsor //Then;
        //Call //DeleteCC;
//Else;
    //Call //AppendToCC;
;
//End //If;
;
//If //Me.CheckCCEmail //Then;
    //Call //CheckEmailAddresses;
//End //If;
;
;
;
}
function ChangeButton_Click(){
//DoCmd.SetWarnings //False;
location.href = "Form_ThisLetterClosing.php";//, //, //, //, //, //acDialog;
//Me.ThisLetterClosing.Requery;
//Me.ThisLetterClosing = //Me.ThisLetterClosing.ItemData(0);
//DoCmd.SetWarnings //True;
}
function Check143_Click(){
//End //Sub;
;
//Private //Sub //CheckCCEmail_Click();
//If //Me.CheckCCEmail //Then;
       //Call //CheckEmailAddresses;
//End //If;
;
;
;
;
;
}
function CheckResponseRequired_Click(){
//If //Me.CheckResponseRequired //Then;
    //Me.RequiredByDate.Enabled = //True;
    //Me.FollowUpDate.Enabled = //True;
    //Me.RequiredByDate.SetFocus;
//Else;
    //Me.RequiredByDate = " ";
    //Me.FollowUpDate = " ";
    //Me.RequiredByDate.Enabled = //False;
    //Me.FollowUpDate.Enabled = //False;
//End //If;
}
function PrintLetter(){
//On //Error //GoTo //error_PrintLetter;
//LetterSwitch = //True;
//If //PreviewSwitch //Then;
     //If //CustomLetter //Then;
            //If //Not //MSWordPreview(Me.DisplayLetterName.Caption) //Then //Application.Quit;
            //GoTo //exit_PrintLetter;
     //Else;
   ;
        //DoCmd.OpenReport //LetterName, //acViewPreview;
     ;
     //End //If;
//Else;
        //If //CustomLetter //Then;
        ;
            //If //Me.ChSponsorPostal //Or //Me.CoordPostal //Or //Me.pipostal //_;
                //Or //Me.ReviewerPostal //Or //Me.IRBPostal //Or //Me.OtherPostal //Then;
                    //If //Not //MSWordRoutine(TempTemplatePath, "x"//) //Then;
                        //GoTo //error_PrintLetter;
                    //End //If;
            //End //If;
        //Else;
            //DoCmd.OpenReport //LetterName, //acViewNormal;
    ;
        //End //If;
//End //If;
   ;
;
//exit_PrintLetter:;
    //Exit //Sub;
//error_PrintLetter:;
    //MsgBox //Err.Description //& "-Error in Printing--" //& //LetterName;
    //PrintError = //True;
    //GoTo //exit_PrintLetter;
}
function cModify_Click(){
//Call //PreviewTemplate;
}
function Coinvcombo_Click(){
//MsgBox "You cannot send a Letter addressed directly to a Co-Investigator in this version.  If you must, then temporarily select the Investigator's name from the Principal Investigator's box and then print the letter.";
}
function Command135_Click(){
//DoCmd.SetWarnings //False;
//DoCmd.OpenQuery "DeleteClosingFromCurrentLetter";
//DoCmd.OpenQuery "AppendDefaultClosingToCurrent";
//Me.ThisLetterClosing.Requery;
//Me.ThisLetterClosing = //Me.ThisLetterClosing.ItemData(0);
//DoCmd.SetWarnings //True;
}
function CoordCheck_Click(){
//If //IsNull(Me.CoordCombo) //Then;
        //If //GeneralLetterSwitch //Then;
            //DisplayMessage //("You //must //select //a //Coordinator");
        //Else;
               //DisplayMessage //("You //must //return //to //the //Study //and //assign //a //Coordinator");
        //End //If;
        //Me.CoordCheck = //False;
        //Exit //Sub;
//End //If;
//If //Me.CoordCheck //Then;
        //If //CheckAlreadySet //Then;
            //Me.CoordCheck = //False;
            //Exit //Sub;
        //Else;
            //gbAlreadySet = //True;
        //End //If;
      ;
    ;
        //Me.CoordPostal.Enabled = //True;
        //Me.Coordemail.Enabled = //True;
        //Me.CoordPostal = //True;
        //Me.Coordemail = //False;
//Else;
        //gbAlreadySet = //False;
        //Me.CoordPostal.Enabled = //False;
        //Me.Coordemail.Enabled = //False;
        //Me.CoordPostal = //False;
        //Me.Coordemail = //False;
//End //If;
}
function CoordCombo_DblClick(Cancel){
//Call //sViewCoordinator;
//Me.Coordemail = //False;
    //gbEmailSet = //False;
//Me.ListCCs.Requery;
}
function Coordemail_Click(){
//If //Demo //Then;
    //MsgBox "Not available in Trial Version";
    //Me.Coordemail = //False;
//Else;
    //If //Me.Coordemail //Then;
        //Me.CheckCCEmail = //True;
        //If //IsNull(Me.CoordCombo.Column(15)) //Or //Me.CoordCombo.Column(15) = "" //Then;
            //MsgBox "Coordinator's Email Address is Blank";
            //Me.Coordemail = //False;
        //End //If;
    //Else;
        //Me.CheckCCEmail = //False;
    //End //If;
//End //If;
//If //Me.Coordemail //Then //gbEmailSet = //True //Else //gbEmailSet = //False;
;
}
function Envelope_Click(){
//On //Error //GoTo //Error_env;
//If //Demo //Then;
    //MsgBox "Not available in Trial Version";
//Else;
     //If //AddressEnvelope //Then;
        //If //YesNo("Do //you //wish //to //Preview //an //Envelope?") //Then;
            //Call //PrintEnv("Preview");
        //Else;
            //MsgBox "Load Envelope ";
            //Call //PrintEnv("Normal");
            //Call //CCEnvelopes;
        //End //If;
    //End //If;
   ;
     ;
//End //If;
;
//exit_Env:;
//Exit //Sub;
;
//Error_env:;
//MsgBox //Err.Description //& "  " //& //Err.Number;
//Resume //exit_Env;
}
function Form_Close(){
//On //Error //GoTo //err_Form_Close;
//Me.TContuingReviewHold = "         ";
//DoCmd.SetWarnings //False;
//DoCmd.OpenQuery "_DeleteCurrentLetterParagraphs";
//DoCmd.SetWarnings //True;
//Me.txtContactModuleID = //Null;
;
//FromICChecklist = //False;
//RenewalLetterSwitch = //False;
//SAELetterSwitch = //False;
//CPALetterSwitch = //False;
//GeneralLetterSwitch = //False;
//EducationLetterSwitch = //False;
;
//LetterManagerSwitch = //False;
//FromPostIRB = //False;
//StudySpecificSwitch = //False;
//IRBInvoiceSwitch = //False;
//LetterSelectedSwitch = //False;
//FollowupLetterSwitch = //False;
//PrintError = //False;
//LettersSentError = //False;
//LetterSwitch = //False;
//If //Not //CustomLetter //Then;
    //If //LetterName = "" //Then;
    //Else;
        //DoCmd.Close //acReport, //LetterName, //acSaveNo;
        //DoCmd.Close //acReport, "Number10Envelope"//, //acSaveNo;
    //End //If;
//End //If;
//GoTo //Exit_Form_Close;
;
//err_Form_Close:;
//Err.Clear;
//Resume //Exit_Form_Close;
;
//Exit_Form_Close:;
}
function Form_Load(){
//If //GeneralLetterSwitch //Then;
    //Me.RecordSource = "DummySelectForGeneralLetter";
    //Me.SponsorCombo.ControlSource = "";
    //Me.PICombo.ControlSource = "";
    //Me.CoordCombo.ControlSource = "";
    //Exit //Sub;
//End //If;
;
    ;
;
}
function Form_Open(Cancel){
//If //isFormLoaded("EducationTrainingForm") //Then;
    //Forms![_LetterManager]![txtContactModuleID] = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(10);
 //End //If;
//If //isFormLoaded("frmEducationModulesDue") //Then;
    //Forms![_LetterManager]![txtContactModuleID] = //Forms![frmeducationmodulesDue]![txtContactModuleID];
//End //If;
//If //IsNull(Me.txtContactModuleID) //Then //Me.txtContactModuleID = //0;
;
;
//SearchContacts = //False;
//gbEmailSet = //False;
//AlreadyCreatedReviewer = //False;
//AlreadyCreatedIRB = //False;
//could initialize i here
//TempTemplatePath = "C:\CC\PROIRB_Dev\TempTemplate.dot";
;
//' empty the CC table for previous letter
//DoCmd.SetWarnings //False;
//DoCmd.RunSQL "delete * from CurrentLetter_CCs";
//DoCmd.SetWarnings //True;
//Me.ListCCs.Requery;
//gbAlreadySet = //False;
//gbLetterFileCreated = //False;
//Me.TextEnvCount = "1";
;
//If //CustomLetter //Then;
    //On //Error //Resume //Next;
    //If //dir(TempTemplatePath) //<> "" //Then;
      //Kill //(TempTemplatePath);
    //End //If;
    //FileCopy //TemplateName, //TempTemplatePath;
    i = Len(dir(TemplateName));
    //HoldDocNAme = //Left$(dir(TemplateName), i //- //4);
    //Me.AddCommentsButton.Enabled = //False;
    ;
//Else;
    //Me.cModify.Enabled = //False;
    //HoldDocNAme = //LetterName;
   //' Me.PreviewLetter.SetFocus
    //Me.CCCoord.Enabled = //False;
    //Me.CCIRBMembers.Enabled = //False;
    //Me.CCPI.Enabled = //False;
    //Me.CCReviewers.Enabled = //False;
    //Me.CCSponsor.Enabled = //False;
    //Me.cmdAddDeleteCC.Enabled = //False;
    //Me.ListCCs.Enabled = //False;
    //Me.CheckCCEmail.Enabled = //False;
    //Me.AddCommentsButton.Enabled = //True;
//End //If;
;
;
;
;
//Me.TextEnvCount = //1;
;
//Me.SponsorCheck = //False;
//Me.CoordCheck = //False;
//Me.PICheck = //False;
//Me.pipostal = //False;
//Me.PIemail = //False;
//Me.Sponsoremail = //False;
//Me.ChSponsorPostal = //False;
//Me.CoordPostal = //False;
//Me.Coordemail = //False;
//Me.CheckResponseRequired = //False;
//LetterSelectedSwitch = //False;
;
//' Letter is  a General letter
 ;
 //If //GeneralLetterSwitch //Then;
    ;
    //Me.SponsorCombo.Locked = //False;
    //Me.PICombo.Locked = //False;
    //Me.CoordCombo.Locked = //False;
    //LetterName = "General Letter";
    //Call //SetLabelCaptions("persons.txt");
    //Me.ReviewerCombo.Visible = //False;
    //Me.ReviewerCheck.Visible = //False;
    //Me.ReviewerEmail.Visible = //False;
    //Me.ReviewerPostal.Visible = //False;
    ;
    //Me.CCSponsor.Enabled = //False;
    //Me.CCPI.Enabled = //False;
    //Me.CCCoord.Enabled = //False;
    //Me.CCReviewers.Enabled = //False;
    //'Me.CCIRBMembers.Enabled = False
    ;
    //Me.IRBText.Visible = //False;
    //Me.TextMeetingDate.Visible = //True;
    //Me.CheckResponseRequired.Enabled = //False;
    //Me.RequiredByDate.Enabled = //False;
    //Me.FollowUpDate.Enabled = //False;
    ;
    //DoCmd.SetWarnings //True;
    //GoTo //ExitSub1;
  ;
//End //If;
;
;
//'Letter is an Education Module Due
 //If //EducationLetterSwitch //Then;
 ;
 ;
 ;
  //Me.CheckResponseRequired.Enabled = //True;
    //Me.RequiredByDate.Enabled = //True;
    //Me.FollowUpDate.Enabled = //True;
    //Me.SponsorCheck.Locked = //True;
    //Me.SponsorCheck = //True;
    //gbAlreadySet = //True;
   //Me.SponsorCombo.Locked = //True;
   //Me.Sponsoremail.Enabled = //True;
   //Me.ChSponsorPostal.Enabled = //True;
   //Me.ChSponsorPostal = //True;
   //' Me.SponsorCombo.Enabled = True
    //Me.Label19.Caption = "Addressee";
    //Me.Label22.Visible = //False;
    //Me.Label24.Visible = //False;
    //Me.Label142.Visible = //False;
    //Me.Label144.Visible = //False;
    ;
    ;
    //Me.Other_Contact.Visible = //False;
    ;
    //Me.PICombo.Visible = //False;
    //Me.CoordCombo.Visible = //False;
    //Me.IRBCombo.Visible = //False;
    //Me.ReviewerCombo.Visible = //False;
    //Me.OtherCombo.Visible = //False;
    //Me.PICheck.Visible = //False;
    //Me.CoordCheck.Visible = //False;
    //Me.IRBCheck.Visible = //False;
    //Me.ReviewerCheck.Visible = //False;
    //Me.OtherCheck.Visible = //False;
    //Me.pipostal.Visible = //False;
    //Me.PIemail.Visible = //False;
    //Me.CoordPostal.Visible = //False;
    //Me.Coordemail.Visible = //False;
    //Me.ReviewerPostal.Visible = //False;
    //Me.ReviewerEmail.Visible = //False;
    //Me.IRBPostal.Visible = //False;
    //Me.IRBEmail.Visible = //False;
    //Me.OtherPostal.Visible = //False;
    //Me.OtherEmail.Visible = //False;
    //LetterName = "Education Letter";
    //Call //SetLabelCaptions("education.txt");
    ;
    //Me.CCSponsor.Enabled = //False;
    //Me.CCPI.Enabled = //False;
    //Me.CCCoord.Enabled = //False;
    //Me.CCReviewers.Enabled = //False;
    //'Me.CCIRBMembers.Enabled = False
    ;
    //Me.IRBText.Visible = //False;
    //Me.TextMeetingDate.Visible = //False;
    ;
    //DoCmd.SetWarnings //True;
    //GoTo //ExitSub1;
  ;
//End //If;
 ;
               ;
             //'Informed Consent Deficiencies
//If //FromICChecklist = //True //Then;
    ;
   //' Me.LetterLabel.Visible = True
    //'Me.LetterLabel.Caption = "Informed Consent Deficiencies"
    //Call //SetLabelCaptions("ICDef.txt");
    //Me.CheckResponseRequired = //True;
    //Me.RequiredByDate.Enabled = //True;
    //Me.RequiredByDate = //Date //+ //14;
    //Me.FollowUpDate.Enabled = //True;
    //Me.FollowUpDate = //Date //+ //7;
    //Me.TextMeetingDate = //Forms![_irb //input //form].AgendaDate;
    //GoTo //ExitSub1;
//Else;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "_DeleteCurrentLetterParagraphs";
        //DoCmd.SetWarnings //True;
//End //If;
;
//'SAE acknowledgment Letter
;
//'If SAELetterSwitch Or CPALetterSwitch Then
    ;
;
 //'   Me.TextMeetingDate.Visible = True
  //'  Me.CheckResponseRequired = False
   //' Me.RequiredByDate.Enabled = False
    //'Me.FollowUpDate.Enabled = False
    //'Me.TextMeetingDate = HoldMeetingDate
    //'If SAELetterSwitch Then Call SetLabelCaptions("SAE.txt")
    //'If CPALetterSwitch Then Call SetLabelCaptions("CPA.txt")
    //'GoTo ExitSub1
//'End If
//'REPLACED THE ABOVE CODE WITH THIS
//If //SAELetterSwitch //Or //CPALetterSwitch //Then;
    ;
;
    //Me.TextMeetingDate.Visible = //True;
    //If //Forms![newmenu]![txtTemplateName] //Like "*Fee*" //Then;
        //Me.CheckResponseRequired = //True;
        //Me.RequiredByDate.Enabled = //True;
        //Me.FollowUpDate.Enabled = //True;
        //Me.RequiredByDate = //DateAdd("d", //-14, //HoldMeetingDate) //'changed 7/23/01
        //Me.FollowUpDate = //DateAdd("d", //-21, //HoldMeetingDate) //' changed 7/23/01
    //Else;
        //Me.CheckResponseRequired = //False;
        //Me.RequiredByDate.Enabled = //False;
        //Me.FollowUpDate.Enabled = //False;
    //End //If;
    ;
    //Me.TextMeetingDate = //HoldMeetingDate;
    //If //SAELetterSwitch //Then //Call //SetLabelCaptions("SAE.txt");
    //If //CPALetterSwitch //Then //Call //SetLabelCaptions("CPA.txt");
    //GoTo //ExitSub1;
//End //If;
;
//'Continuing Review Letter
//If //RenewalLetterSwitch = //True //Then;
    //Me.TContuingReviewHold = "Renewal";
    //Call //SetLabelCaptions("Renewal.txt");
    //Me.TextMeetingDate.Visible = //True;
    //Me.CheckResponseRequired = //True;
    //Me.RequiredByDate.Enabled = //True;
    //Me.FollowUpDate.Enabled = //True;
    //DoCmd.SetWarnings //True;
    //Me.TextMeetingDate = //HoldMeetingDate //'added 7/23
    //Me.RequiredByDate = //DateAdd("d", //-14, //HoldMeetingDate) //'changed 7/23/01
    //Me.FollowUpDate = //DateAdd("d", //-21, //HoldMeetingDate) //' changed 7/23/01
    //GoTo //ExitSub1;
//End //If;
;
                    //'From Post Meeting Agenda
;
//If //FromPostIRB //Then;
   ;
    //Call //SetLabelCaptions("IRBAction.txt");
    //Me.TextMeetingDate.Visible = //True;
    //Me.CheckResponseRequired.Enabled = //True;
    //Me.CheckResponseRequired = //False;
    //Me.RequiredByDate.Enabled = //False;
    //Me.FollowUpDate.Enabled = //False;
    ;
    //Me.TextMeetingDate = //HoldMeetingDate;
    //DoCmd.SetWarnings //True;
    ;
    //GoTo //ExitSub1;
//End //If;
//If //FollowupLetterSwitch //Then;
     //Call //SponsorCheck_Click;
        ;
    //Call //SetLabelCaptions("FU.txt");
    //Me.CheckResponseRequired = //True;
    //Me.RequiredByDate.Enabled = //True;
    //Me.RequiredByDate = //(Date //+ //7);
    //Me.FollowUpDate.Enabled = //True;
    //Me.FollowUpDate = //(Me.RequiredByDate //- //3);
    ;
    //GoTo //ExitSub1;
//End //If;
//If //StudySpecificSwitch //Then;
    //Call //SetLabelCaptions("Study.txt");
    //Me.CheckResponseRequired = //False;
    //Me.RequiredByDate.Enabled = //False;
    //Me.FollowUpDate.Enabled = //False;
//End //If;
//If //IRBInvoiceSwitch //Then;
    //Call //SetLabelCaptions("Invoice.txt");
    //Me.CheckResponseRequired = //True;
    //Me.RequiredByDate.Enabled = //True;
    //Me.FollowUpDate.Enabled = //True;
    //Me.RequiredByDate = //(Date //+ //14);
    //Me.FollowUpDate = //(Me.RequiredByDate //- //5);
//End //If;
    //'Me.LetterLabel.Visible = True
    ;
    //Me.TextMeetingDate.Visible = //True;
    ;
    //DoCmd.SetWarnings //True;
    //Me.TextMeetingDate = //HoldMeetingDate;
    //GoTo //ExitSub1;
;
//ExitSub1:;
;
//'Me.ThisLetterClosing.Requery
;
;
//'Me.ThisLetterClosing = Me.ThisLetterClosing.ItemData(0)
//'MsgBox "who knows"
//'MsgBox [Forms]![_LetterManager]![IRB Code]
//'MsgBox [Forms]![_LetterManager]![LetterLabel].[Caption]
//'If IsNull(Me.ThisLetterClosing) Then
 //'  DoCmd.SetWarnings True
 //' DoCmd.OpenQuery "AppendDefaultClosingToCurrent"
 //'Me.ThisLetterClosing.Requery
 //'Me.ThisLetterClosing = Me.ThisLetterClosing.ItemData(0)
 //'DoCmd.SetWarnings True
 ;
 //If //EducationLetterSwitch //Then;
    //Me.RecordSource = "DummySelectForEducationLetter";
    //Me.SponsorCombo.ControlSource = "";
   ;
    //Me.SponsorCombo.RowSourceType = "Table/Query";
    ;
    //Me.SponsorCombo.BoundColumn = //1;
    //If //isFormLoaded("EducationTrainingForm") //Then;
        //Forms![_LetterManager]![txtContactModuleID] = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(10);
        ;
        //Me.SponsorCombo.RowSource = "qrySelectContactDatafromContactEducation";
    //Else;
        //Me.SponsorCombo.RowSource = "qrySelectContactDataForEducationLetter";
    //End //If;
   //Me.SponsorCombo = //Me.SponsorCombo.ItemData(0);
//'
    //Me.PICombo.ControlSource = "";
    //Me.CoordCombo.ControlSource = "";
    //DoCmd.SetWarnings //False;
    //Me.ThisLetterClosing.Requery;
    //Me.ThisLetterClosing = //Me.ThisLetterClosing.ItemData(0);
    //If //IsNull(Me.ThisLetterClosing) //Then;
        //DoCmd.OpenQuery "AppendDefaultClosingToCurrent";
        //Me.ThisLetterClosing.Requery;
        //Me.ThisLetterClosing = //Me.ThisLetterClosing.ItemData(0);
        //DoCmd.SetWarnings //True;
    //End //If;
 ;
    //Me.ThisLetterClosing.Requery;
    //Me.ThisLetterClosing = //Me.ThisLetterClosing.ItemData(0);
//End //If;
    ;
;
;
;
//Exit //Sub;
;
;
}
function Form_Unload(Cancel){
//On //Error //GoTo //exit_close_err;
//If //Reports(0).Name = //LetterName //Then;
  //MsgBox "You must print or close the Letter First";
  //Cancel = //True;
//End //If;
//unload_exit:;
//Exit //Sub;
//exit_close_err:;
//If //Err.Number = //2457 //Then;
    //Resume //unload_exit;
    //Else;
    //MsgBox "PRO_IRB Error #070-Error in Closing Form or Report" //& " " //& //Err.Number;
    //Resume //unload_exit;
//End //If;
;
;
;
;
;
}
function SetUpSendTo(r){
//On //Error //GoTo //err_SetUpSendTo;
//'Display Name
//Me.TextLine7 = "Dear " //& //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(6, //r)) //& " " //& //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(7, //r)) //& ":";
        ;
//'Format City State Zip
//'msgbox Forms("_LetterManager")(SendToCntrl).Column(11)
//If //Not //Forms("_LetterManager")(SendToCntrl).Column(11, //r) = "" //Then;
 //Me.TextLine5 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(11, //r)) //& ", " //& //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(12, //r)) //_;
           //& ". " //& //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(13, //r)) //& "   " //& //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(14, //r));
//Else  //' no city state line
    //Me.TextLine5 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(5, //r));
    //Me.TextLine4 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(4, //r));
    //Me.TextLine3 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(3, //r));
    //Me.TextLine2 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(2, //r));
    //Me.TextLine1 = "  ";
    //GoTo //Again;
//End //If;
;
//Me.TextLine4 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(5, //r));
//Me.TextLine3 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(4, //r));
//Me.TextLine2 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(3, //r));
//Me.TextLine1 = //RTrim$(Forms("_LetterManager")(SendToCntrl).Column(2, //r));
;
//Again:;
//If //Me.TextLine1 = "" //Then;
    //Me.TextLine1 = //Me.TextLine2;
    //Me.TextLine2 = //Me.TextLine3;
    //Me.TextLine3 = //Me.TextLine4;
    //Me.TextLine4 = //Me.TextLine5;
    //Me.TextLine5 = "  ";
    //GoTo //Again;
//End //If;
;
//If //Me.TextLine2 = "" //Then;
        //Me.TextLine2 = //Me.TextLine3;
        //Me.TextLine3 = //Me.TextLine4;
        //Me.TextLine4 = //Me.TextLine5;
        //Me.TextLine5 = "  ";
        //GoTo //Again;
//End //If;
;
//If //Me.TextLine3 = "" //Then;
        //Me.TextLine3 = //Me.TextLine4;
        //Me.TextLine4 = //Me.TextLine5;
        //Me.TextLine5 = "  ";
        //GoTo //Again;
//End //If;
//If //Me.TextLine4 = "" //Then;
        //Me.TextLine4 = //Me.TextLine5;
        //Me.TextLine5 = "  ";
        //GoTo //Again;
//End //If;
//exit_SetUpSendTo:;
    //Exit //Sub;
//err_SetUpSendTo:;
;
//If //Err.Description //Like "*null*" //Then;
    //Resume //exit_SetUpSendTo;
//Else;
    //MsgBox "Error in Setup Send to  " //& //Err.Description;
        //Resume //exit_SetUpSendTo;
//End //If;
}
function UpdateSendTo(stDocName){
//On //Error //GoTo //Err_Preview_Click;
var ErrorMessage = '';
ErrorMessage = "Error in Generating Letter";
//DoCmd.SetWarnings //False;
//'MsgBox Me.Protocol_Number___Title
//DoCmd.OpenQuery "_deleteCurrentLetterSendTo";
//DoCmd.OpenQuery "_deletecurrentletterprotocol";
//'msgbox Me.TextEnvCount
//DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
//DoCmd.OpenQuery "_appendProtocoltoCurrentLetter";
//DoCmd.SetWarnings //True;
;
;
//Exit_Preview_Click:;
    //Exit //Sub;
;
//Err_Preview_Click:;
    //MsgBox ErrorMessage //& " " //& //Err.Description //& " " //& //Err.Number;
    //Resume //Exit_Preview_Click;
    ;
}
function SendThem(){
//On //Error //GoTo //Err_SendThem;
    //could initialize i here
//DoCmd.SetWarnings //False;
//CreatedSentLetters = //False;
;
;
//If //Me.ChSponsorPostal //Then;
    //Me.strSentTO = "Sponsor Post";
     //SendToCntrl = "SponsorCombo";
     //If //CustomLetter //Then;
        //Call //ExportDataSource(Me.SponsorCombo.Column(0, //Me.SponsorCombo.ListIndex));
        //Call //PrintCustomLetter;
    //Else;
        //Call //DoItAll_Letter(Me.SponsorCombo.ListIndex);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedSponsor //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedSponsor = //True;
        //End //If;
    //End //If;
//End //If;
;
//If //Me.Sponsoremail //Then;
    //Me.strSentTO = "Sponsor Email";
    //EmailAddress = //Me.SponsorCombo.Column(14, //Me.SponsorCombo.ListIndex);
    //SendToCntrl = "SponsorCombo";
    //If //CustomLetter //Then;
        //Call //CustomEmails(EmailAddress, //Me.SponsorCombo.Column(0, //Me.SponsorCombo.ListIndex));
;
    //Else;
        //Call //DoItAll_Email(Me.SponsorCombo.ListIndex, //EmailAddress);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedSponsor //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedSponsor = //True;
        //End //If;
    //End //If;
//End //If;
;
//'PI letters and Email
;
//If //Me.pipostal //Then;
    //Me.strSentTO = "PI Post";
    //SendToCntrl = "PICombo";
    //If //CustomLetter //Then;
        //Call //ExportDataSource(Me.PICombo.Column(0, //Me.PICombo.ListIndex));
        //Call //PrintCustomLetter;
    ;
    //Else;
        //Call //DoItAll_Letter(Me.PICombo.ListIndex);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedPI //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedPI = //True;
        //End //If;
    //End //If;
;
//End //If;
;
//If //Me.PIemail //Then;
    //Me.strSentTO = "PI Email";
    //EmailAddress = //Me.PICombo.Column(15, //Me.PICombo.ListIndex);
    //SendToCntrl = "PICombo";
    //If //CustomLetter //Then;
        //Call //CustomEmails(EmailAddress, //Me.PICombo.Column(0, //Me.PICombo.ListIndex));
;
    //Else;
        //Call //DoItAll_Email(Me.PICombo.ListIndex, //EmailAddress);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedPI //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedPI = //True;
        //End //If;
    //End //If;
;
    ;
//End //If;
;
//'Coord letter and Email
;
//If //Me.CoordPostal //Then;
    //Me.strSentTO = "Coord Post";
    //SendToCntrl = "CoordCombo";
    //If //CustomLetter //Then;
        //Call //ExportDataSource(Me.CoordCombo.Column(0, //Me.CoordCombo.ListIndex));
        //Call //PrintCustomLetter;
    //Else;
        //Call //DoItAll_Letter(Me.CoordCombo.ListIndex);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedCoord //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedCoord = //True;
        //End //If;
    //End //If;
;
    ;
//End //If;
;
//If //Me.Coordemail //Then;
    //Me.strSentTO = "Coord Email";
    //EmailAddress = //Me.CoordCombo.Column(15, //Me.CoordCombo.ListIndex);
    //SendToCntrl = "CoordCombo";
    //If //CustomLetter //Then;
        //Call //CustomEmails(EmailAddress, //Me.CoordCombo.Column(0, //Me.CoordCombo.ListIndex));
        ;
    //Else;
        //Call //DoItAll_Email(Me.CoordCombo.ListIndex, //EmailAddress);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedCoord //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedCoord = //True;
        //End //If;
    //End //If;
//End //If;
;
//'Other Contact letter and Email
;
//If //Me.OtherPostal //Then;
    //Me.strSentTO = "Contact Post";
    //SendToCntrl = "OtherCombo";
    //If //CustomLetter //Then;
        //Call //ExportDataSource(Me.OtherCombo.Column(0, //Me.OtherCombo.ListIndex));
        //Call //PrintCustomLetter;
    //Else;
        //Call //DoItAll_Letter(Me.OtherCombo.ListIndex);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedOther //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedOther = //True;
        //End //If;
    //End //If;
//End //If;
;
//If //Me.OtherEmail //Then;
    //Me.strSentTO = "Contact Email";
    //EmailAddress = //Me.OtherCombo.Column(15, //Me.OtherCombo.ListIndex);
    //SendToCntrl = "OtherCombo";
    //If //CustomLetter //Then;
        //Call //CustomEmails(EmailAddress, //Me.OtherCombo.Column(0, //Me.OtherCombo.ListIndex));
;
    //Else;
        //Call //DoItAll_Email(Me.OtherCombo.ListIndex, //EmailAddress);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedOther //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedOther = //True;
        //End //If;
    //End //If;
;
;
//End //If;
//'Reviewer Contact letter and Email
//'Reviewers
;
;
;
//If //Me.ReviewerPostal //Then;
    //If //Me.ReviewerCombo.ListIndex = //-1 //Then;
        //DisplayMessage //("You //must //select //a //Reviewer");
        //Exit //Sub;
    //End //If;
    //Me.strSentTO = "Reviewers Post";
    //SendToCntrl = "ReviewerCombo";
    //If //CustomLetter //Then;
        //Call //ExportDataSource(Me.ReviewerCombo.Column(0, //Me.ReviewerCombo.ListIndex));
        //Call //PrintCustomLetter;
        ;
    //Else;
        //Call //DoItAll_Letter(Me.ReviewerCombo.ListIndex);
        ;
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedReviewer //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedReviewer = //True;
        //End //If;
    //End //If;
    ;
//End //If;
;
;
//If //Me.ReviewerEmail //Then;
    //If //Me.ReviewerCombo.ListIndex = //-1 //Then;
        //DisplayMessage //("You //must //select //a //Reviewer");
        //Exit //Sub;
    //End //If;
    //Me.strSentTO = "Reviewer Email";
    //EmailAddress = //Me.ReviewerCombo.Column(15, //Me.ReviewerCombo.ListIndex);
    //SendToCntrl = "ReviewerCombo";
    //If //CustomLetter //Then;
        //Call //CustomEmails(EmailAddress, //Me.ReviewerCombo.Column(0, //Me.ReviewerCombo.ListIndex));
    //Else;
        //Call //DoItAll_Email(Me.ReviewerCombo.ListIndex, //EmailAddress);
    //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedReviewer //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedReviewer = //True;
        //End //If;
    //End //If;
    ;
//End //If;
;
;
//'IRB Members
//If //Not //Me.IRBCheck //Then //Exit //Sub;
//If //Not //Me.IRBPostal //Then //GoTo //IrbEmailRoutine;
//Me.strSentTO = "Members Post";
//SendToCntrl = "IRBCombo";
//If //CustomLetter //Then;
   //Call //ExportDataSource(Me.IRBCombo.Column(0, //Me.IRBCombo.ListIndex));
   //Call //PrintCustomLetter;
//Else;
    ;
   //Call //DoItAll_Letter(Me.IRBCombo.ListIndex);
//End //If;
//If //PreviewSwitch //Then;
//Else;
        //If //AlreadyCreatedIRB //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedIRB = //True;
        //End //If;
//End //If;
    ;
    ;
;
//IrbEmailRoutine:;
//If //Not //Me.IRBEmail //Then //GoTo //Exit_SendThem;
;
;
//If //Me.IRBEmail //Then;
    //Me.strSentTO = "Member Email";
    //EmailAddress = //Me.IRBCombo.Column(15, //Me.IRBCombo.ListIndex);
    //SendToCntrl = "IRBCombo";
        //If //CustomLetter //Then;
            //Call //CustomEmails(EmailAddress, //Me.IRBCombo.Column(0, //Me.IRBCombo.ListIndex));
            ;
        //Else;
            //Call //DoItAll_Email(Me.IRBCombo.ListIndex, //EmailAddress);
        //End //If;
    //If //PreviewSwitch //Then;
    //Else;
        //If //AlreadyCreatedIRB //Then;
        //Else;
            //Call //CreateLettersSent;
            //AlreadyCreatedIRB = //True;
        //End //If;
    //End //If;
//End //If;
;
;
//Exit_SendThem:;
//DoCmd.SetWarnings //True;
//gbLetterFileCreated = //False;
;
//Exit //Sub;
;
//Err_SendThem:;
//MsgBox //Err.Description;
//Resume //Exit_SendThem;
;
;
}
function IRBAdd_Click(){
//Call //sIRBAdd;
}
function IRBCheck_Click(){
//If //Me.IRBCheck //Then;
  //If //CheckAlreadySet //Then;
       //Me.IRBCheck = //False;
       //Exit //Sub;
  //Else;
      //gbAlreadySet = //True;
  //End //If;
       ;
    //Me.IRBPostal.Enabled = //True;
    //Me.IRBEmail.Enabled = //True;
    //Me.IRBPostal = //True;
    //Me.IRBEmail = //False;
   //If //Me.IRBCombo.ListIndex = //-1 //Then;
       //DisplayMessage //("You //must //Select //a //Member");
       //Me.IRBCheck = //False;
       //Me.IRBEmail = //False;
       //Me.IRBPostal = //False;
       //gbAlreadySet = //False;
       //Exit //Sub;
   //End //If;
;
;
//Else;
    //gbAlreadySet = //False;
    //Me.IRBPostal.Enabled = //False;
    //Me.IRBEmail.Enabled = //False;
    //Me.IRBPostal = //False;
    //Me.IRBEmail = //False;
        ;
//End //If;
;
}
function IRBCombo_DblClick(Cancel){
//Call //sViewIRB;
//Me.ListCCs.Requery;
//Me.IRBEmail = //False;
   //gbEmailSet = //False;
}
function IRBEmail_Click(){
//If //Demo //Then;
    //MsgBox "Not available in Trial Version";
    //Me.IRBEmail = //False;
//Else;
    //If //Me.IRBEmail //Then;
         //Call //IRBEmailCheck;
    //Else;
        //Me.CheckCCEmail = //False;
    //End //If;
//End //If;
}
function OtherAdd_Click(){
//Call //sOtherAdd;
}
function OtherCheck_Click(){
//If //Me.OtherCheck //Then;
    //If //IsNull(Me.OtherCombo) //Then;
        //DisplayMessage //("You //must //select //a //Contact");
        //Me.OtherCheck = //False;
        //Exit //Sub;
    //End //If;
//End //If;
;
    //If //Me.OtherCheck //Then;
        //If //CheckAlreadySet //Then;
            //Me.OtherCheck = //False;
            //Exit //Sub;
        //Else;
            //gbAlreadySet = //True;
        //End //If;
    //Me.OtherEmail.Enabled = //True;
    //Me.OtherPostal.Enabled = //True;
    //Me.OtherPostal = //True;
    //Me.OtherEmail = //False;
//Else;
    //gbAlreadySet = //False;
   //Me.OtherEmail.Enabled = //False;
   //Me.OtherPostal.Enabled = //False;
   //Me.OtherEmail = //False;
   //Me.OtherPostal = //False;
//End //If;
;
}
function OtherCombo_DblClick(Cancel){
//Call //sViewOther;
//Me.ListCCs.Requery;
//Me.OtherEmail = //False;
    //gbEmailSet = //False;
;
}
function OtherEmail_Click(){
//If //Demo //Then;
    //MsgBox "Not available in Trial Version";
    //Me.OtherEmail = //False;
//Else;
    //If //Me.OtherEmail //Then;
        //Me.CheckCCEmail = //True;
        //If //IsNull(Me.OtherCombo.Column(15)) //Or //Me.OtherCombo.Column(15) = "" //Then;
            //MsgBox "Contacts's Email Address is Blank";
            //Me.OtherEmail = //False;
        //End //If;
    //Else;
        //Me.CheckCCEmail = //False;
    //End //If;
//End //If;
//If //Me.OtherEmail //Then //gbEmailSet = //True //Else //gbEmailSet = //False;
;
}
function PICheck_Click(){
//If //Me.PICheck //Then;
    //If //IsNull(Me.PICombo) //Then;
        //DisplayMessage //("You //must //select //an //Investigator");
        //Me.PICheck = //False;
        //Exit //Sub;
    //End //If;
//End //If;
//If //Me.PICheck //Then;
    //If //CheckAlreadySet //Then;
        //Me.PICheck = //False;
        //Exit //Sub;
    //Else;
        //gbAlreadySet = //True;
    //End //If;
      ;
    //Me.PIemail.Enabled = //True;
    //Me.pipostal.Enabled = //True;
  ;
    //Me.pipostal = //True;
    //Me.PIemail = //False;
//Else;
    //gbAlreadySet = //False;
   //Me.PIemail.Enabled = //False;
   //Me.pipostal.Enabled = //False;
   //Me.pipostal = //False;
   //Me.PIemail = //False;
//End //If;
}
function PICombo_DblClick(Cancel){
//Call //sViewInvestigator;
//Me.PIemail = //False;
    //gbEmailSet = //False;
//Me.ListCCs.Requery;
}
function PIemail_Click(){
//If //Demo //Then;
    //MsgBox "Not available in Trial Version";
    //Me.PIemail = //False;
//Else;
;
    //If //Me.PIemail //Then;
        //Me.CheckCCEmail = //True;
        ;
        //If //IsNull(PICombo.Column(15)) //Or //PICombo.Column(15) = "" //Then;
            //MsgBox "P.I. Email Address is Blank";
            //Me.PIemail = //False;
        //End //If;
    //Else;
    //Me.CheckCCEmail = //False;
    //End //If;
//End //If;
;
//If //Me.PIemail //Then //gbEmailSet = //True //Else //gbEmailSet = //False;
}
function PreviewLetter_Click(){
//If //LetterName = "" //Then;
    //Else;
    //DoCmd.Close //acReport, //LetterName, //acSaveNo;
//End //If;
//'If FollowupLetterSwitch Then
 //'   MsgBox "The Follow-Up Letter is currently in development", vbInformation
  //'  Exit Sub
//'End If
//If //Me.CheckResponseRequired //Then;
    //If //Not //CheckDates //Then;
        //Me.RequiredByDate.SetFocus;
        //Exit //Sub;
    //End //If;
//End //If;
;
//If //Not //gbAlreadySet //Then;
    //MsgBox "Please select To Whom to send the letter to";
    //Exit //Sub;
//End //If;
;
//If //(Not //CustomLetter) //And //(gbEmailSet) //Then;
    //MsgBox "You cannot preview an Email-You can view it in the Email before sending";
    //Exit //Sub;
//End //If;
;
;
;
//If //Me.CheckResponseRequired = //True //Then;
    //If //IsNull(Me.RequiredByDate) //Or //IsNull(Me.FollowUpDate) //Then;
        //MsgBox "Required by and Followup dates cannot be blank";
        //Exit //Sub;
    //End //If;
//End //If;
//PreviewSwitch = //True;
//Call //LetterControl;
;
;
}
function PrintLetterButton_Click(){
//If //LetterName = "" //Then;
        //Else;
        //DoCmd.Close //acReport, //LetterName, //acSaveNo;
    //End //If;
    //If //Me.CheckResponseRequired //Then;
        //If //Not //CheckDates //Then;
            //Exit //Sub;
        //End //If;
    //End //If;
;
    //If //Not //gbAlreadySet //Then;
        //MsgBox "Please select To Whom to send the letter to";
        //Exit //Sub;
    //End //If;
;
    //If //Me.CheckResponseRequired = //True //Then;
        //If //IsNull(Me.RequiredByDate) //Or //IsNull(Me.FollowUpDate) //Then;
            //MsgBox "Required by and Followup dates cannot be blank";
            //Exit //Sub;
        //End //If;
    //End //If;
    //PreviewSwitch = //False;
    //Call //LetterControl;
;
}
function RequiredByDate_Exit(Cancel){
//If //IsNull(Me.RequiredByDate) //Or //Me.RequiredByDate = " " //Then;
    //Exit //Sub;
//Else;
    //Me.FollowUpDate = //DateAdd("d", //-15, //Me.RequiredByDate);
    //Me.FollowUpDate.SetFocus;
//End //If;
}
function ReturnBut_Click(){
//On //Error //GoTo //Err_ReturnBut_Click;
//'GoTo Err_ReturnBut_Click
//If //Not //CustomLetter //Then;
    //If //Reports(0).Name = //LetterName //Then;
        //MsgBox "You must print or close the Letter First";
        //Exit //Sub;
    //End //If;
//Else;
    //GoTo //Exit_ReturnBut_Click;
//End //If;
;
//Err_ReturnBut_Click:;
//If //Err.Number = //2457 //Then;
    //GoTo //Exit_ReturnBut_Click;
//Else;
   //MsgBox "PRO_IRB Error #070-Error in Closing Form or Report" //& //Err.Number;
   //Resume //Exit_ReturnBut_Click;
//End //If;
    ;
;
//Exit_ReturnBut_Click:;
//DoCmd.Close;
    ;
}
function ReviewerAdd_Click(){
//Call //OpenForms("ReviewerSelect");
//Me.ReviewerCombo.RowSourceType = "Table/Query";
;
}
function ReviewerCheck_Click(){
//If //IsNull(Me.ReviewerCombo.ItemData(0)) //Then;
        //DisplayMessage //("You //must //return //to //the //Study //and //assign //Reviewers");
        //Me.ReviewerCheck = //False;
        //Exit //Sub;
    //End //If;
  //If //Me.ReviewerCheck //Then;
    //If //CheckAlreadySet //Then;
        //Me.ReviewerCheck = //False;
        //Exit //Sub;
    //Else;
        //gbAlreadySet = //True;
    //End //If;
      ;
    //Me.ReviewerPostal.Enabled = //True;
    //Me.ReviewerEmail.Enabled = //True;
    //Me.ReviewerPostal = //True;
    //Me.ReviewerEmail = //False;
//Else;
    //gbAlreadySet = //False;
    //Me.ReviewerPostal.Enabled = //False;
    //Me.ReviewerEmail.Enabled = //False;
    //Me.ReviewerPostal = //False;
    //Me.ReviewerEmail = //False;
        ;
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
;
}
function ReviewerCombo_DblClick(Cancel){
var stDocName = '';
//On //Error //GoTo //Err_Reviewer;
    stDocName = "ReviewerSelect";
    //vFrom_LetterMAnager = //True;
    location.href = "Form_"+stDocName+".php";//, //, //, //, //, //acDialog;
    //vFrom_LetterMAnager = //False;
    //Me.ReviewerEmail = //False;
    //gbEmailSet = //False;
    //Me.ListCCs.Requery;
    //Me.ReviewerCombo.Requery;
;
//Exit_Reviewer:;
//Exit //Sub;
//Err_Reviewer:;
//MsgBox //Err.Description, //vbInformation;
//Resume //Exit_Reviewer;
}
function ReviewerEmail_Click(){
//If //Demo //Then;
    //MsgBox "Not available in Trial Version";
    //Me.ReviewerEmail = //False;
//Else;
    //If //Me.ReviewerCombo.ListIndex = //-1 //Then;
        //DisplayMessage //("You //must //select //a //Reviewer.");
        //Exit //Sub;
    //End //If;
    //If //Me.ReviewerEmail //Then;
                ;
        //If //IsNull(Me.ReviewerCombo.Column(15, //Me.ReviewerCombo.ListIndex)) //Or //Me.ReviewerCombo.Column(15, //Me.ReviewerCombo.ListIndex) = "" //Then;
            //MsgBox "Reviewer's Email Address is Blank---" //& //Me.ReviewerCombo.Column(1, //Me.ReviewerCombo.ListIndex);
            //Me.ReviewerEmail = //False;
        //Else;
            //Me.CheckCCEmail = //True;
        //End //If;
    //Else;
        //Me.CheckCCEmail = //False;
    //End //If;
//End //If;
//If //Me.ReviewerEmail //Then //gbEmailSet = //True //Else //gbEmailSet = //False;
;
}
function sponsoradd_Click(){
//Call //sSponsorAdd;
;
    ;
}
function SponsorCheck_Click(){
//If //Me.SponsorCheck //Then;
    ;
    //If //IsNull(Me.SponsorCombo) //Then;
        //DisplayMessage //("You //must //select //a //Sponsor");
        //Me.SponsorCheck = //False;
        //Exit //Sub;
    //End //If;
//End //If;
//If //Me.SponsorCheck //Then;
    //If //CheckAlreadySet //Then;
        //Me.SponsorCheck = //False;
        //Exit //Sub;
    //Else;
        //gbAlreadySet = //True;
    //End //If;
    //Me.Sponsoremail.Enabled = //True;
    //Me.ChSponsorPostal.Enabled = //True;
    //Me.ChSponsorPostal = //True;
    //Me.Sponsoremail = //False;
//Else;
    //gbAlreadySet = //False;
    //Me.Sponsoremail.Enabled = //False;
   //Me.ChSponsorPostal.Enabled = //False;
   //Me.Sponsoremail = //False;
   //Me.ChSponsorPostal = //False;
//End //If;
;
}
function SponsorCombo_DblClick(Cancel){
//If //EducationLetterSwitch //Then;
    //MsgBox " You must return to the Administration Menu to Edit or Review the Name and Address Info";
//Else;
;
    //Call //sViewSponsor;
    //Me.ListCCs.Requery;
    //Me.Sponsoremail = //False;
    //gbEmailSet = //False;
//End //If;
}
function Sponsoremail_Click(){
//If //Demo //Then;
    //MsgBox "Not available in Trial Version";
     //Me.Sponsoremail = //False;
//Else;
    //If //Me.Sponsoremail //Then;
        //If //CustomLetter //Then //Me.CheckCCEmail = //True;
        //If //IsNull(Me.SponsorCombo.Column(15)) //Or //Me.SponsorCombo.Column(15) = "" //Then;
            //MsgBox "Sponsor Email Address is Blank";
            //Me.Sponsoremail = //False;
        //Else;
            ;
            //Me.ChSponsorPostal = //True;
        //End //If;
    //Else;
    //Me.CheckCCEmail = //False;
;
    //End //If;
//End //If;
    //If //Me.Sponsoremail //Then //gbEmailSet = //True //Else //gbEmailSet = //False;
;
}
function DoItAll_Letter(r){
//Call //SetUpSendTo(r);
//UpdateSendTo //(LetterName);
//Call //PrintLetter;
}
function DoItAll_Email(r, Rec){
//Call //SetUpSendTo(r);
//UpdateSendTo //(LetterName);
//Call //SendEmail(r, //Rec);
    ;
}
function LetterControl(){
//If //GeneralLetterSwitch //Then;
    //DoCmd.SetWarnings //True;
    //LetterName = "General Letter";
    //Call //SendThem;
    //Exit //Sub;
   ;
//End //If;
;
;
//If //FromICChecklist = //True //Then;
    //DoCmd.SetWarnings //True;
    //LetterName = "Informed Consent Deficiencies";
    ;
    //Call //AddComments;
    //Call //SendThem;
    //Exit //Sub;
//Else;
    //DoCmd.SetWarnings //False;
    //DoCmd.OpenQuery "_deleteallexceptcomment";
    //DoCmd.SetWarnings //True;
    ;
//End //If;
;
//If //RenewalLetterSwitch //Then;
    //DoCmd.SetWarnings //False;
    //Me.LabelLetterName.Caption = "Continuing Review Notice";
    //Me.TextParagraphNumber = //10;
    //DoCmd.OpenQuery "_appendtoCurrentLetter";
    //DoCmd.SetWarnings //True;
    //LetterName = "Continuing Review Notice";
    //Call //AddComments;
    //Call //SendThem;
    //Exit //Sub;
//End //If;
;
;
//If //FromPostIRB = //True //Then;
      //DoCmd.SetWarnings //False;
    //Me.LabelLetterName.Caption = "IRB Action Letter";
    //Me.TextParagraphNumber = //10;
    //DoCmd.OpenQuery "_appendtoCurrentLetter";
    //DoCmd.SetWarnings //True;
    //LetterName = "IRB Action Letter";
    //Call //AddComments;
    //Call //SendThem;
    //Exit //Sub;
//End //If;
;
//'all other
        ;
        //Me.LabelLetterName.Caption = //Me.LetterLabel.Caption;
        //DoCmd.SetWarnings //False;
        //Me.TextParagraphNumber = //10;
        //DoCmd.OpenQuery "_appendtoCurrentLetter";
        //DoCmd.SetWarnings //True;
       ;
        //LetterName = //Me.LetterLabel.Caption;
        //Call //AddComments;
        //Call //SendThem;
//'End If
;
;
;
}
function AddressEnvelope(){//As //Boolean;
//AddressEnvelope = //True;
//DoCmd.SetWarnings //False;
//DoCmd.OpenQuery "_deleteCurrentLetterSendTo";
//Me.TextEnvCount = //1;
//'Sets up Send to whcih is redundant if a letter is previewed or printed at the same time
    //'however the user may come to the letter manager and only print an envelope
//If //Me.SponsorCheck //And //Me.ChSponsorPostal //Then;
    //SendToCntrl = "SponsorCombo";
    //Me.strSentTO = "Sponsor Post";
    //Call //SetUpSendTo(Me.SponsorCombo.ListIndex);
    //DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
    //Exit //Function;
//End //If;
;
//If //Me.PICheck //And //Me.pipostal //Then;
    //SendToCntrl = "PICombo";
    //Me.strSentTO = "PI Post";
    //Call //SetUpSendTo(Me.PICombo.ListIndex);
    //DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
    //Exit //Function;
//End //If;
//If //Me.CoordCheck //And //Me.CoordPostal //Then;
    //SendToCntrl = "CoordCombo";
    //Me.strSentTO = "Coord Post";
    //Call //SetUpSendTo(Me.CoordCombo.ListIndex);
    //DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
    //Exit //Function;
   ;
//End //If;
//If //Me.OtherCheck //And //Me.OtherPostal //Then;
    //SendToCntrl = "OtherCombo";
    //Me.strSentTO = "Other Post";
    //Call //SetUpSendTo(Me.OtherCombo.ListIndex);
    //DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
    //Exit //Function;
//End //If;
;
;
//If //Me.ReviewerCheck //And //Me.ReviewerPostal //Then;
    ;
    //SendToCntrl = "ReviewerCombo";
    //Me.strSentTO = "Reviewer Post";
    //Call //SetUpSendTo(Me.ReviewerCombo.ListIndex);
    //DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
    //Exit //Function;
//End //If;
;
//If //Me.IRBCheck //And //Me.IRBPostal //Then;
    ;
    //SendToCntrl = "irbCombo";
    //Me.strSentTO = "Member Post";
    //Call //SetUpSendTo(Me.IRBCombo.ListIndex);
    //DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
    //Exit //Function;
//End //If;
;
//MsgBox "Please select an addressee"//, //vbInformation;
//AddressEnvelope = //False;
//Exit //Function;
;
;
       ;
;
}
function PrintEnv(View){
//could initialize WasBlank here
//On //Error //GoTo //err_Env;
//LetterName = "Number10Envelope";
//'set letterlabel.caption to any letter to allow the appendcurrentsento query to
//' create sento rows for the envelope
//If //LetterName = " " //Then;
    WasBlank = True;
    //LetterName = "Invoice for IRB Services";
//End //If;
;
;
//If //View = "Preview" //Then;
    //DoCmd.OpenReport "Number10Envelope"//, //acViewPreview;
//Else;
    //DoCmd.OpenReport "Number10Envelope"//, //acViewNormal;
//End //If;
//If WasBlank //Then;
    //LetterName = " ";
    WasBlank = False;
//End //If;
;
;
//exit_Env:;
//Exit //Sub;
//err_Env:;
//MsgBox "Error Printing envelopes";
//Resume //exit_Env;
}
function AddCommentsButton_Click(){
//On //Error //GoTo //Err_AddCommentsButton_Click;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName = "CurrentLetterComment";
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
;
//Exit_AddCommentsButton_Click:;
    //Exit //Sub;
;
//Err_AddCommentsButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_AddCommentsButton_Click;
    ;
}
function CheckDates(){//As //Boolean;
    //If //IsNull(Me.RequiredByDate) //Or //IsNull(Me.FollowUpDate) //_;
        //Or //Me.RequiredByDate = " " //Or //Me.FollowUpDate = " " //Then;
        //MsgBox "Required by or Follow-up date cannot be blank";
        //CheckDates = //False;
        //Exit //Function;
    //Else;
        //CheckDates = //True;
    //End //If;
}
function ExportDataSource(Contact){//As //Boolean;
//On //Error //GoTo //err_Export;
    //could initialize i here
    //could initialize RevCount here
    //DoCmd.SetWarnings //False;
    //DoCmd.RunSQL //("delete //* //from //ExportedName");
     //Me.TextCCs = "";
    //'set up ccs
    //If //Me.ListCCs.ListCount = //0 //Then;
       ;
  //GoTo //Continue_Export:;
    //Else;
        //Me.TextCCs = "cc: ";
    //End //If;
    RevCount = Me.ListCCs.ListCount - 1;
;
    //For i = 0 To RevCount;
        //Me.TextCCs = //Me.TextCCs //& //Me.ListCCs.Column(16, //i);
        //If i //< RevCount //Then //Me.TextCCs = //Me.TextCCs //& ", ";
    //Next i;
    ;
    //Me.TextCCs = //LTrim$(Me.TextCCs);
//Continue_Export:;
    //Forms![newmenu]![ContactID] = //Contact;
    //Call //SetUpSendTo(0);
    ;
    //'Add in here an if statement to test for general letter and bypass the updatesendto procedure
    //If //(Not //GeneralLetterSwitch) //Then;
        //UpdateSendTo //(LetterName);
    //End //If;
   //'If Not AddContactNAme() Then GoTo err_Export
    //DoCmd.SetWarnings //False;
    ;
    ;
    ;
  //' MsgBox [Forms]![NewMenu]![CONTactID]
   //' MsgBox [Forms]![_LetterManager]![LetterLabel].[Caption]
    //'MsgBox Me.IRBText
   //' [Forms]![_LetterManager]![LetterLabel].[Caption]
    //DoCmd.OpenQuery "AppendContactName";
    //'MsgBox [Forms]![_LetterManager]![LetterLabel].[Caption]
    ;
    //Call //UpdateCCs;
    //If //EducationLetterSwitch //Then;
        //DoCmd.TransferText //acExportDelim, //, "ExportCombinedEducationData"//, "C:\CC\PROIRB_Dev\Data Sources\Education.txt"//, //True;
    //Else;
        //If //GeneralLetterSwitch //Then;
            //DoCmd.RunSQL "UPDATE ExportedName SET ExportedName.StudyNo = 'Dummy' WITH OWNERACCESS OPTION;";
            //DoCmd.TransferText //acExportDelim, //, "ExportCombinedForGeneralLetter"//, //DataSourcePath, //True;
            //DoCmd.SetWarnings //True;
            //GoTo //Exit_Export;
        //End //If;
    //End //If;
    //DoCmd.TransferText //acExportDelim, //, "Export_CombinedData"//, //DataSourcePath, //True;
    //DoCmd.SetWarnings //True;
       ;
//Exit_Export:;
//Exit //Function;
    ;
//err_Export:;
//MsgBox //Err.Description //& "  Exporting Data Source";
//Resume //Exit_Export;
}
function sSponsorAdd(){
//On //Error //GoTo //Err_SponsorAdd;
var stDocName = '';
    var stLinkCriteria = '';
    //HoldSponsor = //Me.[Sponsor //ID];
    stDocName = "_Contact Sponsor";
    //HoldSponsorName = " ";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.PICombo.SetFocus;
    //If //IsNull(HoldSponsor) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Sponsor //ID] = //HoldSponsor;
        //Me.SponsorCombo.RowSourceType = "Table/Query";
    //End //If;
;
;
;
//Exit_sponsoradd:;
    //Exit //Sub;
;
//Err_SponsorAdd:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_sponsoradd;
}
function sInvestigatorAdd(){
//On //Error //GoTo //err_sInvestigatorAdd;
  var stDocName = '';
    var stLinkCriteria = '';
    //HoldPI = //Me.[Investigator //ID];
    stDocName = "_Contact PI";
    //HoldPIName = " ";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.CoordCombo.SetFocus;
    //If //IsNull(HoldPI) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Investigator //ID] = //HoldPI;
        //Me.PICombo.RowSourceType = "Table/Query";
    //End //If;
;
  ;
   ;
//Exit_sInvestigatorAdd:;
    //Exit //Sub;
;
//err_sInvestigatorAdd:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_sInvestigatorAdd;
    ;
;
}
function sCoordinatorAdd(){
//On //Error //GoTo //Err_sCoordinatorAdd;
;
    var stDocName = '';
    var stLinkCriteria = '';
    //HoldCoord = //Me.[Coordinator //ID];
    stDocName = "_contact coordinator";
    //HoldCoordName = " ";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.LetterDate.SetFocus;
     //If //IsNull(HoldCoord) //Then;
        //Exit //Sub;
    //Else;
        //Me.[Coordinator //ID] = //HoldCoord;
        //Me.CoordCombo.RowSourceType = "Table/Query";
    //End //If;
;
    ;
//Exit_sCoordinatorAdd:;
    //Exit //Sub;
;
//Err_sCoordinatorAdd:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_sCoordinatorAdd;
    ;
;
}
function sIRBAdd(){
var stDocName = '';
var stLinkCriteria = '';
//On //Error //GoTo //Err_sIRBAdd;
stDocName = "_contactboardmember";
location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.PrintLetterButton.SetFocus;
    ;
    //Me.IRBCombo.RowSourceType = "Table/Query";
    ;
    ;
//Exit_sIRBAdd:;
//Exit //Sub;
;
//Err_sIRBAdd:;
//MsgBox //Err.desc //& " " //& //Err.Number;
//Resume //Exit_sIRBAdd;
;
}
function sOtherAdd(){
var stDocName = '';
    var stLinkCriteria = '';
    ;
    //On //Error //GoTo //Err_sOtherAdd;
    //'stLinkCriteria = "[common contact id]" & " = " & Me.OtherCombo.Column(0)
    stDocName = "_contactother";
    //'HoldPI = Me.[Investigator ID]
    ;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.PrintLetterButton.SetFocus;
    //'Me.[Investigator ID] = HoldPI
    //Me.OtherCombo.RowSourceType = "Table/Query";
    ;
//Exit_sOtherAdd:;
//Exit //Sub;
;
//Err_sOtherAdd:;
//MsgBox //Err.Description, //vbInformation;
//Resume //Exit_sOtherAdd;
}
function sViewSponsor(){
//On //Error //GoTo //Err_sViewSponsor;
 var stDocName = '';
    var stLinkCriteria = '';
    //If //IsNull(Me.SponsorCombo.Column(0, //Me.SponsorCombo.ListIndex)) //Then;
        //MsgBox //("No //Sponsor //to //View");
        //Exit //Sub;
    //End //If;
   ;
    stLinkCriteria = "[common contact id]" //& " = " //& //Me.SponsorCombo.Column(0, //Me.SponsorCombo.ListIndex);
    //HoldSponsor = //Me.SponsorCombo.Column(0, //Me.SponsorCombo.ListIndex);
    stDocName = "_Contact Sponsor";
   ;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    ;
    //Me.PrintLetterButton.SetFocus;
    //'Me![SponsorCombo] = HoldSponsor
    //Me.SponsorCombo.RowSourceType = "Table/Query";
    ;
    ;
//Exit_sViewSponsor:;
//Exit //Sub;
//Err_sViewSponsor:;
//MsgBox //Err.Description, //vbInformation;
//Resume //Exit_sViewSponsor;
}
function sViewInvestigator(){
//On //Error //GoTo //Err_sViewInvestigator;
var stDocName = '';
    var stLinkCriteria = '';
    //If //IsNull(Me.PICombo.Column(0, //Me.PICombo.ListIndex)) //Then;
        //MsgBox //("No //P.I. //to //View");
        //Exit //Sub;
    //End //If;
    stLinkCriteria = "[common contact id]" //& " = " //& //Me.PICombo.Column(0, //Me.PICombo.ListIndex);
    stDocName = "_Contact PI";
    //HoldPI = //Me.PICombo.Column(0, //Me.PICombo.ListIndex);
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    //Me.PrintLetterButton.SetFocus;
    //'Me.[Investigator ID] = HoldPI
    //Me.PICombo.RowSourceType = "Table/Query";
    ;
//Exit_sViewInvestigator:;
//Exit //Sub;
//Err_sViewInvestigator:;
//MsgBox //Err.Description, //vbInformation;
//Resume //Exit_sViewInvestigator;
}
function sViewCoordinator(){
//On //Error //GoTo //Err_sViewCoordinator;
//'If IsNull(Me.[Coordinator ID]) Then
//If //IsNull(Me.CoordCombo.Column(0, //Me.CoordCombo.ListIndex)) //Then;
        //MsgBox //("No //Coordinator //to //View");
        //Exit //Sub;
//End //If;
;
 var stDocName = '';
    var stLinkCriteria = '';
    //'stLinkCriteria = "[common contact id]" & " = " & Me![Coordinator ID]
    stLinkCriteria = "[common contact id]" //& " = " //& //Me.CoordCombo.Column(0, //Me.CoordCombo.ListIndex);
    stDocName = "_Contact coordinator";
    //'HoldCoord = Me.[Coordinator ID]
    //HoldCoord = //Me.CoordCombo.Column(0, //Me.CoordCombo.ListIndex);
    ;
    //HoldCoordName = " ";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    ;
    //Me.PrintLetterButton.SetFocus;
     //'Me.CoordCombo = HoldCoord
      ;
     ;
    //Me.CoordCombo.RowSourceType = "Table/Query";
    ;
//Exit_sViewCoordinator:;
//Exit //Sub;
//Err_sViewCoordinator:;
//MsgBox //Err.Description, //vbInformation;
//Resume //Exit_sViewCoordinator;
}
function sViewIRB(){
//On //Error //GoTo //Err_sViewIRB;
var stDocName = '';
    var stLinkCriteria = '';
    stLinkCriteria = "[common contact id]" //& " = " //& //Me.IRBCombo.Column(0);
    stDocName = "_contactBoardMember";
    //'HoldPI = Me.[Investigator ID]
    //If //IsNull(Me.IRBCombo.Column(0)) //Then;
        //MsgBox "You must first Select an IRB Member";
        //Exit //Sub;
    //End //If;
            ;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    //Me.PrintLetterButton.SetFocus;
    //'Me.[Investigator ID] = HoldPI
    //Me.IRBCombo.RowSourceType = "Table/Query";
    ;
//Exit_sViewIRB:;
//Exit //Sub;
;
//Err_sViewIRB:;
//MsgBox //Err.Description, //vbInformation;
//Resume //Exit_sViewIRB;
;
}
function sViewOther(){
//On //Error //GoTo //Err_sViewOther;
var stDocName = '';
    var stLinkCriteria = '';
    stLinkCriteria = "[common contact id]" //& " = " //& //Me.OtherCombo.Column(0);
    stDocName = "_contactother";
    //'HoldPI = Me.[Investigator ID]
    //If //IsNull(Me.OtherCombo.Column(0)) //Then;
        //MsgBox "You must first Select an 'Other' Contact";
        //Exit //Sub;
    //End //If;
        ;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    //Me.PrintLetterButton.SetFocus;
    //'Me.[Investigator ID] = HoldPI
    //Me.OtherCombo.RowSourceType = "Table/Query";
    ;
//Exit_sViewOther:;
//Exit //Sub;
//Err_sViewOther:;
//MsgBox //Err.Description, //vbInformation;
//Resume //Exit_sViewOther;
;
;
}
function DeleteCC(){
//DoCmd.SetWarnings //False;
var SQLKEY = '';
//DoCmd.RunSQL //("Delete //* //from //CurrentLetter_CCs //where //contactid = " & Me.HoldCCKey)
DoCmd.SetWarnings True
Me.ListCCs.Requery
End Sub

Public Sub AppendToCC()
DoCmd.SetWarnings False

Dim SQLKEY As String
SQLKEY = Me.HoldCCKey
DoCmd.OpenQuery ("//AppendContacttoCC");
;
//Me.ListCCs.Requery;
//DoCmd.SetWarnings //True;
}
function cmdAddDeleteCC_Click(){
//On //Error //GoTo //Err_cmdAddDeleteCC_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName = "CCDeleteAdd";
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //, //acDialog;
    //Me.ListCCs.Requery;
//Exit_cmdAddDeleteCC_Click:;
    //Exit //Sub;
;
//Err_cmdAddDeleteCC_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdAddDeleteCC_Click;
    ;
}
function SetLabelCaptions(DSource){
//Me.DisplayLetterName.Caption = //HoldDocNAme;
//Me.LetterLabel.Caption = //LetterName;
//'MsgBox LetterName
//DataSourcePath = "C:\Cc\ProIRB_Dev\Data Sources\" //& //DSource;
//Me.LabelLetterName.Caption = //LetterName;
}
function CheckAlreadySet(){//As //Boolean;
//If //gbAlreadySet //Then;
        //MsgBox "You can only choose one type of Addressee";
        //CheckAlreadySet = //True;
        //Exit //Function;
//Else;
                //CheckAlreadySet = //False;
//End //If;
;
}
function CustomEmails(Recipient, RecContactID){
//If //Not //Me.TextMeetingDate = " " //Then;
    ;
//'If PreviewSwitch Then    displaymessage ("You cannot Preview an EmailExit Function
//gsMailMessage = //vbCrLf //& //vbCrLf //& //Forms![_LetterManager]![LetterLabel].Caption //& " From: " //& //Forms![newmenu]![Location] //& //vbCrLf //& //vbCrLf //& //_;
"Study Number: " //& //Forms![_LetterManager].IRBText //& //vbCrLf //& //vbCrLf //& "Protocol Title: " //& //Forms![_LetterManager].TxtTitle;
//Else;
     //gsMailMessage = //vbCrLf //& //vbCrLf //& //Forms![_LetterManager]![LetterLabel].Caption //& " From: " //& //Forms![newmenu].Location;
    ;
//End //If;
//gsfiletosend = "C:\CC\Proirb_dev\" //& //Forms![_LetterManager]![LetterLabel].Caption //& ".doc";
//On //Error //GoTo //err_CustomEmails;
;
//If //gbLetterFileCreated //Then;
    //GoTo //DoMail;
//Else;
    //Call //ExportDataSource(RecContactID);
    //gbForEmails = //True;
    //If //Not //MSWordRoutine(TempTemplatePath, "Letter Created for Email.  Click OK "//) //Then;
        //gbLetterFileCreated = //False;
        //GoTo //err_CustomEmails;
    //Else;
        //gbLetterFileCreated = //True;
    //End //If;
//End //If;
;
;
//DoMail:;
//Call //SetUpCCs;
//SaveAsPath = "C:\CC\Proirb_dev\" //& //Forms![_LetterManager]![LetterLabel].Caption //& ".doc";
//'If MSMailRoutine(Recipient, gsEmailCCs, "IRB Notice--" & Me![LetterLabel].Caption, SaveAsPath, gsMailMessage) Then
    //'Exit Function
//'End If
;
;
;
//'Setup for 2007  Custom emails
;
//If //dir("C:\CC\PROIRB_Dev\Outlook.ini") = "Outlook.ini" //Then;
      ;
    //If //Outlook2007Email(Recipient, //gsEmailCCs, "IRB Notice--" //& //Me![LetterLabel].Caption, //SaveAsPath, //gsMailMessage) //Then;
        //Exit //Function;
    //End //If;
//Else;
;
    //If //MSMailRoutine(Recipient, //gsEmailCCs, "IRB Notice--" //& //Me![LetterLabel].Caption, //SaveAsPath, //gsMailMessage) //Then;
        //Exit //Function;
    //End //If;
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
;
;
;
;
//Exit_CustomEmails:;
//Exit //Function;
;
//err_CustomEmails:;
//MsgBox //Err.Description;
//Resume //Exit_CustomEmails;
;
}
function CheckEmailAddresses(){
//could initialize i here
        //For i = 0 To Me.ListCCs.ListCount - 1;
        //If //IsNull(Me.ListCCs.Column(15, //i)) //Or //Me.ListCCs.Column(15, //i) = "" //Then;
            //MsgBox "Email Address is Blank---" //& //Me.ListCCs.Column(16, //i);
           ;
        //End //If;
        //Next i;
}
function SetUpCCs(){
//could initialize i here
//gsEmailCCs = " ";
//'" emails to CC's
//If //Me.CheckCCEmail //Then;
    //For i = 0 To Me.ListCCs.ListCount - 1;
    //If //Me.ListCCs.Column(15, //i) //<> "" //Then;
        //gsCCAddress = //Me.ListCCs.Column(15, //i);
        //gsEmailCCs = //LTrim$(gsEmailCCs //& //gsCCAddress);
        //If i //< //Me.ListCCs.ListCount //- //1 //Then //gsEmailCCs = //gsEmailCCs //& ";";
           ;
    //Else;
        //MsgBox "Email Address Invalid. Email not sent to--" //& //Me.ListCCs.Column(16, //i);
  ;
    //End //If;
    //Next i;
//End //If;
;
}
function PrintCustomLetter(){
//On //Error //GoTo //Err_PrintCustom;
//If //PreviewSwitch //Then;
;
    //If //Not //MSWordPreview(Me.DisplayLetterName.Caption) //Then //GoTo //Err_PrintCustom;
//Else;
;
    //If //Not //MSWordRoutine(TempTemplatePath, "Letter Printing-Pause then click OK"//) //Then;
        //GoTo //Err_PrintCustom;
    //Else;
        ;
        //gbLetterFileCreated = //True;
    //End //If;
//End //If;
;
//Exit_PrintCustom:;
//Exit //Sub;
;
//Err_PrintCustom:;
//MsgBox //Err.Description;
//Resume //Exit_PrintCustom;
;
;
}
function UpdateCCs(){//As //Boolean;
;
    //On //Error //GoTo //err_UpdateCCs;
   ;
    ;
    //If //Me.TextCCs = "cc:" //Then //Exit //Function;
                    ;
    //could initialize db here
    //could initialize rs here
    //could initialize rs2 here
   ;
    //Set db = CurrentDb;
    //Set rs = db.OpenRecordset("ExportedName");
;
    //rs.Edit;
    ;
    ;
    //rs("TextCCs") = //Me.TextCCs;
    //rs.Update;
   //rs.Close;
   //UpdateCCs = //True;
   //db.Close;
;
//exit_UpdateCCs:;
   //Exit //Function;
//err_UpdateCCs:;
    //MsgBox "Error Adding Contact Record--" //& //Err.Description;
    //UpdateCCs = //False;
    //Resume //exit_UpdateCCs;
}
function IRBEmailCheck(){
//If //Me.IRBCombo.ListIndex = //-1 //Then;
    //DisplayMessage //("You //must //select //an //IRB //Member.");
    //Exit //Sub;
//End //If;
;
;
;
//If //IsNull(Me.IRBCombo.Column(15, //Me.IRBCombo.ListIndex)) //Or //Me.IRBCombo.Column(15, //Me.IRBCombo.ListIndex) = "" //Then;
            //MsgBox "Member's Email Address is Blank---" //& //Me.IRBCombo.Column(1, //Me.IRBCombo.ListIndex);
            //Me.IRBEmail = //False;
//Else;
        //Me.CheckCCEmail = //True;
//End //If;
   ;
//If //Me.IRBEmail //Then //gbEmailSet = //True //Else //gbEmailSet = //False;
;
}
function CCEnvelopes(){
//could initialize i here
//DOCCs:;
//If //Me.ListCCs.ListCount //> //0 //And //Not //Me.CheckCCEmail //Then;
    //For i = 0 To Me.ListCCs.ListCount - 1;
        ;
    //SendToCntrl = "Listccs";
    //Me.strSentTO = "CC Post";
    //Call //SetUpSendTo(i);
    //Me.TextEnvCount = //Me.TextEnvCount //+ //1;
   //DoCmd.OpenQuery "_AppendCurrentLetterSendTo";
   //MsgBox "Load Next Envelope";
   //Call //PrintEnv("Normal");
    //Next i;
    ;
//End //If;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>



    <label id=FormTitle value='Letter Manager' style='position:absolute; left:153; top:6; width:273; height:48; visibility:'>Letter Manager</label>
    <label id=Label3 value='Letter' style='position:absolute; left:49; top:54; width:67; height:28; visibility:'>Letter</label>
    <input type='checkbox' id=SponsorCheck value='False' style='position:absolute; left:133; top:186; width:24; height:0'>

    <label id=Label19 value='Sponsor' style='position:absolute; left:1; top:180; width:127; height:24; visibility:'>Sponsor</label>
    <input type='checkbox' id=PICheck value='True' style='position:absolute; left:133; top:219; width:18; height:0'>

    <label id=Label22 value='PI' style='position:absolute; left:1; top:216; width:127; height:24; visibility:'>PI</label>
    <input type='checkbox' id=CoordCheck value='False' style='position:absolute; left:133; top:252; width:24; height:0'>

    <label id=Label24 value='Coordinator' style='position:absolute; left:1; top:246; width:127; height:24; visibility:'>Coordinator</label>
    <label id=Label25 value='Addressee' style='position:absolute; left:255; top:148; width:123; height:28; visibility:'>Addressee</label>
    <input type='checkbox' id=Sponsoremail value='=False' style='position:absolute; left:474; top:187; width:24; height:19'>
    <input type='checkbox' id=PIemail value='=False' style='position:absolute; left:474; top:219; width:24; height:19'>
    <input type='checkbox' id=Coordemail value='=False' style='position:absolute; left:474; top:252; width:24; height:19'>
    <input type='checkbox' id=ChSponsorPostal value='=False' style='position:absolute; left:504; top:187; width:27; height:21'>
    <input type='checkbox' id=pipostal value='=False' style='position:absolute; left:504; top:219; width:27; height:21'>
    <input type='checkbox' id=CoordPostal value='=False' style='position:absolute; left:504; top:252; width:27; height:21'>
    <input type='text' id=FollowUpDate value='empty' style='position:absolute; left:694; top:102; width:114; height:0'>

    <label id=Label88 value='Date to Follow Up ' style='position:absolute; left:523; top:102; width:168; height:24; visibility:'>Date to Follow Up </label>
    <input type='checkbox' id=CheckResponseRequired value='False' style='position:absolute; left:225; top:105; width:18; height:18'>

    <label id=Label90 value='Response Required' style='position:absolute; left:37; top:102; width:175; height:24; visibility:'>Response Required</label>
    <input type='text' id=RequiredByDate value='empty' style='position:absolute; left:403; top:102; width:114; height:25'>

    <label id=Label92 value='Date Required ' style='position:absolute; left:258; top:102; width:138; height:24; visibility:'>Date Required </label>
    <input type='text' id=LetterDate value='empty' style='position:absolute; left:979; top:130; width:0; height:25'>

    <label id=Label94 value='Letter Date' style='position:absolute; left:856; top:130; width:123; height:28; visibility:'>Letter Date</label>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:447; top:7; width:681; height:30'>
    <label id=MeetingLabel value='Meeting Date' style='position:absolute; left:832; top:88; width:147; height:28; visibility:'>Meeting Date</label>
    <!--     <input type='text' id=TextMeetingDate value='empty' style='position:absolute; left:979; top:88; width:0; height:30'> -->
    <input type='text' id=IRBText value='empty' style='position:absolute; left:981; top:51; width:0; height:30'>

    <label id=Label52 value='Study#' style='position:absolute; left:889; top:51; width:79; height:28; visibility:'>Study#</label>
    <input type='button' id=sponsoradd value='Add' style='position:absolute; left:555; top:180; width:58; height:27' onclick='sponsoradd_Click();'>
    <input type='button' id=addPI value='Add' style='position:absolute; left:555; top:214; width:57; height:27' onclick='addPI_Click();'>
    <input type='button' id=coordadd value='Add' style='position:absolute; left:556; top:247; width:57; height:27'>
    <input type='button' id=AddCommentsButton value='Add  Comments to Letter  ' style='position:absolute; left:729; top:459; width:373; height:43' onclick='AddCommentsButton_Click();'>
    <label id=Label131 value='Closing and Signatory' style='position:absolute; left:384; top:504; width:231; height:28; visibility:'>Closing and Signatory</label>
    <input type='button' id=ChangeButton value='Change Closing' style='position:absolute; left:933; top:537; width:165; height:31' onclick='ChangeButton_Click();'>
    <input type='button' id=Command135 value='Reset to Default' style='position:absolute; left:793; top:537; width:139; height:31' onclick='Command135_Click();'>
    <select name='ThisLetterClosing' Size='24' style='position:absolute; left:142; top:538; width:645; height:27'>

    <label id=Label133 value='For this Letter' style='position:absolute; left:10; top:538; width:127; height:24; visibility:'>For this Letter</label>
    <select name='DefaultClosing' Size='23' style='position:absolute; left:142; top:579; width:645; height:27'>

    <label id=Label130 value='Default' style='position:absolute; left:67; top:573; width:70; height:24; visibility:'>Default</label>
    <label id=DisplayLetterName value='Sample Second Reminder' style='position:absolute; left:123; top:51; width:558; height:39; visibility:'>Sample Second Reminder</label>
    <input type='checkbox' id=ReviewerCheck value='False' style='position:absolute; left:133; top:288; width:24; height:0'>

    <label id=Label142 value='Reviewers' style='position:absolute; left:1; top:282; width:127; height:24; visibility:'>Reviewers</label>
    <input type='checkbox' id=IRBCheck value='False' style='position:absolute; left:130; top:343; width:24; height:22'>

    <label id=Label144 value='IRB \015\012Members' style='position:absolute; left:0; top:336; width:129; height:40; visibility:'>IRB \015\012Members</label>
    <input type='checkbox' id=OtherCheck value='False' style='position:absolute; left:133; top:412; width:24; height:22'>

    <label id=Other Contact value='Other Contact' style='position:absolute; left:1; top:409; width:127; height:22; visibility:'>Other Contact</label>
    <input type='checkbox' id=ReviewerEmail value='=False' style='position:absolute; left:474; top:291; width:24; height:19'>
    <input type='checkbox' id=ReviewerPostal value='=False' style='position:absolute; left:504; top:291; width:27; height:21'>
    <input type='button' id=ReviewerAdd value='Add' style='position:absolute; left:556; top:286; width:57; height:27' onclick='ReviewerAdd_Click();'>
    <input type='checkbox' id=IRBEmail value='=False' style='position:absolute; left:474; top:345; width:24; height:18'>
    <input type='checkbox' id=IRBPostal value='=False' style='position:absolute; left:504; top:343; width:27; height:19'>
    <input type='button' id=IRBAdd value='Add' style='position:absolute; left:556; top:340; width:57; height:27' onclick='IRBAdd_Click();'>
    <input type='checkbox' id=OtherEmail value='=False' style='position:absolute; left:474; top:409; width:24; height:18'>
    <input type='checkbox' id=OtherPostal value='=False' style='position:absolute; left:504; top:409; width:27; height:19'>
    <input type='button' id=OtherAdd value='Add' style='position:absolute; left:556; top:406; width:57; height:27' onclick='OtherAdd_Click();'>

    <select name='ReviewerCombo' Size='32' style='position:absolute; left:163; top:288; width:294; height:43'>
    <select name='IRBCombo' Size='36' style='position:absolute; left:163; top:342; width:294; height:57'>
    <input type='checkbox' id=CCSponsor value='=False' style='position:absolute; left:760; top:198; width:22; height:0'>

    <label id=Label180 value='Sponsor' style='position:absolute; left:678; top:192; width:78; height:24; visibility:'>Sponsor</label>
    <input type='checkbox' id=CCPI value='=False' style='position:absolute; left:760; top:231; width:16; height:0'>

    <label id=Label182 value='PI' style='position:absolute; left:729; top:228; width:27; height:24; visibility:'>PI</label>
    <input type='checkbox' id=CCCoord value='False' style='position:absolute; left:760; top:264; width:22; height:0'>

    <label id=Label184 value='Coordinator' style='position:absolute; left:649; top:258; width:106; height:24; visibility:'>Coordinator</label>
    <input type='checkbox' id=CCReviewers value='False' style='position:absolute; left:760; top:292; width:22; height:0'>

    <label id=Label186 value='All Reviewers' style='position:absolute; left:631; top:286; width:124; height:24; visibility:'>All Reviewers</label>
    <input type='checkbox' id=CCIRBMembers value='False' style='position:absolute; left:760; top:319; width:24; height:22'>

    <label id=Label188 value='All IRB Members' style='position:absolute; left:609; top:312; width:148; height:22; visibility:'>All IRB Members</label>
    <select name='ListCCs' Size='48' style='position:absolute; left:786; top:198; width:309; height:178'>
    <label id=Label195 value='CC:' style='position:absolute; left:798; top:169; width:123; height:28; visibility:'>CC:</label>
    <input type='button' id=cmdAddDeleteCC value='Add or Delete CCs' style='position:absolute; left:793; top:378; width:301; height:40' onclick='cmdAddDeleteCC_Click();'>
    <input type='button' id=PrintLetterButton value='&Send/Print' style='position:absolute; left:540; top:613; width:220; height:40' onclick='PrintLetterButton_Click();'>
    <input type='button' id=Envelope value='Print &Envelope(s)' style='position:absolute; left:760; top:613; width:220; height:40' onclick='Envelope_Click();'>
    <input type='button' id=ReturnBut value='&Return' style='position:absolute; left:981; top:613; width:127; height:40' onclick='ReturnBut_Click();'>
    <input type='button' id=PreviewLetter value='&Preview' style='position:absolute; left:345; top:613; width:195; height:40' onclick='PreviewLetter_Click();'>
    <input type='button' id=cModify value='&Modify (This Letter Only)' style='position:absolute; left:42; top:613; width:301; height:40' onclick='cModify_Click();'>
    <label id=Label28 value='  E-\015\012MAIL' style='position:absolute; left:447; top:133; width:49; height:43; visibility:'>  E-\015\012MAIL</label>
    <label id=Label35 value='Mail' style='position:absolute; left:505; top:151; width:42; height:25; visibility:'>Mail</label>
    <select name='SponsorCombo' Size='4' style='position:absolute; left:163; top:181; width:294; height:0'>
    <select name='PICombo' Size='5' style='position:absolute; left:163; top:216; width:294; height:0'>
    <select name='CoordCombo' Size='6' style='position:absolute; left:163; top:249; width:294; height:0'>
    <select name='OtherCombo' Size='40' style='position:absolute; left:163; top:408; width:294; height:0'>
    <input type='checkbox' id=CheckCCEmail value='=False' style='position:absolute; left:940; top:174; width:24; height:19'>
    <label id=Label200 value='Email CCs' style='position:absolute; left:972; top:174; width:99; height:21; visibility:'>Email CCs</label>
    <!--     <input type='text' id=TextEnvCount value='empty' style='position:absolute; left:214; top:15; width:67; height:0'> -->

    <label id=Label202 value='Text201:' style='position:absolute; left:70; top:15; width:70; height:24; visibility:'>Text201:</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:49; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <select name='Coinvcombo' Size='57' style='position:absolute; left:163; top:444; width:294; height:57' onclick='Coinvcombo_Click();'>
    <input type='checkbox' id=CCCOINV value='False' style='position:absolute; left:760; top:348; width:24; height:22'>

    <label id=Label211 value='Co-Investigators' style='position:absolute; left:610; top:342; width:147; height:24; visibility:'>Co-Investigators</label>
    <label id=Label212 value='Co-Investigators' style='position:absolute; left:1; top:453; width:154; height:22; visibility:'>Co-Investigators</label>
    <!--     <input type='text' id=SponsorText value='empty' style='position:absolute; left:660; top:150; width:210; height:18'> -->

    <label id=Label214 value='Text213:' style='position:absolute; left:516; top:150; width:70; height:24; visibility:'>Text213:</label>
    <!--     <input type='text' id=CoordinatorText value='empty' style='position:absolute; left:666; top:228; width:18; height:0'> -->

    <label id=Label216 value='Text215:' style='position:absolute; left:732; top:48; width:70; height:24; visibility:'>Text215:</label>


    <!--     <input type='text' id=TextLine1 value='empty' style='position:absolute; left:516; top:66; width:348; height:30'> -->
    <!--     <input type='text' id=TextLine2 value='empty' style='position:absolute; left:90; top:72; width:360; height:30'> -->
    <!--     <input type='text' id=TextLine3 value='empty' style='position:absolute; left:90; top:114; width:360; height:30'> -->
    <!--     <input type='text' id=TextLine4 value='empty' style='position:absolute; left:90; top:156; width:360; height:30'> -->
    <!--     <input type='text' id=TextLine5 value='empty' style='position:absolute; left:90; top:198; width:360; height:30'> -->
    <!--     <input type='text' id=TextLine6 value='empty' style='position:absolute; left:90; top:124; width:360; height:30'> -->
    <!--     <label id=Label75 value='Send to Lines Holding Text for append query' style='position:absolute; left:706; top:4; width:276; height:40; visibility:hidden'>Send to Lines Holding Text for append query</label> -->
    <!--     <input type='text' id=TextLine7 value='empty' style='position:absolute; left:528; top:118; width:348; height:30'> -->
    <!--     <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:558; top:144; width:0; height:0'> -->

    <!--     <label id=Label77 value='Protocol#/Title:' style='position:absolute; left:414; top:144; width:118; height:24; visibility:hidden'>Protocol#/Title:</label> -->
    <!--     <input type='text' id=Last Renewal Date value='empty' style='position:absolute; left:624; top:198; width:0; height:0'> -->

    <!--     <label id=Label78 value='Last IRB Study Renewal Date:' style='position:absolute; left:480; top:198; width:225; height:24; visibility:hidden'>Last IRB Study Renewal Date:</label> -->
    <!--     <input type='text' id=TextParagraphNumber value='empty' style='position:absolute; left:96; top:6; width:360; height:48'> -->

    <!--     <label id=LTextParagraph value='Text79:' style='position:absolute; left:0; top:6; width:61; height:24; visibility:hidden'>Text79:</label> -->
    <!--     <input type='text' id=Text85 value='empty' style='position:absolute; left:582; top:21; width:42; height:30'> -->

    <!--     <label id=Label86 value='Text85:' style='position:absolute; left:438; top:21; width:61; height:24; visibility:hidden'>Text85:</label> -->
    <!--     <label id=LabelLetterName value='Education Letter' style='position:absolute; left:672; top:168; width:276; height:36; visibility:hidden'>Education Letter</label> -->
    <!--     <input type='text' id=strSentTO value='empty' style='position:absolute; left:846; top:60; width:156; height:18'> -->

    <!--     <label id=Label110 value='strSentTO' style='position:absolute; left:738; top:60; width:79; height:24; visibility:hidden'>strSentTO</label> -->
    <!--     <input type='text' id=ExpirationDate value='empty' style='position:absolute; left:624; top:12; width:0; height:30'> -->

    <!--     <label id=Label112 value='Expiration\015\012       Date' style='position:absolute; left:486; top:6; width:132; height:48; visibility:hidden'>Expiration\015\012       Date</label> -->
    <!--     <input type='text' id=InternalNumber value='empty' style='position:absolute; left:486; top:39; width:102; height:30'> -->
    <!--     <label id=LetterLabel value='Education Letter' style='position:absolute; left:0; top:39; width:558; height:39; visibility:hidden'>Education Letter</label> -->
    <!--     <select name='TestLetterCombo' Size='14' style='position:absolute; left:541; top:45; width:162; height:30'> -->

    <!--     <label id=Label118 value='Combo117:' style='position:absolute; left:571; top:21; width:88; height:24; visibility:hidden'>Combo117:</label> -->
    <!--     <input type='text' id=TextTopMargin value='empty' style='position:absolute; left:316; top:55; width:138; height:25'> -->

    <!--     <label id=Label196 value='TopMargin' style='position:absolute; left:172; top:55; width:84; height:24; visibility:hidden'>TopMargin</label> -->
    <!--     <input type='text' id=Text136 value='empty' style='position:absolute; left:529; top:48; width:90; height:0'> -->

    <!--     <label id=Label137 value='Text136:' style='position:absolute; left:385; top:48; width:70; height:24; visibility:hidden'>Text136:</label> -->
    <!--     <select name='LetterCombo' Size='17' style='position:absolute; left:0; top:21; width:372; height:24'> -->
    <!--     <input type='text' id=HoldCCKey value='empty' style='position:absolute; left:852; top:36; width:177; height:27'> -->

    <!--     <label id=Label194 value='Text193:' style='position:absolute; left:708; top:36; width:70; height:24; visibility:hidden'>Text193:</label> -->
    <!--     <input type='text' id=TextCCs value='empty' style='position:absolute; left:472; top:99; width:246; height:40'> -->

    <!--     <label id=Label198 value='Text197:' style='position:absolute; left:328; top:99; width:70; height:24; visibility:'>Text197:</label> -->
    <!--     <input type='text' id=TContuingReviewHold value='empty' style='position:absolute; left:912; top:193; width:93; height:16'> -->

    <!--     <label id=Label206 value='TextContactModuleID' style='position:absolute; left:820; top:180; width:87; height:24; visibility:'>TextContactModuleID</label> -->
    <!--     <input type='text' id=txtContactModuleID value='empty' style='position:absolute; left:930; top:96; width:52; height:39'> -->

    <!--     <label id=Label208 value='Text207:' style='position:absolute; left:786; top:96; width:70; height:24; visibility:'>Text207:</label> -->
    <!--     <input type='text' id=TxtTitle value='empty' style='position:absolute; left:240; top:102; width:0; height:76'> -->

    <!--     <label id=Label218 value='Protocol / Title:' style='position:absolute; left:96; top:102; width:117; height:24; visibility:'>Protocol / Title:</label> -->
  </body>
</html>
