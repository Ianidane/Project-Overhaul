<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function chkSecondary_Click(){
//If //Me.chkSecondary //Then;
        //Me.lblCHKSecondary.ForeColor = //16711935;
        //'Me.cmdDuplicate.Visible = False
    //Else;
        //Me.lblCHKSecondary.ForeColor = //0;
        //'Me.cmdDuplicate.Visible = True
//End //If;
;
}
function cmdDuplicate_Click(){
//On //Error //GoTo //err_cmdDuplicate_Click;
//If //Me.NewRecord //Then;
    //MsgBox "You must Save Data first.";
    //Exit //Sub;
//End //If;
//If //Me.chkSecondary //Then;
    //If //Not //YesNo("This //is //a //Secondary //SAE. //Are //you //sure //you //want //to //Continue?") //Then //Exit //Sub;
//End //If;
;
//'CancelUpdate = False
//'DataChanged = False
//'If SaveSAE = "DataError" Then Exit Sub
//Call //SetUPAgendaItem;
//'_____________________________-
;
;
;
    var stDocName = '';
    var stLinkCriteria = '';
    //CancelUpdate = //True;
    stDocName = "DrugSearch_SAE";
    //FromSAE = //True;
    //'DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
    //'Forms![_SAEForm].Requery
;
//exit_cmdDuplicate_Click:;
    //Exit //Sub;
;
//err_cmdDuplicate_Click:;
    //MsgBox //Err.Description;
    //Resume //exit_cmdDuplicate_Click;
    ;
}
function EventType_AfterUpdate(){
//If //Me.OnSite = //True //Then;
    //Me.Label1.ForeColor = //vbRed;
    //PutOnAgenda_Flag = //True;
//Else;
    //Me.Label1.ForeColor = //vbCyan;
    //PutOnAgenda_Flag = //False;
//End //If;
//If //Me.EventType.Column(1) = //True //Then;
    //Me.Label1.ForeColor = //vbRed;
    //PutOnAgenda_Flag = //True;
//End //If;
;
}
function FollowUpReportFlag_Click(){
//If //Me.FollowUpReportFlag //Then;
    //Me.TextFUNumber.Visible = //True;
    //'Me.TextOrigSAE.Visible = True
//Else;
    //Me.TextFUNumber.Visible = //False;
    //'Me.TextOrigSAE.Visible = False
//End //If;
;
}
function Form_Activate(){
//DoCmd.Restore;
}
function Form_AfterInsert(){
//DataChanged = //True;
}
function Form_AfterUpdate(){
//'MsgBox "afterupdate"
//On //Error //GoTo //Form_AfterUpdate_Error;
//DataChanged = //True;
//Me.cmdDuplicate.Enabled = //True;
//If //CyberIRB_Flag //Then //'bridge stuff
    //If //isFormLoaded("CyberBridge") //Then //'bridge stuff
        //gclsCyberIRB.ProcessedByProIRB "SAE"//, //Me.txtcyberrefnum //'bridge stuff
        //Form_CyberBridge_CyberNeedToProcess.RefreshGrid //'bridge stuff   Terry 080902
    //End //If;
     //'******** New 8/21/2010 updating Cyber
;
             //If //CyberIRB_Flag //Then;
                //If //IsNull(Me.txtcyberrefnum) //Then //Me.txtcyberrefnum = "NR" //& //Me.IRB_ //& //Me.SAE_Number;
                //If //UpdateCyberIRBSAE(Me.txtcyberrefnum) //Then   //' check to see if the bolupdatecyber is true or not
                    //UpdateCyberSAE //Me.[IRB#], //Me.[SAE //Number];
                //End //If;
                //Form_CyberBridge_CyberNeedToProcess.RefreshGrid    //'Bridge stuff Terry 090108
            //End //If;
 //'********
;
//End //If //'bridge stuff
//Form_AfterUpdate_Error:;
}
function Form_Close(){
//DataChanged = //False;
;
//Me.Filter = "";
;
;
;
;
//If //isFormLoaded("_IRB //Input //Form") //Then   //'FOR CHLA
    //DoCmd.Close //acForm, "_IRB Input Form"//, //acSaveYes;
    location.href = "Form__irb input form.php"; //'5/17/07
 //End //If;
//Exit_Form_Close:;
//DoCmd.Close //acReport, "STUDY SAEs"//, //acSaveNo;
//Exit_Close:;
//Exit //Sub;
;
;
}
function Form_BeforeUpdate(Cancel){
//'MsgBox "beforeupdate"
//On //Error //GoTo //Error_Up;
;
;
//If //IsNull(Me.Description_of_Event) //Or //Me.Description_of_Event = " " //Then;
   //Me.Description_of_Event = "None ";
//End //If;
    ;
//If //CancelUpdate //Then;
    //Me.Undo;
    //CancelUpdate = //False;
    //Exit //Sub;
//End //If;
;
;
;
;
;
//'DataChanged = True
//exit_SAEup:;
//Exit //Sub;
;
//Error_Up:;
    //MsgBox " Error trying to Update SAE Files    " //& //Err.Description //& "  " //& //Err.Number;
    //Resume //exit_SAEup;
;
}
function Form_Current(){
//Me.date_recieved_by.Visible = //True;
//If //Me.NewRecord //Then //Me.cmdDuplicate.Enabled = //False;
;
//If //Me.chkSecondary //Then;
        //Me.lblCHKSecondary.ForeColor = //16711935;
//'        Me.cmdDuplicate.Visible = False
    //Else;
        //Me.lblCHKSecondary.ForeColor = //0;
 //'       Me.cmdDuplicate.Visible = True
//End //If;
;
;
//If //Me.FollowUpReportFlag //Then;
    //Me.TextFUNumber.Visible = //True;
    //'Me.TextOrigSAE.Visible = True
//Else;
    //Me.TextFUNumber.Visible = //False;
    //'Me.TextOrigSAE.Visible =
//End //If;
   ;
//If //Me.OnSite = //True //Then;
    //Me.Label1.ForeColor = //vbRed;
    //PutOnAgenda_Flag = //True;
//Else;
    //Me.Label1.ForeColor = //vbBlue;
    //PutOnAgenda_Flag = //False;
//End //If;
;
//If //Me.EventType.Column(1) = //True //Then;
    //Me.Label1.ForeColor = //vbRed;
    //PutOnAgenda_Flag = //True;
//End //If;
;
}
function Form_Filter(Cancel, FilterType){
//Filter = "";
;
}
function Form_Load(){
//If //Me.chkSecondary //Then;
        //Me.lblCHKSecondary.ForeColor = //16711935;
        //'Me.cmdDuplicate.Visible = False
    //Else;
        //Me.lblCHKSecondary.ForeColor = //0;
        //'Me.cmdDuplicate.Visible = True
//End //If;
;
}
function Form_Open(Cancel){
//PutOnAgenda_Flag = //False;
//HoldOther1 = "";
//Me.OrderByOn = //False;
//If //Forms![newmenu]![HoldReadOnly] //And //Not //SAEUser //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.AddNew.Enabled = //False;
    //Me.Save_Data.Enabled = //False;
    //Me.SAELetterButton.Enabled = //False;
//End //If;
;
;
}
function NextButton_Click(){
//On //Error //GoTo //EndofFile;
//CancelUpdate = //True;
;
;
             //If //CyberIRB_Flag //Then;
                //If //IsNull(Me.txtcyberrefnum) //Then //Me.txtcyberrefnum = "NR" //& //Me.IRB_ //& //Me.SAE_Number;
                //If //UpdateCyberIRBSAE(Me.txtcyberrefnum) //Then   //' check to see if the bolupdatecyber is true or not
                    //UpdateCyberSAE //Me.[IRB#], //Me.[SAE //Number];
                //End //If;
                //Form_CyberBridge_CyberNeedToProcess.RefreshGrid    //'Bridge stuff Terry 090108
            //End //If;
;
;
;
//DoCmd.RunCommand //acCmdRecordsGoToNext;
//exit_next:;
//Exit //Sub;
//EndofFile:;
//MsgBox "End of SAEs"//, //vbInformation;
//GoTo //exit_next;
}
function OnSite_AfterUpdate(){
//If //Me.OnSite = //True //Then;
    //Me.Label1.ForeColor = //vbRed;
    //PutOnAgenda_Flag = //True;
//Else;
    //Me.Label1.ForeColor = //vbCyan;
    //PutOnAgenda_Flag = //False;
//End //If;
;
//If //Me.EventType.Column(1) = //True //Then;
    //Me.Label1.ForeColor = //vbRed;
    //PutOnAgenda_Flag = //True;
//End //If;
;
}
function preview_print_Click(){
//On //Error //GoTo //Err_preview_print_Click;
//DataChanged = //False;
;
//If //SaveSAE = "DataError" //Then //Exit //Sub;
;
;
//MyFilter = "";
;
    var stDocName = '';
;
;
   //MyFilter = //Me.Filter;
   //'MsgBox Me.Filter
    stDocName = "SaeReport";
//'    Me.Visible = False
//'DoCmd.Close
//'DoCmd.SelectObject acForm, "_irb input form"
   //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
    ;
 ;
   ;
//Exit_preview_print_Click:;
    //Exit //Sub;
//Err_preview_print_Click:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_preview_print_Click;
        ;
 ;
;
}
function PreviousButton_Click(){
//On //Error //GoTo //BegofFile;
//CancelUpdate = //True;
//DoCmd.RunCommand //acCmdRecordsGoToPrevious;
;
//exit_prev:;
//Exit //Sub;
;
//BegofFile:;
//MsgBox "Beginning of SAEs"//, //vbInformation;
//GoTo //exit_prev;
;
;
}
function Returnmenu_Click(){
//CancelUpdate = //True;
//DoCmd.Close;
//DoCmd.SelectObject //acForm, "_irb input form" //'5/17/07
}
function Save_Data_Click(){
//CancelUpdate = //False;
//DataChanged = //True;
//'Me.NewRecord = True
//If //SaveSAE = "DataError" //Then //Exit //Sub;
;
//exit_save_data:;
;
//CancelUpdate = //True;
//'DoCmd.Close
//Exit //Sub;
;
;
   ;
}
function AddNew_Click(){
//On //Error //GoTo //Err_AddNew_Click;
//Me.AllowAdditions = //True;
//Me.Text106.SetFocus;
//DoCmd.GoToRecord //, //, //acNewRec;
;
//Exit_AddNew_Click:;
    //Exit //Sub;
;
//Err_AddNew_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_AddNew_Click;
    ;
}
function CheckIfOn(){//As //Boolean;
;
//On //Error //GoTo //ErrorCheckif;
//could initialize hope here
//could initialize dbs here
//Set dbs = CurrentDb;
//could initialize GoForIt here
;
//sqlSTR = "SELECT [agenda records].[IRB#], [agenda records].MeetingDate, [agenda records].[IRB Action Category]" //_;
//& " FROM [agenda records]WHERE [agenda records].[IRB#]= " //& "'" //& //[Forms]![_SAEForm]![IRB#] //& "'" //& " AND [agenda records].MeetingDate <> " //& "#" //& //[Forms]![_SAEForm]![MeetingDate] //& "#" //& " AND [agenda records].[IRB Action Category] Is Null ";
;
//Set GoForIt = dbs.OpenRecordset(sqlSTR);
;
//If //GoForIt.RecordCount //> //0 //Then;
    //GoForIt.MoveFirst;
    //'msgbox GoForIt("[IRB#]")
    //If //YesNo("The //Study //is //already //on //the //agenda //for " & GoForIt![MeetingDate] & "//.  //Do //you //still //want //to //put //the //SAE //on //the //date //you //selected?") //Then //'done for CHLA
        //CheckIfOn = //False;
    //Else;
        //CheckIfOn = //True;
    //End //If;
    //GoForIt.Close;
    //Exit //Function;
//Else;
    //GoForIt.Close;
    //CheckIfOn = //False;
//End //If;
//exit_checkif:;
//Exit //Function;
;
//ErrorCheckif:;
//MsgBox "Error Checking for correct agenda date" //& //Err.Description //& "  " //& //Err.Number;
//Resume //exit_checkif;
;
}
function SAELetterButton_Click(){
//On //Error //GoTo //Err_SAELetterButton_Click;
;
//LetterName = "SAE Acknowledgment";
;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName = "_LetterManager";
    //SAELetterSwitch = //True;
    //If //Not //IsNull(Me.MeetingDate) //Then;
        //HoldMeetingDate = //Me.MeetingDate;
    //Else;
        //HoldMeetingDate = "";
    //End //If;
    //HoldIRBNo = //Me.IRB_;
    //Forms![newmenu]![SourceNumber] = //Me.SAE_Number;
    stLinkCriteria = "[IRB#]=" //& "'" //& //HoldIRBNo //& "'";
    ;
    //CancelUpdate = //False;
    //DataChanged = //False;
    //If //SaveSAE = "DataError" //Then //Exit //Sub;
    ;
    //If //Me.NewRecord //Then;
        //MsgBox "No data has been entered";
        //Exit //Sub;
    //End //If;
    ;
    //If //GetTemplateLocation("SAE") //Then;
        ;
        //If //TemplateName = "" //Then //Exit //Sub;
        //DoCmd.SetWarnings //False;
        //'MsgBox Me.SAE_Number
        //'MsgBox [Forms]![_SAEForm]![IRB#]
        //DoCmd.OpenQuery //("ExportSAEwithCreateTable");
        //DoCmd.SetWarnings //True;
        ;
        //CustomLetter = //True;
    //Else;
        ;
        //'msgbox "Do PROIRB SAE Letter"
        //CustomLetter = //False;
    //End //If;
   //' DoCmd.Close
    location.href = "Form_"+stDocName+".php";//, //acNormal, //, //stLinkCriteria, //, //acWindowNormal;
;
//Exit_SAELetterButton_Click:;
    //Exit //Sub;
;
//Err_SAELetterButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_SAELetterButton_Click;
    ;
}
function DoAddAgendaSAE(){
//On //Error //GoTo //err_DoAddSae;
//If //Me.IRB_Code //Like "*Tulane*" //Then //GoTo //AddTulane;
//If //AddAgendaRecord(Me.IRB_, //Me.SAE_Number, //_;
//Me.MeetingDate, "SAE"//, //Me.Category, //Me.condition1, //Me.Condition2, //_;
//Me.Description_of_Event, //_;
"z "//, " z"//, "z "//, " z"//, "z "//, //0, //Date, //0, //Forms![newmenu]![UserName], //_;
//Forms![newmenu]![Location], //_;
//IIf(IsNull(Me.Location_if_Off_Site), " "//, " Site-" //& //Me.Location_if_Off_Site //& ";"//) //& //_;
//IIf((Me.FollowUpReportFlag = //False), " "//, "F/U Report Number-" //& //Me.TextFUNumber //& ";"//) //& //_;
//IIf(IsNull(Me.TextOrigSAE), " "//, "Orig SAE#-" //& //Me.TextOrigSAE //& ";"//) //& //_;
//IIf(IsNull(Me.date_recieved_by), " "//, "Date Received;" //& //Me.date_recieved_by //& ";"//) //& //_;
//IIf(IsNull(Me.Date_Event_Occured), " ;"//, "Event Date " //& //Me.Date_Event_Occured //& ";"//) //& //_;
//IIf(IsNull(Me.RelatedCombo), " "//, " Event-" //& //Me.RelatedCombo //& ";"//) //& //_;
//IIf(IsNull(Me.Text106), " "//, " IND-" //& //Me.Text106 //& ";"//) //& //_;
//IIf(IsNull(Me.Medwatch__), " "//, " MedWatch-" //& //Me.Medwatch__ //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Identifier), " "//, " Reference-" //& //Me.Patient_Identifier //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Age), " "//, " PT, Age; " //& //Me.Patient_Age //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Sex), " "//, " PT, Sex; " //& //Me.Patient_Sex //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Status), " "//, " Pt Status;" //& //Me.Patient_Status //& ";"//) //& //_;
//IIf((Me.Are_the_Risks_Altered_ = //False), ":"//, "-Risks Altered" //& ";"//) //& //_;
//IIf((Me.New_Consent_Required_ = //False), ":"//, "-New Consent Required" //& ";"//) //& //_;
//IIf(IsNull(Me.TextAware), ":"//, "-Date PI Aware" //& //Me.TextAware //& ";"//) //& //_;
//IIf(IsNull(Me.Date_Signed), ":"//, "Date PI Signed-" //& //Me.Date_Signed //& ";"//), " z"//, "z "//, //0, //Me.txtcyberrefnum) //Then //'bridgestuff
//End //If;
//GoTo //exit_DoAddSae:;
;
//AddTulane:  //'Copied the above add sunstituted blanks where where tulane didn't want any
;
//If //AddAgendaRecord(Me.IRB_, //Me.SAE_Number, //_;
//Me.MeetingDate, "SAE"//, //Me.Category, //Me.condition1, //Me.Condition2, //_;
//Me.Description_of_Event, //_;
"z "//, " z"//, "z "//, " z"//, "z "//, //0, //Date, //0, //Forms![newmenu]![UserName], //_;
//Forms![newmenu]![Location], //_;
//IIf(IsNull(Me.Location_if_Off_Site), " "//, " Site-" //& //Me.Location_if_Off_Site //& ";"//) //& //_;
//IIf((Me.FollowUpReportFlag = //False), " "//, "F/U Report Number-" //& //Me.TextFUNumber //& ";"//) //& //_;
//IIf(IsNull(Me.TextOrigSAE), " "//, "Orig SAE#-" //& //Me.TextOrigSAE //& ";"//) //& //_;
//IIf(IsNull(Me.date_recieved_by), " "//, "Date Received;" //& //Me.date_recieved_by //& ";"//) //& //_;
//IIf(IsNull(Me.Date_Event_Occured), "No Date of Event;"//, "Event Date " //& //Me.Date_Event_Occured //& ";"//) //& //_;
//IIf(IsNull(Me.RelatedCombo), " "//, " Event-" //& //Me.RelatedCombo //& ";"//) //& //_;
//IIf(IsNull(Me.Text106), " "//, " IND-" //& //Me.Text106 //& ";"//) //& //_;
//IIf(IsNull(Me.Medwatch__), " "//, " MedWatch-" //& //Me.Medwatch__ //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Identifier), " "//, " Reference-" //& //Me.Patient_Identifier //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Age), " "//, " PT, Age; " //& //Me.Patient_Age //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Sex), " "//, " PT, Sex; " //& //Me.Patient_Sex //& ";"//) //& //_;
//IIf(IsNull(Me.Patient_Status), " "//, " Pt Status;" //& //Me.Patient_Status //& ";"//), " z"//, "z "//, //0, //Me.txtcyberrefnum) //Then;
;
//End //If;
;
;
//exit_DoAddSae:;
//Exit //Sub;
//err_DoAddSae:;
//MsgBox "Error trying to add SAE Agenda Item--Notify PRO_IRB Support" //& //Err.Description //& " " //& //Err.Source;
//Resume //exit_DoAddSae;
}
function SaveSAE(){//As //String;
//On //Error //GoTo //Err_SaveSAE;
;
//If //IsNull(Me.date_recieved_by.Value) //Then;
     //MsgBox "Date Received cannot be blank"//, //vbInformation;
     //SaveSAE = "DataError";
     //Exit //Function;
//End //If;
//If //IsNull(Me.EventType) //Then;
     //MsgBox "Event Type cannot be blank"//, //vbInformation;
     //SaveSAE = "DataError";
    //Exit //Function;
//End //If;
       ;
       //'Deletes If Sae HAS agenda record previously on file
 ;
//If //Me.NewRecord //Then;
 //'Me.SAE_Number = 1403 '(IsNull(DMax(" [Sae]![SAE Number]", "SAE")), 1, DMax(" [Sae]![SAE Number]", "SAE") + 1)
  //Me.SAE_Number = //IIf(IsNull(DMax(" //[Sae]![SAE //Number]", "SAE"//)), //1, //DMax(" //[Sae]![SAE //Number]", "SAE"//) //+ //1);
  //'MsgBox "Please record the following SAE Internal # to your document---" & Me![SAE Number], vbInformation
  //If //IsNull(Me.txtcyberrefnum) //Then //Me.txtcyberrefnum = "NR" //& //Me.IRB_ //& //Me.[SAE //Number];
//End //If;
//On //Error //Resume //Next;
//Save_SAE_Record:;
;
//DoCmd.RunCommand //acCmdSaveRecord;
//If //Err.Number //<> //0 //Then;
    //If //Err.Description //Like "*dupl*" //Then;
        //Me.SAE_Number = //IIf(IsNull(DMax(" //[Sae]![SAE //Number]", "SAE"//)), //1, //DMax(" //[Sae]![SAE //Number]", "SAE"//) //+ //1);
        //' MsgBox "The SAE Number has been changed to " & Me.SAE_Number
        //Err.Clear;
        //GoTo //Save_SAE_Record;
    //Else;
        //MsgBox "Error trying to save SAE Data--" //& //Err.Number //& "--" //& //Err.Description //' done chla
        //SaveSAE = "DataError";
        //Exit //Function;
    //End //If;
//End //If;
 ;
 ;
 //'MsgBox "Please record the following SAE Internal # to your document---" & Me![SAE Number], vbInformation
;
//If //DataChanged //Then;
//MsgBox "Please record the following SAE Internal # to your document---" //& //Me![SAE //Number], //vbInformation;
    //DataChanged = //False;
;
    //If //Not //IsNull(Me.MeetingDate) //Then;
        //MsgBox "The original Agenda item and any Meeting Minutes information for this SAE has been deleted. You can re-enter Meeting Information(if any) via the Post Meeting Action button on the Main Menu.";
        //If //Not //Delete_Agenda_Record(Me.SAE_Number) //Then;
           //GoTo //Exit_SaveSAE;
        //End //If;
    //End //If;
;
    ;
    //If //Me.StatusText //Like "*exem*" //Then;
        //MsgBox "SAEs for Exempt Studies will not appear on the Agenda.  However the SAE will be recorded.";
        //GoTo //Exit_SaveSAE  //'if study is exempt no agenda record for SAE
    //End //If;
;
//DoItAgainSae:;
    //If //PutOnAgenda_Flag //Then  //'means SAE goes on the Agenda
        //Call //SetUPAgendaItem;
        //Call //AssignAgendaDate;
        ;
        //If //Not //ReassignSwitch //Then;
        //'IsNull(Me.MeetingDate) Then
               //MsgBox //("The //SAE //has //NOT //been //placed //on //the //Agenda");
        //Else;
            //Me.MeetingDate = //HoldMeetingDate;
            //If //CheckIfOn //Then //GoTo //DoItAgainSae //'check if on the agenda for a different date
                //Forms![newmenu]![SourceNumber] = //Me.SAE_Number;
                //Call //DoAddAgendaSAE;
                //Forms![_irb //input //form]![AgendaDate] = //Me.MeetingDate;
            //End //If;
        //End //If;
       //TempSource = //Forms![_SAEForm].SAE_Number;
;
    //Forms![_SAEForm].AllowEdits = //True;
    //Forms![_SAEForm].DataEntry = //False;
    //Forms![_SAEForm].RecordSource = "_SelectSAEsForRequery"   //'done for CHLA
//End //If;
//Exit_SaveSAE:;
 //'******** New 8/21/2010 updating Cyber
//'DoCmd.RunCommand acCmdSaveRecord
             //If //CyberIRB_Flag //Then;
                //If //UpdateCyberIRBSAE(Me.txtcyberrefnum) //Then   //' check to see if the bolupdatecyber is true or not
                    //UpdateCyberSAE //Me.[IRB#], //Me.[SAE //Number];
                //End //If;
                //Form_CyberBridge_CyberNeedToProcess.RefreshGrid    //'Bridge stuff Terry 090108
            //End //If;
 //'********
;
//SaveSAE = "SaveOK";
;
//Exit //Function;
//Err_SaveSAE:;
;
//SaveSAE = "DataError";
//MsgBox "Error trying to save SAE Data  - Click OK and resave the data making note of the SAE Number.   " //& //Err.Number //& " " //& //Err.Description //' done chla
//Resume //Exit_SaveSAE;
}
function Menu_DatasheetView(){
//'Me.AllowEdits = False
//'Me.AllowAdditions = False
//'Me.AllowDeletions = False
//'DoCmd.OpenForm "_saeform", acDesign
//'Me.DefaultView = 2
;
location.href = "Form__saeform.php";//, //acFormDS, //, //, //acFormReadOnly, //acWindowNormal;
//Me.AllowEdits = //False;
//Me.AllowAdditions = //False;
//Me.AllowDeletions = //False;
;
}
function Menu_Formview(){
//Me.AllowEdits = //True;
//Me.AllowAdditions = //True;
//Me.AllowDeletions = //True;
location.href = "Form__saeform.php";//, //acNormal, //, //, //acFormEdit, //acWindowNormal;
}
function Delete_Agenda_Record(SAENumber){//As //Boolean;
//On //Error //GoTo //err_Delete_Agenda_Record;
//could initialize db here
//could initialize rs here
var sqlSTR = '';
//Set db = CurrentDb;
;
            //'this deletes the agenda record previously on file for this sae
;
    sqlSTR = "select [agenda records].[IRB#], [agenda records].MeetingDate, [agenda records].SourceNumber, [agenda records].[Record Type]" //_;
            //& "FROM [agenda records]WHERE ([agenda records].SourceNumber)=" //& "'" //& //[Forms]![_SAEForm]![SAE //Number] //& "'" //& " AND ([agenda records].[Record Type])=" //& "'SAE'";
    //Set rs = db.OpenRecordset(sqlSTR);
    ;
    //If //Not //rs.RecordCount //> //0 //Then;
            //'MsgBox "Could not find Original Agenda Record.  Processing will continue."
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
;
//exit_Delete_Agenda_Record:;
//Delete_Agenda_Record = //True;
//Exit //Function;
;
//err_Delete_Agenda_Record:;
//MsgBox "Error Deleting Previous Agenda Record.  New SAE data not saved.  Return to Study Input Form and try again. " //& //Err.Number //& "   " //& //Err.Description;
//Delete_Agenda_Record = //False;
//Resume //exit_Delete_Agenda_Record;
}
function SetUPAgendaItem(){
//Me.Category = "Adverse Event";
    //If //Me.[On //Site //?] = //True //Then;
        //Me.condition1 = "On Site SAE";
    //Else;
        //Me.condition1 = "Off Site SAE";
    //End //If;
    //If //IsNull(Me.EventType) //Then //Me.EventType = "None";
    ;
    //Me.Condition2 = //Me.EventType;
    //HoldIRBNo = //Me.IRB_;
    //If //Me.SecondaryFlag //Then;
        //Me.condition1 = "Secondary--" //& //condition1;
    //Else;
        //Me.condition1 = "Primary--" //& //condition1;
    //End //If;
    ;
;
;
}
function AssignAgendaDate(){
location.href = "Form_ReassignAgendaRecord.php";//, //, //, //, //, //acDialog;
    ;
        ;
    ;
}
function Command140_Click(){
//On //Error //GoTo //Err_Command140_Click;
//If //Not //SaveSAE = "SaveOK" //Then;
    //MsgBox "You must first Save Data";
    //Exit //Sub;
//End //If;
    var stDocName = '';
    var stLinkCriteria = '';
    //CancelUpdate = //True;
    stDocName = "DrugSearch";
    //FromSAE = //True;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
;
//Exit_Command140_Click:;
    //Exit //Sub;
;
//Err_Command140_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command140_Click;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();Form_Active();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:768; height:64'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <input type='text' id=IRB# value='empty' style='position:absolute; left:972; top:36; width:132; height:30'>

    <label id=IRB# Label value='Study#' style='position:absolute; left:1020; top:6; width:69; height:24; visibility:'>Study#</label>
    <input type='text' id=SAE Number value='empty' style='position:absolute; left:16; top:154; width:120; height:25'>

    <label id=SAE Number Label value='SAE Internal #' style='position:absolute; left:10; top:124; width:133; height:24; visibility:'>SAE Internal #</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:526; top:108; width:588; height:0'>

    <label id=Protocol Number & Title Label value='Protocol#/Title' style='position:absolute; left:526; top:84; width:588; height:24; visibility:'>Protocol#/Title</label>
    <input type='text' id=Date Event Occured value='empty' style='position:absolute; left:12; top:466; width:126; height:31'>

    <label id=Date Event Occured Label value='Date of Event' style='position:absolute; left:12; top:441; width:127; height:24; visibility:'>Date of Event</label>
    <input type='text' id=Description of Event value='empty' style='position:absolute; left:528; top:210; width:589; height:120'>

    <label id=Description of Event Label value='Description of Event' style='position:absolute; left:690; top:180; width:192; height:24; visibility:'>Description of Event</label>
    <input type='text' id=Medwatch # value='empty' style='position:absolute; left:672; top:366; width:210; height:24'>

    <label id=Medwatch # Label value='MedWatch #' style='position:absolute; left:699; top:336; width:118; height:24; visibility:'>MedWatch #</label>
    <input type='text' id=Patient Age value='empty' style='position:absolute; left:399; top:402; width:54; height:24'>
    <input type='checkbox' id=Are the Risks Altered? value='False' style='position:absolute; left:1110; top:354; width:36; height:30'>

    <label id=Are the Risks Altered? Label value='Are the Risks Altered?' style='position:absolute; left:888; top:354; width:199; height:24; visibility:'>Are the Risks Altered?</label>
    <input type='checkbox' id=New Consent Required? value='False' style='position:absolute; left:1110; top:384; width:42; height:30'>

    <label id=New Consent Required? Label value='New Consent Required?' style='position:absolute; left:886; top:384; width:214; height:24; visibility:'>New Consent Required?</label>
    <input type='checkbox' id=Signed by PI ? value='False' style='position:absolute; left:1023; top:490; width:24; height:30'>

    <label id=Signed by PI ? Label value='Signed by PI ?' style='position:absolute; left:882; top:486; width:133; height:24; visibility:'>Signed by PI ?</label>
    <input type='text' id=Date Signed value='empty' style='position:absolute; left:1002; top:516; width:127; height:24'>

    <label id=Date Signed Label value='Date Signed' style='position:absolute; left:882; top:516; width:114; height:24; visibility:'>Date Signed</label>
    <label id=Label46 value='Age' style='position:absolute; left:403; top:376; width:54; height:24; visibility:'>Age</label>
    <label id=Label47 value='Sex' style='position:absolute; left:478; top:376; width:54; height:24; visibility:'>Sex</label>
    <input type='checkbox' id=OnSite value='=False' style='position:absolute; left:114; top:247; width:0; height:0'>

    <label id=LabelOnSite value='Local?' style='position:absolute; left:10; top:247; width:91; height:24; visibility:'>Local?</label>
    <input type='text' id=date recieved by value='empty' style='position:absolute; left:405; top:465; width:126; height:24'>

    <label id=Label56 value='Date Received' style='position:absolute; left:399; top:433; width:136; height:24; visibility:'>Date Received</label>
    <!--     <input type='checkbox' id=Event related value='False' style='position:absolute; left:462; top:18; width:0; height:0'> -->
    <select name='Patient Sex' Size='12' style='position:absolute; left:465; top:400; width:96; height:0'>
    <input type='text' id=Patient Identifier value='empty' style='position:absolute; left:546; top:463; width:247; height:24'>

    <label id=Patient Identifier Label value='Reference #' style='position:absolute; left:546; top:433; width:115; height:24; visibility:'>Reference #</label>
    <input type='button' id=preview_print value='Preview\015\012 SAEs' style='position:absolute; left:913; top:555; width:153; height:48' onclick='preview_print_Click();'>
    <input type='button' id=Returnmenu value='Return' style='position:absolute; left:1066; top:555; width:120; height:48' onclick='Returnmenu_Click();'>
    <input type='button' id=Save Data value='Save Data' style='position:absolute; left:577; top:555; width:150; height:48' onclick='Save Data_Click();'>
    <!--     <input type='text' id=treviewmonth value='empty' style='position:absolute; left:558; top:36; width:54; height:30'> -->
    <label id=Label1 value='Serious Adverse Event' style='position:absolute; left:1; top:73; width:328; height:40; visibility:'>Serious Adverse Event</label>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:298; top:36; width:676; height:30'>

    <label id=Label453 value='IRB' style='position:absolute; left:405; top:6; width:82; height:24; visibility:'>IRB</label>
    <input type='checkbox' id=MedWatch Report Filed value='=False' style='position:absolute; left:594; top:390; width:0; height:0'>

    <label id=Label98 value='MedWatch Report Filed?' style='position:absolute; left:546; top:336; width:121; height:48; visibility:'>MedWatch Report Filed?</label>
    <!--     <input type='text' id=condition1 value='empty' style='position:absolute; left:168; top:54; width:138; height:72'> -->

    <label id=Label134 value='Text133' style='position:absolute; left:276; top:84; width:78; height:24; visibility:'>Text133</label>
    <!--     <input type='text' id=Condition2 value='empty' style='position:absolute; left:324; top:0; width:114; height:15'> -->

    <!--     <label id=Label136 value='Text135' style='position:absolute; left:486; top:48; width:78; height:24; visibility:hidden'>Text135</label> -->
    <!--     <input type='text' id=Category value='empty' style='position:absolute; left:492; top:0; width:186; height:24'> -->

    <!--     <label id=Label138 value='Text137' style='position:absolute; left:555; top:42; width:78; height:24; visibility:hidden'>Text137</label> -->
    <input type='button' id=NextButton value='Next SAE' style='position:absolute; left:226; top:555; width:168; height:48' onclick='NextButton_Click();'>
    <input type='button' id=PreviousButton value='Previous SAE' style='position:absolute; left:52; top:555; width:175; height:48' onclick='PreviousButton_Click();'>
    <input type='button' id=AddNew value='Add New SAE' style='position:absolute; left:393; top:555; width:184; height:48' onclick='AddNew_Click();'>
    <input type='checkbox' id=FollowUpReportFlag value='False' style='position:absolute; left:394; top:121; width:0; height:0'>

    <label id=Label104 value='Is this a Follow-up Report' style='position:absolute; left:150; top:117; width:226; height:24; visibility:'>Is this a Follow-up Report</label>
    <input type='text' id=Text106 value='empty' style='position:absolute; left:16; top:213; width:204; height:25'>

    <label id=Label107 value='IND Report #' style='position:absolute; left:10; top:186; width:121; height:24; visibility:'>IND Report #</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:33; top:40; width:163; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <!--     <select name='MaxSAE' Size='38' style='position:absolute; left:306; top:78; width:0; height:0'> -->

    <label id=Label110 value='SAE Number' style='position:absolute; left:162; top:78; width:76; height:24; visibility:'>SAE Number</label>
    <input type='button' id=SAELetterButton value='Acknowledgment\015\012Letter' style='position:absolute; left:727; top:555; width:186; height:48' onclick='SAELetterButton_Click();'>
    <input type='text' id=MeetingDate value='empty' style='position:absolute; left:648; top:504; width:144; height:24'>

    <label id=Label101 value='On Meeting Date' style='position:absolute; left:486; top:504; width:153; height:24; visibility:'>On Meeting Date</label>
    <!--     <label id=Label118 value='16711680' style='position:absolute; left:168; top:186; width:93; height:24; visibility:hidden'>16711680</label> -->
    <!--     <input type='text' id=StatusText value='empty' style='position:absolute; left:28; top:117; width:159; height:21'> -->

    <label id=Label120 value='Text119' style='position:absolute; left:198; top:186; width:78; height:24; visibility:'>Text119</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:43; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <input type='text' id=TextAware value='empty' style='position:absolute; left:984; top:426; width:127; height:24'>

    <label id=Label122 value='Date PI Aware' style='position:absolute; left:882; top:414; width:99; height:48; visibility:'>Date PI Aware</label>
    <select name='Location if Off Site' Size='6' style='position:absolute; left:268; top:247; width:249; height:0'>

    <label id=Location if Off Site Label value='Location' style='position:absolute; left:178; top:247; width:82; height:24; visibility:'>Location</label>
    <label id=Label124 value='Follow-up Report #' style='position:absolute; left:216; top:150; width:171; height:24; visibility:'>Follow-up Report #</label>
    <select name='EventType' Size='7' style='position:absolute; left:12; top:298; width:480; height:43'>

    <label id=LabelEventType value='Type of Event:' style='position:absolute; left:12; top:274; width:135; height:24; visibility:'>Type of Event:</label>
    <select name='TextFUNumber' Size='3' style='position:absolute; left:387; top:150; width:126; height:24'>
    <label id=Label125 value='Original SAE#' style='position:absolute; left:258; top:184; width:127; height:24; visibility:'>Original SAE#</label>
    <input type='text' id=TextOrigSAE value='empty' style='position:absolute; left:388; top:184; width:124; height:25'>
    <select name='RelatedCombo' Size='8' style='position:absolute; left:156; top:348; width:231; height:84'>

    <label id=Label111 value='Study Related?' style='position:absolute; left:12; top:348; width:139; height:24; visibility:'>Study Related?</label>
    <select name='Patient Status' Size='10' style='position:absolute; left:156; top:466; width:231; height:79'>

    <label id=Patient Status Label value='Patient Status' style='position:absolute; left:156; top:441; width:231; height:24; visibility:'>Patient Status</label>
    <input type='button' id=cmdDuplicate value='D\015\012R\015\012U\015\012G\015\012/\015\012D\015\012E\015\012V\015\012I\015\012' style='position:absolute; left:1146; top:22; width:31; height:516'>
    <input type='checkbox' id=chkSecondary value='False' style='position:absolute; left:342; top:87; width:150; height:0'>

    <label id=lblCHKSecondary value='Secondary SAE' style='position:absolute; left:364; top:84; width:151; height:24; visibility:'>Secondary SAE</label>
    <input type='text' id=txtCyberRefNum value='empty' style='position:absolute; left:252; top:6; width:72; height:24'>

    <label id=Label145 value='refnum:' style='position:absolute; left:168; top:6; width:78; height:24; visibility:'>refnum:</label>
  </body>
</html>
