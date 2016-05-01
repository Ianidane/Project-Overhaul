<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="file:///C|/VBA2PHP/ProIRB_PHP/images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function UpdatePrintMArgins(){
//On //Error //GoTo //err_PrintMargins;
//' This routine restes the top margin on the Letters listed in Availableletters       '
;
;
       //could initialize wrkAdmin here
       //could initialize dbs here
       //could initialize ctrReports here
       //could initialize docReport here
       //could initialize PrtMipString here
       //could initialize PM here
       //could initialize rpt here
       var strName = '';
       var GetMargin = 0;
       //could initialize rst here
       //'Dim sqlstr As String
       //DBEngine.SystemDB = "c:\cc\proirb_dev\security.mdw";
       //' Create a new session and assign it to the database administrator.
       //Set wrkAdmin;
       //Set dbs;
       //' Open the report in Design view and set the report object variable.
       //Application.Echo //False, "Updating Print Margins";
       //DoCmd.Echo //False;
       //DoCmd.SetWarnings //False;
     //sqlSTR = "SELECT lettername, HoldTopMargin FROM AvailableLetters";
      //Set rst;
;
      //rst.MoveLast;
      //rst.MoveFirst;
      //'MsgBox rst.RecordCount
       //Do;
             //If //rst.EOF //Then //Exit //Do;
             //'this loops bu ignores acknowlement letter do to hold over from early databases
             //If //rst.Fields("LetterName") = "Acknowlegement Letter" //Or //(rst.Fields("LetterName") = "Education Letter"//) //Or //(rst.Fields("LetterName") = "Follow-Up"//) //Then;
              //GoTo //SkipRecord;
            //End //If;
             strName;
             //DoCmd.OpenReport //strName, //acDesign;
             //Set rpt;
             //If //IsNull(rst.Fields("holdtopmargin")) //Then;
                 GetMargin;
             //Else;
                  GetMargin;
             //End //If;
            ;
               //PrtMipString.strRGB = //rpt.PrtMip;
             //LSet PM;
             //PM.lngTopMargin = GetMargin;
             //LSet PrtMipString;
             //rpt.PrtMip = //PrtMipString.strRGB;
                    //' Close the design of the report, saving the changes,
                    //' and then preview the report.
             //DoCmd.Close //acReport, //strName, //acSaveYes;
//SkipRecord:;
            //rst.MoveNext;
       //Loop;
      ;
      //Application.Echo //True;
      //DoCmd.Echo //True;
      //DoCmd.SetWarnings //True;
;
      //Set rst;
      //Set dbs;
      //Set wrkAdmin;
      //Exit //Sub;
//err_PrintMargins:;
    //MsgBox "You must set a default Printer.  Then restart PRO_IRB";
    //Application.Quit;
;
;
}
function SelectStudy(){
//On //Error //GoTo //error_select;
    //CancelUpdate = //False;
                //Me.ReviewMonth = "***";
                //Me.RenewalMonth = "***";
                //Me.RenewalCycle = "***";
                //Me.Status = "OPEN";
                //HoldStatus = //Me.Status;
                //HoldLocation = //Me.Location;
                //searchirbs = //False;
                //LetterManagerSwitch = //False;
                ;
                location.href = "Form__SelectStudy.php";//, //, //, //, //, //acDialog;
                //If //NewApplication //Then;
                    //holdcaption = "New Study";
                    //searchirbs = //False;
                    location.href = "Form__irb input form.php";//, //acNormal, //, //, //acFormAdd;
                    //Exit //Sub;
                //Else;
                //End //If;
            ;
            //If //CancelUpdate //Then;
                //CancelUpdate = //False;
                //Exit //Sub;
            //Else;
                //Me.study = //HoldStudy;
                //holdcaption = "Study #  " //& //Me.study;
                location.href = "Form__irb input form.php";//, //acNormal;
            //End //If;
//exit_select:;
//Exit //Sub;
//error_select:;
//MsgBox "Error # 100 Opening Select Study Form";
//Resume //exit_select;
}
function Command0_Click(){
//DoCmd.OpenQuery "xxTextQuery2";
}
function btnOpenBridge_Click(){
//If //CyberIRB_Flag //Then;
    //gclsCyberIRB.CyberProtocolsProcessing //' bridge
//End //If;
;
}
function CBDrugCorrelation_Click(){
location.href = "Form_Drugsearch.php";;
}
function command_79_Click(){
//If //GoButton() //Then //Exit //Sub //Else //Application.Quit;
}
function Command164_Click(){
//Me.ActionGroup.Value = //False;
//Me.SelectionsCombo.Value = //False;
;
;
//If //SelectAll //Then;
    //Exit //Sub;
//Else;
    location.href = "Form_FollowupLetters.php";;
//End //If;
}
function Command79_Click(){
//If //GoButton //Then //Exit //Sub //Else //Application.Quit;
}
function ContinuingReview_Click(){
//Call //SetupButtons;
//HoldRowNumber = //0;
//CancelUpdate = //False;
  //'DoCmd.Close
  location.href = "Form__display studies.php";;
}
function Form_Activate(){
//Me.ActionGroup.Value = //False;
//Me.SelectionsCombo.Value = //False;
//Me.Toggle140.Enabled = //True;
//Me.Toggle140 = //False;
//Me.QuickStudyNumber.SetFocus;
}
function Form_Current(){
//Me.QuickStudyNumber.Requery;
;
//'DoCmd.DeleteObject acTable, "TempAgendaRecords"
}
function Form_Load(){
//If //ExemptOverrideSwitch //Then;
    //Me.holdexemptoverride = "ZZZZZZZZZ";
//Else;
    //Me.holdexemptoverride = "*Exempt*";
//End //If;
;
;
//Call //UpdatePrintMArgins;
;
}
function Form_Open(Cancel){
//RedisplayAgenda = //False;
//BeenClicked = //0;
//If //Not //IsNull(DMax("ReserveDate1", "qryReservedDate"//)) //Then;
    //Me.CBDrugCorrelation.Enabled = //True;
//End //If;
;
//If //Not //CyberIRB_Flag //Then //btnOpenBridge.Visible = //False //'cyberbridge 090117
;
//Me.Location = //Me.Location.ItemData(0);
//Me.TBLocation = //Me.Location;
;
//DoCmd.SetWarnings //False;
//DoCmd.OpenQuery "DeleteTempAgendaRecords";
//DoCmd.OpenQuery "qryMakeTempDrugSearchTable";
;
;
//DoCmd.SetWarnings //True;
;
//RenewAnniversary = //False;
//If //UserEdit //Then;
    //Me.HoldReadOnly = //False;
    ;
//Else;
    //Me.HoldReadOnly = //True;
    //Me.LetterManagerButton.Enabled = //False;
    //Me.Command153.Enabled = //False;
    ;
//End //If;
//Me.UserName = //UserLastName;
//Me.TCurrentUser = //UserLastName;
;
//labelcolor = //4210816;
//FormBgColor = //-2147483633;
//CancelUpdate = //False;
//NewApplication = //False;
//If //Forms![newmenu]![HoldReadOnly] //And //Not //SAEUser //Then;
    //Me.LetterManagerButton.Enabled = //False;
    //Me.Caption = //Me.Caption //& " --Read Only";
//End //If;
//If //SAEUser //Then;
    //Me.LetterManagerButton.Enabled = //False;
    //Me.Caption = //Me.Caption //& " --Read Only and SAE User";
//End //If;
;
//DoCmd.SetWarnings //False;
//could initialize FirstTime here
//If //IsNull(DMax("[LetterName]", "AvailableLetters"//, "[LetterName] = 'General Letter'"//)) //Then;
 ;
    //'set up entries for the General Letter
    //DoCmd.OpenQuery "AppendGeneralLetterToAvailableLetters1";
    //DoCmd.OpenQuery "AppendGeneralLetterParagraps";
    //DoCmd.OpenQuery "AppendGeneralLetterClosings";
//End //If;
;
//'
FirstTime = DMax("[LetterName]", "CurrentLetterClosing1", "[LetterName] = 'General Letter' and [IRBCode] = " && "'" && Me.Location && "'");
//If //IsNull(FirstTime) //Then;
    //'This repopulates the current letter closings with the defaults
    //DoCmd.RunSQL "DELETE CurrentLetterClosing1.* FROM CurrentLetterClosing1 WITH OWNERACCESS OPTION;";
    //DoCmd.OpenQuery "AppendAllToCurrentLetterClosings1";
      ;
//Else;
//End //If;
;
   //DoCmd.OpenQuery "DeleteNullPreMeetingActionError";
//DoCmd.SetWarnings //True;
;
;
;
;
;
}
function Exit_Application_Click(){
//If //YesNo("Are //you //certain //you //want //to //Exit?") //Then;
        //GoTo //Continue_exit;
    //Else;
        //Exit //Sub;
    //End //If;
//Continue_exit:;
;
//If //Me.HoldReadOnly //Then //GoTo //Exit_Shell;
//If //YesNo("Do //you //wish //to //back //up //the //Database //at //this //time?") //Then;
;
    //Else;
        //GoTo //Exit_Shell;
    //End //If;
//On //Error //GoTo //err_shell;
    ;
    var shellpath = '';
    //could initialize try here
    //Set try;
    var LocBackupMDB = '';
    //could initialize AccessPAth here
    //could initialize one here
    //could initialize two here
    //could initialize three here
    //could initialize four here
    //could initialize oApp here
    //could initialize SetPathBack here
    //'oApp.Close
    ;
    AccessPAth = SysCmd(acSysCmdAccessDir) && "msaccess.exe";
    one = strQuote && AccessPAth && strQuote;
    //If //SysCmd(acSysCmdRuntime) //Then;
        two = " /runtime /wrkgrp " && strQuote && "C:\Cc\proirb_dev\security.mdw" && strQuote;
    //Else;
        two = " /wrkgrp " && strQuote && "C:\Cc\proirb_dev\security.mdw" && strQuote;
    //End //If;
    three = " /user client /pwd " && strQuote && strQuote && " " && strQuote && "C:\Cc\proirb_dev\PROIRB_backup.mdb" && strQuote;
    four = one && two && three;
    //Shell //four, //vbNormalFocus;
//Exit_Shell:;
    //Application.Quit //acQuitSaveAll  //'done for chla
//err_shell:;
//MsgBox "Cannot locate " //& four;
//Resume //Exit_Shell;
}
function IRBActivityButton_Click(){
//' ginny - 12/17/03
//' open the EducationMaintenance form
    location.href = "Form_EducationMaintenance.php";//, //acNormal;
;
//On //Error //GoTo //err_IRBActivityButton_Click;
//Exit_IRBActivityButton_Click:;
    //Exit //Sub;
;
//err_IRBActivityButton_Click:;
    //MsgBox "Error opening Education Maintenance Form: " //& //Err.Number //& ",  " //& //Err.Description, //vbInformation;
    //Resume //Exit_IRBActivityButton_Click;
;
}
function LetterManagerButton_Click(){
var stDocName = '';
var stLinkCriteria = '';
var MinStudy = '';
//'If IsNull(DMin("[IRB#]", "IRB - Research  Database")) Then
 //'   DisplayMessage ("You will need at least one Study on File to create a General Letter")
  //'  Exit Sub'
//'Else
 //'   MinStudy = DMin("[IRB#]", "IRB - Research  Database")
//'End If
;
//If //GetTemplateLocation("GeneralLetters") //Then;
        //If //TemplateName = "" //Then;
            //Exit //Sub;
        //Else;
            //CustomLetter = //True;
        //End //If;
//Else;
        //CustomLetter = //False;
//End //If;
;
//GeneralLetterSwitch = //True;
;
stDocName;
//'stLinkCriteria = "[IRB#]=" & "'" & MinStudy & "'"
;
;
location.href = "Form_"+stDocName+".php";//, //acNormal;
;
;
;
;
;
;
;
;
}
function location_AfterUpdate(){
//Me.TBLocation = //Me.Location;
//If //Me.Location = "***" //Then;
    //Me.Toggle140.Enabled = //True;
    //If //Me.ActionGroup.Value = //False //Then;
        //MsgBox "You may select All locations only when searching.  Click OK then Select the search item..then click ALL or select '***'.";
        //Me.Location = //Me.Location.ItemData(0);
        //Me.Toggle140.Enabled = //False;
        //Exit //Sub;
    //Else;
        //Me.Toggle140 = //True;
    //End //If;
//Else;
    //RenewAnniversary = //False;
   ;
//End //If;
;
}
function location_Enter(){
//Me.Toggle140.Enabled = //True;
//Me.Toggle140 = //False;
;
}
function PostResults_Click(){
//Call //SetupButtons;
//AgendaDetail = //False;
//'DoCmd.Close
location.href = "Form__agendamonthcheck.php";;
//'Me.Visible = False
}
function PrelimButton_Click(){
//BeenClicked = //0;
//Call //SetupButtons;
;
location.href = "Form_CheckAgendaRecordStatus1.php";//, //, //, //, //, //acDialog;
;
//Exit //Sub;
}
function QuickStudyNumber_AfterUpdate(){
//Quick_Click;
}
function QuickStudyNumber_KeyDown(KeyCode, Shift){
//Me.QuickStudyNumber.Dropdown;
}
function SelectionsCombo_Click(){
//Me.ActionGroup.Value = //Me.SelectionsCombo.Column(1);
//Me.Toggle140.Enabled = //True;
;
;
}
function Command153_Click(){
//Me.ActionGroup.Value = //False;
//Me.SelectionsCombo.Value = //False;
;
;
//'Codes and Tables Maintenance
;
    //If //UserLastName //Like "*Admin*" //Then;
        //If //SelectAll //Then;
            //Exit //Sub;
        //Else;
          location.href = "Form_CodesandTables.php";;
        //End //If;
    //Else;
        //MsgBox "Current User not Authorized for Administrative Functions"//, //vbInformation;
    //End //If;
            ;
}
function SelectionsCombo_DblClick(Cancel){
//If //GoButton //Then //Exit //Sub //Else //Application.Quit;
//Cancel = //True;
}
function StudyButton_Click(){
//On //Error //GoTo //Err_StudyButton_Click;
//Me.SelectionsCombo.Value = //False;
;
//Me.ActionGroup.Value = //False;
;
//If //SelectAll //Then;
    //Exit //Sub;
//Else;
  //Call //SelectStudy;
//End //If;
;
//Exit_StudyButton_Click:;
    //Exit //Sub;
;
//Err_StudyButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_StudyButton_Click;
    ;
}
function SAEButton_Click(){
//On //Error //GoTo //Err_SAEButton_Click;
//Me.SelectionsCombo.Value = //False;
;
//Me.ActionGroup.Value = //False;
;
//If //SelectAll //Then;
    //Exit //Sub;
//Else;
 //Call //SelectStudy;
//End //If;
   ;
//Exit_SAEButton_Click:;
    //Exit //Sub;
;
//Err_SAEButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_SAEButton_Click;
    ;
}
function CPAButton_Click(){
//On //Error //GoTo //Err_CPAButton_Click;
//Me.ActionGroup.Value = //False;
//Me.SelectionsCombo.Value = //False;
;
;
//If //SelectAll //Then;
    //Exit //Sub;
//Else;
    //Call //SelectStudy;
//End //If;
   ;
;
//Exit_CPAButton_Click:;
    //Exit //Sub;
;
//Err_CPAButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_CPAButton_Click;
    ;
}
function ICButton_Click(){
//On //Error //GoTo //Err_ICButton_Click;
//Me.ActionGroup.Value = //False;
//Me.SelectionsCombo.Value = //False;
;
//If //SelectAll //Then;
    //Exit //Sub;
//Else;
    //Call //SelectStudy;
//End //If;
//Exit_ICButton_Click:;
    //Exit //Sub;
;
//Err_ICButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ICButton_Click;
    ;
}
function LogoutButton_Click(){
//On //Error //GoTo //Err_LogoutButton_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
    //LoginLogoutSwitch = //True;
;
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //Me.TCurrentUser = //UserLastName;
    //LoginLogoutSwitch = //False;
   ;
;
//Exit_LogoutButton_Click:;
    //Exit //Sub;
;
//Err_LogoutButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_LogoutButton_Click;
    ;
}
function Toggle140_Click(){
//If //Toggle140 //Then;
    //Me.Location = "***";
//Else;
     //Me.Location = //Me.Location.ItemData(0);
    //If //Me.ActionGroup.Value = //False //Then;
        //MsgBox "You may select All locations only when searching.  Click OK then Select the search item..then click ALL or select '***'.";
        //Me.Location = //Me.Location.ItemData(0);
//'        Me.Toggle140.Enabled = False
        //Exit //Sub;
    //End //If;
    ;
//End //If;
}
function SelectAll(){//As //Boolean;
//If //Me.Location = "***" //Then;
    //MsgBox "You may select All locations only when searching.  Click OK then Select the search item..then select All Locations or ***.";
    //SelectAll = //True;
//Else;
    //SelectAll = //False;
//End //If;
}
function Quick_Click(){
var stDocName = '';
var stLinkCriteria = '';
//could initialize dbs here
//Set dbs;
//could initialize GoForIt here
//On //Error //GoTo //Err_Quick_Click;
//If //IsNull(Me.QuickStudyNumber) //Or //Me.QuickStudyNumber = "" //Then //Exit //Sub;
//'Quick # is location insensitive.  Check if on file and set location to irb code of selected study
    //sqlSTR = "SELECT [IRB - Research  Database].[IRB Code], [IRB - Research  Database].[IRB#],[IRB - Research  Database].[Study Active?] FROM [IRB - Research  Database]" //_;
            //& "WHERE ([IRB - Research  Database].[IRB#])=" //& "'" //& //[Forms]![newmenu]![QuickStudyNumber] //& "'";
;
//Set GoForIt;
//If //GoForIt.RecordCount //> //0 //Then;
    //Me.Status = //GoForIt![Study //Active?];
    //HoldStatus = //Me.Status;
    //HoldLocation = //GoForIt![IRB //Code];
    //Me.Location = //HoldLocation;
    //searchirbs = //False;
    //LetterManagerSwitch = //False;
    //Me.study = //Me.QuickStudyNumber;
    //HoldStudy = //Me.QuickStudyNumber;
    //Me.QuickStudyNumber = "";
    //holdcaption = "Study #  " //& //Me.study;
    location.href = "Form__irb input form.php";;
    //GoTo //Exit_Quick_Click;
 //Else;
    //MsgBox "Study not on file";
    //Me.QuickStudyNumber.SetFocus;
//End //If;
;
//Exit_Quick_Click:;
    //'GoForIt.Close
    //Exit //Sub;
;
//Err_Quick_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Quick_Click;
    ;
}
function Update_Backend(){
//On //Error //GoTo //Err_Update;
;
;
;
;
//Exit_Update:;
    //Exit //Sub;
//Err_Update:;
    //MsgBox //Err.Description //& "  " //& "Error occurred during Update Backend";
    //Resume //Exit_Update;
}
function SetupButtons(){
//Me.ActionGroup.Value = //False;
//Me.SelectionsCombo.Value = //False;
;
//If //SelectAll //Then;
   //Exit //Sub;
//Else;
   //HoldLocation = //Me.Location;
//End //If;
}
function Command199_Click(){
//On //Error //GoTo //Err_Command199_Click;
;
    //could initialize oApp here
;
    //Set oApp;
    //oApp.Visible = //True;
;
//Exit_Command199_Click:;
    //Exit //Sub;
;
//Err_Command199_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command199_Click;
    ;
}
function GoButton(){//As //Boolean;
//On //Error //GoTo //err_GoButton;
//GoButton = //True;
//could initialize String1 here
//could initialize String2 here
//QueryString = //Me.SelectionsCombo;
//QueryValue = //Me.ActionGroup.Value;
;
//Select //Case //Me.ActionGroup.Value;
//'Select Case 51 'Me.ActionGroup.Value
//Case //4  //'Open studies
             //CancelUpdate = //False;
            //Me.Status = "Open";
            //holdcaption = "Open Studies";
            //HoldLocation = //Me.Location;
            //Me.PI = "***";
            //Me.Sponsor = "***";
            //Me.CoOrd = "***";
           location.href = "Form__IRBSearchNew.php";;
           //Forms![_IRBSearchNew].TimerInterval = //100;
        ;
    ;
//Case //5   //'Closed Studies
            //CancelUpdate = //False;
            //Me.PI = "***";
            //Me.Sponsor = "***";
            //Me.CoOrd = "***";
            //Me.Status = "Closed";
            //holdcaption = "Closed Studies";
            //HoldLocation = //Me.Location;
            location.href = "Form__IRBSearchNew.php";;
            //Forms![_IRBSearchNew].TimerInterval = //100;
   ;
    ;
//Case //6  //'All Studies
             //CancelUpdate = //False;
            //Me.study = "*";
            //Me.PI = "***";
            //Me.Sponsor = "***";
            //Me.CoOrd = "***";
            //Me.Status = "***";
            //HoldStatus = //Me.Status;
            //holdcaption = "All Studies";
            //HoldLocation = //Me.Location;
            //searchirbs = //True;
            location.href = "Form__IRBSearchNew.php";;
            //Forms![_IRBSearchNew].TimerInterval = //100;
;
;
//Case //7  //'Studies for a Principal Investigators
        //CancelUpdate = //False;
        location.href = "Form__SelectPI.php";//, //, //, //, //, //acDialog;
        //If //CancelUpdate //Then;
            //CancelUpdate = //False;
            //Exit //Function;
        //Else;
            //Me.PI = //HoldPI;
            //Me.PIName = //HoldPIName;
           //Me.Sponsor = "***";
            //Me.CoOrd = "***";
           //Me.Status = "***";
              ;
            location.href = "Form__IRBSearchNew.php";;
            //Forms![_IRBSearchNew].TimerInterval = //100;
;
           ;
        //End //If;
;
    ;
//Case //8 //'Studies for a Sponsor
        //CancelUpdate = //False;
        location.href = "Form__SelectSponsors.php";//, //, //, //, //, //acDialog;
        ;
        //If //CancelUpdate //Then;
            //CancelUpdate = //False;
            //Exit //Function;
        //Else;
            //Me.Sponsor = //HoldSponsor;
            //Me.SponsorName = //HoldSponsorName;
            //Me.PI = "***";
            //Me.Status = "***";
            //Me.CoOrd = "***";
            //HoldLocation = //Me.Location;
            //SponsorQuery = //True;
           location.href = "Form__IRBSearchNew.php";;
           //Forms![_IRBSearchNew].TimerInterval = //100;
           ;
        //End //If;
        ;
 //Case //9 //'Studies for a Coordinator
        //CancelUpdate = //False;
        location.href = "Form__SelectCoord.php";//, //, //, //, //, //acDialog;
        ;
        //If //CancelUpdate //Then;
            //CancelUpdate = //False;
            //Exit //Function;
        //Else;
            //Me.CoOrd = //HoldCoord;
            //Me.CoordName = //HoldCoordName;
            //Me.PI = "***";
            //Me.Sponsor = "***";
            //Me.Status = "***";
            //HoldLocation = //Me.Location;
            location.href = "Form__IRBSearchNew.php";;
            //Forms![_IRBSearchNew].TimerInterval = //100;
            ;
        //End //If;
        ;
//Case //17  //'ALL SAEs for a location
        //HoldLocation = //Me.Location;
            //MyFilter = "[IRB Code] like " //& //strQuote //& //[Forms]![newmenu]![Location] //& //strQuote;
                   ;
            location.href = "Form__SAEForm_Search.php";//, //, //, //MyFilter;
            //Forms![_saeform_search].TimerInterval = //100;
            ;
            ;
            ;
//Case //16    //'All CPAs for a location
             //CancelUpdate = //False;
            //HoldLocation = //Me.Location;
                      ;
            //MyFilter = "[IRB Code] like " //& //strQuote //& //[Forms]![newmenu]![Location] //& //strQuote;
                   ;
            location.href = "Form__CPASearchForm.php";//, //, //, //MyFilter;
            //Forms![_CPASearchForm].TimerInterval = //100;
            ;
//Case //11, //12, //13, //14 //' Add Modify Active Study---also use for SAE or CPA  options
                 //CancelUpdate = //False;
                //Me.ReviewMonth = "***";
                //Me.RenewalMonth = "***";
                //Me.RenewalCycle = "***";
                //Me.Status = "OPEN";
                //HoldStatus = //Me.Status;
                //HoldLocation = //Me.Location;
                //searchirbs = //False;
                //LetterManagerSwitch = //False;
                ;
                location.href = "Form__SelectStudy.php";//, //, //, //, //, //acDialog;
                //If //NewApplication //Then;
                    //holdcaption = "New Study";
                    //searchirbs = //False;
                    location.href = "Form__irb input form.php";//, //acNormal, //, //, //acFormAdd;
                    //Exit //Function;
                //Else;
                //End //If;
            ;
            //If //CancelUpdate //Then;
                //CancelUpdate = //False;
                //Exit //Function;
            //Else;
                //Me.study = //HoldStudy;
                //holdcaption = "Study #  " //& //Me.study;
                location.href = "Form__irb input form.php";//, //acNormal, "_hold IRB Input query";
            //End //If;
            ;
;
//Case //18         //'pi  search
            //SearchContacts = //True;
            location.href = "Form__contact PI.php";;
            ;
                ;
//Case //19;
            //'sponsor search
                //'pi  search
            //SearchContacts = //True;
            location.href = "Form__contact sponsor.php";;
//Case //20;
            //'search coordinators
                    //'pi  search
            //SearchContacts = //True;
            location.href = "Form__contact coordinator.php";;
//Case //21;
            //'search board members
            ;
            ;
            //If //IsNull(DMax("[Common //Contact //ID]", "_selectIRBMembers"//)) //Then;
                //MsgBox "No Board Members on File";
                //Exit //Function;
            //Else;
                //SearchContacts = //True;
                location.href = "Form__contactBoardMember.php";;
            //End //If;
//Case //22;
            //'search "Other" members
            //If //IsNull(DMax("[Common //Contact //ID]", "_SelectOtherContact1"//)) //Then;
                //MsgBox "No 'Other' type Contacts on File";
                //Exit //Function;
            //Else;
                //SearchContacts = //True;
                location.href = "Form__ContactOther.php";;
            //End //If;
//Case //23;
            //DoCmd.OpenReport "InvoiceAndPaymentReporting"//, //acViewPreview;
            //Exit //Function;
//Case //24;
            //DoCmd.OpenReport "InvoiceAndPaymentReportingFromDate"//, //acViewPreview;
            //Exit //Function;
            ;
;
;
;
//Case //50;
            //'Special Report1 for Boston College
                //On //Error //GoTo //err_GoButton;
                //DoCmd.OpenReport "rptBCcustom1"//, //acViewPreview;
                //Exit //Function;
//Case //51 //'Special Report2 for Boston College
                //On //Error //GoTo //err_GoButton;
                //'DoCmd.OpenReport "rptBostonCollege2", acViewPreview
                location.href = "Form_frm_FromAndToDates.php";;
                //Exit //Function;
//Case //55;
            //'Special Report for UT Tyler
                //On //Error //GoTo //err_GoButton;
                //DoCmd.OpenReport "RptCustRptUTexas1"//, //acViewPreview;
                //Exit //Function;
//Case //60;
            //'Special Report for Beth Isreal
                //On //Error //GoTo //err_GoButton;
                //DoCmd.OpenReport "ClosedStudiesWithNoClassification"//, //acViewPreview;
                //Exit //Function;
//Case //70      //'Special Report for Liberty made up of two steps. 1 (70) to make a table  "  &
             //'from a query and then display the records in a datasheet for the"  &
              //'user to edit. Then when the datasheet is closed the report is run
              //On //Error //Resume //Next;
              //DoCmd.SetWarnings //False;
              ;
              //If //YesNo("Do //you //wish //to //Select //a //new //set //of //records?") //Then;
                    //DoCmd.OpenQuery "qryDeleteLibertyReportRecords";
                    //DoCmd.OpenQuery "qryCustRptLiberty1_MakeTable";
                    //'DoCmd.OpenQuery "qryCustRptLiberty1_Select", acViewPreview
                    location.href = "Form_frm_CustRptLiberty1_Select.php";//, //acFormDS;
                //Else;
                    //'DoCmd.OpenQuery "qryCustRptLiberty1_Select", acViewPreview
                    location.href = "Form_frm_CustRptLiberty1_Select.php";//, //acFormDS;
              //End //If;
              //DoCmd.SetWarnings //True;
              //Exit //Function;
//Case //71        //'Running the actual Liberty report
                //On //Error //GoTo //err_GoButton;
                ;
                //DoCmd.OpenReport "rptLiberty1"//, //acViewPreview;
                //Exit //Function;
//Case //72         //'Mount Sinai
                //DoCmd.OpenReport "rptSinaiAnnualReport"//, //acViewPreview;
                //Exit //Function;
                ;
//Case //73;
            //'Special Report for NY Dept of Health rpt1
                //On //Error //Resume //Next;
                //DoCmd.OpenReport "rptNYDOHNewStudies"//, //acViewPreview;
                //Exit //Function;
//Case //74;
            //'Special Report for NY Dept of Health rpt2
                //On //Error //Resume //Next;
                //DoCmd.OpenReport "rptNYDOHCPAs"//, //acViewPreview;
                //Exit //Function;
                ;
;
//Case //Else;
;
//MsgBox "Please Choose an Action";
//End //Select;
//exit_GoButton:;
//Exit //Function;
//err_GoButton:;
//If //Err.Number = //2580 //Then;
    //MsgBox " You must select records first";
    //GoTo //exit_GoButton;
//End //If;
;
//MsgBox //Err.Description;
//GoButton = //False;
//Resume //exit_GoButton;
;
}
function Command222_Click(){
//On //Error //GoTo //Err_Command222_Click;
;
;
    //Screen.PreviousControl.SetFocus;
    //DoCmd.FindNext;
;
//Exit_Command222_Click:;
    //Exit //Sub;
;
//Err_Command222_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command222_Click;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();Form_Active();Form_Current();">
	<form method="post">

    <label id='empty' value='empty' style='visibility:'></label></input></input></input>
  <select name='location' Size='1' style='position:absolute; left:195; top:132; width:723; height:36'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRB Meeting Schedule1].[IRB Code] FROM [IRB Meeting Schedule1] ORDER BY [IRB Meeting Schedule1].[IRB Code]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['IRB Code']."' selected".">".$row['IRB Code']."</Option>";
        }
    ?>
    </select>

    <label id=Irb Code_Label value='Select an IRB' style='position: absolute; left: 99px; top: 135px; width: 97px; height: 34;visibility:'>Select an IRB</label>
    <label id=Label80 value='Task Managers' style='position:absolute; left:804; top:211; width:192; height:34; visibility:'>Task Managers</label>
    <label id=Label104 value='Search and Query' style='position:absolute; left:414; top:210; width:240; height:36; visibility:'>Search and Query</label>
    <label id=Label117 value='Add or Modify' style='position:absolute; left:84; top:211; width:174; height:34; visibility:'>Add or Modify</label>
    <button name='Toggle140' type='submit' style='position: absolute; left: 924; top: 132; width: 28px; height: 31px' onclick='Toggle140_Click();'>"All "</button>
    <input type='button' id=Command153 value='A&dministration' style='position:absolute; left:726; top:393; width:382; height:48' onclick='Command153_Click();'></input>
    <input type='button' id=LetterManagerButton value='General &Letter Choices' style='position:absolute; left:726; top:295; width:382; height:48' onclick='LetterManagerButton_Click();'></input>
    <input type='button' id=Command164 value='&Follow Up Manager' style='position:absolute; left:726; top:343; width:382; height:48' onclick='Command164_Click();'></input>
    <select name='SelectionsCombo' Size='7' style='position:absolute; left:360; top:258; width:358; height:211' onclick='SelectionsCombo_Click();'>
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

        $sql = sqlsrv_query($conn, "SELECT [ReportsQueriesForms1].[SelectionName], [ReportsQueriesForms1].[OptionValue], [ReportsQueriesForms1].[Type] FROM ReportsQueriesForms1 ORDER BY [ReportsQueriesForms1].[OptionValue]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['SelectionName']."'".">".$row['SelectionName']."</Option>";
        }
    ?>
    </select>
    <!--     <input type='text' id=DateCreated value='empty' style='position:absolute; left:240; top:447; width:222; height:30'></input> -->

    <label id=Label166 value='Date Created' style='position:absolute; left:162; top:447; width:102; height:24; visibility:'>Date Created</label>
    <!--     <input type='text' id=UserName value='empty' style='position:absolute; left:372; top:457; width:252; height:0'></input> --><!--     <input type='text' id=ReviewMonth value='empty' style='position:absolute; left:522; top:439; width:150; height:30'></input> --><!--     <input type='text' id=SourceNumber value='empty' style='position:absolute; left:474; top:451; width:204; height:0'></input> --><!--     <input type='text' id=RenewalCycle value='empty' style='position:absolute; left:408; top:442; width:132; height:36'></input> --><!--     <input type='text' id=PI value='empty' style='position:absolute; left:336; top:441; width:180; height:0'></input> -->

  <label id=Label114 value='Text113:' style='position:absolute; left:192; top:441; width:70; height:24; visibility:'>Text113:</label>
    <!--     <input type='text' id=CoordName value='empty' style='position:absolute; left:312; top:441; width:126; height:18'></input> -->

    <label id=Label135 value='CoordName' style='position:absolute; left:168; top:441; width:70; height:24; visibility:'>CoordName</label>
    <!--     <input type='text' id=Coord value='empty' style='position:absolute; left:132; top:433; width:138; height:0'></input> --><!--     <input type='text' id=SponsorName value='empty' style='position:absolute; left:498; top:451; width:174; height:0'></input> --><!--     <input type='text' id=Status value='empty' style='position:absolute; left:198; top:439; width:156; height:36'></input> -->

    <label id=Label92 value='Text91:' style='position:absolute; left:54; top:439; width:61; height:24; visibility:'>Text91:</label>
    <!--     <input type='text' id=Text93 value='empty' style='position:absolute; left:492; top:420; width:0; height:54'></input> --><!--     <input type='text' id=Sponsor value='empty' style='position:absolute; left:528; top:447; width:132; height:0'></input> --><!--     <input type='text' id=RenewalMonth value='empty' style='position:absolute; left:408; top:433; width:0; height:42'></input> -->

    <label id=Label90 value='Text89:' style='position:absolute; left:264; top:433; width:61; height:24; visibility:'>Text89:</label>
    <!--     <input type='text' id=study value='empty' style='position:absolute; left:378; top:441; width:0; height:0'></input> -->
    <!--     <input type='text' id=PIName value='empty' style='position:absolute; left:348; top:210; width:102; height:18'></input> -->
    <input type='button' id=StudyButton value='&Study / Protocol' style='position:absolute; left:18; top:255; width:336; height:48' onclick='StudyButton_Click();'></input>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='font-size:10px; position: absolute; left: 49px; top: 35px; width: 166; height: 25;visibility:'>by ProIRB Plus, Inc.</label>
    <label id=Label183 style='font-size: 40px; position: absolute; left: 381px; top: 46px; width: 245px; height: 66;visibility:' value='Main Menu' name="Label183">Main Menu</label>
    <!--     <input type='text' id=aQuickStudyNumber value='empty' style='position:absolute; left:660; top:204; width:150; height:37'></input> -->
    <!--     <input type='text' id=HoldSource value='empty' style='position:absolute; left:66; top:90; width:66; height:0'></input> -->
    <!--     <input type='checkbox' id=HoldFU value='True' style='position:absolute; left:696; top:102; width:96; height:18'></input> -->
    <!--     <input type='checkbox' id=HoldReadOnly value='False' style='position:absolute; left:696; top:444; width:42; height:6'></input> -->

    <!--     <label id=Label196 value='Check195' style='position:absolute; left:719; top:441; width:81; height:24; visibility:hidden'>Check195</label> -->
    <input type='button' id=PrelimButton value='Preliminary &Agenda Items' style='position:absolute; left:18; top:309; width:336; height:48' onclick='PrelimButton_Click();'></input>
    <input type='button' id=ContinuingReview value='&Due for Continuing Review' style='position:absolute; left:18; top:366; width:336; height:48' onclick='ContinuingReview_Click();'></input>
    <input type='button' id=PostResults value='&Post IRB Meeting Actions' style='position:absolute; left:18; top:421; width:336; height:48' onclick='PostResults_Click();'></input>
    <!--     <input type='text' id=MeetingKey value='empty' style='position:absolute; left:171; top:81; width:84; height:22'></input> --><!--     <input type='text' id=CONTactID value='empty' style='position:absolute; left:954; top:186; width:90; height:16'></input> -->
    <!--     <input type='text' id=holdexemptoverride value='empty' style='position:absolute; left:102; top:144; width:60; height:22'></input> -->
    <label id=Label108 value='ProIRB' style='font-size: 28px; position: absolute; left: 2px; top: 2px; width: 130; height: 44px;visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='font-size:10px; position: absolute; left: 91px; top: 2px; width: 18; height: 16;visibility:'>R</label>
    <!--     <input type='text' id=NewMenuMeetingDate value='empty' style='position:absolute; left:304; top:105; width:109; height:16'></input> -->
    <input type='button' id=IRBActivityButton value=' Education Module/Tracking' style='position:absolute; left:726; top:246; width:382; height:49' onclick='IRBActivityButton_Click();'></input>
    <!--     <input type='text' id=TxtPostMeetingDate value='empty' style='position:absolute; left:303; top:22; width:166; height:31'></input> --><!--     <input type='text' id=TBlocation value='empty' style='position:absolute; left:942; top:114; width:0; height:18'></input> -->
    <input type='button' id=CBDrugCorrelation value='Drug/Device Related Module' style='position:absolute; left:726; top:442; width:382; height:48' onclick='CBDrugCorrelation_Click();'></input>
    <!--     <input type='text' id=CoInvString value='empty' style='position:absolute; left:0; top:0; width:93; height:78'></input> -->
    <!--     <input type='text' id=txtTemplateName value='empty' style='position:absolute; left:954; top:0; width:66; height:6'></input> --><!--     <input type='text' id=txtInvoiceType value='empty' style='position:absolute; left:954; top:16; width:66; height:9'></input> --><!--     <input type='text' id=txtAmountInvoiced value='empty' style='position:absolute; left:954; top:57; width:66; height:6'></input> --><!--     <input type='text' id=txtMeetingDate value='empty' style='position:absolute; left:954; top:72; width:66; height:6'></input> --><!--     <input type='text' id=txtPayCPANumber value='empty' style='position:absolute; left:954; top:39; width:66; height:6'></input> --><!--     <input type='text' id=txtInvCPANumber value='empty' style='position:absolute; left:954; top:88; width:66; height:6'></input> --><!--     <input type='text' id=TxtDatePaid value='empty' style='position:absolute; left:954; top:100; width:66; height:6'></input> --><!--     <input type='text' id=TxtIRBcode value='empty' style='position:absolute; left:954; top:93; width:66; height:6'></input> --></input>


    <input type='button' id=exit_application value='&Exit Application' style='position: absolute; left: 917px; top: 492px; width: 205; height: 40' onclick='exit_application_Click();'></input>
    <input type='button' id=command_79 value='&GO' style='position: absolute; left: 782px; top: 492px; width: 133; height: 40' onclick='command_79_Click();'></input>
    <input type='button' id=LogoutButton value='Change User' style='position: absolute; left: 29px; top: 493px; width: 192; height: 40' onclick='LogoutButton_Click();'></input>
    <input type='text' id=TCurrentUser value='empty' style='position: absolute; left: 223px; top: 494px; width: 168; height: 37'></input>
    <input type='button' id=Quick value='&Quick Study #' style='position: absolute; left: 390px; top: 492px; width: 192; height: 40' onclick='Quick_Click();'></input>
    <select name='QuickStudyNumber' Size='1' style='position: absolute; left: 581px; top: 493px; width: 201; height: 39px'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRB - Research  Database1].[Protocol Number & Title], [IRB - Research  Database1].[IRB#], [IRB - Research  Database1].[Date Received] FROM [IRB - Research  Database1] ORDER BY [IRB - Research  Database1].[IRB#]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Protocol Number & Title']."'".">".$row['Protocol Number & Title']."</Option>";
        }
    ?>
    </select>
  <input type='button' id=btnOpenBridge value='Open Bridge' style='position: absolute; left: 838px; top: 8px; width: 174; height: 57' onClick='btnOpenBridge_Click();'></input>
  </form>
  </body>
</html>
