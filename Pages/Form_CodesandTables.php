<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Command2_Click(){
//End //Sub;
;
//Private //Sub //ChangeIRB_Click();
//On //Error //GoTo //err_change;
//TransactionSwitch = //True;
//CancelUpdate = //False;
//Forms![newmenu].Status = "***";
location.href = "Form__SelectStudy.php";//, //, //, //, //, //acDialog;
//TransactionSwitch = //False;
;
;
//If //CancelUpdate = //True //Then;
    //Exit //Sub;
//Else;
//Me.StudyNumber = //HoldStudy;
//Forms![newmenu]![study] = //HoldStudy;
//Me.TBStudyT = //Mid$(Me.StudyNumber, //1, //9) //& "T";
//End //If;
;
//HoldNewIRBCode = " ";
location.href = "Form_ChangeIRB.php";//, //, //, //, //, //acDialog;
;
//If //HoldNewIRBCode = " " //Then;
        //Exit //Sub;
//Else;
;
    //DoCmd.SetWarnings //False;
    //'MsgBox Me.StudyNumber
    //'Exit Sub
//DoCmd.OpenQuery "AppendTransferredToStudyStatus";
//DoCmd.OpenQuery "ChangeStudyNumberTo_T"  //'The Changed IRB Study Number now suffixed everywhere with a "T"
   ;
//DoCmd.OpenQuery "CopyIRBStudyWithout_The_T"  //'Make a copy or the IRB ReseaRCH RECORD without the T to use for posted agenda records
//DoCmd.OpenQuery "UpdateLocationforStudy";
//DoCmd.OpenQuery "UpdatePostedAgendaForStudy";
//DoCmd.OpenQuery "UpdateUnpostedAgendaWithNewIRBCode";
//DoCmd.SetWarnings //True;
location.href = "Form_SelectUnpostedAgendaItems.php";;
//Forms![SelectUnpostedAgendaItems].TimerInterval = //100;
    ;
//End //If;
//Exit //Sub;
;
;
//err_change:;
//MsgBox //Err.Description;
//Exit //Sub;
}
function Closings_Click(){
location.href = "Form_MaintainLetterClosings.php";;
;
}
function CMD_DIV_Click(){
//On //Error //GoTo //Err_CMD_DIV_Click;
    var stDocName = '';
    stDocName;
    location.href = "Form_"+stDocName+".php";;
    ;
//Exit_CMD_DIV_Click:;
    //Exit //Sub;
;
//Err_CMD_DIV_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_CMD_DIV_Click;
    ;
}
function cmdAgendaCategories_Click(){
//On //Error //GoTo //Err_cmdAgendaCategories_Click;
    var stDocName = '';
    stDocName;
   ;
        location.href = "Form_"+stDocName+".php";//, //acFormDS;
  ;
//Exit_cmdAgendaCategories_Click:;
    //Exit //Sub;
;
//Err_cmdAgendaCategories_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdAgendaCategories_Click;
    ;
}
function cmdBenefits_Click(){
//On //Error //GoTo //Err_cmdBenefits_Click;
    var stDocName = '';
    stDocName;
    location.href = "Form_"+stDocName+".php";;
    ;
//Exit_cmdBenefits_Click:;
    //Exit //Sub;
;
//Err_cmdBenefits_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdBenefits_Click;
    ;
}
function Command17_Click(){
//On //Error //GoTo //error_query;
location.href = "Form_IRBMeetingDates.php";//, //acFormDS;
;
//exit_command17:;
//Exit //Sub;
//error_query:;
//MsgBox //Err.Number //& " " //& //Err.Description;
//Resume //exit_command17;
}
function Command18_Click(){
//End //Sub;
;
//Private //Sub //Command19_Click();
//If //Demo //Then;
    location.href = "Form_.php";//, //acViewNormal, //acReadOnly;
//Else;
    //DoCmd.OpenQuery "_AvailableLetters";
//End //If;
;
;
;
;
}
function Command43_Click(){
location.href = "Form_MaintainICTemplate.php";;
;
;
}
function AgendaCat_Click(){
//On //Error //GoTo //Err_AgendaCat_Click;
;
    var stDocName = '';
    stDocName;
        location.href = "Form_"+stDocName+".php";//, //acFormDS;
   ;
;
//Exit_AgendaCat_Click:;
    //Exit //Sub;
;
//Err_AgendaCat_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_AgendaCat_Click;
}
function Command53_Click(){
//'Board Member Maintenance
  //SearchContacts = //False;
            ;
            location.href = "Form__Select Board Member.php";//, //acNormal, //, //, //, //acDialog;
            //If //NewApplication //Then;
                //NewApplication = //False;
                //HoldCoordName = " ";
                location.href = "Form__contactBoardMember.php";//, //acNormal, //, //, //acFormAdd;
                //Forms![_contactBoardMember].AllowFilters = //False;
            //Else;
                    //If //CancelUpdate //Then;
                        //CancelUpdate = //False;
                        //Exit //Sub;
                    //Else;
                        //Me.CoOrd = //HoldCoord;
                        //Me.CoordName = //HoldCoordName;
                        //MyFilter = "[Contact Data].[Common Contact ID]=" //& //Me.CoOrd;
                        location.href = "Form__contactBoardMember.php";//, //acNormal, //, //MyFilter;
                        //Forms![_contactBoardMember].AllowFilters = //False;
                        //Forms![_contactBoardMember].AllowAdditions = //False;
                    //End //If;
            //End //If;
;
}
function Command79_Click(){
//End //Sub;
;
//Private //Sub //CreateTemplates_Click();
//On //Error //GoTo //Err_CreateButton_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
;
//Exit_CreateButton_Click:;
    //Exit //Sub;
;
//Err_CreateButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_CreateButton_Click;
    ;
}
function Command85_Click(){
//End //Sub;
;
//Private //Sub //Command93_Click();
location.href = "Form_drugSearch.php";;
}
function Exempt_Click(){
//On //Error //GoTo //Err_ExemptCodes_Click;
;
    var stDocName = '';
    stDocName;
   ;
        //DoCmd.OpenQuery stDocName;
  ;
;
//Exit_ExemptCodes_Click:;
    //Exit //Sub;
;
//Err_ExemptCodes_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ExemptCodes_Click;
}
function Expedite_Click(){
//On //Error //GoTo //Err_ExpediteCodes_Click;
;
    var stDocName = '';
    stDocName;
    ;
        //DoCmd.OpenQuery stDocName;
   ;
;
//Exit_ExpediteCodes_Click:;
    //Exit //Sub;
;
//Err_ExpediteCodes_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ExpediteCodes_Click;
}
function Form_Load(){
//Me.MaxIRBs = //Me.MaxIRBs.ItemData(0);
//Me.ActualIRBs = //Me.ActualIRBs.ItemData(0);
//Me.MAxUsers = //Me.MAxUsers.ItemData(0);
//Me.ActualUsers = //Me.ActualUsers.ItemData(0);
//SearchContacts = //True;
//If //Me.MaxIRBs = //1 //Then //Me.ChangeIRB.Enabled = //False;
//If //Not //CyberIRB_Flag //Then //cmdCyberSync.Visible = //False;
}
function IRBActions_Click(){
//On //Error //GoTo //Error_action;
;
        //DoCmd.OpenQuery "IRB Actions"//, //acViewNormal;
;
;
;
//Exit_action:;
//Exit //Sub;
;
//Error_action:;
//MsgBox //("Error //in //Action //Code //Query.  //Close //and //restart //PRO_IRB"), //vbCritical;
//Resume //Exit_action;
}
function CoordButton_Click(){
//'Coordinator Maintenance
    //SearchContacts = //False;
 ;
            location.href = "Form__selectCoord.php";//, //acNormal, //, //, //, //acDialog;
            //'If CancelUpdate Then Exit Sub
            //If //NewApplication //Then;
                //NewApplication = //False;
                //HoldCoordName = " ";
                location.href = "Form__contact Coordinator.php";//, //acNormal, //, //, //acFormAdd;
                //Forms![_contact //coordinator].AllowFilters = //False;
            //Else;
                    //If //CancelUpdate //Then;
                        //CancelUpdate = //False;
                        //Exit //Sub;
                    //Else;
                        //Me.CoOrd = //HoldCoord;
                        //Me.CoordName = //HoldCoordName;
                        //'myfilter = "[Contact Data].[Common Contact ID]=" & [Forms]![CodesandTables]![CoOrd]
                        //MyFilter = "[Contact Data].[Common Contact ID]=" //& //Me.CoOrd;
                        location.href = "Form__contact coordinator.php";//, //acNormal, //, //MyFilter;
                        //Forms![_contact //coordinator].AllowFilters = //False;
                        //Forms![_contact //coordinator].AllowAdditions = //False;
                    //End //If;
            //End //If;
            ;
}
function Form_Activate(){
//DoCmd.Restore;
;
}
function Form_Close(){
//DoCmd.Close //acReport, "StudyCPASAEHistory"//, //acSaveNo;
//SearchContacts = //False;
}
function Ignore_Click(){
//DoCmd.Close;
;
}
function IRBNamesButton_Click(){
//On //Error //GoTo //Err_IRBNamesButton_Click;
    var stDocName = '';
    stDocName;
    ;
        location.href = "Form_"+stDocName+".php";//, //acFormDS;
   ;
;
//Exit_IRBNamesButton_Click:;
    //Exit //Sub;
;
//Err_IRBNamesButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_IRBNamesButton_Click;
;
;
;
}
function PIButton_Click(){
//'P.I.
            //'HoldPIName = " "
            //SearchContacts = //False;
            location.href = "Form__selectPI.php";//, //acNormal, //, //, //, //acDialog;
          ;
            //If //NewApplication //Then;
                //NewApplication = //False;
                //HoldPIName = " ";
                location.href = "Form__contact PI.php";//, //acNormal, //, //, //acFormAdd;
                //Forms![_Contact //PI].AllowFilters = //False;
            //Else;
                //If //CancelUpdate //Then;
                    //CancelUpdate = //False;
                    //Exit //Sub;
                //Else;
                    //If //Not //IsNull(HoldPI) //Then;
                    //Me.PI = //HoldPI;
                    //Me.PIName = //HoldPIName;
                    ;
                    ;
                        //MyFilter = "[Common Contact ID]= " //& //[Forms]![CodesandTables]![PI];
                    //End //If;
                    //'MyFilter = "[Contact Data].[Common Contact ID]=" & [Forms]![CodesandTables]![PI]
                    location.href = "Form__contact PI.php";//, //acNormal, //, //MyFilter;
                    //Forms![_Contact //PI].AllowFilters = //False;
                    //Forms![_Contact //PI].AllowAdditions = //False;
                //End //If;
            //End //If;
           ;
}
function Command3_Click(){
//End //Sub;
;
//Private //Sub //PISpecialtyButton_Click();
//On //Error //GoTo //Err_PISpecialtyButton_Click;
;
    var stDocName = '';
;
    stDocName;
    ;
        location.href = "Form_"+stDocName+".php";//, //acFormDS;
 ;
//Exit_PISpecialtyButton_Click:;
    //Exit //Sub;
;
//Err_PISpecialtyButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_PISpecialtyButton_Click;
    ;
}
function sponsorButton_Click(){
//'            Sponsor maintenance
       //SearchContacts = //False;
            location.href = "Form__selectSponsors.php";//, //acNormal, //, //, //, //acDialog;
            //If //NewApplication //Then;
                //NewApplication = //False;
                //HoldSponsorName = " ";
                location.href = "Form__contact Sponsor.php";//, //acNormal, //, //, //acFormAdd;
                //Forms![_Contact //Sponsor].AllowFilters = //False;
            //Else;
                    //If //CancelUpdate //Then;
                        //CancelUpdate = //False;
                        //Exit //Sub;
                    //Else;
                        //Me.Sponsor = //HoldSponsor;
                        //Me.SponsorName = //HoldSponsorName;
                        //MyFilter = "[Contact Data].[Common Contact ID]=" //& //[Forms]![CodesandTables]![Sponsor];
                        location.href = "Form__contact Sponsor.php";//, //acNormal, //, //MyFilter;
                        //Forms![_Contact //Sponsor].AllowFilters = //False;
                        //Forms![_Contact //Sponsor].AllowAdditions = //False;
                    //End //If;
            //End //If;
           ;
}
function Templates_Click(){
//On //Error //GoTo //Err_Templates_Click;
;
    var stDocName = '';
    stDocName;
    ;
        location.href = "Form_"+stDocName+".php";;
    ;
;
;
//Exit_Templates_Click:;
    //Exit //Sub;
;
//Err_Templates_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Templates_Click;
}
function TransactionButton_Click(){
//On //Error //GoTo //err_Trans;
//TransactionSwitch = //True;
//CancelUpdate = //False;
      ;
;
//Forms![newmenu].Status = "OPEN";
//HoldStatus = //Forms![newmenu].Status;
//HoldLocation = //Forms![newmenu].Location;
//searchirbs = //False;
location.href = "Form__SelectStudy.php";//, //, //, //, //, //acDialog;
//TransactionSwitch = //False;
;
//If //CancelUpdate = //True //Then;
    //Exit //Sub;
//Else;
//'Starts here builging transaction records for the selected study in HoldStudy
//'DoCmd.Hourglass True
//Forms![newmenu]![study] = //HoldStudy;
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
//End //If;
//exit_Trans:;
//Exit //Sub;
;
//err_Trans:;
//MsgBox "Error Previewing Trans History";
//Resume //exit_Trans;
  ;
                ;
                ;
}
function Command26_Click(){
//On //Error //GoTo //Err_Command26_Click;
//TransactionSwitch = //True;
//CancelUpdate = //False;
;
        ;
//Forms![newmenu].[ReviewMonth] = "***";
//Forms![newmenu].RenewalMonth = "***";
//Forms![newmenu].RenewalCycle = "***";
//Forms![newmenu].Status = "OPEN";
//HoldStatus = //Forms![newmenu].Status;
//HoldLocation = //Forms![newmenu].Location;
//searchirbs = //False;
location.href = "Form__SelectStudy.php";//, //, //, //, //, //acDialog;
//TransactionSwitch = //False;
;
//If //CancelUpdate = //True //Then;
    //Exit //Sub;
//Else;
    //MyFilter = "[irb#]= " //& "'" //& //HoldStudy //& "'";
    var stDocName = '';
;
    stDocName;
    //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
    //DoCmd.RunCommand //acCmdZoom100;
//End //If;
;
//Exit_Command26_Click:;
    //Exit //Sub;
;
//Err_Command26_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command26_Click;
    ;
}
function Command28_Click(){
//On //Error //GoTo //Err_Command28_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
;
//Exit_Command28_Click:;
    //Exit //Sub;
;
//Err_Command28_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command28_Click;
    ;
}
function EventTypes_Click(){
//On //Error //GoTo //Err_EventTypes_Click;
;
  ;
    ;
        //DoCmd.OpenQuery "SAEEventTypes"//, //acViewNormal;
   ;
        ;
    ;
//Exit_EventTypes_Click:;
    //Exit //Sub;
;
//Err_EventTypes_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_EventTypes_Click;
    ;
}
function Command33_Click(){
//On //Error //GoTo //Err_Command33_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
;
//Exit_Command33_Click:;
    //Exit //Sub;
;
//Err_Command33_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command33_Click;
    ;
}
function cpatypes_Click(){
//On //Error //GoTo //Err_cpatypes_Click;
;
;
    var stDocName = '';
;
    stDocName;
    ;
        location.href = "Form_"+stDocName+".php";;
        //'DoCmd.OpenQuery stDocName disabled until I can desing a form that only allows valid changes
        //'and doesn't allow change to Final report or Progress report and others that may cause a problem
        //'
  ;
;
//Exit_cpatypes_Click:;
    //Exit //Sub;
;
//Err_cpatypes_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cpatypes_Click;
    ;
}
function StatusCodes_Click(){
//On //Error //GoTo //Err_StatusCodes_Click;
;
    var stDocName = '';
    stDocName;
   ;
        //DoCmd.OpenQuery stDocName;
  ;
//Exit_StatusCodes_Click:;
    //Exit //Sub;
;
//Err_StatusCodes_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_StatusCodes_Click;
    ;
}
function ButUsers_Click(){
//On //Error //GoTo //Err_ButUsers_Click;
//If //Demo //Then;
    //MsgBox "Not Available in Trial Version";
    //Exit //Sub;
//End //If;
;
;
;
    var stDocName = '';
    stDocName;
    ;
        location.href = "Form_"+stDocName+".php";//, //acFormDS;
    ;
;
   ;
;
//Exit_ButUsers_Click:;
    //Exit //Sub;
;
//Err_ButUsers_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ButUsers_Click;
    ;
}
function Words_Click(){
location.href = "Form_MaintainLetterParagraphs.php";;
}
function cmdVulPop_Click(){
//On //Error //GoTo //Err_cmdVulPop_Click;
;
   var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
//Exit_cmdVulPop_Click:;
    //Exit //Sub;
;
//Err_cmdVulPop_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdVulPop_Click;
    ;
}
function cmdSites_Click(){
//On //Error //GoTo //Err_cmdSites_Click;
 var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
//Exit_cmdSites_Click:;
    //Exit //Sub;
;
//Err_cmdSites_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdSites_Click;
    ;
}
function cmdClass_Click(){
//On //Error //GoTo //Err_cmdClass_Click;
;
    var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
//Exit_cmdClass_Click:;
    //Exit //Sub;
;
//Err_cmdClass_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdClass_Click;
    ;
}
function cmdDept_Click(){
//On //Error //GoTo //Err_cmdDept_Click;
;
    var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
//Exit_cmdDept_Click:;
    //Exit //Sub;
;
//Err_cmdDept_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdDept_Click;
    ;
}
function cmdDrugs_Click(){
//On //Error //GoTo //Err_cmdDrugs_Click;
;
    var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
;
//Exit_cmdDrugs_Click:;
    //Exit //Sub;
;
//Err_cmdDrugs_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdDrugs_Click;
    ;
}
function cmdInd_Click(){
//On //Error //GoTo //Err_cmdInd_Click;
;
   var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
//Exit_cmdInd_Click:;
    //Exit //Sub;
;
//Err_cmdInd_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdInd_Click;
    ;
}
function cmdRisks_Click(){
//On //Error //GoTo //Err_cmdRisks_Click;
;
  var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
;
//Exit_cmdRisks_Click:;
    //Exit //Sub;
;
//Err_cmdRisks_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdRisks_Click;
    ;
}
function cmdGrant_Click(){
//On //Error //GoTo //Err_cmdGrant_Click;
 var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
//Exit_cmdGrant_Click:;
    //Exit //Sub;
;
//Err_cmdGrant_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdGrant_Click;
    ;
}
function cmdChild_Click(){
//On //Error //GoTo //Err_cmdChild_Click;
;
    var stDocName = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
//Exit_cmdChild_Click:;
    //Exit //Sub;
;
//Err_cmdChild_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdChild_Click;
    ;
}
function CmdReviewers_Click(){
//On //Error //GoTo //Err_cmdReviewers_Click;
;
 //'Reviewer copied from PI
            ;
            //SearchContacts = //False;
            location.href = "Form__SelectOther.php";//, //acNormal, //, //, //, //acDialog;
          ;
            //If //NewApplication //Then;
                //NewApplication = //False;
                //HoldOtherName = " ";
                location.href = "Form__ContactOther.php";//, //acNormal, //, //, //acFormAdd;
                //Forms![_ContactOther].AllowFilters = //False;
            //Else;
                //If //CancelUpdate //Then;
                    //CancelUpdate = //False;
                    //Exit //Sub;
                //Else;
                    //Me.Other = //HoldOther;
                    //Me.OtherName = //HoldOtherName;
                    //MyFilter = "[Common Contact ID]= " //& //[Forms]![CodesandTables]![Other];
                    location.href = "Form__ContactOther.php";//, //acNormal, //, //MyFilter;
                    //Forms![_ContactOther].AllowFilters = //False;
                    //Forms![_ContactOther].AllowAdditions = //False;
                //End //If;
            //End //If;
           ;
//Exit_cmdReviewers_Click:;
    //Exit //Sub;
;
//Err_cmdReviewers_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdReviewers_Click;
    ;
}
function bCreateTemplates_Click(){
//On //Error //GoTo //Err_bCreateTemplates_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";;
;
//Exit_bCreateTemplates_Click:;
    //Exit //Sub;
;
//Err_bCreateTemplates_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_bCreateTemplates_Click;
    ;
}
function cmdCyberSync_click(){
//''Terry 6/6/2009
//SyncAllDataWithCyber;
}
function cmdCyberSyncOLD_Click(){//'June's updates
;
//'June's updates
;
//'Terry 6/6/2009 This sub was moved to ProToCyber
;
var strSQL = '';
var strConnectionString = '';
//could initialize cmd here
//Set cmd;
;
;
//'upload the irb - research database table to the server
;
//10    //DoCmd.SetWarnings //False;
//'upload the irb - research database table to the server
//'if the table is there, delete it first
//'then run the sproc that will load the ProIRBData table
//'  On Error Resume Next
//'    DoCmd.RunSQL "drop table [" & g_CyberServerConnection & "].[IRB - Research  Database1_" & gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") & "]"
;
//'First, let's delete the old cyber tables
;
//20    //cmd.ActiveConnection = //gclsCyberIRB.CyberConnection;
//30    //cmd.CommandText = "PIRB_" //& //gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") //& "_ReSyncDeleteTables";
//40    //cmd.CommandType = //adCmdStoredProc;
//'50    cmd.Parameters("@varCustNum") = gclsCyberIRB.mcolCyberConfigAttributes("CustNum")
//'60    cmd.Parameters("@varCustPrefix") = gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix")
//'
//70    //cmd.Execute;
;
//'On Error GoTo 0
//80    //Debug.Print //Now();
//TestConnectionToDB;
//90        strSQL = "SELECT * INTO [" && g_CyberServerConnection && "].[Pro_ " && gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") && "_IRB - Research  Database1" && "] FROM [IRB - Research  Database1] ;";
    //Debug.Print strSQL;
//100       //DoCmd.RunSQL strSQL;
;
//110   //Debug.Print //Now();
//120       strSQL = "SELECT * INTO [" && g_CyberServerConnection && "].[Pro_ " && gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") && "_Contact Data1_" && "] FROM [Contact Data1] ;";
//130       //DoCmd.RunSQL strSQL;
;
//140   //Debug.Print //Now();
//150       strSQL = "SELECT * INTO [" && g_CyberServerConnection && "].[Pro_ " && gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") && "_CoInvStudy1_" && "] FROM [CoInvStudy1] ;";
//160       //DoCmd.RunSQL strSQL;
;
//170   //Debug.Print //Now();
;
//180       strSQL = "SELECT * INTO [" && g_CyberServerConnection && "].[Pro_ " && gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") && "_ReviewerStudy1_" && "] FROM [ReviewerStudy1] ;";
//190       //DoCmd.RunSQL strSQL;
;
//200   //Debug.Print //Now();
 //Set cmd;
;
      //' run the  stored procedure  - PIRB_UNL_ReSyncProIRBToCyber
//210   //cmd.ActiveConnection = //gclsCyberIRB.CyberConnection;
//220   //cmd.CommandText = "PIRB_" //& //gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") //& "_ReSyncProIRBToCyber";
//230   //cmd.CommandType = //adCmdStoredProc;
//'      '****************************
;
//240   //cmd.Parameters("@varCustNum") = //gclsCyberIRB.mcolCyberConfigAttributes("CustNum");
//250   //cmd.Parameters("@varCustPrefix") = //gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix");
//255   //cmd.Parameters("@varTableName") = //gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix");
//'
//260   //cmd.Execute;
;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Active();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>



    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>




    <label id=Label1 value='Administration' style='position:absolute; left:378; top:15; width:183; height:36; visibility:'>Administration</label>
    <input type='button' id=PIButton value='Investigators' style='position:absolute; left:42; top:118; width:238; height:31' onclick='PIButton_Click();'></input>
    <input type='button' id=CoordButton value='Coordinators' style='position:absolute; left:42; top:190; width:238; height:31' onclick='CoordButton_Click();'></input>
    <input type='button' id=sponsorButton value='Sponsors' style='position:absolute; left:42; top:154; width:238; height:31' onclick='sponsorButton_Click();'></input>
    <input type='button' id=Ignore value='&Return' style='position:absolute; left:786; top:588; width:115; height:42' onclick='Ignore_Click();'></input>
    <input type='button' id=Command17 value='IRB Meeting Dates' style='position:absolute; left:618; top:226; width:238; height:31' onclick='Command17_Click();'></input>
    <input type='button' id=IRBNamesButton value='IRB Names' style='position:absolute; left:42; top:334; width:238; height:31' onclick='IRBNamesButton_Click();'></input>
    <input type='button' id=Words value='Letter Wording' style='position:absolute; left:618; top:298; width:238; height:31' onclick='Words_Click();'></input>
    <input type='button' id=TransactionButton value='Study Trans History' style='position:absolute; left:42; top:456; width:238; height:31' onclick='TransactionButton_Click();'></input>
    <input type='button' id=Command26 value='Study CPA / SAE History' style='position:absolute; left:42; top:420; width:238; height:31' onclick='Command26_Click();'></input>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:133; top:24; width:156; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='button' id=EventTypes value='SAE Event Types' style='position:absolute; left:618; top:370; width:238; height:31' onclick='EventTypes_Click();'></input>
    <input type='button' id=IRBActions value='IRB Action Codes' style='position:absolute; left:618; top:262; width:238; height:31' onclick='IRBActions_Click();'></input>
    <input type='button' id=cpatypes value='CPA Types' style='position:absolute; left:618; top:154; width:238; height:31' onclick='cpatypes_Click();'></input>
    <input type='button' id=StatusCodes value='Study Status Codes' style='position:absolute; left:618; top:406; width:238; height:31' onclick='StatusCodes_Click();'></input>
    <input type='button' id=ChangeIRB value='Change Study IRB' style='position:absolute; left:42; top:492; width:238; height:31' onclick='ChangeIRB_Click();'></input>
    <input type='button' id=ButUsers value='Users' style='position:absolute; left:42; top:298; width:238; height:31' onclick='ButUsers_Click();'></input>
    <input type='button' id=Closings value='Letter Closings' style='position:absolute; left:618; top:334; width:238; height:31' onclick='Closings_Click();'></input>
    <input type='button' id=Command43 value='Consent Requirements' style='position:absolute; left:618; top:118; width:238; height:31' onclick='Command43_Click();'></input>
    <input type='button' id=PISpecialtyButton value='Investigator Specialties' style='position:absolute; left:618; top:190; width:238; height:31' onclick='PISpecialtyButton_Click();'></input>
    <input type='button' id=Command53 value='Board Members' style='position:absolute; left:42; top:226; width:238; height:31' onclick='Command53_Click();'></input>
    <input type='button' id=Expedite value='Expedite Category(ies)' style='position:absolute; left:333; top:372; width:238; height:31' onclick='Expedite_Click();'></input>
    <input type='button' id=Exempt value='Exempt Category(ies)' style='position:absolute; left:333; top:336; width:238; height:31' onclick='Exempt_Click();'></input>
    <input type='button' id=cmdVulPop value='Vulnerable Population(s)' style='position:absolute; left:336; top:552; width:238; height:31' onclick='cmdVulPop_Click();'></input>
    <input type='button' id=cmdSites value='Site(s)' style='position:absolute; left:333; top:228; width:238; height:31' onclick='cmdSites_Click();'></input>
    <input type='button' id=cmdBenefits value='Benefits' style='position:absolute; left:333; top:264; width:238; height:31' onclick='cmdBenefits_Click();'></input>
    <input type='button' id=cmdClass value='Classification(s)' style='position:absolute; left:333; top:120; width:238; height:34' onclick='cmdClass_Click();'></input>
    <input type='button' id=cmdDept value='Dept/Group(s)' style='position:absolute; left:333; top:193; width:238; height:31' onclick='cmdDept_Click();'></input>
    <input type='button' id=cmdDrugs value='Drug(s)/Specific Devices' style='position:absolute; left:333; top:300; width:238; height:31' onclick='cmdDrugs_Click();'></input>
    <input type='button' id=cmdInd value='IND/Device Type(s)' style='position:absolute; left:333; top:444; width:238; height:31' onclick='cmdInd_Click();'></input>
    <input type='button' id=cmdRisks value='Risk(s)' style='position:absolute; left:333; top:480; width:238; height:31' onclick='cmdRisks_Click();'></input>
    <input type='button' id=cmdGrant value='Grant/Project(s)' style='position:absolute; left:333; top:406; width:238; height:31' onclick='cmdGrant_Click();'></input>
    <input type='button' id=cmdChild value='Child Category(ies)' style='position:absolute; left:336; top:516; width:238; height:31' onclick='cmdChild_Click();'></input>
    <input type='button' id=cmdReviewers value='Other Contacts' style='position:absolute; left:42; top:262; width:238; height:31' onclick='cmdReviewers_Click();'></input>
    <label id=Label67 value='Entities' style='position:absolute; left:111; top:76; width:106; height:36; visibility:'>Entities</label>
    <label id=Label68 value='Study Groupings' style='position:absolute; left:340; top:76; width:226; height:36; visibility:'>Study Groupings</label>

    <label id=Label70 value='Codes and Tables' style='position:absolute; left:613; top:76; width:240; height:36; visibility:'>Codes and Tables</label>
    <label id=Label71 value='Reports and Other' style='position:absolute; left:37; top:388; width:244; height:34; visibility:'>Reports and Other</label>



    <input type='button' id=Templates value='Template Locations' style='position:absolute; left:618; top:510; width:238; height:37' onclick='Templates_Click();'></input>
    <input type='button' id=bCreateTemplates value='Create/ ModifyTemplates' style='position:absolute; left:619; top:546; width:238; height:37' onclick='bCreateTemplates_Click();'></input>
    <label id=Label83 value='Templates' style='position:absolute; left:610; top:471; width:247; height:36; visibility:'>Templates</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:49; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <input type='button' id=CMD_DIV value='Divisions' style='position:absolute; left:333; top:159; width:238; height:31' onclick='CMD_DIV_Click();'></input>
    <input type='button' id=cmdAgendaCategories value='Agenda Categories' style='position:absolute; left:618; top:441; width:238; height:31' onclick='cmdAgendaCategories_Click();'></input>
    <input type='button' id=Command93 value='Drug/Device Related' style='position:absolute; left:42; top:526; width:238; height:31' onclick='Command93_Click();'></input>
    <input type='button' id=cmdCyberSync value='Sync with CyberIRB' style='position:absolute; left:30; top:576; width:246; height:39'></input>


    <!--     <input type='text' id=PI value='empty' style='position:absolute; left:402; top:6; width:234; height:30'></input> -->
    <!--     <input type='text' id=SponsorName value='empty' style='position:absolute; left:288; top:30; width:258; height:48'></input> -->

    <label id=Label12 value='Text11:' style='position:absolute; left:144; top:30; width:61; height:24; visibility:'>Text11:</label>
    <!--     <input type='text' id=Coord value='empty' style='position:absolute; left:144; top:90; width:222; height:42'></input> -->

    <label id=Label14 value='Text9:' style='position:absolute; left:0; top:90; width:52; height:24; visibility:'>Text9:</label>
    <!--     <input type='text' id=StudyNumber value='empty' style='position:absolute; left:162; top:24; width:156; height:36'></input> -->

    <label id=Label25 value='Text24:' style='position:absolute; left:72; top:24; width:61; height:24; visibility:'>Text24:</label>
    <!--     <select name='MaxIRBs' Size='4' style='position:absolute; left:300; top:42; width:90; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [Copyright1].[MaxUsers] FROM Copyright1");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MaxUsers']."'".">".$row['MaxUsers']."</Option>";
        }
    ?>
    </select> -->

    <label id=MAxIRBS_Label value='MAxIRBS' style='position:absolute; left:210; top:42; width:70; height:24; visibility:'>MAxIRBS</label>
    <!--     <select name='CurrentUserName_Label' Size='5' style='position:absolute; left:162; top:6; width:79; height:24'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [Users1].[ID], [Users1].[CurrentUserName] FROM Users1");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['ID']."'".">".$row['ID']."</Option>";
        }
    ?>
    </select> -->
    <!--     <input type='text' id=PIName value='empty' style='position:absolute; left:516; top:90; width:330; height:30'></input> -->

    <label id=Label8 value='Text7:' style='position:absolute; left:0; top:66; width:52; height:24; visibility:'>Text7:</label>
    <!--     <input type='text' id=Sponsor value='empty' style='position:absolute; left:144; top:96; width:222; height:42'></input> -->

    <label id=Label10 value='Text9:' style='position:absolute; left:0; top:96; width:52; height:24; visibility:'>Text9:</label>
    <!--     <input type='text' id=CoordName value='empty' style='position:absolute; left:54; top:72; width:258; height:48'></input> -->

    <label id=Label16 value='Text11:' style='position:absolute; left:120; top:114; width:61; height:24; visibility:'>Text11:</label>
    <!--     <select name='ActualIRBs' Size='9' style='position:absolute; left:288; top:72; width:96; height:18'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [Users1].[ID], [Users1].[CurrentUserName] FROM Users1");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['ID']."'".">".$row['ID']."</Option>";
        }
    ?>
    </select> -->

    <label id=ActualIRBS_Label value='ActualIRBS' style='position:absolute; left:252; top:72; width:70; height:24; visibility:'>ActualIRBS</label>
    <label id=MaxUsers_Label value='MaxUsers' style='position:absolute; left:78; top:78; width:79; height:24; visibility:'>MaxUsers</label>
    <!--     <select name='MAxUsers' Size='10' style='position:absolute; left:42; top:24; width:78; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [IRB Meeting Schedule1].[Irb Code] FROM [IRB Meeting Schedule1]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Irb Code']."'".">".$row['Irb Code']."</Option>";
        }
    ?>
    </select> -->
    <!--     <input type='text' id=OtherName value='empty' style='position:absolute; left:516; top:24; width:330; height:30'></input> -->

    <label id=Label77 value='Text7:' style='position:absolute; left:0; top:0; width:52; height:24; visibility:'>Text7:</label>
    <!--     <input type='text' id=Other value='empty' style='position:absolute; left:492; top:0; width:234; height:30'></input> -->
    <!--     <input type='button' id=AgendaCat value='InvisibleAgenda Categories' style='position:absolute; left:561; top:21; width:238; height:30' onclick='AgendaCat_Click();'></input> -->
    <input type='text' id=TBNewIRB value='empty' style='position:absolute; left:414; top:102; width:72; height:12'></input>

    <label id=Label88 value='Text87:' style='position:absolute; left:270; top:102; width:61; height:24; visibility:'>Text87:</label>
    <input type='text' id=TBStudyT value='empty' style='position:absolute; left:594; top:66; width:120; height:12'></input>
    <input type='text' id=StudyIRBCode value='empty' style='position:absolute; left:348; top:60; width:204; height:48'></input>

    <label id=Label92 value='Text91:' style='position:absolute; left:204; top:60; width:61; height:24; visibility:'>Text91:</label>
  </body>
</html>
