<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function bPrepareAgenda_Click(){
//On //Error //GoTo //Err_bPrepareAgenda;
;
//DoCmd.SetWarnings //False;
;
//If //DataChanged = //False //Then //GoTo //Exit_Return;
;
//If //ItemLineNumberChangeFlag = //True //Then;
        //'ItemLineNumberChangeFlag = false  commented out for 7.3
        ;
       //If //YesNo("Do //you //wish //to //save //these //line //numbers?  //If //you //answer //'No' only the line numbers you just changed for this date will be undone. ") Then
            //UseZeros = //False;
                ;
            //Call //UpdateAgendaRecordsWithLineNumbers(UseZeros);
        //Else;
            //ItemLineNumberChangeFlag = //False;
            //'sets line numbers back to increments of 10
            //'Call Resequence
            //DoCmd.OpenQuery "UpdateLineNumberFlagInMeetingDates";
        //End //If;
//Else;
        //'UseZeros = True
        //'Call UpdateAgendaRecordsWithLineNumbers(UseZeros)
            ;
;
//End //If;
         ;
;
//Exit_Return:;
//DoCmd.SetWarnings //True;
    //RedisplayAgenda = //False;
    //NewHoldMeetingKeepTogether = //Me.chkKeepTogether;
;
;
//Continue1:;
//StudySelected = //False;
 ;
//Call //fCheckNewMembers;
    //MyFilter = "MeetingDate = " //& //Me.MeetingDate;
;
    var stDocName = '';
    stDocName;
    //'Forms![_display agenda records].Visible = False
    location.href = "Form_"+stDocName+".php";//, //acNormal;
  ;
    ;
    ;
//Exit_bPrepareAgenda:;
    //Exit //Sub;
;
//Err_bPrepareAgenda:;
//If //Err.Number = //2501 //Then;
    ;
    //MsgBox "No Agenda Items for Printing";
    //Resume //Exit_bPrepareAgenda;
//Else;
    //MsgBox //Err.Number //& "   Error in Printing Agenda";
    //Resume //Exit_bPrepareAgenda;
    ;
//End //If;
;
;
}
function chkKeepTogether_Click(){
//On //Error //GoTo //err_chk;
;
//If //DataChanged //Then;
    //MsgBox "You cannot change the Keep Together option once you have changed line numbers or other data.  Use the Reset feature if you wish to Keep the Studies Together";
    //Exit //Sub;
//Else;
//End //If;
//ItemLineNumberChangeFlag = //True;
//DataChanged = //True;
//Call //ReNumberClickEvent;
//could initialize X here
;
//DoCmd.SetWarnings //False;
//'MsgBox Me.chkKeepTogether
//DoCmd.OpenQuery "SaveKeepTogether";
;
//'docmd.OpenQuery "qryUpdateMeetingDateWithKeepTogether_sELECT", acViewNormal, acReadOnly
;
//could initialize reopen here
reopen;
;
//NewHoldMeetingKeepTogether = //Forms![_display //agenda //records]![chkKeepTogether];
//DoCmd.OpenQuery "qryUpdateMeetingDateWithKeepTogether";
;
//If //Me.chkKeepTogether //Then;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.Line.Locked = //True;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.[Agenda //Category].Locked = //True;
        ;
//Else;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.Line.Locked = //False;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.[Agenda //Category].Locked = //False;
      ;
//End //If;
;
//Me.MeetingKey = //Me.MeetingDate.Column(1);
//Forms![newmenu]![MeetingKey] = //Me.MeetingKey;
//Forms![newmenu]![NewMenuMeetingDate] = //Me.MeetingDate.Column(0);
//[Forms]![newmenu]![TxtPostMeetingDate] = //Me.MeetingDate.Column(0);
//HoldListIndex = //Me.MeetingDate.ListIndex;
;
//RedisplayAgenda = //True;
//DoCmd.Close;
//DoCmd.Close //acForm, "_display agenda records sub form";
//exit_Chk:;
//Exit //Sub;
//err_chk:;
//MsgBox //Err.Description;
//Resume //exit_Chk;
;
;
}
function Command108_Click(){
//End //Sub;
;
//Private //Sub //CmdRenumber_Click();
//DataChange = //False;
;
//Call //ReNumberClickEvent;
 ;
}
function Command81_Click(){
//If //StudySelected //Then;
    //StudySelected = //False;
    //If //YesNo("This //Report //prints //Summary //information //for //ALL //Studies. //Do //you //wish //to //continue?") //Then;
        //GoTo //Continue1;
    //Else;
        //Exit //Sub;
    //End //If;
//End //If;
//Continue1:;
//StudySelected = //False;
//On //Error //GoTo //Err_Print_summary_button_Click;
;
    var stDocName = '';
//MyFilter = "MeetingDate = " //& //Me.MeetingDate;
    stDocName;
        ;
    ;
    //DoCmd.OpenReport //stDocName, //acViewPreview;
   ;
//DoCmd.Close //acForm, "_display agenda records";
//Forms!newmenu.Visible = //False;
;
    ;
//Exit_Err_Print_summary_button_Click:;
    //Exit //Sub;
;
//Err_Print_summary_button_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Err_Print_summary_button_Click;
    ;
}
function Form_Activate(){
//On //Error //Resume //Next;
//DoCmd.Restore;
}
function Form_Close(){
//'SaveHoldListIndex = HoldListIndex
//StudySelected = //False;
//DoCmd.Close //acReport, "AgendaWDetails, acSaveNo";
//DoCmd.Close //acReport, "AgendaSummary"//, //acSaveNo;
//FromIRB = //False;
//LineChangeFlag = //False;
//If //isFormLoaded("_IRB //Input //Form") //Then   //'FOR CHLA
    //Forms![_irb //input //form].Visible = //True;
//End //If;
;
;
}
function Form_Load(){
//On //Error //GoTo //Err_Load;
//could initialize X here
;
//Me.chkKeepTogether = //NewHoldMeetingKeepTogether;
//DataChanged = //False;
//RedisplayAgenda = //False;
;
//Forms![_display //agenda //records]![TextExpansion2] = //Me.MeetingDate.Column(4);
;
//If //Forms![_display //agenda //records]![TextExpansion2] = "YES" //Then;
    //ItemLineNumberChangeFlag = //True;
//Else;
    //ItemLineNumberChangeFlag = //False;
//End //If;
;
;
//Me.MeetingDate.Requery;
;
//If //FromCheck = "noitems" //Then;
    //If //SaveHoldListIndex = //0 //Then;
        //'MsgBox "This one"
        //RedisplayAgenda = //True;
        //DoCmd.Close;
        //DoCmd.Close //acForm, "_display agenda records sub form";
        //Exit //Sub;
    //End //If;
    //HoldListIndex = //SaveHoldListIndex;
//End //If;
;
//Me.MeetingDate.Selected(HoldListIndex) = //True;
//Me.MeetingDate = //Me.MeetingDate.ItemData(HoldListIndex);
//Me.MeetingKey = //Me.MeetingDate.Column(1);
//Forms![newmenu]![MeetingKey] = //Me.MeetingKey;
//'MsgBox Me.MeetingKey
//AgendaDetail = //False;
//If //Not //UserLastName //Like "*Admin*" //Then;
    //Me.ButDelete.Enabled = //False;
//End //If;
;
//If //Forms![newmenu]![HoldReadOnly] //Then;
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.ReassignButton.Enabled = //False;
    //Me.ReAssignALL.Enabled = //False;
    //Me.ReviewDetailForOne.Enabled = //False;
//End //If;
;
//If //Me.chkKeepTogether //Then;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.Line.Locked = //True;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.[Agenda //Category].Locked = //True;
//Else;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.Line.Locked = //False;
        //Forms![_display //agenda //records]![_agenda //sub //form].Form.[Agenda //Category].Locked = //False;
//End //If;
;
;
//Exit_Load:;
    //Exit //Sub;
;
//Err_Load:;
;
;
;
//If //Err.Description //Like "*lock" //And //Me.MeetingDate //< //Now //Then;
    //'MsgBox "There are no Agenda Items for this meeting date." '12/28 commented out
    //HoldListIndex = //DefaultIndex;
    ;
    ;
    //Resume //Exit_Load;
//Else;
    //MsgBox //Err.Number //& " " //& //Err.Description;
     //Resume //Exit_Load;
    ;
//End //If;
;
}
function GetHelp1_Click(){
//could initialize returnvalue here
returnvalue;
//Exit_GetHelp1_Click:;
    //Exit //Sub;
;
}
function Form_Open(Cancel){
//could initialize X here
//'DoCmd.Close acForm, "CheckAgendaRecordStatus"
//On //Error //Resume //Next;
}
function GetHelp_Click(){
//could initialize returnvalue here
returnvalue;
//Exit_GetHelp1_Click:;
    //Exit //Sub;
}
function MeetingDate_AfterUpdate(){
//On //Error //Resume //Next;
//SaveHoldListIndex = //HoldListIndex;
//If //HoldListIndex = //Me.MeetingDate.ListIndex //Then;
    ;
    //Exit //Sub;
//Else;
  ;
  //If //IsNull(Me.MeetingDate) //Then;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "SaveKeepTogether";
;
        //could initialize reopen here
        reopen;
;
        //NewHoldMeetingKeepTogether = //Forms![_display //agenda //records]![chkKeepTogether];
        //'MsgBox Me.chkKeepTogether
        //DoCmd.OpenQuery "qryUpdateMeetingDateWithKeepTogether";
    //Else;
        //Me.MeetingKey = //Me.MeetingDate.Column(1);
        //Forms![newmenu]![MeetingKey] = //Me.MeetingKey;
        //Forms![newmenu]![NewMenuMeetingDate] = //Me.MeetingDate;
        //[Forms]![newmenu]![TxtPostMeetingDate] = //Me.MeetingDate;
        //HoldListIndex = //Me.MeetingDate.ListIndex;
    //End //If;
;
    //DoCmd.SetWarnings //True;
    //RedisplayAgenda = //True;
    //DoCmd.Close;
    //'docmd.Close acForm, "_display agenda records sub form"
//End //If;
//Exit //Sub;
;
;
;
;
;
}
function Print_history_button_Click(){
//If //StudySelected //Then;
    //StudySelected = //False;
    //If //YesNo("Even //though //you //have //selected //a //single //item, //this //Report //prints //the //detail //for //ALL //agenda //items. //Do //you //wish //to //continue?") //Then;
        //GoTo //Continue;
    //Else;
        //Exit //Sub;
    //End //If;
//End //If;
//Continue:;
//If //StudySelected //Then;
    //StudySelected = //False;
    //If //YesNo("This //Report //prints //detail //for //ALL //Studies. //Do //you //wish //to //continue?") //Then;
        //GoTo //Continue1;
    //Else;
        //Exit //Sub;
    //End //If;
//End //If;
//Continue1:;
//StudySelected = //False;
//On //Error //GoTo //Err_Print_history_button_Click;
    var stDocName = '';
    //MyFilter = "MeetingDate = " //& //Me.MeetingDate;
 stDocName;
    //Forms![_display //agenda //records].Visible = //False;
    //DoCmd.OpenReport //stDocName, //acViewPreview;
    ;
;
    ;
//Exit_Print_history_button_Click:;
    //Exit //Sub;
;
//Err_Print_history_button_Click:;
//If //Err.Number = //2501 //Then;
    ;
    //MsgBox "No Agenda Items for Printing";
    //Resume //Exit_Print_history_button_Click;
//Else;
    //MsgBox //Err.Number //& "   Error in Printing Agenda";
    //Resume //Exit_Print_history_button_Click;
    ;
//End //If;
}
function Print_history_button_Small_Click(){
//If //StudySelected //Then;
    //StudySelected = //False;
    //If //YesNo("Even //though //you //have //selected //a //single //item, //this //Report //prints //the //detail //for //ALL //agenda //items. //Do //you //wish //to //continue?") //Then;
        //GoTo //Continue;
    //Else;
        //Exit //Sub;
    //End //If;
//End //If;
//Continue:;
//If //StudySelected //Then;
    //StudySelected = //False;
    //If //YesNo("This //Report //prints //detail //for //ALL //Studies. //Do //you //wish //to //continue?") //Then;
        //GoTo //Continue1;
    //Else;
        //Exit //Sub;
    //End //If;
//End //If;
//Continue1:;
//StudySelected = //False;
//Call //fCheckNewMembers;
;
;
;
;
//On //Error //GoTo //Err_Print_history_button_Small_Click;
    //MyFilter = "MeetingDate = " //& //Me.MeetingDate;
;
    var stDocName = '';
    stDocName;
    //'Forms![_display agenda records].Visible = False
    location.href = "Form_"+stDocName+".php";//, //acNormal;
    ;
    ;
;
;
    ;
//Exit_Print_history_button_Small_Click:;
    //Exit //Sub;
;
//Err_Print_history_button_Small_Click:;
//If //Err.Number = //2501 //Then;
    ;
    //MsgBox "No Agenda Items for Printing";
    //Resume //Exit_Print_history_button_Small_Click;
//Else;
    //MsgBox //Err.Number //& "   Error in Printing Agenda";
    //Resume //Exit_Print_history_button_Small_Click;
    ;
//End //If;
;
}
function MeetingDate_BeforeUpdate(Cancel){
//On //Error //Resume //Next;
//could initialize MyRecCount here
//'MsgBox HoldListIndex
//'MsgBox Me.MeetingDate.ListIndex
//'
//'MsgBox Me.MeetingDate.OldValue
//'SaveHoldListIndex = HoldListIndex
//If //HoldListIndex = //Me.MeetingDate.ListIndex //Then;
    //Exit //Sub;
//Else;
    //Call //UpdateAgendaRecordsWithLineNumbers(False);
    MyRecCount = (DCount("[Record_ID]", "Agenda Records", "[meetingdate] = " && "#" && Me.MeetingDate && "#" && " and [irbcode] = " && strQuote && Me.IRB_Code && strQuote));
    //'MsgBox MyRecCount
    //If MyRecCount;
       //MsgBox "There are no Agenda Items for--" //& //Me.MeetingDate;
    //End //If;
    ;
//End //If;
}
function ReAssignALL_Click(){
//On //Error //Resume //Next;
;
location.href = "Form_ReassignAgendaRecord.php";//, //, //, //, //, //acDialog;
//If //ReassignSwitch = //True //Then;
    //If //YesNo("This //action //will //move //all //Agenda //Items //from //this //meeting //to //the //one //you //selected. //Do //you //wish //to //continue?") //Then;
        ;
        //Me.AgendaRecID = //HoldRecID;
        //Me.AssignAll = "***";
        //Me.NewMeetingDate = //HoldMeetingDate;
        //DoCmd.SetWarnings //False;
        //[Forms]![_display //agenda //records]![MeetingDate] = //[Forms]![_display //agenda //records]![MeetingDate].Column(0);
               ;
        //DoCmd.OpenQuery "ReassignMeetingDate";
        //DoCmd.OpenQuery "ReassignMeetingDateInTempAgenda";
        //DoCmd.OpenQuery "ReassignCPAOnMeetingDate";
        //DoCmd.OpenQuery "ReassignSAEOnMeetingDate";
        //DoCmd.SetWarnings //True;
        //MsgBox //("All //Items //have //been //Reassigned //to //the--" //& //HoldMeetingDate //& "  Meeting"//);
        //'Me.Requery
        //'Me![_agenda sub form].Requery
        //'Me.Requery
        //Forms![_display //agenda //records].[_agenda //sub //form].Form.Requery;
        //'Me![_agenda sub form].Requery
        //DoCmd.SetWarnings //True;
        //RedisplayAgenda = //True;
        //DoCmd.Close;
    //End //If;
//Else;
        //Exit //Sub;
//End //If;
    ;
     ;
}
function ReassignButton_Click(){
//On //Error //Resume //Next;
//If //StudySelected = //True //Then;
        //'removed in 5.00.0005
    //'If Not IsNull(Me![_agenda sub form].Form![TextActionCategory]) Then
     //'   msgbox "You cannot reassign an Item which has already been posted"
     //'Exit Sub
    //'End If
    //StudySelected = //False;
    location.href = "Form_ReassignAgendaRecord.php";//, //, //, //, //, //acDialog;
    //If //ReassignSwitch = //True //Then;
      //If //YesNo("All //Agenda //items //for //Study //Number-" //& //Me.TextIRBno //& "- (NOT JUST THE ITEM YOU SELECTED) will be moved to the new Meeting. Do you wish to continue?"//) //Then;
        //[Forms]![_display //agenda //records]![MeetingDate] = //[Forms]![_display //agenda //records]![MeetingDate].Column(0);
        //Me.AgendaRecID = //HoldRecID;
        //Me.AssignAll = //Me.TextIRBno;
        //Me.NewMeetingDate = //HoldMeetingDate;
        //DoCmd.SetWarnings //False;
        ;
        //DoCmd.OpenQuery "ReassignMeetingDate";
        //DoCmd.OpenQuery "ReassignMeetingDateInTempAgenda";
        //DoCmd.OpenQuery "ReassignCPAOnMeetingDate";
        //DoCmd.OpenQuery "ReassignSAEOnMeetingDate";
        //DoCmd.SetWarnings //True;
        //MsgBox //("Study //#- " & HoldIRBNo & "  //Reassigned //to //the--" //& //HoldMeetingDate //& "  Meeting"//);
        //'Me.Requery
        //Forms![_display //agenda //records].[_agenda //sub //form].Form.Requery;
        //'Me![_agenda sub form].Requery
        //DoCmd.SetWarnings //True;
        //RedisplayAgenda = //True;
        //DoCmd.Close;
      //End //If;
    //Else;
        //Exit //Sub;
    //End //If;
//Else;
    //MsgBox "You must first select a Study";
//End //If;
;
 ;
;
}
function ReturnButton_Click(){
	window.history.back();
//DoCmd.SetWarnings //False;
;
//If //DataChanged = //False //Then //GoTo //Exit_Return;
;
//If //ItemLineNumberChangeFlag = //True //Then;
        //'ItemLineNumberChangeFlag = false  commented out for 7.3
        ;
       //If //YesNo("Do //you //wish //to //save //these //line //numbers?  //If //you //answer //'No' only the line numbers you just changed for this date will be undone. ") Then
            //UseZeros = //False;
                ;
            //Call //UpdateAgendaRecordsWithLineNumbers(UseZeros);
        //Else;
            //ItemLineNumberChangeFlag = //False;
            //'sets line numbers back to increments of 10
            //'Call Resequence
            //DoCmd.OpenQuery "UpdateLineNumberFlagInMeetingDates";
        //End //If;
//Else;
        //'UseZeros = True
        //'Call UpdateAgendaRecordsWithLineNumbers(UseZeros)
            ;
;
//End //If;
         ;
;
//Exit_Return:;
//DoCmd.SetWarnings //True;
    //RedisplayAgenda = //False;
    //NewHoldMeetingKeepTogether = //Me.chkKeepTogether;
       ;
    //DoCmd.Close;
    //DoCmd.Close //acForm, "_display agenda records sub form";
   ;
}
function ButDelete_Click(){
//On //Error //GoTo //Err_ButDelete_Click;
//If //StudySelected = //True //Then;
    //DoCmd.SetWarnings //False;
    //StudySelected = //True;
   ;
    //If //YesNo("Deleting //the //Agenda //Item //will //not //delete //the //underlying //action //that //placed //the //item //on //the //Agenda. //Do //you //still //want //to //delete //the //item?") //Then;
        //Me![HoldNewDate] = "";
        //If //Me![_agenda //sub //form].Form![Agenda //Category] //Like "*Adverse*" //Then;
            //Me.MAXSAE.Requery;
            //Me.MAXSAE = //Me.MAXSAE.ItemData(1);
            //'msgbox Me.MaxSAE
            //DoCmd.OpenQuery "ChangeMeetingDateOnSAE";
        //Else;
            //If //Not //Me![_agenda //sub //form].Form![Agenda //Category] //Like "Initial" //Then;
        ;
                //Me.MAXCPA.Requery;
                //Me.MAXCPA = //Me.MAXCPA.ItemData(1);
                //'msgbox Me.MAXCPA
                //DoCmd.OpenQuery "ChangeMeetingDateOnCPA";
            //End //If;
        //End //If;
        //DoCmd.OpenQuery "DeleteAgendaItem";
        //DoCmd.OpenQuery "DeleteAgendaItemFromTempAgenda";
        //Forms![_display //agenda //records].[_agenda //sub //form].Form.Requery;
    //Else;
        //Exit //Sub;
    //End //If;
//Else;
    //MsgBox "You must first select a Study";
//End //If;
;
  ;
;
//Exit_ButDelete_Click:;
    //Exit //Sub;
;
//Err_ButDelete_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ButDelete_Click;
    ;
}
function ReviewDetailForOne_Click(){
var stDocName = '';
        //Me.MeetingKey = //Me.MeetingDate.Column(1)                            //'added
        //Forms![newmenu]![MeetingKey] = //Me.MeetingKey                //'added
        //HoldMeetingKey = //Me.MeetingKey;
        //Me.MeetingDate = //Me.MeetingDate.ItemData(Me.MeetingDate.ListIndex) //'added
        //Forms![newmenu]![NewMenuMeetingDate] = //Me.MeetingDate          //'added
        //[Forms]![newmenu]![TxtPostMeetingDate] = //Me.MeetingDate         //'added
        //HoldListIndex = //Me.MeetingDate.ListIndex                        //'added
var stLinkCriteria = '';
    //StudySelected = //False;
    //CancelUpdate = //False;
    stDocName;
//'    MyFilter = "[IRB Code]= " & strQuote & [Forms]![NewMenu]![Location] & strQuote & " AND  [agenda records].[MeetingDate] = #" & [Forms]![_display agenda records]![MeetingDate] & "#"
   //' msgbox MyFilter
    //AgendaDetail = //True;
    location.href = "Form__PostMeetingForm.php";//, //, //, //, //, //acDialog;
    //Me.Visible = //False;
    ;
;
}
function cmdRefresh_Click(){
//On //Error //GoTo //Err_cmdRefresh_Click;
;
//If //ItemLineNumberChangeFlag = //True //Then;
//Else;
        //Call //Resequence;
        //DoCmd.SetWarnings //False;
        //Forms![_display //agenda //records]![TextExpansion2] = "NO";
        //'Forms![_display agenda records]![MeetingDate].Requery
        //ItemLineNumberChangeFlag = //True;
        //DataChanged = //True;
        //DoCmd.OpenQuery "UpdateLineNumberFlagInMeetingDates";
        //DoCmd.SetWarnings //True;
        //GoTo //NewRecordSource;
      ;
//End //If;
//NewRecordSource:;
//Forms![_display //agenda //records].[_agenda //sub //form].Form.RecordSource = "SelectTempAgendaRecords";
;
//Exit_cmdRefresh_Click:;
    //Exit //Sub;
;
//Err_cmdRefresh_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdRefresh_Click;
    ;
}
function ReNumberClickEvent(){
//'If YesNo("This action will reset all Items Numbers to multiples of ten based upon the Agenda Category hierarchy.  Do you want to continue?") Then
  ;
        //Call //Resequence;
        //DoCmd.SetWarnings //False;
        //Forms![_display //agenda //records]![TextExpansion2] = "NO";
        //'MsgBox Me.MeetingKey
        //ItemLineNumberChangeFlag = //True;
        //DataChanged = //True;
        //DoCmd.OpenQuery "UpdateLineNumberFlagInMeetingDates";
        //DoCmd.SetWarnings //True;
        //Forms![_display //agenda //records]![MeetingDate].Requery;
        //Forms![_display //agenda //records]![MeetingDate].Selected(HoldListIndex) = //True;
        //Forms![_display //agenda //records].[_agenda //sub //form].Form.RecordSource = "SelectTempAgendaRecords";
//'Else
 //'   Exit Sub
      ;
//'End If
}

    </script>
  </head>
  	<style>
	table, th, td {
    	border: 1px solid black;
    	border-collapse: collapse;
		}
	</style>
  <body onLoad="Form_Load();Form_Open();Form_Active();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    


    <input type='text' id='empty' style='position: absolute; left: 0; top: 0; width: 210; height: 96; visibility: hidden;' value='empty' visibility="hidden"></input>
  
  




    <label id=_display agenda records sub form Label value='Agenda Records' style='position: absolute; left: 0; top: 276; width: 363; height: 28; visibility:; visibility: hidden;'>Agenda Records</label>
    <label id=form_caption value='Preliminary Agenda Items' style='font-size: 28px; position: absolute; left: 741px; top: 21px; width: 307px; height: 45;visibility:'>Preliminary Agenda Items</label>
    <textarea type='text' id=irb code value='' style='position: absolute; left: 696; top: 100px; width: 384; height: 69'>
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
        echo $row['IRB Code'];
        }
    ?>
    </textarea>

  <label id=Label52 value='IRB' style='position: absolute; left: 879px; top: 67px; width: 47px; height: 28;visibility:'>IRB</label>
    <input type='button' id=ReturnButton value='&Return' style='position:absolute; left:873; top:543; width:276; height:39' onclick='ReturnButton_Click();'></input>
    <input type='button' id=ReviewDetailForOne value='&View Agenda Detail Form for All' style='position:absolute; left:873; top:421; width:276; height:60' onclick='ReviewDetailForOne_Click();'></input>
    <!--     <input type='text' id=TextIRBno value='empty' style='position:absolute; left:186; top:111; width:174; height:36'></input> -->

    <label id=Label70 value='Text69:' style='position: absolute; left: 234; top: 87; width: 84; height: 28; visibility:; visibility: hidden;'>Text69:</label>
    <!--     <select name='nextmeeting' Size='7' style='position:absolute; left:619; top:144; width:144; height:42'></select> -->

    <label id=Label23 value='Combo22:' style='position: absolute; left: 583; top: 144; width: 79; height: 24; visibility:; visibility: hidden;'>Combo22:</label>
    <input type='button' id=ReassignButton value='Assign Selected  Study  to Another &Meeting' style='position:absolute; left:873; top:298; width:276; height:60' onclick='ReassignButton_Click();'></input>
    <!--     <input type='text' id=AgendaRecID value='empty' style='position:absolute; left:547; top:81; width:138; height:42'></input> -->
    <!--     <input type='text' id=NewMeetingDate value='empty' style='position:absolute; left:552; top:81; width:132; height:42'></input> -->
    <select name='MeetingDate' Size='5' style='position: absolute; left: 319px; top: 55px; width: 180; height: 107px'>
    <?php
        $serverName = $SVNM;
        $connectionInfo = array( "Database"=>$DB, "UID"=>$UID, "PWD"=>$PWD, 'ReturnDatesAsStrings'=>true);
        $conn = sqlsrv_connect( $serverName, $connectionInfo);

        if( $conn ) {
        }
        else{
            echo "Connection could not be established.<br />";
            die( print_r( sqlsrv_errors(), true));
        }

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate], [IRBMeetingDates1].[key], [IRBMeetingDates1].[IRB Code] FROM [IRBMeetingDates1]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select>

    <label id=Label54 value='Next Meeting Date' style='font-size: 20px; position: absolute; left: 319px; top: 16px; width: 250; height: 25;visibility:'>Next Meeting Date</label>
    <input type='button' id=Command81 value='Preview  All Agenda Items &Summary Report' style='position:absolute; left:873; top:237; width:276; height:60' onclick='Command81_Click();'></input>
    <input type='button' value='bPrepareAgenda' style='position: absolute; left: 873; top: 177px; width: 276; height: 60' onclick='Command81_Click();'></input>
    <input type='button' value='Refresh Sort Sequence' style='position: absolute; left: 6px; top: 77px; width: 220px; height: 40px' onclick='Command81_Click();'></input>
    <input type='button' value='Refresh Item Numbers' style='position: absolute; left: 7px; top: 121px; width: 220px; height: 40px' onclick='Command81_Click();'></input>
    <input type='button' value='Assign ALL to Another Meeting' style='position: absolute; left: 875px; top: 360px; width: 276; height: 60' onclick='Command81_Click();'></input>
    <input type='button' id=ButDelete value='Delete Agenda Item' style='position:absolute; left:873; top:483; width:276; height:60' onclick='ButDelete_Click();'></input>
    <!--     <input type='text' id=SelectedRecordID value='empty' style='position:absolute; left:282; top:132; width:126; height:36'></input> -->

    <!--     <label id=Label84 value='Text83:' style='position:absolute; left:138; top:132; width:84; height:28; visibility:hidden'>Text83:</label> -->
    <!--     <input type='text' id=HoldNewDate value='empty' style='position:absolute; left:156; top:96; width:144; height:18'></input> -->

    <label id=Label86 value='Text85:' style='position: absolute; left: 50px; top: 24px; width: 143px; height: 28;'>Keep all Items for a Study Together</label>
    <!--     <select name='Label88' Size='10' style='position:absolute; left:78; top:168; width:112; height:28'></select> -->
    <!--     <select name='Label92' Size='1' style='position:absolute; left:186; top:12; width:112; height:28'></select> -->
    <input type='checkbox' id=CheckNotOnAgenda value='False' style='position: absolute; left: 6px; top: 26px; width: 26px; height: 27px'></input>
    <table style='width: 857px; position: absolute; left: 11px; top: 183px; height: 398px;'>
  
    <tr>
    <th width="8%">IRB#</th>
    <th width="11%">Agenda Category</th>
    <th width="8%">Agenda Group</th>
    <th width="8%">Item#</th>
    <th width="8%">EXP</th> 
    <th width="7%">Condition1</th>
    <th width="6%">Condition2</th>
    <th width="7%">Investigator ID</th>
    <th width="9%">Renewal</th>
    <th width="9%">Review</th>
    <th width="9%">Class</th>
    <th width="9%">Internal Number</th>
    <th width="9%">Meeting Date</th>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
    </table>
  </body>
</html>
