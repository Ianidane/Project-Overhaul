<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Cancel_Click(){
	window.history.back();
//DoCmd.Close;
;
}
function Form_Close(){
//'RUTGERSISSUE routine to create the AgendaGroup to use for EXP1 in the meeting minutes based on the keep together flag
//DoCmd.SetWarnings //False;
;
//If //DMax("EXPANSION3", "GetKeepTogetherFlagForMeeting"//) //Then;
    //DoCmd.OpenQuery "qryUpdateAgendaRecordsWithLineNumbers";
    //DoCmd.OpenQuery "Update_Minutes_EXP1GroupBasedonMaxClass";
//Else;
    //DoCmd.OpenQuery "Update_Minutes_EXP1GroupBasedonAgendaCategory";
//End //If;
}
function OK_Click(){
//On //Error //GoTo //Error_OK;
//could initialize UseZeros here
//could initialize X here
;
var ThisIRBCode = '';
;
//If //IsNull(Me.priormeeting) //Then;
    //MsgBox "You must have a least one meeting date prior to the one you have selected. Return to the Main Menu-Administration and enter a meeting date ";
    //DoCmd.Close;
    //Exit //Sub;
//End //If;
;
//Forms![newmenu]![TxtPostMeetingDate] = //Me.[irb //next //date];
//HoldMeetingKey = //Me.irb_next_date.Column(1);
//Me.TextMeetingKey = //Me.irb_next_date.Column(1);
//HoldPriorMonth = //Me.priormeeting;
//ThisMeetingDate = //Me.[irb //next //date];
ThisIRBCode;
;
;
//If //Me.irb_next_date.Column(2) = "" //Then;
    //PostKeepTogether = //DMax("[KeepStudyTogether]", "Copyright"//);
//Else;
    //PostKeepTogether = //Me.irb_next_date.Column(2);
//End //If;
;
;
//If //Me.irb_next_date.Column(2) = "" //Then //'never had agenda run
;
    ;
    //If //IsNull(DMax("[Record_ID]", "Agenda Records"//, "[meetingdate] = " //& "#" //& //ThisMeetingDate //& "#" //& " and [irbcode] = " //& //strQuote //& ThisIRBCode //& //strQuote //& " and [AgendaLineNumber] > 0"//)) //Then;
        //'has no line numbers
        //Call //NewLineNumbers(PostKeepTogether);
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "qryUpdateAgendaRecordsWithLineNumbers";
        //'docmd.OpenQuery "QryUpdateMeetingdateexpansionfromagendaCheck"
        //'docmd.OpenQuery "QryUpdateMeetingDateExpansionFromPost"
        //DoCmd.OpenQuery "QryUpdateMeetingDateExpansionFromAgendaCheck";
        //DoCmd.SetWarnings //True;
    //End //If;
    ;
//End //If;
;
//exit_OK:;
    ;
    //DoCmd.Close //acForm, "_agendamonthcheck";
    //On //Error //Resume //Next;
    location.href = "Form__PostMeetingForm.php";;
   ;
      //Exit //Sub;
  ;
        ;
;
   ;
//Error_OK:;
//MsgBox "Error occurred preparing for posting.   " //& //Err.Description //& "  " //& //Err.Number;
//Application.Quit;
;
}

    </script>
  </head>
  <body onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position: absolute; left: 341px; top: 584px; width: 0; height: 0; visibility: hidden;'></input>
    <input type='checkbox' id='empty' value='False' style='position: absolute; left: 177px; top: 437px; width: 0; height: 0; visibility: hidden;'></input>


    <input type='text' id='empty' style='position: absolute; left: 285px; top: 431px; width: 210; height: 96; visibility: hidden;' value='empty' visibility="hidden"></input>
    <select name='empty' style='position: absolute; left: 0; top: 0; width: 0; height: 0; visibility: hidden;'></select>
    <select name='empty' style='position: absolute; left: 8px; top: 233px; width: 102; height: 120; visibility: hidden;' visibility="hidden"></select>


    <button name='empty' type='submit' style='position: absolute; left: 271px; top: 566px; width: 0; height: 0; visibility: hidden;'></button>




    <label id=Label8 value='Verify the IRB and the Meeting Date you wish to Post IRB Actions to:' style='position: absolute; left: 146px; top: 56px; width: 470px; height: 28px;visibility:'>Verify the IRB and the Meeting Date you wish to Post IRB Actions to:</label>
    <input type='button' id=OK value='&OK' style='position:absolute; left:370; top:330; width:133; height:48' onclick='OK_Click();'></input>
    <input type='button' id=Cancel value='&Cancel' style='position:absolute; left:504; top:330; width:133; height:48' onclick='Cancel_Click();'></input>
    <textarea type='text' id=select_irb value='empty' style='position:absolute; left:132; top:100; width:456; height:54'>
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

    <label id=Label13 value='IRB' style='position: absolute; left: 92px; top: 116px; width: 29px; height: 34;visibility:'>IRB</label>
    <!--     <select name='priormeeting' Size='4' style='position:absolute; left:64; top:223; width:144; height:42'></select> -->

    <label id=Label23 value='Combo22:' style='position: absolute; left: 28; top: 223; width: 79; height: 24; visibility:; visibility: hidden;'>Combo22:</label>
    <select name='irb next date' Size='2' style='position: absolute; left: 236px; top: 172px; width: 166; height: 107px'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate], [IRBMeetingDates1].[key], [IRBMeetingDates1].[Expansion3], [IRBMeetingDates1].[IRB Code] FROM [IRBMeetingDates1] ORDER BY [IRBMeetingDates1].[MeetingDate]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
	</select>
	

    <label id=Label10 value='Meeting Date' style='position:absolute; left:90; top:175; width:171; height:34; visibility:'>Meeting Date</label>
    <!--     <input type='text' id=TextMeetingKey value='empty' style='position:absolute; left:48; top:294; width:168; height:24'></input> -->

    <label id=Label25 value='Text24:' style='position: absolute; left: 0; top: 294; width: 61; height: 24; visibility:; visibility: hidden;'>Text24:</label>

  </body>
</html>
