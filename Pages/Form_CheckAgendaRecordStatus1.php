<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function NewCheckTogetherFunction(){
//'MsgBox [Forms]![NewMenu]![TBLocation]
;
//Me.CheckMeetingDate.Requery;
//'added this line for test
//'Forms![CheckAgendaRecordStatus1].CheckMeetingDate = Forms![_irb input form].AgendaDate
//StudySelected = //False;
//Forms![CheckAgendaRecordStatus1]![CheckNextMeetingDate] = //Forms![CheckAgendaRecordStatus1]![CheckNextMeetingDate].ItemData(0);
//If //IsNull(Forms![CheckAgendaRecordStatus1]![CheckNextMeetingDate]) //Then;
        //MsgBox //("You //should //add //current //Meeting //Dates //before //using //this //feature.");
        //NoMeetingDates = //True;
        //Exit //Function;
    //Else;
        //NoMeetingDates = //False;
//End //If;
//'MsgBox "NextMeetingDate---" & Forms![CheckAgendaRecordStatus1]![CheckNextMeetingDate]
//If //RedisplayAgenda = //False //Then;
    //'selects the next meeting date
    //Me.CheckMeetingDate.Selected(1) = //True;
;
    //could initialize ctlList here
    //varItem = //1;
    ;
    //' Return Control object variable pointing to list box.
    //Set ctlList;
    //' Enumerate through selected items.
    //While //varItem //< //ctlList.ListCount //+ //1;
        //' Print value of bound column.
        //If //DateValue(ctlList.ItemData(varItem)) = //DateValue(Forms![CheckAgendaRecordStatus1]![CheckNextMeetingDate]) //Then;
            //HoldListIndex = //varItem;
            //GoTo //exit_While;
        //End //If;
        //varItem = //varItem //+ //1;
    //Wend;
//exit_While:;
    //DefaultIndex = //varItem;
    //Me.CheckMeetingDate.Selected(HoldListIndex) = //True;
//Else;
    //Me.CheckMeetingDate.Selected(HoldListIndex) = //True;
//End //If;
;
;
;
//Forms![CheckAgendaRecordStatus1].CheckMeetingDate = //Forms![CheckAgendaRecordStatus1].CheckMeetingDate.ItemData(HoldListIndex);
;
;
;
//'MsgBox Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(3) & "-" & Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(4)
//If //Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(4) = "yes" //Then;
    //ItemLineNumberChangeFlag = //True;
//Else;
    //ItemLineNumberChangeFlag = //False;
//End //If;
//Forms![newmenu]![NewMenuMeetingDate] = //Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(0);
    //Forms![newmenu]![TxtPostMeetingDate] = //Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(0);
    //Forms![newmenu]![MeetingKey] = //Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(1);
    //HoldListIndex = //Me.CheckMeetingDate.ListIndex;
    ;
//If //Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(3) = "" //Or //IsNull(Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(3)) //Then;
        //'first time thru this month
        //could initialize tempkt here
        tempkt;
    //If //Not //DMax("[KeepStudyTogether]", "Copyright"//) //Then;
        //NewHoldMeetingKeepTogether = //False;
    //Else;
        //NewHoldMeetingKeepTogether = //True;
    //End //If;
//Else;
    //'MsgBox Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(3) & "  NewChecktogether"
    //NewHoldMeetingKeepTogether = //Forms![CheckAgendaRecordStatus1].CheckMeetingDate.Column(3);
    ;
//End //If;
;
//could initialize reopen here
reopen;
 //X = //CheckAgendaRecordStatus([Forms]![newmenu]![TxtPostMeetingDate], //Forms![newmenu]![NewMenuMeetingDate], //Forms![newmenu]![Location], //NewHoldMeetingKeepTogether, //ItemLineNumberChangeFlag, //NewHoldMeetingKeepTogether, //reopen);
;
//Exit //Function;
}
function Form_Load(){
//Me.CheckMeetingDate.Requery;
//Me.CheckNextMeetingDate.Requery;
}
function Form_Open(Cancel){
//could initialize Looper here
//Me.CheckMeetingDate.Requery;
//Me.CheckNextMeetingDate.Requery;
//'SaveHoldListIndex = HoldListIndex
;
//Call //NewCheckTogetherFunction;
//If //NoMeetingDates //Then //GoTo //Exit_Check;
//'Me.Visible = False
location.href = "Form__display agenda records.php";//, //acNormal, //, //, //, //acDialog;
;
//While //RedisplayAgenda = //True;
        //'DoCmd.OpenForm "dummy", acNormal, , , , acHidden
        //'DoCmd.Close acForm, "dummy"
    //SendKeys "{Enter}";
    //MsgBox "Please click OK to Continue";
  ;
    //Call //NewCheckTogetherFunction;
    location.href = "Form__display agenda records.php";//, //acNormal, //, //, //, //acDialog;
//Wend;
  ;
//Exit_Check:;
;
    //DoCmd.Close //acForm, "CheckAgendaRecordStatus1";
    //Exit //Sub;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>
    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <select name='CheckNextMeetingDate' style='position:absolute; left:60; top:54; width:0; height:42'>
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

        $sql = sqlsrv_query($conn, "SELECT [NextMeetingDate].[MinOfMeetingDate] FROM NextMeetingDate");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MinOfMeetingDate']."'".">".$row['MinOfMeetingDate']."</Option>";
        }
    ?>
    </select>

    <label id=Label23 value='Combo22:' style='position:absolute; left:24; top:54; width:79; height:24; visibility:'>Combo22:</label>
    <select name='CheckMeetingDate' Size='1' style='position:absolute; left:534; top:84; width:156; height:132'>    </select>

    <label id=Label3 value='List2:' style='position:absolute; left:390; top:84; width:46; height:24; visibility:'>List2:</label>
  </body>
</html>
