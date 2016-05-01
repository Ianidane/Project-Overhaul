<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Cancel_Click(){
//OKSwitch = //False;
//ReassignSwitch = //False;
//DoCmd.Close;
;
}
function Form_Open(Cancel){
//ReassignSwitch = //True;
//If //IsNull(HoldIRBNo) //Or //HoldIRBNo = "" //Then;
    //Me.TStudyno = "** ALL **";
//Else;
    //Me.TStudyno = //HoldIRBNo;
//End //If;
//If //RenewalLetterSwitch = //True //Then;
    //Me.irb_next_date = //HoldMeetingDate;
//End //If;
}
function OK_Click(){
//ReassignSwitch = //True;
//If //IsNull(Me.irb_next_date) //Or //Me.irb_next_date = "" //Then;
    //MsgBox "Please choose a date";
    //Exit //Sub;
//Else;
    //HoldMeetingDate = //Me.irb_next_date;
//End //If;
;
//DoCmd.Close;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:210; height:96'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:102; height:120'>    </select>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>




    <label id=Label8 value='Schedule Study Number:' style='position:absolute; left:3; top:102; width:351; height:34; visibility:'>Schedule Study Number:</label>
    <input type='button' id=OK value='&OK' style='position:absolute; left:370; top:306; width:133; height:48' onclick='OK_Click();'></input>
    <input type='button' id=Cancel value='&Cancel' style='position:absolute; left:504; top:306; width:133; height:48' onclick='Cancel_Click();'></input>
    <input type='text' id=select_irb value='empty' style='position:absolute; left:289; top:13; width:358; height:58'></input>

    <label id=Label13 value='IRB' style='position:absolute; left:223; top:13; width:58; height:40; visibility:'>IRB</label>
    <!--     <select name='priormeeting' Size='4' style='position:absolute; left:204; top:210; width:144; height:42'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].[Irb Code])=[Forms]![NewMenu]![location])) ORDER BY [IRBMeetingDates1].[MeetingDate]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select> -->

    <label id=Label23 value='Combo22:' style='position:absolute; left:168; top:210; width:79; height:24; visibility:'>Combo22:</label>
    <input type='text' id=TStudyno value='empty' style='position:absolute; left:349; top:105; width:285; height:36'></input>
    <select name='irb next date' Size='2' style='position:absolute; left:468; top:156; width:166; height:0'>
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

    <label id=Label10 value='To the Meeting Date of:' style='position:absolute; left:156; top:156; width:294; height:34; visibility:'>To the Meeting Date of:</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:34; top:39; width:168; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:40; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>

  </body>
</html>
