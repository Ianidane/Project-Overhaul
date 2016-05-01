<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function ChangeIRBButton_Click(){
//On //Error //GoTo //Err_ChangeIRBButton_Click;
//If //IsNull(Me.NewIRB) //Then;
    //MsgBox //("You //must //select //a //New //IRB //to //change //the //Study //to");
    //Exit //Sub;
//Else;
//HoldNewIRBCode = //Me.NewIRB;
//Forms![CodesandTables]![StudyIRBCode] = //Me.IRB_Code;
//End //If;
;
;
;
//DoCmd.Close;
;
;
;
;
;
//Exit_ChangeIRBButton_Click:;
    //Exit //Sub;
;
//Err_ChangeIRBButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ChangeIRBButton_Click;
    ;
}
function Command14_Click(){
//On //Error //GoTo //Err_Command14_Click;
;
//HoldNewIRBCode = " ";
    ;
    //DoCmd.Close;
;
//Exit_Command14_Click:;
    //Exit //Sub;
;
//Err_Command14_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command14_Click;
    ;
}
function Form_Open(Cancel){
//'If IsNull(Me.Approval_Date___original_) Then
 //'   If Demo Then
  //'      Me.ChangeIRBButton.Enabled = False
   //' End If
    //'Exit Sub'
    ;
//'Else
 //'   MsgBox "Study has already been approved--You cannot change IRB at this time"
  //'  DoCmd.Close
//'End If
;
}
function NewIRB_Change(){
//'MsgBox Me.NewIRB & "  " & Me.NewMeeting
//Me.NewMeeting.Requery;
//If //IsNull(Me.NewMeeting.ItemData(0)) //Then;
    //MsgBox "There are no meeting dates for this IRB.  Return to the Administration Form and enter Meeting Dates";
   //Me.NewIRB = " ";
    //Exit //Sub;
    //'MsgBox Me.NewIRB & "  " & Me.NewMeeting
//End //If;
;
//HoldNewIRBCode = //Me.NewIRB;
//[Forms]![CodesandTables]![TBNewIRB] = //Me.NewIRB;
;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <input type='text' id=StudyNumber value='empty' style='position:absolute; left:162; top:12; width:0; height:0'></input>

    <label id=Label0 value='  IRB #:' style='position:absolute; left:72; top:12; width:82; height:28; visibility:'>  IRB #:</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:505; top:18; width:360; height:76'></input>

    <label id=Label1 value='Protocol / Title:' style='position:absolute; left:336; top:36; width:168; height:28; visibility:'>Protocol / Title:</label>
    <label id=Label7 value='Change From:' style='position:absolute; left:93; top:162; width:153; height:27; visibility:'>Change From:</label>
    <label id=Label8 value='Change To:' style='position:absolute; left:93; top:265; width:129; height:27; visibility:'>Change To:</label>
    <select name='NewIRB' Size='6' style='position:absolute; left:96; top:307; width:661; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].[IRB Code])=[Forms]![ChangeIRB]![NewIRB])) ORDER BY [IRBMeetingDates1].[MeetingDate]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select>

    <label id=Label10 value='New IRB' style='position:absolute; left:270; top:265; width:97; height:27; visibility:'>New IRB</label>
    <!--     <select name='NewMeeting' Size='7' style='position:absolute; left:600; top:354; width:0; height:30'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [IRB Meeting Schedule1].[IRB Code], [IRB Meeting Schedule1].[Sort seq] FROM [IRB Meeting Schedule1] WHERE ((([IRB Meeting Schedule1].[IRB Code])<>\"***\" And ([IRB Meeting Schedule1].[IRB Code])<>[Forms]![ChangeIRB]![IRB Code])) ORDER BY [IRB Meeting Schedule1].[Sort seq]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['IRB Code']."'".">".$row['IRB Code']."</Option>";
        }
    ?>
    </select> -->

    <label id=Label12 value='Meeting Date:' style='position:absolute; left:685; top:259; width:153; height:24; visibility:'>Meeting Date:</label>
    <input type='button' id=ChangeIRBButton value='Change IRB' style='position:absolute; left:360; top:354; width:172; height:51' onclick='ChangeIRBButton_Click();'></input>
    <input type='button' id=Command14 value='&Cancel' style='position:absolute; left:754; top:354; width:108; height:51' onclick='Command14_Click();'></input>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:99; top:201; width:658; height:25'></input>

    <label id=Label6 value='Current IRB' style='position:absolute; left:241; top:162; width:126; height:27; visibility:'>Current IRB</label>
    <!--     <select name='MaxOfMeetingDate' Size='5' style='position:absolute; left:684; top:201; width:0; height:25'>    </select> -->

    <label id=Label5 value='Meeting Date:' style='position:absolute; left:684; top:162; width:153; height:24; visibility:'>Meeting Date:</label>
    <!--     <input type='text' id=Approval Date  (original) value='empty' style='position:absolute; left:576; top:109; width:0; height:27'></input> -->

    <label id=Label15 value='Original Approval Date:' style='position:absolute; left:432; top:109; width:171; height:24; visibility:'>Original Approval Date:</label>
    <input type='text' id=Study Active? value='empty' style='position:absolute; left:162; top:60; width:0; height:30'></input>

    <label id=Label2 value='Active?:' style='position:absolute; left:63; top:60; width:91; height:28; visibility:'>Active?:</label>
    <input type='text' id=Study Status value='empty' style='position:absolute; left:162; top:112; width:252; height:30'></input>

    <label id=Label3 value='Study Status:' style='position:absolute; left:10; top:112; width:144; height:28; visibility:'>Study Status:</label>
  </body>
</html>
