<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function CancelButton_Click(){
//DoCmd.Close;
}
function Form_AfterUpdate(){
//'Me.New_IRBCode = Me.TBIRB_New
//DoCmd.SetWarnings //False;
//DoCmd.OpenQuery "UpdateSAEAgendaDAteforStudy";
//DoCmd.OpenQuery "UpdateCPAAgendaDateforStudy";
//DoCmd.SetWarnings //True;
}
function Form_BeforeUpdate(Cancel){
//'If Me.MeetingDate = Me.Text20 Then
 //'   MsgBox "You must select a New Meeting Date for the Agenda Item at the new IRB"
  //'  Cancel = True
   //' Exit Sub
//'End If
;
;
//'DoCmd.SetWarnings True
;
//'DoCmd.OpenQuery "UpdateAgendaForStudy"
//'DoCmd.OpenQuery "UpdateSAEAgendaDAteforStudy"
//'DoCmd.OpenQuery "UpdateCPAAgendaDatefor Study"
//'DoCmd.OpenQuery "UpdateLocationforStudy"
//'DoCmd.OpenQuery "UpdateIRBCodeAgendaRecordsChangeIRB"
//'DoCmd.SetWarnings True
;
}
function Form_Current(){
//Me.MeetingDate.Requery;
;
}
function Form_Open(Cancel){
//Me.TBIRB_New = "Select a New *" //& //Forms![CodesandTables]![TBNewIRB] //& "* " //& "Meeting Date for each item";
//'Me.MeetingDate.Requery
}
function Form_Timer(){
//On //Error //GoTo //Err_timer;
//Me.TimerInterval = //0;
//If //Me.RecordsetClone.RecordCount //<> //0 //Then;
    //MsgBox "There are unposted Agenda Items for this Study which must be assigned to a meeting date of the IRB they are being Transferred to.";
//Else;
    //DoCmd.Close //acForm, "SelectUnpostedAgendaItems";
    //'MsgBox "there are no unposted"
//End //If;
//exit_timer:;
//Exit //Sub;
//Err_timer:;
//MsgBox //Err.Description //& " " //& //Err.Number;
//GoTo //exit_timer;
}
function MeetingDate_AfterUpdate(){
//'Me.New_IRBCode = Forms![CodesandTables]![TBNewIRB]
}
function New_IRBCode_BeforeUpdate(Cancel){
//Me.MeetingDate.Requery;
;
//If //IsNull(Me.MeetingDate.ItemData(0)) //Then;
    //MsgBox "There are no meeting dates for this IRB.  Return to the Administration Form and enter Meeting Dates";
    ;
;
    //'Me.MeetingDate = Me.MeetingDate.OldValue
    //Cancel = //True;
    //Exit //Sub;
   ;
//Else;
    //Me.MeetingDate = //Me.MeetingDate.ItemData(0);
 ;
   ;
//End //If;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Current();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <label id=IRB#_Label value='IRB#' style='position:absolute; left:6; top:135; width:96; height:24; visibility:'>IRB#</label>
    <label id=SourceNumber_Label value='SourceNumber' style='position:absolute; left:108; top:135; width:144; height:24; visibility:'>SourceNumber</label>
    <label id=MeetingDate_Label value='Select New\015\012MeetingDate' style='position:absolute; left:469; top:108; width:141; height:52; visibility:'>Select New\015\012MeetingDate</label>
    <label id=Agenda Category_Label value='Agenda Category' style='position:absolute; left:670; top:135; width:136; height:24; visibility:'>Agenda Category</label>
    <label id=Agenda Condition1_Label value='Agenda Condition1' style='position:absolute; left:816; top:135; width:142; height:24; visibility:'>Agenda Condition1</label>
    <label id=Agenda Condition2_Label value='Agenda Condition2' style='position:absolute; left:966; top:135; width:130; height:24; visibility:'>Agenda Condition2</label>
    <!--     <label id=IRBCode_Label value='Transferring From IRB' style='position:absolute; left:258; top:108; width:153; height:52; visibility:hidden'>Transferring From IRB</label> -->
    <label id=Label19 value='From Meeting Date' style='position:absolute; left:235; top:108; width:153; height:52; visibility:'>From Meeting Date</label>
    <input type='text' id=TBIRB_New value='empty' style='position:absolute; left:243; top:36; width:528; height:30'></input>
    <input type='button' id=CancelButton value='Return' style='position:absolute; left:988; top:42; width:108; height:42' onclick='CancelButton_Click();'></input>


    <input type='text' id=SourceNumber value='empty' style='position:absolute; left:108; top:6; width:108; height:25'></input>
    <!--     <select name='IRB Action Category' Size='3' style='position:absolute; left:678; top:6; width:96; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].[IRB Code])=[Forms]![CodesandTables]![TBNewIRB])) ORDER BY [IRBMeetingDates1].[MeetingDate] WITH OWNERACCESS OPTION");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select> -->
    <select name='Agenda Category' Size='4' style='position:absolute; left:672; top:6; width:136; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].[IRB Code])=[Forms]![CodesandTables]![TBNewIRB])) ORDER BY [IRBMeetingDates1].[MeetingDate] WITH OWNERACCESS OPTION");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select>
    <select name='Agenda Condition1' Size='5' style='position:absolute; left:819; top:6; width:142; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].[IRB Code])=[Forms]![CodesandTables]![TBNewIRB])) ORDER BY [IRBMeetingDates1].[MeetingDate] WITH OWNERACCESS OPTION");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select>
    <select name='Agenda Condition2' Size='6' style='position:absolute; left:969; top:6; width:130; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].[MeetingDate] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].[IRB Code])=[Forms]![CodesandTables]![TBNewIRB])) ORDER BY [IRBMeetingDates1].[MeetingDate] WITH OWNERACCESS OPTION");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select>
    <select name='MeetingDate' Size='2' style='position:absolute; left:466; top:6; width:0; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRB Conditions1].[Condition] FROM [IRB Conditions] ORDER BY [IRB Conditions1].[Condition]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Condition']."'".">".$row['Condition']."</Option>";
        }
    ?>
    </select>
    <input type='text' id=IRB# value='empty' style='position:absolute; left:6; top:6; width:96; height:25'></input>
    <!--     <input type='text' id=New_IRBCode value='empty' style='position:absolute; left:232; top:6; width:328; height:0'></input> -->
    <input type='text' id=Text20 value='empty' style='position:absolute; left:234; top:6; width:0; height:0'></input>

  </body>
</html>
