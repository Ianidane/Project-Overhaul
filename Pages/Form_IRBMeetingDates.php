<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Form_AfterUpdate(){
//Exit_Update:;
//Exit //Sub;
//error_dup:;
//MsgBox //Err.Number;
//Resume //Exit_Update;
;
}
function Form_BeforeUpdate(Cancel){
//On //Error //GoTo //error_dup2;
//If //Me.NewRecord //Then //Exit //Sub;
//If //YesNo("If //you //change //the //meeting //date, //all //items //scheduled //for //the //original //date //will //be //changed //to //reflect //the //new //meeting //date." //_;
//& "Do you wish to continue?"//) //Then;
    //If //YesNo("Still //want //to //change //meeting //date //from  " & Me.MeetingDate.OldValue & "  //TO  " & Me.MeetingDate.Value & "  //?") //Then;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "UpdateAgendaDate";
        //DoCmd.OpenQuery "UpdateSAEonMeetingDate";
        //DoCmd.OpenQuery "UpdateCPAonMeetingDate";
        //DoCmd.SetWarnings //True;
    //Else;
        //Me.Undo;
    //End //If;
//Else;
//Me.Undo;
//End //If;
//exit_update1:;
//Exit //Sub;
//error_dup2:;
//MsgBox //Err.Number //& " - " //& //Err.Description;
//Resume //exit_update1;
}
function Form_Delete(Cancel){
//On //Error //GoTo //ErrorCheckif2;
//could initialize sql1 here
//could initialize sql2 here
//could initialize sql3 here
//could initialize hope here
//could initialize dbs here
//Set dbs;
//could initialize GoForIt here
//'SELECT [IRB - Research  Database].[IRB Code], [agenda records].MeetingDate, [agenda records].[IRB#], [IRB - Research  Database].[IRB#]
//'FROM [agenda records] INNER JOIN [IRB - Research  Database] ON [agenda records].[IRB#] = [IRB - Research  Database].[IRB#]
//'WHERE ((([IRB - Research
//'Database].[IRB Code])=[Forms]![IRBMeetingDates]![Irb Code]) AND
//'(([agenda records].MeetingDate)=[Forms]![IRBMeetingDates]![MeetingDate]));
;
sql1;
sql2 = " WHERE [IRB - Research  Database].[IRB Code]= " && "'" && [Forms]![IRBMeetingDates]![IRB Code];
sql3 = "' AND [agenda records].MeetingDate = " && "#" && [Forms]![IRBMeetingDates]![MeetingDate] && "#";
//sqlSTR = sql1 //& sql2 //& sql3;
//Set GoForIt;
;
//If //GoForIt.RecordCount //> //0 //Then;
    //GoForIt.MoveFirst;
    //MsgBox "There are agenda records for this date. You cannot delete a meeting date when it has items associated with it. "//, //vbInformation;
    //Cancel = //True;
//End //If;
;
//exit_checkif2:;
//Exit //Sub;
;
//ErrorCheckif2:;
//MsgBox "Error Checking for agenda records before deleting a meeting date  " //& //Err.Description //& //Err.HelpContext;
//Resume //exit_checkif2;
;
;
}
function Form_Open(Cancel){
//If //Demo //Then;
    //Me.IRB_Code.Locked = //True;
    //Me.MeetingDate.Locked = //True;
//End //If;
//MeetingDatesChanged = //False //'bridgestuff
}
function Form_Unload(Cancel){
//If //CyberIRB_Flag //Then //'bridgestuff
    //If //MeetingDatesChanged //Then //'bridgestuff
//'       MsgBox "I'm going to update CyberIRB with the new dates. It may take a minute or so....", vbInformation, "Updating CyberIRB"
//'60        DoCmd.SetWarnings False
//'    Dim strSQL As String
//'    Dim cmd As ADODB.Command
//'
//'10    Set cmd = New ADODB.Command
    ;
         //gclsCyberIRB.ReplaceCyberRelatedTableWithProTable "IRBMeetingDates1" //'bridgestuff
        ;
//'70    cmd.ActiveConnection = gclsCyberIRB.CyberConnection
//'80    cmd.CommandText = "toolRenameAnyTable"
//'90    cmd.CommandType = adCmdStoredProc
//'50    cmd.Parameters("@strTableName") = gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") & "ProIRBMeetingDates1"
//'100   cmd.Execute
//'
//'140           Debug.Print Now()
//'150       strSQL = "SELECT * INTO [" & g_CyberServerConnection & "].[" & gclsCyberIRB.mcolCyberConfigAttributes("CustPrefix") & "ProIRBMeetingDates1" & "] FROM [IRBMeetingDates1] ;"
//'160       Debug.Print strSQL
//'170       DoCmd.RunSQL strSQL
;
        ;
        ;
        ;
        ;
    //End //If //'bridgestuff
//End //If;
}
function MeetingDate_AfterUpdate(){
//MeetingDatesChanged = //True //'bridgestuff
}
function MeetingDate_BeforeUpdate(Cancel){
//If //IsNull(Me.MeetingDate) //Or //Me.MeetingDate = " " //Then;
    //MsgBox "Date cannot be blank.  If you want to Delete, select the row then Right Mouse Delete";
    //Me.Undo;
    //Exit //Sub;
//End //If;
//Call //CheckIFDupDate;
    ;
}
function CheckIFDupDate(){//As //Boolean;
//Me.CheckDupCombo.Requery;
//If //Me.CheckDupCombo.ListCount //> //0 //Then;
    //MsgBox "The meeting date is already present"//, //vbInformation;
    //Me.Undo;
    //CheckIFDupDate = //True;
    //Exit //Function;
//Else;
    //CheckIFDupDate = //False;
//End //If;
//exit_checkif:;
//Exit //Function;
;
//ErrorCheckif:;
//MsgBox "Error Checking for duplicate agenda date" //& //Err.Description //& "  " //& //Err.Number;
//Resume //exit_checkif;
;
;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>




    <input type='text' id=MeetingDate value='empty' style='position:absolute; left:168; top:48; width:114; height:0'></input>

    <label id=MeetingDate Label value='MeetingDate' style='position:absolute; left:6; top:48; width:156; height:25; visibility:'>MeetingDate</label>
    <!--     <select name='CheckDupCombo' Size='2' style='position:absolute; left:744; top:18; width:60; height:36'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [IRBMeetingDates1].[Irb Code], [IRBMeetingDates1].[MeetingDate] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].[Irb Code])=[Forms]![IRBMeetingDates1]![Irb Code]) And (([IRBMeetingDates1].[MeetingDate])=[Forms]![IRBMeetingDates1]![MeetingDate]))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Irb Code']."'".">".$row['Irb Code']."</Option>";
        }
    ?>
    </select> -->

    <!--     <label id=CheckDupCombo_Label value='CheckDupCombo' style='position:absolute; left:600; top:18; width:70; height:24; visibility:hidden'>CheckDupCombo</label> -->
    <input type='text' id=Irb Code value='empty' style='position:absolute; left:168; top:12; width:421; height:0'></input>

    <label id=Irb Code Label value='IRBCode' style='position:absolute; left:6; top:12; width:156; height:25; visibility:'>IRBCode</label>

  </body>
</html>
