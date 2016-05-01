<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function AddNew_Click(){
//NewApplication = //True;
//CancelUpdate = //False;
//DoCmd.Close;
;
}
function ButtonCancel_Click(){
//DoCmd.Close;
//CancelUpdate = //True;
}
function ButtonOk_Click(){
//If //PISelected = //False //Then;
    ;
    //MsgBox "You must first Select a PI";
    //Exit //Sub;
//Else;
    //PISelected = //False;
    //HoldPI = //Me.ComboPI;
    //HoldPIName = //Me.ComboPI.Column(1);
;
    //CancelUpdate = //False;
    //DoCmd.Close;
//End //If;
;
}
function ComboPI_Click(){
//PISelected = //True;
}
function ComboPI_DblClick(Cancel){
//HoldPI = //Me.ComboPI;
//HoldPIName = //Me.ComboPI.Column(1);
//DoCmd.Close;
;
}
function Form_Open(Cancel){
//If //Forms![newmenu]![ActionGroup].Value = //7 //Then;
    //Me.AddNew.Enabled = //False;
//End //If;
//Me.ComboPI.SetFocus;
//'dropdown
}
function Command12_Click(){
//On //Error //GoTo //Err_Command12_Click;
;
;
    //Screen.PreviousControl.SetFocus;
    //DoCmd.FindNext;
;
//Exit_Command12_Click:;
    //Exit //Sub;
;
//Err_Command12_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command12_Click;
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


    <button name='ButtonCancel' type='submit' style='position:absolute; left:390; top:330; width:108; height:51' onclick='ButtonCancel_Click();'>"&Return"</button>
    <button name='AddNew' type='submit' style='position:absolute; left:60; top:330; width:180; height:51' onclick='AddNew_Click();'>"&Add New P.I."</button>
    <input type='button' id=Buttonok value='&Select' style='position:absolute; left:240; top:330; width:150; height:51' onclick='Buttonok_Click();'></input>
    <select name='ComboPI' style='position:absolute; left:108; top:66; width:324; height:234' onclick='ComboPI_Click();'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_Select PIs for irb#].[Common Contact ID], [_Select PIs for irb#].[display name] FROM [_Select PIs for irb#] ORDER BY [_Select PIs for irb#].[display name]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select>

    <label id=display name_Label value='Select Investigator:' style='position:absolute; left:54; top:12; width:403; height:40; visibility:'>Select Investigator:</label>
  </body>
</html>
