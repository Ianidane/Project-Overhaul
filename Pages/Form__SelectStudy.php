<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function AddButton_Click(){
//If //StudySelected //Then;
    //If //YesNo("Are //you //sure //you //want //to //ADD //a //New //Study?") //Then;
        //GoTo //AddStudy;
    //Else;
        //Exit //Sub;
    //End //If;
//End //If;
//AddStudy:;
//NewApplication = //True;
//CancelUpdate = //False;
//DoCmd.Close;
;
}
function ButtonCancel_Click(){
//CancelUpdate = //True;
//DoCmd.Close;
;
}
function ButtonOK2_Click(){
//'MsgBox DCount("[irb#]", "_SelectStudy")
;
//If //StudySelected = //False //Then;
    //MsgBox "You must first Select a Study"//, //vbInformation;
    //Exit //Sub;
//Else;
    //If //Not //IsNull(Me.AgendaDate) //Then;
        //HoldMeetingDate = //Me.AgendaDate;
    //Else;
        //HoldMeetingDate = "";
    //End //If;
    //CancelUpdate = //False;
    //DoCmd.Close;
//End //If;
}
function Command26_Click(){
//End //Sub;
;
//Private //Sub //Form_Close();
//Me.AddButton.Enabled = //True;
//Me.Switch.Enabled = //True;
}
function Form_Open(Cancel){
//Me.Requery;
//If //LetterManagerSwitch //Then //Me.AddButton.Enabled = //False;
//Me![_selectSubform].Form.DatasheetForeColor = //0;
//Me![_selectSubform].Form.DatasheetFontWeight = //600;
//NewApplication = //False;
//CancelUpdate = //False;
//StudySelected = //False;
//If //TransactionSwitch //Or //Forms![newmenu]![HoldReadOnly] //Then;
    //Me.AddButton.Enabled = //False;
    //'Me.Switch.Enabled = False
//End //If;
;
;
    ;
;
;
}
function Switch_Click(){
//If //HoldStatus = "Open" //Then;
    //[Forms]![newmenu]![Status] = "Closed";
    //HoldStatus = "Closed";
    //Me.Switch.Caption = "Switch to Open Studies";
//Else;
    //[Forms]![newmenu]![Status] = "Open";
    //HoldStatus = "Open";
    //Me.Switch.Caption = "Switch to Closed Studies";
//End //If;
;
    ;
//Me![_selectSubform].Requery;
//Me.TextStatus = //HoldStatus;
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


    <button name='ButtonCancel' type='submit' style='position:absolute; left:751; top:552; width:108; height:45' onclick='ButtonCancel_Click();'>"&Cancel"</button>
    <label id=Title value='List of:' style='position:absolute; left:66; top:25; width:76; height:28; visibility:'>List of:</label>
    <input type='text' id=TextStatus value='empty' style='position:absolute; left:144; top:21; width:145; height:34'></input>
    <input type='text' id=TextLocation value='empty' style='position:absolute; left:144; top:69; width:648; height:34'></input>
    <input type='button' id=ButtonOK2 value='& Select  Study' style='position:absolute; left:586; top:552; width:168; height:45' onclick='ButtonOK2_Click();'></input>
    <input type='button' id=switch value='Switch to &Closed Studies' style='position:absolute; left:102; top:552; width:288; height:45' onclick='switch_Click();'></input>

    <label id=Label25 value='Studies at:' style='position:absolute; left:30; top:72; width:112; height:28; visibility:'>Studies at:</label>
    <input type='button' id=AddButton value='&Add a New Study' style='position:absolute; left:391; top:552; width:202; height:45' onclick='AddButton_Click();'></input>
    <!--     <select name='AgendaDate' Size='7' style='position:absolute; left:630; top:399; width:156; height:28'>
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

        $sql = sqlsrv_query($conn, "SELECT [agenda records1].[IRB#], Max([agenda records1].[MeetingDate]) AS MaxOfMeetingDate FROM [agenda records1] GROUP BY [agenda records1].[IRB#] HAVING ((([agenda records1].[IRB#])=Forms![_SelectStudy]![_SelectSubform].Form![IRB#]) And ((Max([agenda records1].MeetingDate))>Date())) ORDER BY Max([agenda records1].[MeetingDate])");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['IRB#']."'".">".$row['IRB#']."</Option>";
        }
    ?>
    </select> -->

    <!--     <label id=Label488 value='On Agenda' style='position:absolute; left:624; top:195; width:103; height:24; visibility:hidden'>On Agenda</label> -->
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:690; top:39; width:166; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:618; top:0; width:130; height:43; visibility:'>ProIRB</label>
    <input type='text' id=Text29 value='empty' style='position:absolute; left:478; top:123; width:145; height:34'></input>
    <label id=Label30 value='Highest Number Used' style='position:absolute; left:232; top:126; width:237; height:28; visibility:'>Highest Number Used</label>
  </body>
</html>
