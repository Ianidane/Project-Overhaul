<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function AddNew_Click(){
//If //CoordSelected //Then;
    //MsgBox "You have selected an existing Coordinator then selected that you wish to Add" //_;
    //& " a New Coordinator.  Please click OK and Chose one option or the other.";
    //CoordSelected = //False;
    //Me.Combocoord.Selected(Me.Combocoord.ListIndex) = //False;
    //Exit //Sub;
//Else;
    //NewApplication = //True;
    //CancelUpdate = //False;
    //DoCmd.Close;
//End //If;
;
;
;
;
;
;
;
;
;
;
}
function ButtonCancel_Click(){
//DoCmd.Close;
//CancelUpdate = //True;
}
function ButtonOk_Click(){
//If //CoordSelected = //False //Then;
    ;
    //MsgBox "You must first Select a Coordinator";
    //Exit //Sub;
//Else;
    ;
    //HoldCoord = //Me.Combocoord;
    //HoldCoordName = //Me.Combocoord.Column(1);
;
    //CancelUpdate = //False;
    //DoCmd.Close;
//End //If;
;
;
}
function Combocoord_Click(){
//CoordSelected = //True;
}
function Form_Open(Cancel){
//CoordSelected = //False;
;
//Me.Combocoord.SetFocus;
//'Me.Combocoord.Dropdown
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


    <button name='ButtonOk' type='submit' style='position:absolute; left:276; top:330; width:151; height:51' onclick='ButtonOk_Click();'>"&Select"</button>
    <button name='ButtonCancel' type='submit' style='position:absolute; left:427; top:330; width:106; height:51' onclick='ButtonCancel_Click();'>"&Return"</button>
    <button name='AddNew' type='submit' style='position:absolute; left:3; top:330; width:280; height:51' onclick='AddNew_Click();'>"&Add New Coordinator"</button>
    <select name='Combocoord' style='position:absolute; left:90; top:72; width:324; height:246' onclick='Combocoord_Click();'>    </select>

    <label id=display name_Label value='Select Study Coordinator:' style='position:absolute; left:72; top:18; width:403; height:40; visibility:'>Select Study Coordinator:</label>
  </body>
</html>