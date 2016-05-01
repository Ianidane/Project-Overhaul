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
//If //OtherSelected = //False //Then;
    ;
    //MsgBox "You must first Select a Contact";
    //Exit //Sub;
//Else;
    //OtherSelected = //False;
    //HoldOther = //Me.ComboOther;
    //HoldOtherName = //Me.ComboOther.Column(1);
    //CancelUpdate = //False;
    //DoCmd.Close;
//End //If;
;
}
function ComboOther_DblClick(Cancel){
//HoldOther = //Me.ComboOther;
//HoldOtherName = //Me.ComboOther.Column(1);
//DoCmd.Close;
;
}
function ComboOther_Click(){
//OtherSelected = //True;
}
function Form_Open(Cancel){
//If //Forms![newmenu]![ActionGroup].Value = //7 //Then;
    //Me.AddNew.Enabled = //False;
//End //If;
//Me.ComboOther.SetFocus;
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


    <button name='ButtonCancel' type='submit' style='position:absolute; left:435; top:330; width:108; height:51' onclick='ButtonCancel_Click();'>"&Return"</button>
    <button name='AddNew' type='submit' style='position:absolute; left:49; top:330; width:232; height:51' onclick='AddNew_Click();'>"&Add New Contact"</button>
    <input type='button' id=Buttonok value='&Select' style='position:absolute; left:285; top:330; width:150; height:51' onclick='Buttonok_Click();'></input>
    <select name='ComboOther' style='position:absolute; left:6; top:72; width:592; height:234' onclick='ComboOther_Click();'>    </select>

    <label id=display name_Label value='Select Other Contact:' style='position:absolute; left:144; top:24; width:306; height:40; visibility:'>Select Other Contact:</label>
  </body>
</html>
