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
//If //SponsorSelected = //False //Then;
    ;
    //MsgBox "You must first Select a Sponsor"//, //vbInformation;
    //Exit //Sub;
//Else;
    //SponsorSelected = //False;
    //HoldSponsor = //Me.ComboSponsor;
    //HoldSponsorName = //Me.ComboSponsor.Column(1);
;
    //CancelUpdate = //False;
    //DoCmd.Close;
//End //If;
;
;
}
function Combosponsor_Click(){
//SponsorSelected = //True;
}
function ComboSponsor_DblClick(Cancel){
//SponsorSelected = //False;
    //If //IsNull(Me.ComboSponsor) //Or //Me.ComboSponsor = "" //Then;
        //Exit //Sub;
    //Else;
        //HoldSponsor = //Me.ComboSponsor;
        //HoldSponsorName = //Me.ComboSponsor.Column(1);
    //End //If;
;
    //CancelUpdate = //False;
    //DoCmd.Close;
}
function Form_Open(Cancel){
//If //Forms![newmenu]![ActionGroup].Value = //8 //Then;
    //Me.AddNew.Enabled = //False;
//End //If;
//Me.ComboSponsor.SetFocus;
;
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
    //MsgBox //Err.Description, //vbInformation;
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


    <button name='ButtonCancel' type='submit' style='position:absolute; left:438; top:330; width:108; height:51' onclick='ButtonCancel_Click();'>"&Return"</button>
    <button name='AddNew' type='submit' style='position:absolute; left:91; top:330; width:240; height:51' onclick='AddNew_Click();'>"&Add New Sponsor"</button>
    <input type='button' id=Buttonok value='&Select' style='position:absolute; left:330; top:330; width:108; height:51' onclick='Buttonok_Click();'></input>
    <select name='ComboSponsor' Size='1' style='position:absolute; left:94; top:90; width:451; height:222' onclick='ComboSponsor_Click();'>    </select>

    <label id=display name_Label value='Select Study Sponsor:' style='position:absolute; left:76; top:36; width:403; height:40; visibility:'>Select Study Sponsor:</label>
  </body>
</html>
