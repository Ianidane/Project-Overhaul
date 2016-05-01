<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function DontShowAgain_Click(){
//DoCmd.RunCommand //acCmdSaveRecord;
;
;
}
function ExitButton_Click(){
//Application.Quit;
}
function Form_Open(Cancel){
//On //Error //GoTo //err_db;
//could initialize FirstTime here
;
;
//ExemptOverrideSwitch = //Me.CheckOverride;
//'note that the FirstTimeRun field in the FirstTime table is set to true in the FrontEnd distributed with an update
//If //Me.DontShowAgain = //True //Then;
location.href = "Form_Frm_display_generated_number.php";;
   ;
//Else;
//'run queries to set up Current Letterclosings in Front end
;
;
;
//End //If;
//exit_open:;
//Exit //Sub;
;
//err_db:;
//MsgBox " PRO_IRB error Opening Flash Screen---" //& //Err.Description //& "  " //& //Err.Number;
//GoTo //exit_open;
}
function PROIRBButton_Click(){
//Me.DontShowAgain = //True;
  //DoCmd.Close;
   ;
   location.href = "Form_frm_display_generated_number.php";;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>



    <input type='text' id=Text2 value='empty' style='position:absolute; left:378; top:162; width:486; height:222'></input>
    <label id=Label10 value='Confidentiality Agreement' style='position:absolute; left:504; top:126; width:264; height:28; visibility:'>Confidentiality Agreement</label>
    <!--     <input type='checkbox' id=DontShowAgain value='False' style='position:absolute; left:60; top:381; width:0; height:0'></input> -->

    <!--     <label id=Label11 value='Dont show me this again' style='position:absolute; left:82; top:378; width:226; height:28; visibility:hidden'>Dont show me this again</label> -->
    <input type='button' id=PROIRBButton value='I have read and accept the terms as stated above.' style='position:absolute; left:375; top:402; width:492; height:40' onclick='PROIRBButton_Click();'></input>
    <input type='button' id=ExitButton value='&Exit\015\012Application' style='position:absolute; left:744; top:462; width:126; height:48' onclick='ExitButton_Click();'></input>

    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:534; top:54; width:193; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <!--     <input type='checkbox' id=CheckOverride value='False' style='position:absolute; left:69; top:313; width:42; height:16'></input> -->

    <label id=Label21 value='Check20' style='position:absolute; left:92; top:310; width:72; height:24; visibility:'>Check20</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:469; top:16; width:130; height:39; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:591; top:22; width:18; height:16; visibility:'>R</label>
  </body>
</html>
