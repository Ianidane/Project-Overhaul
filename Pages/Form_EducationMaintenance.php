<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function btnReturn_Click(){
	window.history.back();
//could initialize Response here
;
//On //Error //GoTo //err_btnReturn;
;
//Exit_Return:;
    //DoCmd.Close;
    //Exit //Sub;
//err_btnReturn:;
    //MsgBox "Error in return function.   " //& //Err.Description;
//Resume //Exit_Return;
;
}
function cmdReminders_Click(){
//DoCmd.Close;
location.href = "Form_frmEducationModulesDue.php";;
;
}
function cndMaintain_Click(){
//DoCmd.Close;
    //FromMaintenanceFormSwitch = //True;
    location.href = "Form_AssignModule.php";//, //acNormal;
    //Forms![AssignModule].btnEdit.Visible = //True;
    //Forms![AssignModule].btnAdd.Visible = //True;
    //Forms![AssignModule].btnDelete.Visible = //True;
    //Forms![AssignModule].btnContact.Visible = //True;
    //Forms![AssignModule].btnShow.Visible = //False;
    //Forms![AssignModule].btnSelect.Visible = //False;
    //Forms![AssignModule].Caption = "Complete List of Available Modules";
    //Forms![AssignModule].lblModules.Caption = "Complete List of Available Modules";
    //Forms![AssignModule].lstAvailable.RowSourceType = "Table/Query";
    //Forms![AssignModule].lstAvailable.RowSource = "qryShowAllModules";
    //Forms![AssignModule].lstAvailable.SetFocus;
    //Forms![AssignModule].lstAvailable.Requery;
    ;
//On //Error //GoTo //err_cndMaintain_Click;
//Exit_cndMaintain_Click:;
    //Exit //Sub;
;
//err_cndMaintain_Click:;
    //MsgBox "Error opening Assign Module Form: " //& //Err.Number //& ",  " //& //Err.Description, //vbInformation;
    //Resume //Exit_cndMaintain_Click;
    ;
}
function GetHelp_Click(){
//could initialize returnvalue here
returnvalue;
//Exit_GetHelp_Click:;
    //Exit //Sub;
;
;
}

    </script>
  </head>
  <body>


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <label id=lblTitle value='Education Module/Tracking' style='position: absolute; left: 327px; top: 76px; width: 187px; height: 40;visibility:'>Education Module/Tracking</label>
    <input type='button' id=cmdReminders value='Prepare Reminders' style='position:absolute; left:258; top:171; width:325; height:48' onclick='cmdReminders_Click();'></input>
    <input type='button' id=cndMaintain value='Module Maintenance' style='position:absolute; left:258; top:114; width:325; height:48' onclick='cndMaintain_Click();'></input>
    <input type='button' id=btnReturn value='Return' style='position:absolute; left:585; top:240; width:132; height:48' onclick='btnReturn_Click();'></input>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='font-size:10px; position: absolute; left: 49px; top: 35px; width: 166; height: 25;visibility:'>by ProIRB Plus, Inc.</label>
    <label id=Label108 value='ProIRB' style='font-size:28px; position: absolute; left: 2px; top: 2px; width: 130; height: 43;visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='font-size:10px; position: absolute; left: 91px; top: 2px; width: 18; height: 16;visibility:'>R</label>
    <input type='button' id=GetHelp value='?' style='position: absolute; left: 611px; top: 143px; width: 57; height: 57' onclick='GetHelp_Click();'></input>
  </body>
</html>
