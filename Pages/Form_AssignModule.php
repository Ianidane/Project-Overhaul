<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function btnAdd_Click(){
//'Add a new module to the Core Module table
//On //Error //GoTo //Err_btnAdd_Click;
;
location.href = "Form_EditCoreModule.php";//, //acNormal;
//Forms![EditCoreModule].lblModules.Caption = "Add Modules";
//Forms![EditCoreModule].Caption = "Add Modules";
//Forms![EditCoreModule].btnEdit.Enabled = //False;
//Forms![EditCoreModule].cmdAdd.Enabled = //True;
//Forms![EditCoreModule].txtModuleName.Locked = //False;
    ;
//Exit_btnAdd_Click:;
    //Exit //Sub;
//Err_btnAdd_Click:;
    //MsgBox //Err.Description;
//Resume //Exit_btnAdd_Click;
;
;
}
function btnClose_Click(){
//'Close form and return to AreaDetails Form
//On //Error //GoTo //Err_cmdClose_Click;
;
    //DoCmd.Close;
;
//Exit_cmdClose_Click:;
    //Exit //Sub;
;
//Err_cmdClose_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdClose_Click;
;
}
function btnContact_Click(){
//On //Error //GoTo //Err_btnContact_Click;
;
    //If //Me.lstAvailable.ItemsSelected.Count = //0 //Then;
        //If //MsgBox("You //must //make //a //selection", //vbOKOnly, "Selection Required"//) = //vbOK //Then;
            //Exit //Sub;
        //End //If;
    //End //If;
;
location.href = "Form_ContactList.php";//, //acNormal;
//Forms![ContactList].txtModuleName = //Me.lstAvailable.Column(1);
    ;
//Exit_btnContact_Click:;
    //Exit //Sub;
//Err_btnContact_Click:;
    //MsgBox //Err.Description;
//Resume //Exit_btnContact_Click;
;
}
function btnDelete_Click(){
//On //Error //GoTo //Err_btnDelete_Click;
//could initialize Msg, here
//could initialize db here
//could initialize rstCoreModules here
//could initialize Found here
;
//DoCmd.SetWarnings //False;
;
//If //Me.lstAvailable.ItemsSelected.Count = //0 //Then;
    //If //MsgBox("You //must //select //a //row //to //delete.", //vbOKOnly, "Selection Required"//) = //vbOK //Then;
            //Exit //Sub;
    //End //If;
//End //If;
;
//strModuleName = //Me.lstAvailable.Column(1);
//strContactID = //Me.lstAvailable.Column(0);
;
//Msg = "Are you sure you want to remove " //& //strModuleName //& "?";
//Style = //vbOKCancel //+ //vbQuestion //+ //vbDefaultButton1;
//Title = "Delete Selection?";
//Response = //MsgBox(Msg, //Style, //Title);
    //If //Response = //vbOK //Then;
//'Delete the selected record from the CoreModule and CoreModuleRequiredPerType tables
        //Set db;
        //Set //rstRequiredModules = //db.OpenRecordset("CoreModulesRequiredPerType1");
        Found;
        //With //rstRequiredModules;
            //If //.RecordCount = //0 //Then //GoTo //DeleteCore;
            //.MoveFirst;
            //Do //Until //.EOF;
               //If //((![module //Name] = //strModuleName) //And //_;
                    //(![Contact //Type] = //strContactID)) //Then;
                    //.Delete;
                    Found;
                     //MsgBox "Module " //& //strModuleName //_;
                    //& " has been deleted.";
               //End //If;
               //.MoveNext;
            //Loop;
            //.Close;
        //End //With;
        ;
        //'if the module is not found in the CoreModulePerType table then
        //'delete it from the CoreModule table
//DeleteCore:;
        //If //Not //(Found) //Then;
            //Set rstCoreModules;
            //With rstCoreModules;
                //.MoveFirst;
                //Do //Until //.EOF;
                    //If //((![module //Name] = //strModuleName)) //Then;
                        //.Delete;
                        //MsgBox "Module " //& //strModuleName //_;
                        //& " has been permanently deleted.";
                    //End //If;
                    //.MoveNext;
                //Loop;
                //.Close;
            //End //With;
        //End //If;
        ;
        //db.Close;
            ;
    //Else;
        //Exit //Sub;
    //End //If;
;
//Me.lstAvailable.Requery;
//DoCmd.SetWarnings //True;
;
//Exit_DeleteButton_Click:;
    //Exit //Sub;
;
//Err_btnDelete_Click:;
    //MsgBox //Err.Description //& " " //& //Err.Number;
    //Resume //Exit_DeleteButton_Click;
;
}
function btnEdit_Click(){
//On //Error //GoTo //Err_btnEdit_Click;
;
    //If //Me.lstAvailable.ItemsSelected.Count = //0 //Then;
        //If //MsgBox("You //must //make //a //selection", //vbOKOnly, "Selection Required"//) = //vbOK //Then;
            //Exit //Sub;
        //End //If;
    //End //If;
;
location.href = "Form_EditCoreModule.php";//, //acNormal;
//Forms![EditCoreModule].lblModules.Caption = "Edit Modules";
//Forms![EditCoreModule].Caption = "Edit Modules";
//Forms![EditCoreModule].btnEdit.Enabled = //True;
//Forms![EditCoreModule].cmdAdd.Enabled = //False;
//Forms![EditCoreModule].txtModuleName.Locked = //False;
//Forms![EditCoreModule].txtModuleName = //Me.lstAvailable.Column(1);
//Forms![EditCoreModule].txtModuleNameOrig = //Me.lstAvailable.Column(1);
//Forms![EditCoreModule].txtModuleDescription = //Me.lstAvailable.Column(2);
//Forms![EditCoreModule].txtCertificationDescription = //Me.lstAvailable.Column(5);
//Forms![EditCoreModule].txtCertificationMethod = //Me.lstAvailable.Column(4);
//Forms![EditCoreModule].txtFrequency = //Me.lstAvailable.Column(3);
;
;
//Forms![EditCoreModule].TxtCE_Credits = //Me.lstAvailable.Column(6);
//Exit_btnEdit_Click:;
    //Exit //Sub;
//Err_btnEdit_Click:;
    //MsgBox //Err.Description;
//Resume //Exit_btnEdit_Click;
;
}
function btnSelect_Click(){
//'Add selection to contact module table
;
//On //Error //GoTo //Err_btnSelect_Click;
;
    //DoCmd.SetWarnings //False;
    //If //Me.lstAvailable.ItemsSelected.Count = //0 //Then;
        //If //MsgBox("You //must //make //a //selection", //vbOKOnly, "Selection Required"//) = //vbOK //Then;
            //Exit //Sub;
        //End //If;
    //End //If;
    ;
//'Add the selected record to the ContactModule table
    //Set //db = //CurrentDb();
    //Set //rst = //db.OpenRecordset("contactmodule1");
    //With //rst;
        //.AddNew;
        //![ContactID] = //strContact;
        //![module //Name] = //lstAvailable.Column(1);
        //![Module //Description] = //lstAvailable.Column(2);
        //![Frequency] = //lstAvailable.Column(3);
        //![Certification //Method] = //lstAvailable.Column(4);
        //.Update;
        //.Close;
    //End //With;
    ;
    //db.Close;
    ;
//'Redisplay the list with the newly selected item
    //Forms!EducationTrainingForm!lstContactTrainingModules.Requery;
//'Notify the user of the successfully added record
    //MsgBox "Module " //& //lstAvailable.Column(1) //& " has been added to:   '" //_;
           //& //Forms![EducationTrainingForm]![txtName];
//DoCmd.SetWarnings //True;
    ;
//Exit_btnSelect_Click:;
    //Exit //Sub;
    ;
//Err_btnSelect_Click:;
    //If //Err.Number = //3022 //Then;
        //MsgBox "You can't add duplicates";
        //Resume //Exit_btnSelect_Click;
    //End //If;
    //MsgBox //Err.Description;
    //Resume //Exit_btnSelect_Click;
;
}
function btnShow_Click(){
//'Me.btnSelect.Enabled = False
    //Me.btnEdit.Enabled = //True;
    //Me.btnAdd.Enabled = //True;
    //Me.btnDelete.Enabled = //True;
    //Me.btnContact.Enabled = //True;
    //Me.Caption = "Complete List of Available Modules";
    //Me.lblModules.Caption = "Complete List of Available Modules";
    //Me.lstAvailable.RowSourceType = "Table/Query";
    //Me.lstAvailable.RowSource = "qryShowAllModules";
    //Me.lstAvailable.SetFocus;
    //Me.lstAvailable.Requery;
    //Me.btnShow.Enabled = //False;
;
}
function Command42_Click(){
//End //Sub;
;
//Private //Sub //Form_Activate();
//'requery lists after add or delete
//On //Error //GoTo //Err_Form_Activate;
 ;
//Me.lstAvailable.Requery;
;
//Exit_Form_Activate:;
    //Exit //Sub;
//Err_Form_Activate:;
    //MsgBox //Err.Description;
//Resume //Exit_Form_Activate;
;
}
function Form_Close(){
//Me.btnShow.Enabled = //False;
    //Me.btnEdit.Enabled = //False;
    //Me.btnAdd.Enabled = //False;
    //Me.btnDelete.Enabled = //False;
    //Me.btnContact.Enabled = //False;
}
function Form_Open(Cancel){
//If //FromMaintenanceFormSwitch //Then;
    //FromMaintenanceFormSwitch = //False;
//Else;
    //Forms![AssignModule].btnSelect.Visible = //True;
    //Forms![AssignModule].btnShow.Visible = //True;
    //Forms![AssignModule].btnEdit.Visible = //False;
    //Forms![AssignModule].btnAdd.Visible = //False;
    //Forms![AssignModule].btnDelete.Visible = //False;
    //Forms![AssignModule].btnContact.Visible = //False;
    //Forms![AssignModule].lblModules.Caption = "Modules Available for this Contact";
    //Forms![AssignModule].Caption = "Modules Available for this Contact";
    //Forms![AssignModule].lstAvailable.RowSourceType = "Table/Query";
    //Forms![AssignModule].lstAvailable.RowSource = "qryAssignModules";
   ;
//End //If;
;
}
function lstAvailable_AfterUpdate(){
//'requery lists after change
//On //Error //GoTo //Err_lstAvailable_AfterUpdate;
;
//Me.lstAvailable.Requery;
;
//Exit_lstAvailable_AfterUpdate:;
   //Exit //Sub;
//Err_lstAvailable_AfterUpdate:;
    //MsgBox //Err.Description;
//Resume //Exit_lstAvailable_AfterUpdate;
;
}
function lstAvailable_DblClick(Cancel){
//'    If (Me.btnEdit.Enabled) Then
 //'       btnEdit_Click
  //'  End If
    ;
}

    </script>
  </head>
  <body onLoad="Form_Open();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <select name='lstAvailable' style='position:absolute; left:18; top:114; width:954; height:264'>    </select>
    <input type='button' id=btnSelect value='Add Module to Selected Contact' style='position:absolute; left:6; top:379; width:301; height:40' onclick='btnSelect_Click();'></input>
    <input type='button' id=btnShow value='Show All Core Modules' style='position:absolute; left:306; top:379; width:220; height:40' onclick='btnShow_Click();'></input>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:48; top:46; width:193; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='button' id=btnClose value='&Return' style='position:absolute; left:847; top:478; width:93; height:40' onclick='btnClose_Click();'></input>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:49; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <label id=lblModules value='Modules Available for this Contact' style='position:absolute; left:288; top:72; width:390; height:30; visibility:'>Modules Available for this Contact</label>
    <input type='button' id=btnAdd value='Add New Core Module' style='position:absolute; left:631; top:478; width:214; height:40' onclick='btnAdd_Click();'></input>
    <input type='button' id=btnDelete value='Delete Module As Contact Type Requirement' style='position:absolute; left:525; top:418; width:453; height:40' onclick='btnDelete_Click();'></input>
    <input type='button' id=btnEdit value='Edit Core Module' style='position:absolute; left:460; top:478; width:171; height:40' onclick='btnEdit_Click();'></input>
    <input type='button' id=btnContact value='Assign  Module As Contact Type Requirement' style='position:absolute; left:525; top:379; width:453; height:40' onclick='btnContact_Click();'></input>
    <label id=Label40 value='Education Core Module Maintenance' style='position:absolute; left:255; top:22; width:451; height:34; visibility:'>Education Core Module Maintenance</label>
  </body>
</html>
