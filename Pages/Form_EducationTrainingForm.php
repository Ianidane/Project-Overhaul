<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function btnAssign_Click(){
location.href = "Form_AssignModule.php";//, //, //, //, //, //acDialog;
    ;
//On //Error //GoTo //err_btnAssign_click;
//Exit_btnAssign_click:;
    //Exit //Sub;
;
//err_btnAssign_click:;
    //MsgBox "Error opening Assign Module Form: " //& //Err.Number //& ",  " //& //Err.Description, //vbInformation;
    //Resume //Exit_btnAssign_click;
   ;
}
function btnDelete_Click(){
//On //Error //GoTo //Err_btnDelete_Click;
//could initialize Msg, here
//could initialize db here
;
//DoCmd.SetWarnings //False;
;
//If //Me.lstContactTrainingModules.ItemsSelected.Count = //0 //Then;
    //If //MsgBox("You //must //make //a //selection", //vbOKOnly, "Selection Required"//) = //vbOK //Then;
            //Exit //Sub;
    //End //If;
//End //If;
;
//strModuleName = //Me.lstContactTrainingModules.Column(0);
//strContactID = //strContact;
//Msg = "Are you sure you want to remove " //& //strModuleName //& "?";
//Style = //vbOKCancel //+ //vbQuestion //+ //vbDefaultButton1;
//Title = "Delete Selection?";
//Response = //MsgBox(Msg, //Style, //Title);
    //If //Response = //vbOK //Then;
//'Delete the selected record from the ContactModule table
        //Set db;
        //Set //rst = //db.OpenRecordset("contactmodule1");
        //With //rst;
            //.MoveFirst;
            //Do //Until //.EOF;
                //If //((![ContactID] = //strContactID) //_;
                    //And //(![module //Name] = //Me.lstContactTrainingModules.Column(0))) //Then;
                    //.Delete;
                    //MsgBox "Module " //& //strModuleName //_;
                    //& " has been deleted for this contact.";
                //End //If;
                //.MoveNext;
            //Loop;
            //.Close;
        //End //With;
    ;
        //db.Close;
            ;
    //Else;
        //Exit //Sub;
    //End //If;
;
//Me.lstContactTrainingModules.Requery;
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
function btnEditModule_Click(){
//DoCmd.SetWarnings //False;
;
    //If //Me.lstContactTrainingModules.ItemsSelected.Count = //0 //Then;
        //If //MsgBox("You //must //make //a //selection", //vbOKOnly, "Selection Required"//) = //vbOK //Then;
            //Exit //Sub;
        //End //If;
    //End //If;
    ;
    //strModuleName = //Me.lstContactTrainingModules.Column(0);
    ;
    location.href = "Form_EditContactModule.php";//, //acNormal, //, //, //, //acDialog;
    ;
//On //Error //GoTo //err_btnAssign_click;
//Exit_btnAssign_click:;
    //Exit //Sub;
//DoCmd.SetWarnings //True;
;
//err_btnAssign_click:;
    //MsgBox "Error opening Assign Module Form: " //& //Err.Number //& ",  " //& //Err.Description, //vbInformation;
    //Resume //Exit_btnAssign_click;
;
}
function btnLetter_Click(){
//On //Error //GoTo //Err_btnLetter_Click;
//If //Demo //Then;
    //DemoMsg;
    //Exit //Sub;
//End //If;
;
//If //OKSwitch = //False //Then;
    //MsgBox "You must first select a Module";
    //Exit //Sub;
//End //If;
;
//LetterName = "Education Letter";
var stDocName = '';
stDocName;
//EducationLetterSwitch = //True;
    ;
    //If //GetTemplateLocation("Education") //Then;
        ;
        //If //TemplateName = "" //Then //Exit //Sub;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery //("ExportContactEducationModuleWCreateTable");
        //DoCmd.SetWarnings //True;
        //CustomLetter = //True;
        location.href = "Form_"+stDocName+".php";//, //, //, //, //, //acDialog;
        //DoCmd.Close //acForm, "EducationTrainingForm"//, //acSaveYes;
    //Else;
        //Exit //Sub;
    //End //If;
   ;
    ;
;
//Exit_btnLetter_Click:;
    //Exit //Sub;
;
//Err_btnLetter_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_btnLetter_Click;
;
}
function Form_Close(){
//' Probably should clear all fields before closing form
}
function btnReturn_Click(){
//could initialize Response here
;
//On //Error //GoTo //err_btnReturn;
;
//CancelUpdate = //True;
//Exit_Return:;
    //DoCmd.Close;
    //Exit //Sub;
//err_btnReturn:;
    //MsgBox "Error in return function.   " //& //Err.Description;
//Resume //Exit_Return;
;
}
function Form_Open(Cancel){
//If //IsLoaded("_contact //pi") //Then;
        //Forms![EducationTrainingForm].txtName.Value = //Forms![_Contact //PI].[display //name];
        //Forms![EducationTrainingForm].txtIRBName.Value = //Forms![_Contact //PI].[IRB //Code];
        //Forms![EducationTrainingForm].txtContactType.Value = //Forms![_Contact //PI].Contact_Type;
        //Forms![EducationTrainingForm].txtResearchEducation.Value = //Forms![_Contact //PI].Education;
//End //If;
//If //IsLoaded("_contact //coordinator") //Then;
        //Forms![EducationTrainingForm].txtName.Value = //Forms![_contact //coordinator].[display //name];
        //Forms![EducationTrainingForm].txtIRBName.Value = //Forms![_contact //coordinator].[IRB //Code];
        //Forms![EducationTrainingForm].txtContactType.Value = //Forms![_contact //coordinator].Contact_Type;
        //Forms![EducationTrainingForm].txtResearchEducation.Value = //Forms![_contact //coordinator].Education;
//End //If;
//If //IsLoaded("_contactBoardMember") //Then;
        //Forms![EducationTrainingForm].txtName.Value = //Forms![_contactBoardMember].[display //name];
        //Forms![EducationTrainingForm].txtIRBName.Value = //Forms![_contactBoardMember].[IRB //Code];
        //Forms![EducationTrainingForm].txtContactType.Value = //Forms![_contactBoardMember].Contact_Type;
        //Forms![EducationTrainingForm].txtResearchEducation.Value = //Forms![_contactBoardMember].Education;
//End //If;
//If //IsLoaded("_ContactOther") //Then;
        //Forms![EducationTrainingForm].txtName.Value = //Forms![_ContactOther].[display //name];
        //Forms![EducationTrainingForm].txtIRBName.Value = //Forms![_ContactOther].[IRB //Code];
        //Forms![EducationTrainingForm].txtContactType.Value = //Forms![_ContactOther].Contact_Type;
        //Forms![EducationTrainingForm].txtResearchEducation.Value = //Forms![_ContactOther].Education;
//End //If;
;
;
;
;
;
;
;
    //Me.txtContactID = //strContact;
    //Me.lstContactTrainingModules.Requery;
    //OKSwitch = //False;
;
}
function lstContactTrainingModules_Click(){
//'MsgBox Me.lstContactTrainingModules.Column(9)
//Me.HoldContactModuleID = //Me.lstContactTrainingModules.Column(10);
//OKSwitch = //True;
}
function lstContactTrainingModules_DblClick(Cancel){
//btnEditModule_Click;
}

    </script>
  </head>
  <body onLoad="Form_Open();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:210; height:96'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:102; height:120'>    </select>

    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>


    <label id=lblTitle value='Education Training' style='position:absolute; left:330; top:6; width:348; height:40; visibility:'>Education Training</label>
    <input type='text' id=txtName value='empty' style='position:absolute; left:24; top:132; width:336; height:30'></input>
    <label id=lblName value='Individual's Name' style='position:absolute; left:24; top:96; width:159; height:24; visibility:'>Individual's Name</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:142; top:27; width:193; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='text' id=txtResearchEducation value='empty' style='position:absolute; left:24; top:475; width:684; height:84'></input>

    <label id=lblResearchEducation value='Prior Versions Research Education' style='position:absolute; left:24; top:444; width:372; height:28; visibility:'>Prior Versions Research Education</label>
    <input type='text' id=txtContactType value='empty' style='position:absolute; left:372; top:132; width:174; height:30'></input>

    <label id=lblContact value='Contact Type' style='position:absolute; left:372; top:96; width:126; height:24; visibility:'>Contact Type</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:49; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <label id=lblIRBName value='IRB Name' style='position:absolute; left:558; top:96; width:144; height:24; visibility:'>IRB Name</label>
    <input type='text' id=txtIRBName value='empty' style='position:absolute; left:558; top:132; width:384; height:30'></input>
    <input type='button' id=btnDelete value='Delete Module' style='position:absolute; left:819; top:259; width:151; height:40' onclick='btnDelete_Click();'></input>
    <input type='button' id=btnReturn value='Return' style='position:absolute; left:819; top:526; width:151; height:40' onclick='btnReturn_Click();'></input>
    <input type='button' id=btnAssign value='&Assign Module' style='position:absolute; left:819; top:219; width:151; height:40' onclick='btnAssign_Click();'></input>
    <input type='button' id=btnEditModule value='&Edit Module' style='position:absolute; left:819; top:300; width:151; height:40' onclick='btnEditModule_Click();'></input>
    <input type='button' id=btnLetter value='Send Letter' style='position:absolute; left:819; top:340; width:124; height:40' onclick='btnLetter_Click();'></input>
    <select name='lstContactTrainingModules' Size='3' style='position:absolute; left:24; top:220; width:786; height:204' onclick='lstContactTrainingModules_Click();'>    </select>

    <label id=Contact Training Modules_Label value='Contact Training Modules' style='position:absolute; left:25; top:180; width:273; height:28; visibility:'>Contact Training Modules</label>
    <!--     <input type='text' id=txtContactID value='empty' style='position:absolute; left:780; top:78; width:84; height:30'></input> -->
    <!--     <input type='text' id=HoldContactModuleID value='empty' style='position:absolute; left:468; top:48; width:138; height:30'></input> -->

    <label id=Label72 value='Text71:' style='position:absolute; left:324; top:48; width:73; height:24; visibility:'>Text71:</label>
  </body>
</html>
