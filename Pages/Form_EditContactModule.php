<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function cmdUpdate_Click(){
//On //Error //GoTo //Err_cmdUpdate_Click;
//could initialize Msg, here
;
//DoCmd.SetWarnings //False;
;
//'Make sure module name is not blank
//If //((Me.txtModuleName = ""//) //Or //(IsNull(Me.txtModuleName))) //Then;
    //MsgBox //("The //module //name //must //not //be //blank.");
    //Exit //Sub;
//End //If;
;
//Msg = "Are you sure you want to change this module?";
//Style = //vbOKCancel //+ //vbQuestion //+ //vbDefaultButton1;
//Title = "Update Selection?";
//Response = //MsgBox(Msg, //Style, //Title);
    //If //Response = //vbOK //Then;
    //'Update the selected record in the ContactModule table
        //Set //db = //CurrentDb();
        //Set //rst = //db.OpenRecordset("contactmodule1");
        //With //rst;
            //.MoveFirst;
            //Do //Until //.EOF;
                //If //((![ContactID] = //strContactID) //_;
                    //And //(![module //Name] = //strModuleName)) //Then;
                    //.Edit;
                    //![module //Name] = //Me.txtModuleName;
                    //If //Me.txtFrequency //<> "" //Or //IsNull(Me.txtFrequency) //Then;
                        //![Frequency] = //Me.txtFrequency;
                    //End //If;
                    //If //Me.txtCertificationMethod //<> "" //Or //IsNull(Me.txtCertificationMethod) //Then;
                        //![Certification //Method] = //Me.txtCertificationMethod;
                    //End //If;
                    //If //Me.txtExpirationDate //<> "" //Or //IsNull(Me.txtExpirationDate) //Then;
                        //![Expiration //Date] = //Me.txtExpirationDate;
                    //End //If;
                    //If //Me.txtLastCertificationDate //<> "" //Or //IsNull(Me.txtLastCertificationDate) //Then;
                        //![Last //Certification] = //Me.txtLastCertificationDate;
                    //End //If;
                    //If //Me.txtRemarks //<> "" //Or //IsNull(Me.txtRemarks) //Then;
                        //![Remarks] = //Me.txtRemarks;
                    //End //If;
                    //If //Me.txtModuleDescription //<> "" //Or //IsNull(Me.txtModuleDescription) //Then;
                        //![Module //Description] = //Me.txtModuleDescription;
                    //End //If;
                        //If //Me.txtLastReminderDate //<> "" //Or //IsNull(Me.txtLastReminderDate) //Then;
                        //![Last //Reminder //Date] = //Me.txtLastReminderDate;
                    //End //If;
                    //If //Me.txtLastReminderName //<> "" //Or //IsNull(Me.txtLastReminderName) //Then;
                        //![Reminder //Letter //Name] = //Me.txtLastReminderName;
                    //End //If;
                    //If //Me.TxtCE_Credits //<> "" //Or //IsNull(Me.TxtCE_Credits) //Then;
                        //![CE_Credits] = //Me.TxtCE_Credits;
                    //End //If;
                    //.Update;
                //'Notify the user of the successfully added record
                    //MsgBox "Module " //& //strModuleName //& " has been updated for this contact.";
                    //End //If;
                //.MoveNext;
            //Loop;
            ;
            //.Close;
        //End //With;
    ;
        //db.Close;
    ;
        //'Redisplay the list with the updated module
        //Forms!EducationTrainingForm!lstContactTrainingModules.Requery;
    //End //If;
;
//DoCmd.SetWarnings //True;
    ;
//Exit_cmdUpdate_Click:;
    //Exit //Sub;
//Err_cmdUpdate_Click:;
    //If //Err.Number = //3022 //Then;
        //MsgBox "You can't add duplicates";
        //Resume //Exit_cmdUpdate_Click;
    //End //If;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdUpdate_Click;
;
}
function cmdClose_Click(){
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
function Form_Open(Cancel){
//strContactID = //strContact;
    //Forms![EditContactModule].txtModuleName = //strModuleName;
    //'Forms![EditContactModule].strModuleName = strModuleName
    //Forms![EditContactModule].txtModuleDescription = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(1);
    //Forms![EditContactModule].txtLastCertificationDate = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(2);
    //Forms![EditContactModule].txtExpirationDate = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(3);
    //Forms![EditContactModule].txtLastReminderDate = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(4);
    //Forms![EditContactModule].txtCertificationMethod = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(5);
    //Forms![EditContactModule].txtFrequency = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(6);
    //Forms![EditContactModule].txtLastReminderName = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(7);
    //Forms![EditContactModule].txtRemarks = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(8);
    //Forms![EditContactModule].TxtCE_Credits = //Forms![EducationTrainingForm].lstContactTrainingModules.Column(11);
    //Me.Refresh;
    ;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <input type='button' id=cmdUpdate value='&Change' style='position:absolute; left:502; top:444; width:102; height:0' onclick='cmdUpdate_Click();'></input>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:136; top:24; width:165; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='button' id=cmdClose value='&Return' style='position:absolute; left:606; top:444; width:102; height:0' onclick='cmdClose_Click();'></input>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:49; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <label id=lblModules value='Edit Module ' style='position:absolute; left:54; top:60; width:546; height:30; visibility:'>Edit Module </label>
    <input type='text' id=txtModuleName value='empty' style='position:absolute; left:138; top:108; width:162; height:30'></input>

    <label id=lblModule Name value='Module Name' style='position:absolute; left:12; top:108; width:84; height:42; visibility:'>Module Name</label>
    <input type='text' id=txtLastCertificationDate value='empty' style='position:absolute; left:558; top:168; width:162; height:30'></input>

    <label id=lblLast Certification Date value='Last Certification Date' style='position:absolute; left:420; top:168; width:120; height:60; visibility:'>Last Certification Date</label>
    <input type='text' id=txtExpirationDate value='empty' style='position:absolute; left:558; top:234; width:162; height:30'></input>

    <label id=lblExpiration Date value='Expiration Date' style='position:absolute; left:420; top:234; width:114; height:42; visibility:'>Expiration Date</label>
    <input type='text' id=txtModuleDescription value='empty' style='position:absolute; left:138; top:228; width:192; height:84'></input>

    <label id=lblDescription value='Module Description' style='position:absolute; left:12; top:228; width:108; height:48; visibility:'>Module Description</label>
    <select name='txtCertificationMethod' Size='1' style='position:absolute; left:138; top:168; width:162; height:30'>    </select>

    <label id=lblCertification Method value='Certification Method' style='position:absolute; left:12; top:168; width:120; height:42; visibility:'>Certification Method</label>
    <input type='text' id=txtRemarks value='empty' style='position:absolute; left:138; top:336; width:198; height:84'></input>

    <label id=lblRemarks value='Remarks' style='position:absolute; left:12; top:336; width:108; height:30; visibility:'>Remarks</label>
    <select name='lblFrequency' Size='4' style='position:absolute; left:420; top:108; width:96; height:24'>
        <Option value='0'>0</option>
        <Option value='1 year'>1 year</option>
        <Option value='2 years'>2 years</option>
        <Option value='3 years'>3 years</option>
        <Option value='4 years'>4 years</option>
        <Option value='5 years'>5 years</option>
        <Option value='6 years'>6 years</option>
        <Option value='7"
                        " years'>7"
                        " years</option>
        <Option value='8 years'>8 years</option>
        <Option value='9 years'>9 years</option>
        <Option value='10 years'>10 years</option>
    </select>
    <input type='text' id=txtLastReminderDate value='empty' style='position:absolute; left:558; top:294; width:162; height:30'></input>

    <label id=lblLastReminderDate value='Last Reminder Date' style='position:absolute; left:420; top:294; width:138; height:48; visibility:'>Last Reminder Date</label>
    <input type='text' id=txtLastReminderName value='empty' style='position:absolute; left:558; top:360; width:162; height:30'></input>

    <label id=lblLastReminderName value='Last Reminder Name' style='position:absolute; left:420; top:360; width:138; height:48; visibility:'>Last Reminder Name</label>
    <input type='text' id=TxtCE_Credits value='empty' style='position:absolute; left:138; top:444; width:174; height:30'></input>

    <label id=Label49 value='Ce_Credits' style='position:absolute; left:12; top:444; width:100; height:24; visibility:'>Ce_Credits</label>
  </body>
</html>
