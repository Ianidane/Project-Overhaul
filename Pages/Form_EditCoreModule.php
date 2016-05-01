<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function btnEdit_Click(){
//On //Error //GoTo //Err_btnEdit_Click;
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
//Msg = "Are you sure you want to change '" //& //Me.txtModuleName //& "' module?";
//Style = //vbOKCancel //+ //vbQuestion //+ //vbDefaultButton1;
//Title = "Update Selection?";
//Response = //MsgBox(Msg, //Style, //Title);
    //If //Response = //vbOK //Then;
    //'Update the selected record in the ContactModule and the CoreModulesRequiredPerType tables
        //Set //db = //CurrentDb();
        //Set //rst = //db.OpenRecordset("coremodules1");
        //With //rst;
            //.MoveFirst;
            //Do //Until //.EOF;
                //'MsgBox ![module Name]
                //If //((![module //Name] = //Me.txtModuleNameOrig)) //Then;
                    //.Edit;
                    //![module //Name] = //Me.txtModuleName;
                    //If //Me.txtFrequency //<> "" //Then;
                        //![Frequency] = //Me.txtFrequency;
                    //End //If;
                    //If //Me.txtCertificationMethod //<> "" //Then;
                        //![Certification //Method] = //Me.txtCertificationMethod;
                    //End //If;
                    //If //Me.txtCertificationDescription //<> "" //Then;
                        //![Certification //Description] = //Me.txtCertificationDescription;
                    //End //If;
                    //If //Me.txtModuleDescription //<> "" //Then;
                        //![Module //Description] = //Me.txtModuleDescription;
                    //End //If;
                    //If //Me.TxtCE_Credits //<> "" //Then;
                        //![Core_CE_Credits] = //Me.TxtCE_Credits;
                    //End //If;
                        ;
                    //.Update;
                //'Notify the user of the successfully added record
                    //MsgBox "Module " //& //Me.txtModuleName //& " has been updated.";
                //End //If;
                //.MoveNext;
            //Loop;
            ;
            //.Close;
        //End //With;
    ;
        //db.Close;
    ;
    //End //If;
;
//DoCmd.SetWarnings //True;
    ;
//Exit_btnEdit_Click:;
    //Exit //Sub;
//Err_btnEdit_Click:;
    //If //Err.Number = //3022 //Then;
        //MsgBox "You can't add duplicates";
        //Resume //Exit_btnEdit_Click;
    //End //If;
    //MsgBox //Err.Description;
    //Resume //Exit_btnEdit_Click;
;
}
function cmdADD_Click(){
//'Add module to CoreModules table
//On //Error //GoTo //Err_cmdAdd_Click;
;
//DoCmd.SetWarnings //False;
;
//'Make sure module name is not blank
//If //((Me.txtModuleName = ""//) //Or //(IsNull(Me.txtModuleName))) //Then;
    //MsgBox //("The //module //name //must //not //be //blank.");
    //Exit //Sub;
//End //If;
;
;
//Set //db = //CurrentDb();
//Set //rst = //db.OpenRecordset("coremodules1");
  ;
//With //rst;
    //.AddNew;
    //![module //Name] = //Me.txtModuleName;
    //If //Me.txtModuleDescription //<> "" //Then;
        //![Module //Description] = //Me.txtModuleDescription;
    //End //If;
    //If //Me.txtFrequency //<> "" //Then;
      //![Frequency] = //Me.txtFrequency;
    //End //If;
    //If //Me.txtCertificationMethod //<> "" //Then;
      //![Certification //Method] = //Me.txtCertificationMethod;
    //End //If;
    //If //Me.txtCertificationDescription //<> "" //Then;
      //![Certification //Description] = //Me.txtCertificationDescription;
    //End //If;
    //.Update;
    //.Close;
    //MsgBox "Module '" //& //Me.txtModuleName //& "' has been added to the module list.";
//End //With;
;
//db.Close;
;
//DoCmd.SetWarnings //True;
;
//Exit_cmdAdd_Click:;
    //Exit //Sub;
//Err_cmdAdd_Click:;
//If //Err.Number = //3022 //Then;
        //MsgBox "Module '" //& //Me.txtModuleName //& "' is already exists.";
        //Resume //Exit_cmdAdd_Click;
    //End //If;
    //MsgBox //Err.Description;
//Resume //Exit_cmdAdd_Click;
;
}
function cmdClose_Click(){
//'Close form and return to AreaDetails Form
//On //Error //GoTo //Err_cmdClose_Click;
;
    //Me.txtCertificationDescription = "";
    //Me.txtCertificationMethod = "";
    //Me.txtFrequency = "";
    //Me.txtModuleDescription = "";
    //Me.txtModuleName = "";
    ;
    //'Redisplay the list with the updated module
    //Forms![AssignModule].lstAvailable.Requery;
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
//Me.txtModuleNameOrig = //Me.txtModuleName;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <input type='button' id=cmdAdd value='&Add' style='position:absolute; left:514; top:384; width:102; height:0' onclick='cmdAdd_Click();'></input>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:136; top:24; width:165; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='button' id=cmdClose value='&Return' style='position:absolute; left:618; top:384; width:102; height:0' onclick='cmdClose_Click();'></input>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:49; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <label id=lblModules value='Edit Module ' style='position:absolute; left:177; top:66; width:406; height:30; visibility:'>Edit Module </label>
    <input type='text' id=txtModuleName value='empty' style='position:absolute; left:133; top:132; width:510; height:36'></input>

    <label id=lblModule Name value='Module Name' style='position:absolute; left:294; top:102; width:147; height:24; visibility:'>Module Name</label>
    <input type='text' id=txtCertificationDescription value='empty' style='position:absolute; left:132; top:264; width:268; height:72'></input>

    <label id=lblCompletionDescription value='Certification Description' style='position:absolute; left:6; top:264; width:114; height:60; visibility:'>Certification Description</label>
    <select name='txtCertificationMethod' Size='2' style='position:absolute; left:129; top:198; width:268; height:30'>    </select>

    <label id=lblCertification Method value='Certification Method' style='position:absolute; left:3; top:198; width:120; height:42; visibility:'>Certification Method</label>
    <input type='text' id=txtModuleDescription value='empty' style='position:absolute; left:534; top:264; width:220; height:84'></input>

    <label id=lblDescription value='Module Description' style='position:absolute; left:408; top:264; width:108; height:48; visibility:'>Module Description</label>
    <select name='lblFrequency' Size='1' style='position:absolute; left:408; top:198; width:96; height:24'>
        <Option value='0nce'>0nce</option>
        <Option value='1 year'>1 year</option>
        <Option value='2 years'>2 years</option>
        <Option value='3 years'>3 years</option>
        <Option value='4 years'>4 years</option>
        <Option value='5 years'>5 years</option>
        <Option value='6 years'>6 years</option>
        <Option value='"
                        "7 years'>"
                        "7 years</option>
        <Option value='8 years'>8 years</option>
        <Option value='9 years'>9 years</option>
        <Option value='10 years'>10 years</option>
    </select>
    <input type='button' id=btnEdit value='Save' style='position:absolute; left:411; top:384; width:102; height:0' onclick='btnEdit_Click();'></input>
    <!--     <input type='text' id=txtModuleNameOrig value='empty' style='position:absolute; left:133; top:30; width:186; height:36'></input> -->

    <label id=Label40 value='Module Name' style='position:absolute; left:294; top:0; width:147; height:24; visibility:'>Module Name</label>
    <input type='text' id=TxtCE_Credits value='empty' style='position:absolute; left:132; top:360; width:0; height:42'></input>

    <label id=Label42 value='CE_Credits' style='position:absolute; left:6; top:360; width:102; height:24; visibility:'>CE_Credits</label>
  </body>
</html>
