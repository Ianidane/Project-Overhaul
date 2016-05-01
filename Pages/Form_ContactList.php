<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function btnReturn_Click(){
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
function cmdADD_Click(){
//'Add module to CoreModulesRequiredPerType table
//could initialize rst here
//could initialize db here
;
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
//If //((Me.lstContactType = ""//)) //Then;
    //MsgBox //("The //contact //type //must //not //be //blank.");
    //Exit //Sub;
//End //If;
;
;
//Set db;
//Set rst;
  ;
//With rst;
    //.AddNew;
    //![module //Name] = //Me.txtModuleName;
    //If //Me.lstContactType //<> "" //Then;
      //![Contact //Type] = //Me.lstContactType;
    //End //If;
    //.Update;
    //.Close;
    //MsgBox "Module '" //& //Me.txtModuleName //& "' has been added to the " //_;
            //& "required list for the '" //& //Me.lstContactType //& "' contact type.";
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
        //MsgBox "Module '" //& //Me.txtModuleName //& "' already exists " //_;
            //& "for the '" //& //Me.lstContactType //& "'.";
        //Resume //Exit_cmdAdd_Click;
    //End //If;
    //MsgBox //Err.Description;
//Resume //Exit_cmdAdd_Click;
;
}

    </script>
  </head>
  <body>


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <label id=lblTitle value='ADD CORE MODULE TO CONTACT TYPE' style='position:absolute; left:72; top:24; width:604; height:40; visibility:'>ADD CORE MODULE TO CONTACT TYPE</label>
    <input type='button' id=cmdAdd value='Add  Module to Contact Type' style='position:absolute; left:306; top:204; width:268; height:40' onclick='cmdAdd_Click();'></input>
    <input type='button' id=btnReturn value='Return' style='position:absolute; left:576; top:204; width:132; height:40' onclick='btnReturn_Click();'></input>
    <select name='lstContactType' Size='2' style='position:absolute; left:522; top:102; width:174; height:30'>
        <Option value='Board Member'>Board Member</option>
        <Option value='Coordinator'>Coordinator</option>
        <Option value='P.I.'>P.I.</option>
        <Option value='Other'>Other</option>
    </select>
    <input type='text' id=txtModuleName value='empty' style='position:absolute; left:48; top:102; width:450; height:30'></input>

    <label id=lblModuleName value='Module Name:' style='position:absolute; left:60; top:72; width:111; height:24; visibility:'>Module Name:</label>
  </body>
</html>
