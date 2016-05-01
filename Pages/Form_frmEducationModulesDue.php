<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function ButToday_Click(){
//'101019Me.UserDate.Value = Date
//Me.DatesCombo = //Date;
//Call //ReDisplay;
}
function cmdSelective_Click(){
//On //Error //GoTo //Err_cmdSelective_Click;
;
//If //Not //OKSwitch //Then;
    //MsgBox //("You //must //select //a //module //for //the //Contact //you //wish //a //report //on");
    //Exit //Sub;
//End //If;
//OKSwitch = //False;
    var stDocName = '';
;
    stDocName;
    ;
    //If //IsNull(Me.[frmEducationModulesDueSubForm].Form.Filter) //Or //Me.[frmEducationModulesDueSubForm].Form.Filter = "" //Then;
  //Forms!frmeducationmodulesDue.SendToContactID = //Forms!frmeducationmodulesDue.SendToContactID //'took out in to test
    //Forms!frmeducationmodulesDue.SendDisplayName = //Forms!frmeducationmodulesDue.SendDisplayName //'took out in to test
    ;
    //DoCmd.OpenReport //stDocName, //acPreview, //, "(ContactID) Like " //& //strQuote //& //Forms!frmeducationmodulesDue.SendToContactID //& //strQuote;
//'took out in to test
        ;
    //Else;
        //DoCmd.OpenReport //stDocName, //acPreview, //, //Me.[frmEducationModulesDueSubForm].Form.Filter //& "AND (contactid) like " //& //strQuote //& //Forms!frmeducationmodulesDue.SendToContactID //& //strQuote;
        ;
    //End //If;
//Exit_cmdSelective_Click:;
    //Exit //Sub;
;
//Err_cmdSelective_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdSelective_Click;
}
function Form_Close(){
//'RenewalLetterSwitch = False
}
function Form_Current(){
//Call //ReDisplay;
//Me.frmEducationModulesDueSubForm.Form.Requery;
;
//'If Demo Then
 //'   Me![frmEducationModulesDueSubForm].Form![received].Locked = True
  //'  Me![frmEducationModulesDueSubForm].Form![DateReceivedFollowUP].Locked = True
//'Else
 //'   If UserEdit Then
       //' Me![frmEducationModulesDueSubForm].Form![received].Locked = False
        //'Me![frmEducationModulesDueSubForm].Form![DateReceivedFollowUP].Locked = False
  //'  Else
   //'     Me![frmEducationModulesDueSubForm].Form![received].Locked = True
    //'    Me![frmEducationModulesDueSubForm].Form![DateReceivedFollowUP].Locked = True
    //'End If
//'End If
;
}
function Form_Load(){
//'101019Me.UserDate.Value = Date
//Me.DatesCombo = //Date;
//Me.txtReportHeading = "MODULES DUE TO EXPIRE BY--" //& //Me.DatesCombo;
 ;
}
function Form_Open(Cancel){
//Me.ReturnButton.SetFocus;
//OKSwitch = //False;
//Me.FUALL = //False;
;
//If //Forms![newmenu]![HoldReadOnly] //Or //AgendaDetail //Then;
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.[frmEducationModulesDueSubForm].Locked = //True;
    //Me.LettersButton.Enabled = //False;
    //Me.AllowDeletions = //False;
    ;
//End //If;
//Me.txtReportHeading = "MODULES DUE TO EXPIRE BY--" //& //Me.DatesCombo;
;
}
function LettersButton_Click(){
//On //Error //GoTo //Err_LettersButton_Click;
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
        //DoCmd.OpenQuery //("ExportEducationwithCreateTable");
        //DoCmd.SetWarnings //True;
        //CustomLetter = //True;
        location.href = "Form_"+stDocName+".php";//, //acNormal;
        ;
    //Else;
        //Exit //Sub;
    //End //If;
   ;
    ;
;
//Exit_LettersButton_Click:;
    //Exit //Sub;
;
//Err_LettersButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_LettersButton_Click;
;
;
;
;
}
function ReturnButton_Click(){
//DoCmd.Close;
}
function ShowOption_AfterUpdate(){
//'101019If Me.UserDate.Value = "12/31/2099" Then
//If //Me.DatesCombo.Value = "12/31/2099" //Then //'101019
//'101019        Me.UserDate.Value = Date
        //Me.DatesCombo = //Date;
//End //If;
//Call //ReDisplay;
}
function ReDisplay(){
//Me.frmEducationModulesDueSubForm.Form.Requery;
//'Forms![frmEducationModulesDue].[frmEducationModulesDueSubForm].Form.Requery
//Me.frmEducationModulesDueSubForm.Form.FilterOn = //True;
//'Forms![frmEducationModulesDue].[frmEducationModulesDueSubForm].Form.FilterOn = True
//Select //Case //Me.ShowOption;
 //Case //1 //'Show Open
     //Me.FUALL = //False;
    //Forms![newmenu]![HoldFU] = //True;
    //Me.[frmEducationModulesDueSubForm].Form.Filter = "(([Expiration Date])<=[Forms]![frmEducationModulesDue]![userdate])";
    //'MsgBox Me.UserDate
    //Me.txtReportHeading = "MODULES DUE TO EXPIRE BY--" //& //Me.DatesCombo;
 //Case //2  //'Show Open and Closed
    //Forms![newmenu]![HoldFU] = //True;
    //Me.FUALL = //False;
     //Me.[frmEducationModulesDueSubForm].Form.Filter = "(([Expiration Date])<=[Forms]![frmEducationModulesDue]![userdate]) and  ( not isnull([Last Reminder Date]))";
     //Me.txtReportHeading = "MODULES TO EXPIRE BY--" //& //Me.DatesCombo //& " FOR WHICH A LETTER HAS BEEN SENT";
 //Case //3 //'Show Past Due
    //Forms![newmenu]![HoldFU] = //True;
    //Me.FUALL = //False;
//'101019    Me.UserDate.Value = Now
    //Me.DatesCombo = //Now;
    //Me.[frmEducationModulesDueSubForm].Form.Filter = "(([Expiration Date]) <=[Forms]![frmEducationModulesDue]![userdate])";
    //Me.txtReportHeading = "MODULES WHICH ARE OR WILL BE PAST DUE BY--" //& //Me.DatesCombo;
//Case //4  //'Show all letters sent in file
    //Forms![newmenu]![HoldFU] = //True;
    //Me.FUALL = //True;
//'101019    Me.UserDate.Value = "12/31/2099"
    //Me.DatesCombo = "12/31/2099";
    //Me.[frmEducationModulesDueSubForm].Form.Filter = "";
    //Me.txtReportHeading = "All MODULES ON FILE FOR ACTIVE CONTACT(s) ";
     //'(([Expiration Date]) <=[Forms]![frmEducationModulesDue]![userdate])"
//End //Select;
 //Me.[frmEducationModulesDueSubForm].Form.Requery;
;
;
;
;
}
function UserDate_AfterUpdate(){
//'Me.DatesCombo = Me.UserDate
//'Call ReDisplay
//'End Sub
;
//Private //Sub //ReportButton_Click();
//On //Error //GoTo //Err_ReportButton_Click;
;
;
    var stDocName = '';
;
    stDocName;
    //'Forms!frmEducationModulesDue.SendDisplayName = "***"  commented for test
    ;
    ;
   //' DoCmd.OpenReport stDocName, acPreview, , Me.Filter   'put back in to test
    //If //IsNull(Me.[frmEducationModulesDueSubForm].Form.Filter) //Or //Me.[frmEducationModulesDueSubForm].Form.Filter = "" //Then //'put back in to test
        //DoCmd.OpenReport //stDocName, //acPreview;
    //Else //'put back in to test
        //DoCmd.OpenReport //stDocName, //acPreview, //, //Me.[frmEducationModulesDueSubForm].Form.Filter //'put back in to test
    //End //If   //'put back in to test
//Exit_ReportButton_Click:;
    //Exit //Sub;
;
//Err_ReportButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ReportButton_Click;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:210; height:96'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:102; height:120'>    </select>



    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <input type='button' id=LettersButton value='Send Reminder Letter(s)' style='position:absolute; left:675; top:564; width:277; height:45' onclick='LettersButton_Click();'></input>
    <input type='button' id=ReturnButton value='&Return' style='position:absolute; left:952; top:564; width:121; height:45' onclick='ReturnButton_Click();'></input>
    <input type='text' id=DatesCombo value='empty' style='position:absolute; left:517; top:190; width:180; height:34'></input>
    <label id=Label147 value='Will Expire On Or Before' style='position:absolute; left:523; top:121; width:169; height:64; visibility:'>Will Expire On Or Before</label>
    <label id=Label151 value='-----Completed-----' style='position:absolute; left:882; top:264; width:157; height:24; visibility:'>-----Completed-----</label>
    <label id=Label152 value='Education Modules (For Active Contacts) Due to Expire' style='position:absolute; left:175; top:12; width:781; height:40; visibility:'>Education Modules (For Active Contacts) Due to Expire</label>


    <label id=Label174 value='Modules (For Active Contacts) Display Options' style='position:absolute; left:21; top:90; width:490; height:28; visibility:'>Modules (For Active Contacts) Display Options</label>
    <input type='radio' style='position:absolute; left:25; top:128; width:0; height:0'></input>

    <label id=Label177 value='Show Only Due to Expire by Date Selected' style='position:absolute; left:52; top:126; width:444; height:28; visibility:'>Show Only Due to Expire by Date Selected</label>
    <input type='radio' style='position:absolute; left:25; top:164; width:0; height:0'></input>

    <label id=Label179 value='Show Due to Expire With Letter Sent' style='position:absolute; left:52; top:162; width:378; height:28; visibility:'>Show Due to Expire With Letter Sent</label>
    <input type='radio' style='position:absolute; left:25; top:200; width:0; height:0'></input>

    <label id=Label181 value='Show Past Expiration Date Only' style='position:absolute; left:52; top:198; width:334; height:28; visibility:'>Show Past Expiration Date Only</label>
    <input type='radio' style='position:absolute; left:25; top:235; width:0; height:0'></input>

    <label id=Label183 value='Show all (Active Contact) Modules on File' style='position:absolute; left:52; top:235; width:436; height:28; visibility:'>Show all (Active Contact) Modules on File</label>
    <input type='button' id=ButToday value='Switch to Today' style='position:absolute; left:529; top:226; width:163; height:40' onclick='ButToday_Click();'></input>
    <input type='button' id=ReportButton value='All Modules Listed Report' style='position:absolute; left:381; top:564; width:292; height:45' onclick='ReportButton_Click();'></input>
    <!--     <input type='checkbox' id=FUALL value='False' style='position:absolute; left:438; top:57; width:138; height:12'></input> -->

    <label id=Label191 value='Check190' style='position:absolute; left:461; top:54; width:109; height:28; visibility:'>Check190</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:76; top:39; width:166; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:43; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <!--     <input type='text' id=SendToContactID value='empty' style='position:absolute; left:262; top:48; width:118; height:27'></input> -->

    <label id=Label193 value='Text192:' style='position:absolute; left:118; top:48; width:96; height:28; visibility:'>Text192:</label>
    <!--     <input type='text' id=txtContactModuleID value='empty' style='position:absolute; left:858; top:18; width:118; height:27'></input> -->

    <label id=Label195 value='Text192:' style='position:absolute; left:714; top:18; width:96; height:28; visibility:'>Text192:</label>
    <input type='button' id=cmdSelective value='Modules for Selected Contact' style='position:absolute; left:48; top:564; width:331; height:45' onclick='cmdSelective_Click();'></input>
    <!--     <input type='text' id=SendDisplayName value='empty' style='position:absolute; left:168; top:12; width:126; height:18'></input> -->

    <label id=Label200 value='Text199:' style='position:absolute; left:24; top:12; width:96; height:28; visibility:'>Text199:</label>
    <input type='text' id=txtReportHeading value='empty' style='position:absolute; left:54; top:270; width:714; height:24'></input>
    <input type='button' id=cmdCalDate value='&Hire Date Calendar' style='position:absolute; left:696; top:192; width:33; height:33'></input>
  </body>
</html>
