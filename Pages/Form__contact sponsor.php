<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function addsponsor_Click(){
//On //Error //GoTo //Err_addsponsor_Click;
    //CancelUpdate = //False;
    //DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
    //'DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    //HoldSponsorName = //Me.Company_Name;
    //HoldSponsor = //Me.Common_Contact_ID;
    //MsgBox "Sponsor Data Saved";
    //DataSaved = //True;
    //Forms![newmenu]![Sponsor] = //Me.Common_Contact_ID;
   //Forms![newmenu]![Sponsor].Requery;
   //SearchContacts = //False;
    //DoCmd.Close;
    ;
;
//Exit_addsponsor_Click:;
    //Exit //Sub;
;
//Err_addsponsor_Click:;
//MsgBox "Error Save Data Function--Data Not Saved" //& " " //& //Err.Number //& " " //& //Err.Description, //vbInformation;
    //Resume //Exit_addsponsor_Click;
    ;
}
function DeleteButton_Click(){
//On //Error //GoTo //Err_DeleteButton_Click;
;
//DoCmd.SetWarnings //False;
    //DoCmd.RunCommand //acCmdDeleteRecord;
    ;
    //DoCmd.SetWarnings //True;
    //DoCmd.Close;
    //MsgBox "Sponsor record Deleted";
;
//Exit_DeleteButton_Click:;
    //Exit //Sub;
;
//Err_DeleteButton_Click:;
//If //Err.Number = //3200 //Then;
    //MsgBox "You cannot delete a Sponsor who is assigned to a Study(s).  Use Search 'Studies for Sponsor.' to " //_;
     //& "determine the Study(s).";
     //Resume //Exit_DeleteButton_Click;
//Else;
    ;
    //MsgBox //Err.Description //& " " //& //Err.Number;
    //Resume //Exit_DeleteButton_Click;
//End //If;
;
}
function Form_BeforeUpdate(Cancel){
//could initialize Response here
//If //Me.NewRecord //And //Me.Company_Name = " " //Then;
    //Me.Undo;
    //Exit //Sub;
//End //If;
//If //CancelUpdate //Then;
    //Cancel = //True;
    ;
       ;
//End //If;
;
;
}
function Form_Current(){
//If //Me.NewRecord //Then;
    //Me.Company_Name = //HoldSponsorName;
//End //If;
}
function Form_Filter(Cancel, FilterType){
//Me.FilterOn = //True;
//Me.Filter = "";
}
function Form_Load(){
//Me.[IRB //Code] = //Forms![newmenu]![Location];
}
function Form_Open(Cancel){
//If //SearchContacts //Then;
    //Me.AllowFilters = //True;
    //Me.AllowEdits = //False;
    //Me.AllowAdditions = //False;
    //Me.NavigationButtons = //True;
    //Me.addsponsor.Enabled = //False;
    //Call //adhToggleControl(Me, //True);
    ;
    //Me.LabelReadOnly.Visible = //False;
    //Me.DeleteButton.Enabled = //False;
    ;
//Else;
    //Me.PreviewReport.Enabled = //False;
//End //If;
}
function Ignore_Click(){
//On //Error //GoTo //Err_Ignore_Click;
//CancelUpdate = //True;
//DoCmd.Close;
;
;
;
//Exit_Ignore_Click:;
    //Exit //Sub;
;
//Err_Ignore_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Ignore_Click;
    ;
}
function PreviewReport_Click(){
//On //Error //GoTo //Err_PreviewReport_Click;
;
    var stDocName = '';
;
    stDocName;
    //LetterName = "_ContactReportSponsors";
    //If //Me.FilterOn //Then;
        //MyFilter = //Me.Filter;
    //Else;
        //MyFilter = "";
    //End //If;
    //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
;
//Exit_PreviewReport_Click:;
    //Exit //Sub;
;
//Err_PreviewReport_Click:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_PreviewReport_Click;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();Form_Current();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:210; height:96'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:102; height:120'>    </select>
    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>


    <label id=Label1 value='Sponsor Data' style='position:absolute; left:564; top:6; width:246; height:40; visibility:'>Sponsor Data</label>
    <input type='text' id=Common Contact ID value='empty' style='position:absolute; left:18; top:190; width:114; height:24'></input>

    <label id=Label2 value='System ID' style='position:absolute; left:18; top:160; width:100; height:24; visibility:'>System ID</label>
    <select name='Contact Type' Size='1' style='position:absolute; left:18; top:127; width:0; height:24'>
        <Option value='Co-Ordinator'>Co-Ordinator</option>
        <Option value='Sponsor'>Sponsor</option>
        <Option value='P.I.'>P.I.</option>
    </select>

    <label id=Label4 value='Type' style='position:absolute; left:60; top:97; width:51; height:24; visibility:'>Type</label>
    <input type='text' id=Company Name value='empty' style='position:absolute; left:346; top:96; width:609; height:30'></input>

    <label id=Label5 value='Company Name:' style='position:absolute; left:190; top:96; width:145; height:24; visibility:'>Company Name:</label>
    <input type='text' id=Last Name value='empty' style='position:absolute; left:346; top:189; width:0; height:30'></input>

    <label id=Label6 value='Last Name:' style='position:absolute; left:408; top:159; width:105; height:24; visibility:'>Last Name:</label>
    <input type='text' id=First Name value='empty' style='position:absolute; left:570; top:189; width:0; height:30'></input>

    <label id=Label7 value='First Name:' style='position:absolute; left:624; top:159; width:105; height:24; visibility:'>First Name:</label>
    <input type='text' id=Middle value='empty' style='position:absolute; left:810; top:189; width:60; height:30'></input>

    <label id=Label8 value='Middle:' style='position:absolute; left:804; top:159; width:70; height:24; visibility:'>Middle:</label>
    <input type='text' id=Suffix value='empty' style='position:absolute; left:886; top:189; width:78; height:30'></input>

    <label id=Label9 value='Suffix:' style='position:absolute; left:892; top:159; width:63; height:24; visibility:'>Suffix:</label>
    <input type='text' id=Street Address 1 value='empty' style='position:absolute; left:346; top:226; width:336; height:30'></input>

    <label id=Label10 value='Street Address 1:' style='position:absolute; left:180; top:226; width:156; height:24; visibility:'>Street Address 1:</label>
    <input type='text' id=Street Address 2 value='empty' style='position:absolute; left:346; top:262; width:336; height:30'></input>

    <label id=Label11 value='Street Address 2:' style='position:absolute; left:180; top:262; width:156; height:24; visibility:'>Street Address 2:</label>
    <input type='text' id=City value='empty' style='position:absolute; left:346; top:298; width:255; height:30'></input>

    <label id=Label12 value='City:' style='position:absolute; left:289; top:298; width:46; height:24; visibility:'>City:</label>
    <input type='text' id=State value='empty' style='position:absolute; left:672; top:298; width:120; height:30'></input>

    <label id=Label13 value='State:' style='position:absolute; left:612; top:298; width:60; height:24; visibility:'>State:</label>
    <input type='text' id=Zip value='empty' style='position:absolute; left:852; top:298; width:114; height:30'></input>

    <label id=Label14 value='Zip:' style='position:absolute; left:804; top:298; width:42; height:24; visibility:'>Zip:</label>
    <input type='text' id=Main Phone # value='empty' style='position:absolute; left:346; top:354; width:123; height:30'></input>

    <label id=Label16 value='Main Phone #:' style='position:absolute; left:202; top:354; width:133; height:24; visibility:'>Main Phone #:</label>
    <input type='text' id=Main # Ext value='empty' style='position:absolute; left:534; top:354; width:87; height:30'></input>

    <label id=Label17 value='Ext:' style='position:absolute; left:478; top:354; width:42; height:24; visibility:'>Ext:</label>
    <input type='text' id=Cell Phone value='empty' style='position:absolute; left:346; top:390; width:123; height:30'></input>

    <label id=Label18 value='Cell Phone:' style='position:absolute; left:229; top:390; width:106; height:24; visibility:'>Cell Phone:</label>
    <input type='text' id=Pager Number value='empty' style='position:absolute; left:552; top:391; width:318; height:30'></input>

    <label id=Label19 value='Pager:' style='position:absolute; left:478; top:391; width:64; height:24; visibility:'>Pager:</label>
    <input type='text' id=FaxNumber value='empty' style='position:absolute; left:774; top:354; width:123; height:30'></input>

    <label id=Label20 value='Fax:' style='position:absolute; left:721; top:354; width:45; height:24; visibility:'>Fax:</label>
    <input type='text' id=EmailAddress value='empty' style='position:absolute; left:346; top:427; width:301; height:30'></input>

    <label id=Label21 value='Email:' style='position:absolute; left:276; top:427; width:60; height:24; visibility:'>Email:</label>
    <label id=Label22 value='Contact Name:' style='position:absolute; left:201; top:189; width:135; height:24; visibility:'>Contact Name:</label>
    <input type='button' id=addsponsor value='&Save Data' style='position:absolute; left:700; top:502; width:133; height:48' onclick='addsponsor_Click();'></input>
    <input type='button' id=Ignore value='&Return' style='position:absolute; left:834; top:502; width:133; height:48' onclick='Ignore_Click();'></input>
    <input type='button' id=PreviewReport value='Preview\015\012& Report' style='position:absolute; left:565; top:502; width:133; height:48' onclick='PreviewReport_Click();'></input>
    <input type='checkbox' id=Active value='=True' style='position:absolute; left:33; top:228; width:0; height:0'></input>

    <label id=Label30 value='Active' style='position:absolute; left:56; top:225; width:63; height:24; visibility:'>Active</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:34; top:39; width:193; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <select name='Label31' Size='9' style='position:absolute; left:736; top:229; width:52; height:24'>
        <Option value='Co-Ordinator'>Co-Ordinator</option>
        <Option value='Sponsor'>Sponsor</option>
        <Option value='P.I.'>P.I.</option>
    </select>
    <input type='button' id=DeleteButton value='Delete Sponsor' style='position:absolute; left:396; top:502; width:169; height:48' onclick='DeleteButton_Click();'></input>
    <input type='text' id=Country value='empty' style='position:absolute; left:774; top:427; width:189; height:30'></input>

    <label id=Label32 value='Country:' style='position:absolute; left:687; top:427; width:79; height:24; visibility:'>Country:</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:42; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <!--     <label id=LabelReadOnly value='Data is Read Only' style='position:absolute; left:576; top:54; width:234; height:24; visibility:hidden'>Data is Read Only</label> -->
  </body>
</html>
