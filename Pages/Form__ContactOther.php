<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function addother_Click(){
//On //Error //GoTo //err_addother_click;
//CancelUpdate = //False;
//Call //Checkfields;
//If //FieldError //Then //GoTo //Exit_Save;
//DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
//MsgBox "Other Contact data saved";
//DataSaved = //True;
//DoCmd.SetWarnings //False;
//DoCmd.SetWarnings //True;
//HoldOther = //Me![Common //Contact //ID];
;
;
//DoCmd.Close;
;
//Exit_Save:;
    //Exit //Sub;
//err_addother_click:;
//MsgBox "Error Save Data Function--Data Not Saved" //& " " //& //Err.Number //& " " //& //Err.Description, //vbInformation;
        //Resume //Exit_Save;
   ;
}
function btnEducationTraining_Click(){
//If //Me.Last_Name = " " //Then;
            //MsgBox "Last Name cannot be blank";
            //Exit //Sub;
        //End //If;
        //strContact = //Me.Common_Contact_ID;
        location.href = "Form_EducationTrainingForm.php";//, //acNormal;
        //Forms![EducationTrainingForm].txtName.Value = //Me.[display //name];
        //Forms![EducationTrainingForm].txtIRBName.Value = //Me.[IRB //Code];
        //Forms![EducationTrainingForm].txtContactType.Value = //Me.Contact_Type;
        //Forms![EducationTrainingForm].txtResearchEducation.Value = //Me.Education;
}
function DeleteButton_Click(){
//On //Error //GoTo //Err_DeleteButton_Click;
;
//DoCmd.SetWarnings //False;
    //DoCmd.RunCommand //acCmdDeleteRecord;
    ;
    //DoCmd.SetWarnings //True;
    //DoCmd.Close;
    //MsgBox "Other Contact record Deleted";
;
//Exit_DeleteButton_Click:;
    //Exit //Sub;
;
//Err_DeleteButton_Click:;
//If //Err.Number = //3200 //Then;
    //MsgBox "You cannot delete a Contact who is assigned to a Study(s).  Use Search 'Studies for a P.I.' to " //_;
     //& "determine the Study(s).";
     //Resume //Exit_DeleteButton_Click;
//Else;
    ;
    //MsgBox //Err.Description //& " " //& //Err.Number;
    //Resume //Exit_DeleteButton_Click;
//End //If;
}
function Form_BeforeUpdate(Cancel){
//could initialize Response here
//If //Me.NewRecord //And //Me.Last_Name = " " //Then;
    //Me.Undo;
    //Exit //Sub;
//End //If;
//If //CancelUpdate //Then;
      //CancelUpdate = //False;
      //Me.Undo;
      //Exit //Sub;
//Else;
    Response;
        //& " " //& "  Do you wish to Save the Changes?"//, //vbYesNo);
    //If Response;
        //Me.Undo;
        //MsgBox "The New Changes were not saved!"//, //vbInformation;
    //End //If;
//End //If;
;
}
function Form_Close(){
//Me.Filter = "";
//CancelUpdate = //False;
//SearchContacts = //False;
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
//Call //adhToggleControl(Me, //True);
    //Me.NavigationButtons = //True;
     //Me.AllowFilters = //True;
    //Me.AllowAdditions = //False;
    //Me.AllowEdits = //False;
    //Me.addother.Enabled = //False;
    //Me.DeleteButton.Enabled = //False;
    //Me.LabelReadOnly.Visible = //False;
    //Me.btnEducationTraining.Enabled = //False;
//Else;
//Me.PreviewReport.Enabled = //False;
//End //If;
;
;
//DataSaved = //False;
//CancelUpdate = //False;
//'Me.PreviewReport.Enabled = False
;
}
function Ignore_Click(){
//could initialize Response here
;
//On //Error //GoTo //err_return;
;
//CancelUpdate = //True;
//Exit_Return:;
    //DoCmd.Close;
    //Exit //Sub;
//err_return:;
    //MsgBox "Error in return function.   " //& //Err.Description;
//Resume //Exit_Return;
;
}
function PreviewReport_Click(){
//On //Error //GoTo //Err_PreviewReport_Click;
;
    var stDocName = '';
    //If //FilterOn //Then;
        //MyFilter = //Me.Filter;
    //Else;
        //MyFilter = "";
    //End //If;
;
;
    stDocName;
    //LetterName = "_OtherContacts";
    //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
    //DoCmd.Close //acForm, "_contactother";
;
//Exit_PreviewReport_Click:;
    //Exit //Sub;
;
//Err_PreviewReport_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_PreviewReport_Click;
}
function Checkfields(){
//FieldError = //False;
//If //IsNull(Me.Last_Name) //Or //Me.Last_Name = " " //Then;
    //MsgBox "Last Name cannot be blank";
    //FieldError = //True;
//End //If;
    ;
;
;
;
;
;
;
;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:210; height:96'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:102; height:120'>    </select>
    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <label id=Label1 value='\"Other\" Contact Data' style='position:absolute; left:468; top:3; width:469; height:40; visibility:'>\"Other\" Contact Data</label>
    <input type='text' id=Common Contact ID value='empty' style='position:absolute; left:12; top:145; width:118; height:24'></input>

    <label id=Label2 value='Contact ID' style='position:absolute; left:12; top:115; width:118; height:24; visibility:'>Contact ID</label>
    <input type='text' id=Company Name value='empty' style='position:absolute; left:352; top:262; width:382; height:30'></input>

    <label id=Label5 value='Company Name:' style='position:absolute; left:198; top:262; width:145; height:24; visibility:'>Company Name:</label>
    <input type='text' id=Last Name value='empty' style='position:absolute; left:352; top:106; width:0; height:30'></input>

    <label id=Label6 value='Last Name:' style='position:absolute; left:408; top:75; width:105; height:24; visibility:'>Last Name:</label>
    <input type='text' id=First Name value='empty' style='position:absolute; left:570; top:106; width:0; height:30'></input>

    <label id=Label7 value='First Name:' style='position:absolute; left:624; top:75; width:105; height:24; visibility:'>First Name:</label>
    <input type='text' id=Middle value='empty' style='position:absolute; left:837; top:106; width:60; height:30'></input>

    <label id=Label8 value='Middle:' style='position:absolute; left:834; top:75; width:70; height:24; visibility:'>Middle:</label>
    <input type='text' id=Suffix value='empty' style='position:absolute; left:927; top:106; width:78; height:30'></input>

    <label id=Label9 value='Suffix:' style='position:absolute; left:936; top:75; width:63; height:24; visibility:'>Suffix:</label>
    <input type='text' id=Street Address 1 value='empty' style='position:absolute; left:352; top:142; width:336; height:30'></input>

    <label id=Label10 value='Street Address 1:' style='position:absolute; left:187; top:142; width:156; height:24; visibility:'>Street Address 1:</label>
    <input type='text' id=Street Address 2 value='empty' style='position:absolute; left:352; top:178; width:336; height:30'></input>

    <label id=Label11 value='Street Address 2:' style='position:absolute; left:187; top:178; width:156; height:24; visibility:'>Street Address 2:</label>
    <input type='text' id=City value='empty' style='position:absolute; left:352; top:220; width:292; height:30'></input>

    <label id=Label12 value='City:' style='position:absolute; left:297; top:220; width:46; height:24; visibility:'>City:</label>
    <input type='text' id=State value='empty' style='position:absolute; left:711; top:219; width:126; height:30'></input>

    <label id=Label13 value='State:' style='position:absolute; left:648; top:219; width:60; height:24; visibility:'>State:</label>
    <input type='text' id=Zip value='empty' style='position:absolute; left:895; top:220; width:190; height:30'></input>

    <label id=Label14 value='Zip:' style='position:absolute; left:852; top:219; width:42; height:24; visibility:'>Zip:</label>
    <input type='text' id=Main Phone # value='empty' style='position:absolute; left:352; top:303; width:123; height:30'></input>

    <label id=Label16 value='Main Phone #:' style='position:absolute; left:210; top:303; width:133; height:24; visibility:'>Main Phone #:</label>
    <input type='text' id=Main # Ext value='empty' style='position:absolute; left:538; top:303; width:84; height:30'></input>

    <label id=Label17 value='Ext:' style='position:absolute; left:492; top:303; width:42; height:24; visibility:'>Ext:</label>
    <input type='text' id=Cell Phone value='empty' style='position:absolute; left:352; top:345; width:123; height:30'></input>

    <label id=Label18 value='Cell Phone:' style='position:absolute; left:237; top:345; width:106; height:24; visibility:'>Cell Phone:</label>
    <input type='text' id=Pager Number value='empty' style='position:absolute; left:886; top:303; width:199; height:30'></input>

    <label id=Label19 value='Pager:' style='position:absolute; left:816; top:303; width:64; height:24; visibility:'>Pager:</label>
    <input type='text' id=FaxNumber value='empty' style='position:absolute; left:685; top:303; width:123; height:30'></input>

    <label id=Label20 value='Fax:' style='position:absolute; left:636; top:303; width:45; height:24; visibility:'>Fax:</label>
    <input type='text' id=EmailAddress value='empty' style='position:absolute; left:799; top:261; width:286; height:30'></input>

    <label id=Label21 value='Email:' style='position:absolute; left:738; top:261; width:60; height:24; visibility:'>Email:</label>
    <label id=Label22 value='Person's Name:' style='position:absolute; left:174; top:105; width:169; height:24; visibility:'>Person's Name:</label>
    <!--     <select name='checkpicoord' Size='15' style='position:absolute; left:18; top:270; width:132; height:30'>
    <?php
        $serverName = $SVNM;
        $connectionInfo = array( "Database"=>$DB, "UID"=>$UID, "PWD"=>$PWD);
        $conn = sqlsrv_connect( $serverName, $connectionInfo);

        if( $conn ) {
        }
        else{
            echo "Connection could not be established.<br />";
            die( print_r( sqlsrv_errors(), true));
        }

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_PICoordinator].[PIInvestigatorID], [_PICoordinator].[PICoordinatorID] FROM _PICoordinator WHERE ((([_PICoordinator].[PIInvestigatorID])=[Forms]![_contact PI]![Common Contact ID]))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['PIInvestigatorID']."'".">".$row['PIInvestigatorID']."</Option>";
        }
    ?>
    </select> -->

    <label id=checkpicoord_Label value='checkpicoord' style='position:absolute; left:0; top:229; width:93; height:24; visibility:'>checkpicoord</label>
    <input type='checkbox' id=Active value='=True' style='position:absolute; left:27; top:190; width:0; height:0'></input>

    <label id=Label47 value='Active' style='position:absolute; left:50; top:187; width:63; height:24; visibility:'>Active</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:142; top:27; width:193; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='text' id=AlternatePH value='empty' style='position:absolute; left:352; top:384; width:123; height:30'></input>

    <label id=Label49 value='Alternate Ph:' style='position:absolute; left:223; top:384; width:120; height:24; visibility:'>Alternate Ph:</label>
    <input type='text' id=AlternateFax value='empty' style='position:absolute; left:627; top:384; width:123; height:30'></input>

    <label id=Label50 value='Alternate Fax:' style='position:absolute; left:492; top:384; width:127; height:24; visibility:'>Alternate Fax:</label>
    <input type='text' id=ForeignPH value='empty' style='position:absolute; left:859; top:343; width:226; height:30'></input>

    <label id=Label51 value='Foreign Ph:' style='position:absolute; left:744; top:343; width:106; height:24; visibility:'>Foreign Ph:</label>
    <select name='Label31' Size='6' style='position:absolute; left:774; top:144; width:52; height:24'>
        <Option value='Dr.'>Dr.</option>
        <Option value='Mr.'>Mr.</option>
        <Option value='Miss.'>Miss.</option>
        <Option value='Mrs.'>Mrs.</option>
        <Option value='Ms.'>Ms.</option>
        <Option value='ENS'>ENS</option>
        <Option value='LTJG'>LTJG</option>
        <Option value='LT'>LT</option>
        <Option value='LCDR'>LCDR</option>
        <Option value='CD"
                        "R'>CD"
                        "R</option>
        <Option value='CAPT'>CAPT</option>
        <Option value='RADM'>RADM</option>
        <Option value='MAJ'>MAJ</option>
        <Option value='LT COL'>LT COL</option>
        <Option value='COL'>COL</option>
        <Option value='Sirs'>Sirs</option>
        <Option value='sir or Madam'>sir or Madam</option>
    </select>
    <input type='text' id=Education value='empty' style='position:absolute; left:498; top:433; width:585; height:76'></input>

    <label id=Label53 value='Research Education:' style='position:absolute; left:306; top:433; width:187; height:24; visibility:'>Research Education:</label>
    <label id=Label54 value='Country:' style='position:absolute; left:492; top:343; width:79; height:23; visibility:'>Country:</label>
    <input type='text' id=Text55 value='empty' style='position:absolute; left:582; top:343; width:108; height:29'></input>
    <label id=Label57 value='Voice Mail:' style='position:absolute; left:762; top:384; width:103; height:24; visibility:'>Voice Mail:</label>
    <input type='text' id=Text58 value='empty' style='position:absolute; left:870; top:384; width:214; height:30'></input>
    <input type='text' id=Contact Type value='empty' style='position:absolute; left:12; top:85; width:144; height:22'></input>

    <label id=Label4 value='Type' style='position:absolute; left:54; top:55; width:51; height:24; visibility:'>Type</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:49; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <!--     <label id=LabelReadOnly value='Data is Read Only' style='position:absolute; left:591; top:48; width:234; height:24; visibility:hidden'>Data is Read Only</label> -->
  </body>
</html>
