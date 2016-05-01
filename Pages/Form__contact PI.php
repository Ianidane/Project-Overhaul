<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function addPI_Click(){
//On //Error //GoTo //err_addpi_click;
//CancelUpdate = //False;
//Call //Checkfields;
//If //FieldError //Then //GoTo //Exit_Save;
//DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
//MsgBox "PI data saved";
//DataSaved = //True;
//DoCmd.SetWarnings //False;
//If //IsNull(Me.PICoordCombo) //Then //GoTo //SkipPICoord;
//If //IsNull(Me.checkpicoord) //Then;
    //'MsgBox "Append It " & Me.PICoordCombo & Me.checkpicoord
    //DoCmd.OpenQuery "PI_AppendCoordinator";
//Else;
    //'MsgBox "Update It" & Me.PICoordCombo
    //DoCmd.OpenQuery "PI_UpdateCoordinator";
//End //If;
//SkipPICoord:;
//DoCmd.SetWarnings //True;
//HoldPI = //Me![Common //Contact //ID];
//Forms![newmenu]![PI] = //Me.Common_Contact_ID;
//Forms![newmenu]![PI].Requery;
//Exit_Save:;
//If //Not //SearchContacts //Then;
    //DoCmd.Close;
    //Exit //Sub;
//Else;
    //Exit //Sub;
//End //If;
//err_addpi_click:;
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
        ;
        location.href = "Form_EducationTrainingForm.php";//, //, //, //, //, //acDialog;
//'Forms![_Contact PI].Visible = False
}
function cmdUseDeptAddress_Click(){
//On //Error //GoTo //Err_cmdUseDeptAddress_Click;
;
    var stDocName = '';
;
    stDocName;
    //FromPIForm = //True;
    location.href = "Form_"+stDocName+".php";//, //acNormal, //, //, //, //acDialog;
;
//Exit_cmdUseDeptAddress_Click:;
    //Exit //Sub;
;
//Err_cmdUseDeptAddress_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdUseDeptAddress_Click;
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
    //MsgBox "P.I. record Deleted";
;
//Exit_DeleteButton_Click:;
    //Exit //Sub;
;
//Err_DeleteButton_Click:;
//If //Err.Number = //3200 //Then;
    //MsgBox "You cannot delete a P.I. who is assigned to a Study(s).  Use Search 'Studies for a P.I.' to " //_;
     //& "determine the Study(s).";
     //Resume //Exit_DeleteButton_Click;
//Else;
    ;
    //MsgBox //Err.Description //& " " //& //Err.Number;
    //Resume //Exit_DeleteButton_Click;
//End //If;
}
function Form_BeforeUpdate(Cancel){
//If //Me.NewRecord //And //Me.Last_Name = " " //Then;
    //Me.Undo;
    //Exit //Sub;
//End //If;
;
}
function Form_Close(){
//Me.Filter = "";
//CancelUpdate = //False;
//SearchContacts = //False;
}
function Form_Current(){
//If //Me.NewRecord //And //Not //IsNull(HoldPIName) //Then;
    //Me.Active.Enabled = //False;
   //Me.Last_Name = //HoldPIName;
//End //If;
//Me.PICoordCombo.Requery;
}
function Form_Filter(Cancel, FilterType){
//Me.FilterOn = //True;
//Me.Filter = "";
}
function Form_Load(){
//If //Not //GeneralLetterSwitch //Then //Me.[IRB //Code] = //Forms![newmenu]![Location];
}
function Form_Open(Cancel){
//If //SearchContacts //Then;
    //Me.addpi.Enabled = //True;
    //Me.AllowFilters = //True;
    //'Me.OrderBy = False
    //Me.AllowAdditions = //False;
    //Me.NavigationButtons = //True;
    //Me.DeleteButton.Enabled = //False;
    //Me.btnEducationTraining.Enabled = //False;
    //Me.LabelReadOnly.Visible = //False;
    //'Call adhToggleControl(Me, True)
    ;
    ;
//Else;
    //Me.PreviewReport.Enabled = //False;
//End //If;
//DataSaved = //False;
//CancelUpdate = //False;
;
;
}
function Ignore_Click(){
//Me.Undo;
;
//Exit_Return:;
    //DoCmd.Close;
    //Exit //Sub;
//err_return:;
    //MsgBox "Error in return function.   " //& //Err.Description;
//Resume //Exit_Return;
;
}
function PICoordCombo_NotInList(NewData, Response){
//MsgBox "Press 'Esc' to clear";
}
function PRD_BeforeUpdate(Cancel){
//'Coast added field to form and changed  query ExportNameSelect
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
    //LetterName = "_PIs";
    //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
    //DoCmd.RunCommand //acCmdZoom75;
    //'DoCmd.Close acForm, "_contact PI"
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
function AddCoordButton_Click(){
//On //Error //GoTo //Err_AddCoordButton_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, //stLinkCriteria, //acFormAdd, //acDialog;
    //If //IsNull(HoldCoord) //Then;
        //GoTo //Exit_AddCoordButton_Click;
    //Else;
        //Me.PICoordCombo = //HoldCoord;
        //Me.PICoordCombo.Requery;
    //End //If;
//Exit_AddCoordButton_Click:;
    //Exit //Sub;
;
//Err_AddCoordButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_AddCoordButton_Click;
    ;
}
function DataChanged(){//As //Boolean;
//could initialize intObjType here
;
intObjType;
//strObjName = //Application.CurrentObjectName;
//intObjState = //SysCmd(acSysCmdGetObjectState, //intObjType, //strObjName);
//If //intObjState //<> //0 //Then;
    //DataChanged = //True;
//Else;
    //DataChanged = //False;
    ;
//End //If;
;
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



    <label id=Label1 value='Investigator Data' style='position:absolute; left:471; top:6; width:463; height:40; visibility:'>Investigator Data</label>
    <!--     <input type='text' id=Common Contact ID value='empty' style='position:absolute; left:12; top:199; width:118; height:24'></input> -->

    <label id=Label2 value='PI System ID' style='position:absolute; left:12; top:169; width:118; height:24; visibility:'>PI System ID</label>
    <select name='Contact Type' Size='1' style='position:absolute; left:12; top:85; width:0; height:24'>
        <Option value='Co-Ordinator'>Co-Ordinator</option>
        <Option value='Sponsor'>Sponsor</option>
        <Option value='P.I.'>P.I.</option>
    </select>

    <label id=Label4 value='Type' style='position:absolute; left:54; top:55; width:51; height:24; visibility:'>Type</label>
    <input type='text' id=Company Name value='empty' style='position:absolute; left:13; top:259; width:513; height:30'></input>

    <label id=Label5 value='Department/Company Name' style='position:absolute; left:18; top:231; width:270; height:24; visibility:'>Department/Company Name</label>
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
    <input type='text' id=Zip value='empty' style='position:absolute; left:895; top:220; width:114; height:30'></input>

    <label id=Label14 value='Zip:' style='position:absolute; left:852; top:219; width:42; height:24; visibility:'>Zip:</label>
    <input type='text' id=Main Phone # value='empty' style='position:absolute; left:237; top:303; width:123; height:30'></input>

    <label id=Label16 value='Main Phone #:' style='position:absolute; left:94; top:303; width:133; height:24; visibility:'>Main Phone #:</label>
    <input type='text' id=Main # Ext value='empty' style='position:absolute; left:423; top:303; width:84; height:30'></input>

    <label id=Label17 value='Ext:' style='position:absolute; left:376; top:303; width:42; height:24; visibility:'>Ext:</label>
    <input type='text' id=Cell Phone value='empty' style='position:absolute; left:352; top:345; width:123; height:30'></input>

    <label id=Label18 value='Cell Phone:' style='position:absolute; left:237; top:345; width:106; height:24; visibility:'>Cell Phone:</label>
    <input type='text' id=Pager Number value='empty' style='position:absolute; left:771; top:303; width:232; height:30'></input>

    <label id=Label19 value='Pager:' style='position:absolute; left:700; top:303; width:64; height:24; visibility:'>Pager:</label>
    <input type='text' id=FaxNumber value='empty' style='position:absolute; left:570; top:303; width:123; height:30'></input>

    <label id=Label20 value='Fax:' style='position:absolute; left:520; top:303; width:45; height:24; visibility:'>Fax:</label>
    <input type='text' id=EmailAddress value='empty' style='position:absolute; left:799; top:259; width:0; height:30'></input>

    <label id=Label21 value='Email:' style='position:absolute; left:738; top:261; width:60; height:24; visibility:'>Email:</label>
    <label id=Label22 value='Investigator Name:' style='position:absolute; left:174; top:105; width:169; height:24; visibility:'>Investigator Name:</label>
    <select name='Specialty1' Size='27' style='position:absolute; left:352; top:423; width:204; height:30'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_select coordinators for irb#].[Common Contact ID], [_select coordinators for irb#].[display name], [_select coordinators for irb#].[Contact Type] FROM [_select coordinators for irb#] WHERE ((([_select coordinators for irb#].[Contact Type])=\"Coordinator\"))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select>

    <label id=Label28 value='Specialty 1:' style='position:absolute; left:234; top:423; width:109; height:24; visibility:'>Specialty 1:</label>
    <select name='Specialty2' Size='28' style='position:absolute; left:718; top:423; width:204; height:30'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [PISpecialty1].[Specialty] FROM PISpecialty1 ORDER BY [PISpecialty1].[Specialty]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Specialty']."'".">".$row['Specialty']."</Option>";
        }
    ?>
    </select>

    <label id=Label29 value='Specialty 2:' style='position:absolute; left:604; top:423; width:109; height:24; visibility:'>Specialty 2:</label>
    <!--     <select name='checkpicoord' Size='23' style='position:absolute; left:64; top:367; width:132; height:30'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [PISpecialty1].[Specialty] FROM PISpecialty1 ORDER BY [PISpecialty1].[Specialty]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Specialty']."'".">".$row['Specialty']."</Option>";
        }
    ?>
    </select> -->

    <label id=checkpicoord_Label value='checkpicoord' style='position:absolute; left:46; top:327; width:93; height:24; visibility:'>checkpicoord</label>
    <select name='PICoordCombo' Size='30' style='position:absolute; left:18; top:532; width:246; height:30'>
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
    </select>

    <label id=Common Contact ID_Label value='PI's Normal Coordinator:' style='position:absolute; left:18; top:502; width:214; height:24; visibility:'>PI's Normal Coordinator:</label>
    <input type='checkbox' id=Active value='=True' style='position:absolute; left:25; top:126; width:0; height:0'></input>

    <label id=Label47 value='Active?' style='position:absolute; left:48; top:123; width:73; height:24; visibility:'>Active?</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:127; top:21; width:193; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='text' id=AlternatePH value='empty' style='position:absolute; left:352; top:384; width:123; height:30'></input>

    <label id=Label49 value='Alternate Ph:' style='position:absolute; left:223; top:384; width:120; height:24; visibility:'>Alternate Ph:</label>
    <input type='text' id=AlternateFax value='empty' style='position:absolute; left:627; top:384; width:123; height:30'></input>

    <label id=Label50 value='Alternate Fax:' style='position:absolute; left:492; top:384; width:127; height:24; visibility:'>Alternate Fax:</label>
    <input type='text' id=ForeignPH value='empty' style='position:absolute; left:859; top:343; width:150; height:30'></input>

    <label id=Label51 value='Foreign Ph:' style='position:absolute; left:744; top:343; width:106; height:24; visibility:'>Foreign Ph:</label>
    <select name='Label31' Size='7' style='position:absolute; left:774; top:144; width:52; height:24'>
        <Option value='Co-Ordinator'>Co-Ordinator</option>
        <Option value='Sponsor'>Sponsor</option>
        <Option value='P.I.'>P.I.</option>
    </select>
    <input type='text' id=Education value='empty' style='position:absolute; left:546; top:474; width:508; height:93'></input>

    <label id=Label53 value='Research Education:' style='position:absolute; left:354; top:474; width:187; height:24; visibility:'>Research Education:</label>
    <label id=Label54 value='Country:' style='position:absolute; left:492; top:343; width:79; height:23; visibility:'>Country:</label>
    <input type='text' id=Text55 value='empty' style='position:absolute; left:582; top:343; width:108; height:29'></input>
    <label id=Label57 value='Voice Mail:' style='position:absolute; left:762; top:384; width:103; height:24; visibility:'>Voice Mail:</label>
    <input type='text' id=Text58 value='empty' style='position:absolute; left:870; top:384; width:138; height:30'></input>
    <input type='button' id=cmdUseDeptAddress value='Use Dept Address' style='position:absolute; left:535; top:259; width:190; height:30' onclick='cmdUseDeptAddress_Click();'></input>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:42; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <input type='text' id=PRD value='empty' style='position:absolute; left:39; top:454; width:153; height:22'></input>

    <label id=Label61 value='Projected Rotation Date' style='position:absolute; left:0; top:423; width:210; height:24; visibility:'>Projected Rotation Date</label>
    <!--     <label id=LabelReadOnly value='Data is Read Only' style='position:absolute; left:576; top:51; width:234; height:24; visibility:hidden'>Data is Read Only</label> -->
  </body>
</html>
