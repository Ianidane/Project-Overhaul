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
;
    //CancelUpdate = //False;
    //DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
    //MsgBox "Coordinator Data saved";
//DataSaved = //True;
;
    //HoldCoord = //Me![Common //Contact //ID];
    //Forms![newmenu]![CoOrd] = //Me.Common_Contact_ID;
   //Forms![newmenu]![CoOrd].Requery;
    //DoCmd.Close;
;
//Exit_addpi_Click:;
    //Exit //Sub;
;
//err_addpi_click:;
    //MsgBox "Error Save Data Function--Data Not Saved" //& " " //& //Err.Number //& " " //& //Err.Description, //vbInformation;
    //Resume //Exit_addpi_Click;
    ;
}
function btnEducationTraining_Click(){
//If //Me.Last_Name = " " //Then;
            //MsgBox "Last Name cannot be blank";
            //Exit //Sub;
        //End //If;
        //strContact = //Me.Common_Contact_ID;
        location.href = "Form_EducationTrainingForm.php";//, //, //, //, //, //acDialog;
        //'Forms![EducationTrainingForm].txtName.Value = Me.[display name]
        //'Forms![EducationTrainingForm].txtIRBName.Value = Me.[IRB Code]
        //'Forms![EducationTrainingForm].txtContactType.Value = Me.Contact_Type
        //'Forms![EducationTrainingForm].txtResearchEducation.Value = Me.Education
}
function DeleteButton_Click(){
//On //Error //GoTo //Err_DeleteButton_Click;
;
//DoCmd.SetWarnings //False;
    //DoCmd.RunCommand //acCmdDeleteRecord;
    ;
    //DoCmd.SetWarnings //True;
    //DoCmd.Close;
    //MsgBox "Coordinator record Deleted";
;
//Exit_DeleteButton_Click:;
    //Exit //Sub;
;
//Err_DeleteButton_Click:;
//If //Err.Number = //3200 //Then;
    //MsgBox "You cannot delete a Coordinator who is assigned to a Study(s).  Use Search 'All Studies' then Filter the Coordinator " //_;
     //& " to determine the Study(s).";
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
//If //Me.NewRecord //And //Me.Last_Name = " " //Then;
    //Me.Undo;
    //Exit //Sub;
//End //If;
//If //CancelUpdate //Then;
    //Cancel = //True;
    //End //If;
;
;
}
function Form_Close(){
//CancelUpdate = //False;
//SearchContacts = //False;
}
function Form_Current(){
//Exit //Sub;
//On //Error //GoTo //Err_Form_Current;
var strSQL = '';
//could initialize rst here
//could initialize db here
//If //HoldCoordName = "" //Then //Exit //Sub;
//If //Me.NewRecord //Then;
    //Me.Last_Name = //HoldCoordName;
    //Set db;
    strSQL = "SELECT [Contact Data].*, [Contact Data].[Common Contact ID] FROM [Contact Data]WHERE ((([Contact Data].[Common Contact ID])= " && [Forms]![_contact coordinator]![Common Contact ID] && "));";
    //Set rst;
    //With rst;
          //Me.Street_Address_1 = //![Street //Address //1];
          //Me.Street_Address_2 = //![Street //Address //2];
          //Me.City = //![City];
          //Me.State = //![State];
          //Me.Zip = //![Zip];
          //Me.Country = //![Country];
          //Me.Main_Phone__ = //![Main //Phone //#];
          //Me.Main___Ext = //![Main //# //Ext];
          //Me.Cell_Phone = //![Cell //Phone];
          //Me.Pager_Number = //![Pager //Number];
          //Me.FaxNumber = //![FaxNumber];
          //Me.EmailAddress = //![EmailAddress];
          //Me.[Voice //Mail] = //![Voice //Mail];
          //Me.Active = //![Active];
          //Me.AlternatePH = //![AlternatePH];
          //Me.AlternateFax = //![AlternateFax];
          //Me.ForeignPH = //![ForeignPH];
          //.Close;
    //End //With;
    //db.Close;
    //MsgBox "The data appearing in the Coordinator's Form has been copied from the PI data.  You may change if appropiate.  The PI's data will not change.  Only the Coordinator's";
//End //If;
    ;
;
//Exit_Form_Current:;
//Exit //Sub;
//Err_Form_Current:;
//MsgBox //Err.Description;
//Resume //Exit_Form_Current;
;
;
;
;
;
;
;
;
;
;
;
;
;
;
;
;
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
    //Call //adhToggleControl(Me, //True);
    //Me.addpi.Enabled = //False;
    //Me.NavigationButtons = //True;
     //Me.AllowFilters = //True;
    //Me.AllowAdditions = //False;
    //Me.AllowEdits = //False;
    //Me.DeleteButton.Enabled = //False;
    //Me.btnEducationTraining.Enabled = //False;
    //Me.LabelReadOnly.Visible = //True;
//Else;
//Me.PreviewReport.Enabled = //False;
//End //If;
}
function Ignore_Click(){
//CancelUpdate = //True;
//DoCmd.Close;
;
}
function PreviewReport_Click(){
//On //Error //GoTo //Err_PreviewReport_Click;
;
    var stDocName = '';
;
    stDocName;
    //LetterName = "_Coordinators";
    //If //Me.FilterOn //Then;
        //MyFilter = //Me.Filter;
    //Else;
        //MyFilter = "";
    //End //If;
    ;
    ;
    ;
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
  <body onLoad="Form_Load();Form_Open();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:210; height:96'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:102; height:120'>    </select>
    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <label id=Label1 value=' Coordinator Data' style='position:absolute; left:492; top:7; width:252; height:40; visibility:'> Coordinator Data</label>
    <input type='text' id=Common Contact ID value='empty' style='position:absolute; left:12; top:60; width:114; height:24'></input>

    <label id=Label2 value='System ID' style='position:absolute; left:18; top:30; width:100; height:24; visibility:'>System ID</label>
    <select name='Contact Type' Size='1' style='position:absolute; left:66; top:94; width:96; height:24'>
        <Option value='Coordinator'>Coordinator</option>
        <Option value='Sponsor'>Sponsor</option>
        <Option value='P.I.'>P.I.</option>
    </select>

    <label id=Label4 value='Type' style='position:absolute; left:6; top:94; width:51; height:24; visibility:'>Type</label>
    <input type='text' id=Last Name value='empty' style='position:absolute; left:184; top:124; width:0; height:30'></input>

    <label id=Label6 value='Last Name:' style='position:absolute; left:238; top:94; width:105; height:24; visibility:'>Last Name:</label>
    <input type='text' id=First Name value='empty' style='position:absolute; left:399; top:124; width:0; height:30'></input>

    <label id=Label7 value='First Name:' style='position:absolute; left:453; top:94; width:105; height:24; visibility:'>First Name:</label>
    <input type='text' id=Middle value='empty' style='position:absolute; left:670; top:124; width:60; height:30'></input>

    <label id=Label8 value='Middle:' style='position:absolute; left:664; top:94; width:70; height:24; visibility:'>Middle:</label>
    <input type='text' id=Suffix value='empty' style='position:absolute; left:757; top:124; width:78; height:30'></input>

    <label id=Label9 value='Suffix:' style='position:absolute; left:765; top:96; width:63; height:24; visibility:'>Suffix:</label>
    <input type='text' id=Street Address 1 value='empty' style='position:absolute; left:184; top:163; width:336; height:30'></input>

    <label id=Label10 value='Street Address 1:' style='position:absolute; left:16; top:163; width:156; height:24; visibility:'>Street Address 1:</label>
    <input type='text' id=Street Address 2 value='empty' style='position:absolute; left:184; top:199; width:336; height:30'></input>

    <label id=Label11 value='Street Address 2:' style='position:absolute; left:16; top:199; width:156; height:24; visibility:'>Street Address 2:</label>
    <input type='text' id=City value='empty' style='position:absolute; left:184; top:237; width:315; height:30'></input>

    <label id=Label12 value='City:' style='position:absolute; left:126; top:237; width:46; height:24; visibility:'>City:</label>
    <input type='text' id=State value='empty' style='position:absolute; left:576; top:237; width:126; height:30'></input>

    <label id=Label13 value='State:' style='position:absolute; left:510; top:237; width:60; height:24; visibility:'>State:</label>
    <input type='text' id=Zip value='empty' style='position:absolute; left:774; top:237; width:144; height:30'></input>

    <label id=Label14 value='Zip:' style='position:absolute; left:720; top:237; width:42; height:24; visibility:'>Zip:</label>
    <label id=Label22 value='Coordinator: ' style='position:absolute; left:54; top:124; width:118; height:24; visibility:'>Coordinator: </label>
    <input type='checkbox' id=Active value='=True' style='position:absolute; left:730; top:364; width:0; height:0'></input>

    <label id=Label27 value='Active' style='position:absolute; left:753; top:361; width:63; height:24; visibility:'>Active</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:139; top:7; width:130; height:42; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:261; top:13; width:18; height:16; visibility:'>R</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:220; top:46; width:193; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='text' id=Company Name value='empty' style='position:absolute; left:184; top:276; width:420; height:30'></input>

    <label id=Label5 value='Company Name:' style='position:absolute; left:27; top:276; width:145; height:24; visibility:'>Company Name:</label>
    <input type='text' id=Main Phone # value='empty' style='position:absolute; left:184; top:318; width:123; height:30'></input>

    <label id=Label16 value='Main Phone #:' style='position:absolute; left:39; top:318; width:133; height:24; visibility:'>Main Phone #:</label>
    <input type='text' id=Main # Ext value='empty' style='position:absolute; left:372; top:318; width:84; height:30'></input>

    <label id=Label17 value='Ext:' style='position:absolute; left:316; top:318; width:42; height:24; visibility:'>Ext:</label>
    <input type='text' id=Cell Phone value='empty' style='position:absolute; left:184; top:360; width:123; height:30'></input>

    <label id=Label18 value='Cell Phone:' style='position:absolute; left:66; top:358; width:106; height:24; visibility:'>Cell Phone:</label>
    <input type='text' id=Pager Number value='empty' style='position:absolute; left:727; top:318; width:190; height:30'></input>

    <label id=Label19 value='Pager:' style='position:absolute; left:655; top:318; width:64; height:24; visibility:'>Pager:</label>
    <input type='text' id=FaxNumber value='empty' style='position:absolute; left:519; top:318; width:123; height:30'></input>

    <label id=Label20 value='Fax:' style='position:absolute; left:471; top:318; width:45; height:24; visibility:'>Fax:</label>
    <input type='text' id=EmailAddress value='empty' style='position:absolute; left:678; top:276; width:240; height:30'></input>

    <label id=Label21 value='Email:' style='position:absolute; left:606; top:276; width:60; height:24; visibility:'>Email:</label>
    <select name='Specialty1' Size='24' style='position:absolute; left:184; top:438; width:204; height:30'>
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

    <label id=Label28 value='Specialty 1:' style='position:absolute; left:63; top:438; width:109; height:24; visibility:'>Specialty 1:</label>
    <select name='Specialty2' Size='25' style='position:absolute; left:543; top:438; width:204; height:30'>
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

    <label id=Label29 value='Specialty 2:' style='position:absolute; left:429; top:438; width:109; height:24; visibility:'>Specialty 2:</label>
    <input type='text' id=AlternatePH value='empty' style='position:absolute; left:184; top:402; width:123; height:30'></input>

    <label id=Label49 value='Alternate Ph:' style='position:absolute; left:52; top:399; width:120; height:24; visibility:'>Alternate Ph:</label>
    <input type='text' id=AlternateFax value='empty' style='position:absolute; left:460; top:400; width:123; height:30'></input>

    <label id=Label50 value='Alternate Fax:' style='position:absolute; left:316; top:400; width:127; height:24; visibility:'>Alternate Fax:</label>
    <input type='text' id=ForeignPH value='empty' style='position:absolute; left:442; top:360; width:190; height:30'></input>

    <label id=Label51 value='Foreign Ph:' style='position:absolute; left:316; top:360; width:106; height:24; visibility:'>Foreign Ph:</label>
    <select name='Label30' Size='7' style='position:absolute; left:604; top:168; width:52; height:24'>
        <Option value='Coordinator'>Coordinator</option>
        <Option value='Sponsor'>Sponsor</option>
        <Option value='P.I.'>P.I.</option>
    </select>
    <input type='text' id=Education value='empty' style='position:absolute; left:214; top:483; width:564; height:78'></input>
    <label id=Label53 value='Research Education:' style='position:absolute; left:22; top:481; width:187; height:24; visibility:'>Research Education:</label>
    <label id=Label57 value='Voice Mail:' style='position:absolute; left:634; top:400; width:103; height:24; visibility:'>Voice Mail:</label>
    <input type='text' id=Text58 value='empty' style='position:absolute; left:750; top:400; width:168; height:30'></input>
    <!--     <label id=LabelReadOnly value='Data is Read Only' style='position:absolute; left:498; top:54; width:234; height:24; visibility:hidden'>Data is Read Only</label> -->
  </body>
</html>
