<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Active_BeforeUpdate(Cancel){
//DoCmd.SetWarnings //False;
//If //Not //Active //Then //DoCmd.OpenQuery "DeleteInactiveBoardMembers";
//DoCmd.SetWarnings //True;
}
function addPI_Click(){
//On //Error //GoTo //err_addpi_click;
;
    //CancelUpdate = //False;
    //DoCmd.DoMenuItem //acFormBar, //acRecordsMenu, //acSaveRecord, //, //acMenuVer70;
    //MsgBox "Member Data saved";
//DataSaved = //True;
    //HoldCoord = //Me![Common //Contact //ID];
    //DoCmd.Close;
;
//Exit_addpi_Click:;
    //Exit //Sub;
;
//err_addpi_click:;
    //MsgBox //Err.Number //& "   " //& //Err.Description, //vbInformation;
    //Resume //Exit_addpi_Click;
    ;
}
function Command38_Click(){
//End //Sub;
;
//' 11/21 - ginny
//' Open Education form with the information from this form
//Private //Sub //btnEducationTraining_Click();
        //If //Me.Last_Name = " " //Then;
            //MsgBox "Last Name cannot be blank";
            //Exit //Sub;
        //End //If;
        //strContact = //Me.Common_Contact_ID;
        location.href = "Form_EducationTrainingForm.php";//, //, //, //, //, //acDialog;
        //'Forms![EducationTrainingForm].txtName.Value = Me.[display name]
       //' Forms![EducationTrainingForm].txtIRBName.Value = Me.IRB_Code
        //'Forms![EducationTrainingForm].txtContactType.Value = Me.Contact_Type
        //'Forms![EducationTrainingForm].txtResearchEducation.Value = Me.Education
}
function Form_BeforeUpdate(Cancel){
//could initialize Response here
//If //Me.NewRecord //And //Me.Last_Name = " " //Then;
    //Me.Undo;
    //Exit //Sub;
//End //If;
//If //CancelUpdate //Then;
      //CancelUpdate = //False;
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
//CancelUpdate = //False;
//SearchContacts = //False;
}
function Form_Current(){
//If //HoldCoordName = "" //Then //Exit //Sub;
//If //Me.NewRecord //Then;
    //Me.Last_Name = //HoldCoordName;
//End //If;
;
}
function Form_Filter(Cancel, FilterType){
//Me.FilterOn = //True;
//Me.Filter = "";
}
function Form_Open(Cancel){
//If //SearchContacts //Then;
//Call //adhToggleControl(Me, //True);
    //Me.NavigationButtons = //True;
     //Me.AllowFilters = //True;
    //Me.AllowAdditions = //False;
    //Me.AllowEdits = //False;
    //Me.addpi.Enabled = //False;
    //Me.DeleteButton.Enabled = //False;
    //Me.LabelReadOnly.Visible = //False;
    //Me.btnEducationTraining.Enabled = //False;
//Else;
//Me.PreviewReport.Enabled = //False;
//End //If;
;
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
    //LetterName = "_Board Members";
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
function DeleteButton_Click(){
//On //Error //GoTo //Err_DeleteButton_Click;
;
//DoCmd.SetWarnings //False;
    //DoCmd.RunCommand //acCmdDeleteRecord;
    ;
    //DoCmd.SetWarnings //True;
    //DoCmd.Close;
    //MsgBox "Member record Deleted";
;
//Exit_DeleteButton_Click:;
    //Exit //Sub;
;
//Err_DeleteButton_Click:;
    //MsgBox //Err.Description //& " " //& //Err.Number;
    //Resume //Exit_DeleteButton_Click;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:210; height:96'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:102; height:120'>    </select>
    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <label id=Label1 value='Board Member     Information' style='position:absolute; left:1; top:3; width:264; height:78; visibility:'>Board Member     Information</label>
    <input type='text' id=Common Contact ID value='empty' style='position:absolute; left:174; top:513; width:114; height:29'></input>

    <label id=Label2 value='System ID' style='position:absolute; left:67; top:513; width:100; height:24; visibility:'>System ID</label>
    <input type='text' id=Last Name value='empty' style='position:absolute; left:174; top:123; width:0; height:30'></input>

    <label id=Label6 value='Last Name:' style='position:absolute; left:238; top:93; width:105; height:24; visibility:'>Last Name:</label>
    <input type='text' id=First Name value='empty' style='position:absolute; left:400; top:123; width:0; height:30'></input>

    <label id=Label7 value='First Name:' style='position:absolute; left:454; top:93; width:105; height:24; visibility:'>First Name:</label>
    <input type='text' id=Middle value='empty' style='position:absolute; left:628; top:123; width:60; height:30'></input>

    <label id=Label8 value='Middle:' style='position:absolute; left:622; top:93; width:70; height:24; visibility:'>Middle:</label>
    <input type='text' id=Suffix value='empty' style='position:absolute; left:706; top:123; width:141; height:30'></input>

    <label id=Label9 value='Suffix:' style='position:absolute; left:718; top:93; width:63; height:24; visibility:'>Suffix:</label>
    <input type='text' id=Street Address 1 value='empty' style='position:absolute; left:174; top:202; width:336; height:30'></input>

    <label id=Label10 value='Street Address 1:' style='position:absolute; left:12; top:202; width:156; height:24; visibility:'>Street Address 1:</label>
    <input type='text' id=Street Address 2 value='empty' style='position:absolute; left:174; top:243; width:336; height:30'></input>

    <label id=Label11 value='Street Address 2:' style='position:absolute; left:12; top:243; width:156; height:24; visibility:'>Street Address 2:</label>
    <input type='text' id=City value='empty' style='position:absolute; left:174; top:279; width:0; height:30'></input>

    <label id=Label12 value='City:' style='position:absolute; left:121; top:279; width:46; height:24; visibility:'>City:</label>
    <input type='text' id=State value='empty' style='position:absolute; left:499; top:279; width:126; height:30'></input>

    <label id=Label13 value='State:' style='position:absolute; left:432; top:279; width:60; height:24; visibility:'>State:</label>
    <input type='text' id=Zip value='empty' style='position:absolute; left:700; top:279; width:114; height:30'></input>

    <label id=Label14 value='Zip:' style='position:absolute; left:652; top:279; width:42; height:24; visibility:'>Zip:</label>
    <label id=Label22 value='Board Member:' style='position:absolute; left:31; top:121; width:136; height:24; visibility:'>Board Member:</label>
    <input type='checkbox' id=Active value='=True' style='position:absolute; left:330; top:516; width:0; height:0'></input>

    <label id=Label27 value='Active' style='position:absolute; left:353; top:513; width:63; height:24; visibility:'>Active</label>
    <input type='text' id=Company Name value='empty' style='position:absolute; left:174; top:162; width:342; height:30'></input>

    <label id=Label5 value='Representing:' style='position:absolute; left:40; top:162; width:127; height:24; visibility:'>Representing:</label>
    <input type='text' id=Main Phone # value='empty' style='position:absolute; left:174; top:349; width:123; height:30'></input>

    <label id=Label16 value='Main Phone #:' style='position:absolute; left:34; top:349; width:133; height:24; visibility:'>Main Phone #:</label>
    <input type='text' id=Main # Ext value='empty' style='position:absolute; left:354; top:349; width:84; height:30'></input>

    <label id=Label17 value='Ext:' style='position:absolute; left:304; top:349; width:42; height:24; visibility:'>Ext:</label>
    <input type='text' id=Cell Phone value='empty' style='position:absolute; left:174; top:390; width:123; height:30'></input>

    <label id=Label18 value='Cell Phone:' style='position:absolute; left:61; top:390; width:106; height:24; visibility:'>Cell Phone:</label>
    <input type='text' id=Pager Number value='empty' style='position:absolute; left:688; top:349; width:192; height:30'></input>

    <label id=Label19 value='Pager:' style='position:absolute; left:616; top:349; width:64; height:24; visibility:'>Pager:</label>
    <input type='text' id=FaxNumber value='empty' style='position:absolute; left:492; top:349; width:123; height:30'></input>

    <label id=Label20 value='Fax:' style='position:absolute; left:442; top:349; width:45; height:24; visibility:'>Fax:</label>
    <input type='text' id=EmailAddress value='empty' style='position:absolute; left:585; top:390; width:291; height:30'></input>

    <label id=Label21 value='Email:' style='position:absolute; left:519; top:390; width:60; height:24; visibility:'>Email:</label>
    <select name='Specialty1' Size='22' style='position:absolute; left:174; top:472; width:250; height:30'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRB Meeting Schedule1].[Irb Code] FROM [IRB Meeting Schedule1] WHERE ((Not ([IRB Meeting Schedule1].[Irb Code])=\"***\"))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Irb Code']."'".">".$row['Irb Code']."</Option>";
        }
    ?>
    </select>

    <label id=Label28 value='Specialty 1:' style='position:absolute; left:58; top:472; width:109; height:24; visibility:'>Specialty 1:</label>
    <select name='Specialty2' Size='23' style='position:absolute; left:585; top:472; width:291; height:30'>
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

    <label id=Label29 value='Specialty 2:' style='position:absolute; left:469; top:472; width:109; height:24; visibility:'>Specialty 2:</label>
    <input type='text' id=AlternatePH value='empty' style='position:absolute; left:174; top:430; width:123; height:30'></input>

    <label id=Label49 value='Alternate Ph:' style='position:absolute; left:48; top:430; width:120; height:24; visibility:'>Alternate Ph:</label>
    <input type='text' id=AlternateFax value='empty' style='position:absolute; left:585; top:430; width:123; height:30'></input>

    <label id=Label50 value='Alternate Fax:' style='position:absolute; left:451; top:430; width:127; height:24; visibility:'>Alternate Fax:</label>
    <input type='text' id=ForeignPH value='empty' style='position:absolute; left:585; top:513; width:291; height:30'></input>

    <label id=Label51 value='Foreign Ph:' style='position:absolute; left:472; top:513; width:106; height:24; visibility:'>Foreign Ph:</label>
    <select name='Label30' Size='6' style='position:absolute; left:615; top:166; width:73; height:24'>
        <Option value='PERM'>PERM</option>
        <Option value='ALT'>ALT</option>
    </select>
    <select name='IRB Code' style='position:absolute; left:280; top:24; width:718; height:30'>
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
    <select name='ALT' Size='8' style='position:absolute; left:711; top:214; width:142; height:42'>
        <Option value='PERM'>PERM</option>
        <Option value='ALT'>ALT</option>
    </select>

    <label id=Label33 value='Permanent or\015\012Alternate Member' style='position:absolute; left:528; top:214; width:166; height:40; visibility:'>Permanent or\015\012Alternate Member</label>
    <input type='text' id=Education value='empty' style='position:absolute; left:202; top:550; width:652; height:73'></input>
    <label id=Label53 value='Research Education:' style='position:absolute; left:10; top:550; width:187; height:24; visibility:'>Research Education:</label>
    <input type='text' id=Country value='empty' style='position:absolute; left:174; top:313; width:0; height:30'></input>

    <label id=Label36 value='Country:' style='position:absolute; left:88; top:313; width:79; height:24; visibility:'>Country:</label>
    <!--     <select name='Contact Type' Size='28' style='position:absolute; left:379; top:70; width:0; height:0'>
        <Option value='Coordinator'>Coordinator</option>
        <Option value='Sponsor'>Sponsor</option>
        <Option value='P.I.'>P.I.</option>
        <Option value='Board Member'>Board Member</option>
    </select> -->

    <label id=Label37 value='Contact Type:' style='position:absolute; left:235; top:70; width:129; height:24; visibility:'>Contact Type:</label>
    <!--     <label id=LabelReadOnly value='Data is Read Only' style='position:absolute; left:1; top:93; width:223; height:24; visibility:hidden'>Data is Read Only</label> -->
  </body>
</html>
