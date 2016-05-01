<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function DeleteButton_Click(){
//On //Error //GoTo //Err_DeleteButton_Click;
;
//If //YesNo("This //action //will //delete //this //CPA //and //the //Associated //Agenda //item //(if //any).  //Do //you //want //to //proceed?") //Then;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "DeleteCPAfromAgenda";
        //DoCmd.OpenQuery "DeleteCPAfromCPA1";
        //DoCmd.Requery;
        //DoCmd.SetWarnings //True;
    //End //If;
        ;
    ;
//Exit_DeleteButton_Click:;
    //Exit //Sub;
;
//Err_DeleteButton_Click:;
    //MsgBox //Err.Description //& //Err.Number;
    //Resume //Exit_DeleteButton_Click;
    ;
}
function Form_Close(){
//Me.Filter = "";
}
function Form_Current(){
//If //Me.NewRecord //And //Me.FilterOn = //True //Then;
    //MsgBox "There are no CPAs meeting the filter criteria";
    //DoCmd.RunCommand //acCmdRecordsGoToFirst;
    //Me.Filter = "";
    //Me.FilterOn = //False;
    //Exit //Sub;
//Else;
    //If //Me.NewRecord //And //Me.RecordsetClone.RecordCount //> //0 //Then;
        //MsgBox "You cannot ADD a new CPA from this Form";
        //DoCmd.RunCommand //acCmdRecordsGoToFirst;
        //Exit //Sub;
    //End //If;
//End //If;
//Me.tSearchLocation.Caption = //[Forms]![newmenu].[Location];
//Me.QueryName = //Forms![newmenu]![SelectionsCombo];
;
//If //Not //IsNull(Me.Place_on_agenda) //And //Me.Place_on_agenda //Then;
       //Me.on_agenda_label.ForeColor = //vbRed;
    //Else;
    //Me.on_agenda_label.ForeColor = //vbBlack;
//End //If;
;
   ;
;
}
function Form_Filter(Cancel, FilterType){
//Me.Filter = "";
}
function Form_Open(Cancel){
//Call //adhToggleControl(Me, //True);
//Me.LabelReadOnly.Visible = //False;
//Me.Filter = "";
//'On Error GoTo err_open
//Me.tSearchLocation.Caption = //Forms![newmenu]![Location];
//Me.QueryName = //Forms![newmenu]![SelectionsCombo];
;
//If //UserLastName //Like "*ADMIN*" //Then;
    //Me.DeleteButton.Visible = //True;
  ;
//Else;
    //Me.DeleteButton.Visible = //False;
//End //If;
;
}
function Form_Timer(){
//On //Error //GoTo //Err_timer;
//Me.TimerInterval = //0;
//If //Me.RecordsetClone.RecordCount = //0 //Then;
    //MsgBox "There are no CPAs for this IRB.";
    //DoCmd.Close;
//End //If;
//exit_timer:;
//Exit //Sub;
//Err_timer:;
//MsgBox //Err.Description //& " " //& //Err.Number;
//GoTo //exit_timer;
}
function preview_print_Click(){
//If //YesNo("By //clicking //Yes //you //will //preview //(and //be " _
& "//able //to //print) //ALL //CPAs //for //ALL //Studies //in //your" //_;
//& " database OR those CPA records returned based upon " //_;
//& "your filter criteria."//) //Then;
;
//Else;
    //Exit //Sub;
//End //If;
    ;
   //If //Me.Filter = "" //Or //IsNull(Me.Filter) //Then;
        //MyFilter = "";
   //Else;
        //MyFilter = //Me.Filter;
   //End //If;
    //DoCmd.OpenReport "CPAPrintFullDetail2"//, //acViewPreview, //, //MyFilter;
}
function Returnmenu_Click(){
//On //Error //Resume //Next;
//CancelUpdate = //True;
//closeswitch = //True;
//DoCmd.Close;
;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:768; height:64'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <input type='text' id=tCPAnumber value='empty' style='position:absolute; left:18; top:201; width:168; height:30'></input>

    <label id=SAE Number Label value='CPA Number' style='position:absolute; left:12; top:165; width:178; height:24; visibility:'>CPA Number</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:624; top:144; width:504; height:0'></input>

    <label id=Protocol Number & Title Label value='Protocol#/Title' style='position:absolute; left:612; top:120; width:138; height:24; visibility:'>Protocol#/Title</label>
    <input type='text' id=Date of Change value='empty' style='position:absolute; left:246; top:201; width:154; height:30'></input>

    <label id=Date Change Label value='Date of Change' style='position:absolute; left:228; top:165; width:183; height:24; visibility:'>Date of Change</label>
    <input type='checkbox' id=Signed by PI ? value='False' style='position:absolute; left:456; top:498; width:24; height:30'></input>

    <label id=Signed by PI ? Label value='Signed ' style='position:absolute; left:426; top:456; width:73; height:24; visibility:'>Signed </label>
    <input type='text' id=Date Signed value='empty' style='position:absolute; left:306; top:492; width:96; height:30'></input>

    <label id=Date Signed Label value='Date Signed' style='position:absolute; left:288; top:456; width:114; height:24; visibility:'>Date Signed</label>
    <input type='text' id=date recieved by value='empty' style='position:absolute; left:426; top:199; width:144; height:31'></input>

    <label id=Label56 value='Date Received ' style='position:absolute; left:426; top:165; width:142; height:24; visibility:'>Date Received </label>
    <input type='button' id=preview_print value='&Preview CPAs' style='position:absolute; left:831; top:540; width:183; height:48' onclick='preview_print_Click();'></input>
    <input type='button' id=Returnmenu value='&Return' style='position:absolute; left:1014; top:540; width:120; height:48' onclick='Returnmenu_Click();'></input>
    <input type='checkbox' id=Place on agenda value='False' style='position:absolute; left:36; top:549; width:0; height:0'></input>

    <label id=on agenda label value='Place on Meeting Agenda' style='position:absolute; left:58; top:546; width:228; height:24; visibility:'>Place on Meeting Agenda</label>
    <input type='text' id=Other 1 value='empty' style='position:absolute; left:630; top:228; width:498; height:59'></input>

    <label id=Label108 value='Other 1' style='position:absolute; left:534; top:252; width:72; height:24; visibility:'>Other 1</label>
    <input type='text' id=other2 value='empty' style='position:absolute; left:633; top:300; width:495; height:65'></input>

    <label id=Label110 value='Other 2' style='position:absolute; left:540; top:318; width:72; height:24; visibility:'>Other 2</label>
    <label id=Label1 value='Change In Procedure' style='position:absolute; left:42; top:78; width:520; height:54; visibility:'>Change In Procedure</label>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:700; top:42; width:294; height:30'></input>

    <label id=Label453 value='Study IRB' style='position:absolute; left:742; top:9; width:136; height:24; visibility:'>Study IRB</label>
    <select name='CpaActions' Size='2' style='position:absolute; left:279; top:276; width:210; height:163'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCT [Cpa1].[IRB#] FROM Cpa1 WITH OWNERACCESS OPTION");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['IRB#']."'".">".$row['IRB#']."</Option>";
        }
    ?>
    </select>

    <label id=CPAAction_Label value='Pre-Meeting Action Taken' style='position:absolute; left:267; top:240; width:231; height:24; visibility:'>Pre-Meeting Action Taken</label>
    <input type='checkbox' id=CheckExpedite value='False' style='position:absolute; left:36; top:507; width:174; height:18'></input>

    <label id=Label132 value='Expedite' style='position:absolute; left:59; top:504; width:82; height:30; visibility:'>Expedite</label>
    <!--     <input type='text' id=condition1 value='empty' style='position:absolute; left:318; top:258; width:138; height:72'></input> -->

    <label id=Label134 value='Text133' style='position:absolute; left:663; top:234; width:78; height:24; visibility:'>Text133</label>
    <!--     <input type='text' id=Condition2 value='empty' style='position:absolute; left:264; top:228; width:114; height:66'></input> -->

    <label id=Label136 value='Text135' style='position:absolute; left:609; top:204; width:78; height:24; visibility:'>Text135</label>
    <!--     <input type='text' id=Category value='empty' style='position:absolute; left:132; top:258; width:186; height:54'></input> -->

    <label id=Label138 value='Text137' style='position:absolute; left:477; top:234; width:78; height:24; visibility:'>Text137</label>
    <select name='CPATypeList' Size='20' style='position:absolute; left:9; top:270; width:240; height:222'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [CPAActionCodes1].[CPAAction], [CPAActionCodes1].[Study Status], [StudyStatusCodes].[StudyActiveCode] FROM StudyStatusCodes INNER JOIN CPAActionCodes1 ON [StudyStatusCodes].[Study Status]=[CPAActionCodes1].[Study Status] WHERE ((([CPAActionCodes1].[SortSeq])<500)) ORDER BY [CPAActionCodes1].[SortSeq]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['CPAAction']."'".">".$row['CPAAction']."</Option>";
        }
    ?>
    </select>
    <label id=Label143 value='Type of Change' style='position:absolute; left:60; top:240; width:144; height:24; visibility:'>Type of Change</label>
    <input type='text' id=MeetingDate value='empty' style='position:absolute; left:390; top:546; width:150; height:25'></input>

    <label id=Label145 value='For:' style='position:absolute; left:300; top:546; width:78; height:24; visibility:'>For:</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:15; top:40; width:153; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <label id=tSearchLocation value='Search Location' style='position:absolute; left:187; top:42; width:288; height:30; visibility:'>Search Location</label>
    <label id=Label151 value='Search Location' style='position:absolute; left:211; top:9; width:318; height:24; visibility:'>Search Location</label>
    <input type='text' id=QueryName value='empty' style='position:absolute; left:484; top:42; width:207; height:30'></input>
    <label id=Label515 value='Query' style='position:absolute; left:571; top:9; width:58; height:24; visibility:'>Query</label>
    <label id=Label161 value='Investigator' style='position:absolute; left:612; top:388; width:109; height:24; visibility:'>Investigator</label>
    <label id=Label167 value='Coordinator' style='position:absolute; left:612; top:433; width:109; height:24; visibility:'>Coordinator</label>
    <select name='Coordinator ID' Size='25' style='position:absolute; left:735; top:433; width:228; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [CPATypes1].[CPAType], [CPATypes1].[Agenda Category], [CPATypes1].[PlaceOnAgenda] FROM CPATypes1 ORDER BY [CPATypes1].[CPAType] DESC");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['CPAType']."'".">".$row['CPAType']."</Option>";
        }
    ?>
    </select>
    <select name='Investigator ID' Size='24' style='position:absolute; left:735; top:391; width:228; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [Contact Data1]![Last Name] & \", \" & [Contact Data1]![First Name] & \" \ " & [Contact Data1]![Middle] & \" \" & [Contact Data1]![Suffix] AS [display name] FROM [Contact Data1] WHERE ((([Contact Data1].[Contact Type])=\"Coordinator\"))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Last Name']."'".">".$row['Last Name']."</Option>";
        }
    ?>
    </select>
    <select name='Study Active?' Size='15' style='position:absolute; left:718; top:84; width:102; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [Contact Data1]![Last Name] & \", \" & [Contact Data1]![First Name] & \" \ " & [Contact Data1]![Middle] & \" \" & [Contact Data1]![Suffix] AS [display name] FROM [Contact Data1] WHERE ((([Contact Data1].[Contact Type])=\"P.I.\"))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Last Name']."'".">".$row['Last Name']."</Option>";
        }
    ?>
    </select>

    <label id=Label69 value='Active?:' style='position:absolute; left:628; top:84; width:79; height:24; visibility:'>Active?:</label>
    <select name='Study Status' Size='14' style='position:absolute; left:969; top:84; width:162; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCT [StudyStatusCodes1].[StudyActiveCode] FROM StudyStatusCodes1");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['StudyActiveCode']."'".">".$row['StudyActiveCode']."</Option>";
        }
    ?>
    </select>

    <label id=Label70 value='Study Status:' style='position:absolute; left:837; top:84; width:123; height:24; visibility:'>Study Status:</label>
    <select name='Sponsor ID' Size='22' style='position:absolute; left:735; top:468; width:228; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [StudyStatusCodes1].[Study Status] FROM StudyStatusCodes1");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Study Status']."'".">".$row['Study Status']."</Option>";
        }
    ?>
    </select>

    <label id=Label153 value='Sponsor:' style='position:absolute; left:637; top:468; width:84; height:24; visibility:'>Sponsor:</label>
    <input type='button' id=DeleteButton value='Delete CPA' style='position:absolute; left:648; top:540; width:183; height:48' onclick='DeleteButton_Click();'></input>
    <select name='studyNo' Size='11' style='position:absolute; left:1000; top:42; width:126; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT [_select sponsors for irb#].[Company Name], [_select sponsors for irb#].[Contact Type] FROM [_select sponsors for irb#] WHERE ((([_select sponsors for irb#].[Contact Type])=\"Sponsor\")) ORDER BY [_select sponsors for irb#].[Company Name]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Company Name']."'".">".$row['Company Name']."</Option>";
        }
    ?>
    </select>

    <label id=Label112 value='Study #' style='position:absolute; left:1033; top:9; width:75; height:24; visibility:'>Study #</label>
    <label id=Label174 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:43; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <!--     <label id=LabelReadOnly value='Data is Read Only' style='position:absolute; left:199; top:136; width:234; height:24; visibility:hidden'>Data is Read Only</label> -->
  </body>
</html>
