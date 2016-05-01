<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Form_Activate(){
//If //Me.FilterOn = //True //Then;
//Else;
//Me.Filter = "";
//End //If;
}
function Form_ApplyFilter(Cancel, ApplyType){
//'MsgBox Me.Filter
}
function Form_Current(){
//If //Me.NewRecord //And //Me.FilterOn = //True //Then;
    //MsgBox "There are no SAEs meeting the filter criteria";
    //DoCmd.RunCommand //acCmdRecordsGoToFirst;
    //Me.Filter = "";
    //Me.FilterOn = //False;
    //Exit //Sub;
//Else;
    //If //Me.NewRecord //And //Me.RecordsetClone.RecordCount //> //0 //Then;
        //MsgBox "You cannot add an SAE from this form";
        //DoCmd.RunCommand //acCmdRecordsGoToFirst;
        //Exit //Sub;
    //End //If;
//End //If;
//Me.tSearchLocation.Caption = //[Forms]![newmenu].[Location];
//Me.QueryName = //Forms![newmenu]![SelectionsCombo];
;
;
;
//If //Me.OnSite = //True //Then;
    //Me.LabelOnSite.ForeColor = //vbRed;
    //Me.Label1.ForeColor = //vbRed;
//Else;
    //Me.LabelOnSite.ForeColor = //labelcolor;
    //Me.Label1.ForeColor = //vbBlue;
//End //If;
;
//If //Me.EventType = "Death" //Then;
    //Me.Label1.ForeColor = //vbRed;
    //Me.LabelEventType.ForeColor = //vbRed;
//Else;
    //Me.LabelEventType.ForeColor = //labelcolor;
//End //If;
//exit_curr:;
//Exit //Sub;
;
//err_current:;
//MsgBox "Error occured in On Current event of SAE Form"//, //vbInformation;
//GoTo //exit_curr;
;
}
function Form_Filter(Cancel, FilterType){
//Me.Filter = "";
;
}
function NextButton_Click(){
//On //Error //GoTo //EndofFile;
//CancelUpdate = //True;
//DoCmd.RunCommand //acCmdRecordsGoToNext;
//exit_next:;
//Exit //Sub;
//EndofFile:;
//MsgBox "End of SAEs"//, //vbInformation;
//GoTo //exit_next;
}
function Form_Open(Cancel){
//Call //adhToggleControl(Me, //True);
;
//Me.LabelReadOnly.Visible = //False;
;
//Me.Filter = "";
//Me.tSearchLocation.Caption = //[Forms]![newmenu].[Location];
//Me.QueryName = //Forms![newmenu]![SelectionsCombo];
//If //UserLastName //Like "*ADMIN*" //Then //Me.DeleteButton.Visible = //True //Else //Me.DeleteButton.Visible = //False;
;
;
}
function Form_Timer(){
//On //Error //GoTo //Err_timer;
//Me.TimerInterval = //0;
//If //Me.RecordsetClone.RecordCount = //0 //Then;
    //MsgBox "There are no SAEs for this IRB.";
    //DoCmd.Close;
//End //If;
//exit_timer:;
//Exit //Sub;
//Err_timer:;
//MsgBox //Err.Description //& " " //& //Err.Number;
//GoTo //exit_timer;
}
function OnSite_AfterUpdate(){
//If //Me.OnSite = //True //Then;
    //Me.LabelOnSite.ForeColor = //vbRed;
    //Me.Label1.ForeColor = //vbRed;
   ;
//Else;
    //Me.LabelOnSite.ForeColor = //labelcolor;
    //Me.Label1.ForeColor = //vbCyan;
    ;
//End //If;
;
//'If Me.EventType.Column(1) = True Then
//'    Me.Label1.ForeColor = vbRed
   ;
  //'  Me.LabelEventType.ForeColor = vbRed
//'Else
   //' Me.LabelEventType.ForeColor = labelcolor
    ;
//'End If
;
}
function preview_print_Click(){
//On //Error //GoTo //Err_preview_print_Click;
//If //YesNo("By //clicking //Yes //you //will //preview //(and //be " _
& "//able //to //print) //ALL //SAEs //for //ALL //Studies //in //your" //_;
//& " database OR those SAE records returned based upon " //_;
//& "your filter criteria."//) //Then;
;
//Else;
    //Exit //Sub;
//End //If;
;
var stDocName = '';
//If //Me.Filter = "" //Or //IsNull(Me.Filter) //Then;
        //MyFilter = "";
   //Else;
        //MyFilter = //Me.Filter;
        //'MsgBox Me.Filter
        ;
//End //If;
  ;
    stDocName;
;
   //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
    //'DoCmd.RunCommand acCmdZoom100
    ;
    //GoTo //Exit_preview_print_Click;
   ;
//Exit_preview_print_Click:;
    //Exit //Sub;
//Err_preview_print_Click:;
    //MsgBox //Err.Description, //vbInformation;
    //Resume //Exit_preview_print_Click;
        ;
 ;
;
}
function PreviousButton_Click(){
//On //Error //GoTo //BegofFile;
//CancelUpdate = //True;
//DoCmd.RunCommand //acCmdRecordsGoToPrevious;
;
//exit_prev:;
//Exit //Sub;
;
//BegofFile:;
//MsgBox "Beginning of SAEs"//, //vbInformation;
//GoTo //exit_prev;
;
;
}
function Returnmenu_Click(){
//'MsgBox Me.OrderBy
    //'Me.OrderBy = ""
  //'  MsgBox Me.OrderBy
//DoCmd.Close;
}
function DeleteButton_Click(){
//On //Error //GoTo //Err_DeleteButton_Click;
    ;
    //If //YesNo("This //action //will //delete //this //SAE //and //the //Associated //Agenda //item //(if //any).  //Do //you //want //to //proceed?") //Then;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery "DeleteSAEAgendaItem";
        //DoCmd.OpenQuery "DeleteSAEfromSAE1";
        //MsgBox //Me.Filter;
        //Me.Filter = "";
        //DoCmd.Requery;
        //'DoCmd.RunCommand acCmdDeleteRecord
        //DoCmd.SetWarnings //True;
    //End //If;
//Exit_DeleteButton_Click:;
   ;
    //Exit //Sub;
;
//Err_DeleteButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_DeleteButton_Click;
    ;
}
function butPIList_Click(){
//On //Error //GoTo //Err_butPIList_Click;
;
   //Me.txtHoldFilter = //Me.Filter;
   ;
   //'MsgBox Me.OrderBy
   ;
//DoCmd.OpenReport "SAEs List by PI"//, //acPreview, //, //Me.txtHoldFilter;
;
;
//Exit_butPIList_Click:;
    //Exit //Sub;
;
//Err_butPIList_Click:;
    ;
    //MsgBox //Err.Description;
    //Resume //Exit_butPIList_Click;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Active();Form_Current();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:768; height:64'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>



    <input type='text' id=IRB# value='empty' style='position:absolute; left:979; top:42; width:135; height:30'></input>

    <label id=IRB# Label value='Study#' style='position:absolute; left:1033; top:6; width:69; height:24; visibility:'>Study#</label>
    <input type='text' id=SAE Number value='empty' style='position:absolute; left:12; top:159; width:120; height:30'></input>

    <label id=SAE Number Label value='SAE Internal #' style='position:absolute; left:6; top:123; width:133; height:24; visibility:'>SAE Internal #</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:526; top:108; width:588; height:0'></input>

    <label id=Protocol Number & Title Label value='Protocol#/Title' style='position:absolute; left:526; top:84; width:588; height:24; visibility:'>Protocol#/Title</label>
    <input type='text' id=Date Event Occured value='empty' style='position:absolute; left:12; top:378; width:126; height:30'></input>

    <label id=Date Event Occured Label value='Date of Event' style='position:absolute; left:12; top:348; width:127; height:24; visibility:'>Date of Event</label>
    <input type='text' id=Description of Event value='empty' style='position:absolute; left:528; top:210; width:589; height:78'></input>

    <label id=Description of Event Label value='Description of Event' style='position:absolute; left:690; top:180; width:192; height:24; visibility:'>Description of Event</label>
    <input type='text' id=Location if Off Site value='empty' style='position:absolute; left:198; top:378; width:192; height:30'></input>

    <label id=Location if Off Site Label value='Location if Off Site' style='position:absolute; left:198; top:348; width:169; height:24; visibility:'>Location if Off Site</label>
    <input type='text' id=Medwatch # value='empty' style='position:absolute; left:642; top:352; width:144; height:30'></input>

    <label id=Medwatch # Label value='Medwatch #' style='position:absolute; left:522; top:354; width:109; height:24; visibility:'>Medwatch #</label>
    <input type='text' id=Patient Age value='empty' style='position:absolute; left:955; top:462; width:54; height:24'></input>
    <input type='checkbox' id=Are the Risks Altered? value='False' style='position:absolute; left:1104; top:378; width:36; height:30'></input>

    <label id=Are the Risks Altered? Label value='Are the Risks Altered?' style='position:absolute; left:889; top:372; width:199; height:24; visibility:'>Are the Risks Altered?</label>
    <input type='checkbox' id=New Consent Required? value='False' style='position:absolute; left:1104; top:408; width:42; height:30'></input>

    <label id=New Consent Required? Label value='New Consent Required?' style='position:absolute; left:876; top:402; width:214; height:24; visibility:'>New Consent Required?</label>
    <input type='checkbox' id=Signed by PI ? value='False' style='position:absolute; left:342; top:492; width:24; height:30'></input>

    <label id=Signed by PI ? Label value='Signed by PI ?' style='position:absolute; left:192; top:486; width:133; height:24; visibility:'>Signed by PI ?</label>
    <input type='text' id=Date Signed value='empty' style='position:absolute; left:312; top:522; width:132; height:30'></input>

    <label id=Date Signed Label value='Date Signed' style='position:absolute; left:192; top:522; width:114; height:24; visibility:'>Date Signed</label>
    <label id=Label46 value='Age' style='position:absolute; left:957; top:438; width:54; height:24; visibility:'>Age</label>
    <label id=Label47 value='Sex' style='position:absolute; left:1043; top:438; width:54; height:24; visibility:'>Sex</label>
    <input type='text' id=date recieved by value='empty' style='position:absolute; left:12; top:450; width:144; height:24'></input>

    <label id=Label56 value='Date Received' style='position:absolute; left:12; top:414; width:136; height:24; visibility:'>Date Received</label>
    <!--     <input type='checkbox' id=Event related value='False' style='position:absolute; left:1104; top:351; width:0; height:0'></input> -->

    <!--     <label id=Label85 value='Event Related to Study?' style='position:absolute; left:870; top:342; width:217; height:24; visibility:hidden'>Event Related to Study?</label> -->
    <select name='Patient Status' Size='3' style='position:absolute; left:198; top:450; width:231; height:0'>
        <Option value='CLosed'>CLosed</option>
        <Option value='Open'>Open</option>
    </select>

    <label id=Patient Status Label value='Patient Status' style='position:absolute; left:198; top:414; width:231; height:24; visibility:'>Patient Status</label>
    <input type='text' id=Patient Identifier value='empty' style='position:absolute; left:642; top:426; width:231; height:24'></input>

    <label id=Patient Identifier Label value='Reference' style='position:absolute; left:534; top:426; width:97; height:24; visibility:'>Reference</label>
    <label id=Label1 value='SAE's' style='position:absolute; left:28; top:76; width:93; height:40; visibility:'>SAE's</label>
    <input type='checkbox' id=MedWatch Report Filed value='=False' style='position:absolute; left:763; top:396; width:0; height:0'></input>

    <label id=Label98 value='MedWatch Report Filed' style='position:absolute; left:552; top:390; width:210; height:24; visibility:'>MedWatch Report Filed</label>
    <!--     <input type='text' id=condition1 value='empty' style='position:absolute; left:712; top:84; width:138; height:72'></input> -->

    <label id=Label134 value='Text133' style='position:absolute; left:778; top:108; width:78; height:24; visibility:'>Text133</label>
    <!--     <input type='text' id=Condition2 value='empty' style='position:absolute; left:628; top:87; width:114; height:66'></input> -->

    <label id=Label136 value='Text135' style='position:absolute; left:646; top:147; width:78; height:24; visibility:'>Text135</label>
    <!--     <input type='text' id=Category value='empty' style='position:absolute; left:330; top:3; width:186; height:54'></input> -->

    <label id=Label138 value='Text137' style='position:absolute; left:450; top:27; width:78; height:24; visibility:'>Text137</label>
    <input type='text' id=MeetingDate value='empty' style='position:absolute; left:18; top:634; width:120; height:30'></input>

    <label id=Label101 value='OnMeetingDate' style='position:absolute; left:12; top:604; width:141; height:24; visibility:'>OnMeetingDate</label>
    <input type='checkbox' id=FollowUpReportFlag value='False' style='position:absolute; left:411; top:135; width:0; height:0'></input>

    <label id=Label104 value='Is this a Followup Report?' style='position:absolute; left:162; top:127; width:231; height:24; visibility:'>Is this a Followup Report?</label>
    <input type='text' id=Text106 value='empty' style='position:absolute; left:12; top:219; width:147; height:30'></input>

    <label id=Label107 value='IND Report #' style='position:absolute; left:12; top:192; width:121; height:24; visibility:'>IND Report #</label>
    <label id=tSearchLocation value='Search Location' style='position:absolute; left:178; top:42; width:303; height:30; visibility:'>Search Location</label>
    <label id=Label110 value='Search Location' style='position:absolute; left:234; top:7; width:262; height:24; visibility:'>Search Location</label>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:684; top:42; width:292; height:30'></input>

    <label id=Label114 value='This Study's IRB' style='position:absolute; left:747; top:12; width:150; height:24; visibility:'>This Study's IRB</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:7; top:37; width:156; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='text' id=QueryName value='empty' style='position:absolute; left:490; top:42; width:189; height:30'></input>
    <label id=Label515 value='Query' style='position:absolute; left:517; top:10; width:186; height:24; visibility:'>Query</label>
    <input type='text' id=Patient Sex value='empty' style='position:absolute; left:1030; top:462; width:96; height:24'></input>
    <label id=LabelOnSite value='On Site ?' style='position:absolute; left:34; top:255; width:88; height:24; visibility:'>On Site ?</label>
    <label id=LabelEventType value='Type of Event:' style='position:absolute; left:28; top:283; width:135; height:24; visibility:'>Type of Event:</label>
    <label id=Label111 value='Study Related?' style='position:absolute; left:24; top:318; width:139; height:24; visibility:'>Study Related?</label>
    <input type='checkbox' id=OnSite value='False' style='position:absolute; left:142; top:258; width:0; height:0'></input>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:42; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <input type='text' id=OriginalSAENumber value='empty' style='position:absolute; left:402; top:201; width:120; height:30'></input>

    <label id=Label133 value='OriginalSAENumber' style='position:absolute; left:219; top:201; width:174; height:24; visibility:'>OriginalSAENumber</label>
    <input type='text' id=FU_ReportNumber value='empty' style='position:absolute; left:400; top:160; width:120; height:30'></input>

    <label id=Label135 value='FU_ReportNumber' style='position:absolute; left:228; top:160; width:165; height:24; visibility:'>FU_ReportNumber</label>
    <!--     <label id=LabelReadOnly value='All Data is Read Only' style='position:absolute; left:256; top:244; width:235; height:24; visibility:hidden'>All Data is Read Only</label> -->
    <select name='Combo140' Size='32' style='position:absolute; left:639; top:312; width:258; height:25'>    </select>

    <label id=Label141 value='Drug 1' style='position:absolute; left:564; top:312; width:66; height:24; visibility:'>Drug 1</label>
    <input type='text' id=RelatedCombo value='empty' style='position:absolute; left:172; top:315; width:249; height:24'></input>
    <input type='text' id=EventType value='empty' style='position:absolute; left:174; top:282; width:249; height:24'></input>
    <input type='checkbox' id=cbSecondaryFlag value='False' style='position:absolute; left:198; top:90; width:240; height:18'></input>

    <label id=Label144 value='Secondary Flag' style='position:absolute; left:221; top:87; width:187; height:24; visibility:'>Secondary Flag</label>
    <select name='PIList' Size='33' style='position:absolute; left:960; top:504; width:168; height:174'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [StudyStatusCodes1].[Study Status] FROM StudyStatusCodes1");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Study Status']."'".">".$row['Study Status']."</Option>";
        }
    ?>
    </select>

    <label id=PiForFiltering value='Use this Field\015\012to Select PI(s)-->\015\012 when Filtering' style='position:absolute; left:804; top:552; width:150; height:79; visibility:'>Use this Field\015\012to Select PI(s)-->\015\012 when Filtering</label>
    <input type='button' id=preview_print value='Preview SAEs' style='position:absolute; left:877; top:697; width:133; height:48' onclick='preview_print_Click();'></input>
    <input type='button' id=Returnmenu value='Return' style='position:absolute; left:1009; top:697; width:120; height:48' onclick='Returnmenu_Click();'></input>
    <!--     <input type='button' id=DeleteButton value='Delete SAE' style='position:absolute; left:741; top:697; width:133; height:48' onclick='DeleteButton_Click();'></input> -->
    <input type='button' id=butPIList value='Summary SAE List by PI(s)' style='position:absolute; left:489; top:700; width:249; height:45' onclick='butPIList_Click();'></input>
    <!--     <input type='text' id=txtHoldFilter value='empty' style='position:absolute; left:552; top:594; width:120; height:30'></input> -->
    <input type='text' id=Text159 value='empty' style='position:absolute; left:30; top:562; width:78; height:30'></input>

    <label id=Label160 value='# Days from\015\012Event to\015\012Received' style='position:absolute; left:12; top:486; width:111; height:63; visibility:'># Days from\015\012Event to\015\012Received</label>
    <input type='text' id=DatePIAware value='empty' style='position:absolute; left:313; top:564; width:132; height:30'></input>

    <label id=Label162 value='Date PI Aware' style='position:absolute; left:174; top:564; width:132; height:24; visibility:'>Date PI Aware</label>
    <select name='Investigator ID' Size='41' style='position:absolute; left:730; top:504; width:222; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [Contact Data1]![Last Name] & \", \" & [Contact Data1]![First Name] & \" \" & [Contact Data1]![Middle] & \" \" & [Contact Data1]![Suffix] AS [display name] FROM ([_select saes for irb#1] INNER JOIN [IRB - Research Database] ON [_select saes for irb#1].[IRB#]=[IRB - Research Database].[IRB#]) INNER JOIN [Contact Data1] ON [IRB - Research Database].[Investigator ID]=[Contact Data1].[Common Contact ID] ORDER BY [Contact Data1]![Last Name] & \", \" & [Contact Data1]![First Name] & \" \" & [Contact Data1]![Middle] & \" \" & [Contact Data1]![Suffix]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Last Name']."'".">".$row['Last Name']."</Option>";
        }
    ?>
    </select>

    <label id=Label165 value='Investigator' style='position:absolute; left:586; top:504; width:141; height:24; visibility:'>Investigator</label>
    <input type='text' id=AwareReceived value='empty' style='position:absolute; left:313; top:612; width:78; height:24'></input>

    <label id=Label166 value='# Days from Aware to Received' style='position:absolute; left:174; top:606; width:117; height:63; visibility:'># Days from Aware to Received</label>
    <select name='Study Active?' Size='43' style='position:absolute; left:10; top:720; width:0; height:0'>
        <Option value='CLosed'>CLosed</option>
        <Option value='Open'>Open</option>
    </select>

    <label id=Label167 value='Active?:' style='position:absolute; left:36; top:688; width:79; height:24; visibility:'>Active?:</label>
    <select name='Study Status' Size='44' style='position:absolute; left:172; top:720; width:234; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [Contact Data1].[Common Contact ID], [Contact Data1]![Last Name] & \", \" & [Contact Data1]![First Name] & \" \" & [Contact Data1]![Middle] & \" \" & [Contact Data1]![Suffix] AS [display name] FROM [Contact Data1] WHERE ((([Contact Data1].[Contact Type])=\"P.I.\")) ORDER BY [Contact Data1]![Last Name] & \", \" & [Contact Data1]![First Name] & \" \" & [Contact Data1]![Middle] & \ " \" & [Contact Data1]![Suffix]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select>

    <label id=Label168 value='Study Status:' style='position:absolute; left:207; top:688; width:123; height:24; visibility:'>Study Status:</label>
  </body>
</html>
