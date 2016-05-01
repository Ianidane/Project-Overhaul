<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function CancelButton_Click(){
//DoCmd.Close;
;
//Exit_CancelButton_Click:;
    //Exit //Sub;
;
//Err_CancelButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_CancelButton_Click;
    ;
}
function Command14_Click(){
//End //Sub;
;
//Private //Sub //Command30_Click();
;
//DoCmd.Close;
;
//Exit_CancelButton_Click:;
    //Exit //Sub;
;
//Err_CancelButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_CancelButton_Click;
;
}
function Command31_Click(){
var stDocName = '';
//DoCmd.RunCommand //acCmdSaveRecord;
//MyFilter = "MeetingDate = " //& //Forms![_postmeetingform]![Meeting_date];
stDocName;
//'DoCmd.Close
//DoCmd.OpenReport //stDocName, //acViewPreview;
//If //noRecords = //True //Then;
   //noRecords = //False;
   //DoCmd.Close //acReport, "MeetingMinutes2_large";
//End //If;
;
}
function Command32_Click(){
//End //Sub;
;
;
;
//Private //Sub //CheckSites_Click();
//If //CheckSites //Then;
    //G_PrtSites = //True;
    //'MsgBox "Note that Sites are available on the Small Font Report Only"
//Else;
    //G_PrtSites = //False;
//End //If;
;
}
function CKCoINV_Click(){
//If //CKCoINV //Then //G_PrtCoInv = //True //Else //G_PrtCoInv = //False;
}
function Form_Open(Cancel){
//G_PrtCoInv = //False;
//CancelUpdate = //False;
//G_PrtSites = //False;
;
//If //Forms![newmenu]![HoldReadOnly] //Or //AgendaDetail //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
//End //If;
;
}
function PreviewButton_Click(){
var stDocName = '';
//DoCmd.RunCommand //acCmdSaveRecord;
//MyFilter = "MeetingDate = " //& //Forms![_postmeetingform]![Meeting_date];
stDocName;
//'DoCmd.Close
//DoCmd.OpenReport //stDocName, //acViewPreview;
//If //noRecords = //True //Then;
   //noRecords = //False;
   //DoCmd.Close //acReport, "MeetingMinutes2_large";
//End //If;
;
}
function PreviewButtonSmall_Click(){
var stDocName = '';
//DoCmd.RunCommand //acCmdSaveRecord;
//MyFilter = "MeetingDate = " //& //Forms![_postmeetingform]![Meeting_date];
stDocName;
//'DoCmd.Close
//DoCmd.OpenReport //stDocName, //acViewPreview;
//If //noRecords = //True //Then;
   //noRecords = //False;
   //DoCmd.Close //acReport, "MeetingMinutes2";
//End //If;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>






    <label id=Board Member Attendance Label value='Board Member Attendance' style='position:absolute; left:103; top:273; width:285; height:28; visibility:'>Board Member Attendance</label>
    <label id=Label2 value='Meeting Minutes Preparation' style='position:absolute; left:450; top:4; width:337; height:36; visibility:'>Meeting Minutes Preparation</label>
    <input type='text' id=TextCall value='empty' style='position:absolute; left:619; top:303; width:481; height:60'></input>

    <label id=Label6 value='CALL TO ORDER' style='position:absolute; left:784; top:273; width:187; height:28; visibility:'>CALL TO ORDER</label>
    <input type='text' id=TextApprov value='empty' style='position:absolute; left:619; top:405; width:481; height:60'></input>

    <label id=Label8 value='APPROVAL OF MINUTES' style='position:absolute; left:784; top:375; width:273; height:28; visibility:'>APPROVAL OF MINUTES</label>
    <input type='text' id=TextOld value='empty' style='position:absolute; left:619; top:502; width:481; height:87'></input>

    <label id=Label10 value='OLD BUSINESS' style='position:absolute; left:784; top:469; width:174; height:28; visibility:'>OLD BUSINESS</label>
    <input type='text' id=TextGuests value='empty' style='position:absolute; left:0; top:655; width:613; height:60'></input>

    <label id=Label12 value='GUESTS/STAFF' style='position:absolute; left:153; top:622; width:181; height:28; visibility:'>GUESTS/STAFF</label>
    <input type='text' id=TextNew value='empty' style='position:absolute; left:619; top:655; width:481; height:159'></input>

    <label id=Label16 value='NEW BUSINESS' style='position:absolute; left:780; top:622; width:181; height:28; visibility:'>NEW BUSINESS</label>
    <input type='text' id=QuorumText value='empty' style='position:absolute; left:0; top:747; width:613; height:60'></input>

    <label id=Label18 value='QUORUM NOTES' style='position:absolute; left:186; top:717; width:193; height:28; visibility:'>QUORUM NOTES</label>
    <input type='text' id=Preparer value='empty' style='position:absolute; left:820; top:84; width:289; height:28'></input>
    <select name='ComboChair' style='position:absolute; left:0; top:72; width:510; height:42'>
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

        $sql = sqlsrv_query($conn, "SELECT [Meetings3MonthsPastandAllFuture].[MeetingDate] FROM Meetings3MonthsPastandAllFuture");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
    </select>

    <label id=Label25 value='Chairperson' style='position:absolute; left:0; top:42; width:135; height:28; visibility:'>Chairperson</label>
    <select name='NextDate' Size='17' style='position:absolute; left:198; top:922; width:130; height:0'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW ([IRBMeetingAttendance1]![AddressTo]) & \" - \" & [IRBMeetingAttendance1]![Representing] AS NameandTitle, [IRBMeetingAttendance1].[MeetingKey], [IRBMeetingAttendance1].[Common Contact ID], [IRBMeetingAttendance1].[DisplayName] FROM IRBMeetingAttendance1 WHERE ((([IRBMeetingAttendance1].[MeetingKey])=[Forms]![NewMenu]![MeetingKey])) ORDER BY [IRBMeetingAttendance1].[DisplayName]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['AddressTo']."'".">".$row['AddressTo']."</Option>";
        }
    ?>
    </select>

    <label id=Label211 value='Next Meeting Date' style='position:absolute; left:6; top:922; width:186; height:33; visibility:'>Next Meeting Date</label>
    <input type='text' id=TimeText value='empty' style='position:absolute; left:396; top:922; width:93; height:0'></input>

    <label id=Label27 value='Time' style='position:absolute; left:336; top:922; width:61; height:30; visibility:'>Time</label>
    <input type='text' id=MeetLocation value='empty' style='position:absolute; left:589; top:922; width:514; height:0'></input>

    <label id=Label29 value='Location' style='position:absolute; left:496; top:922; width:90; height:30; visibility:'>Location</label>
    <input type='text' id=MinutesHeadingText value='empty' style='position:absolute; left:51; top:207; width:1020; height:60'></input>

    <label id=Label35 value='Minutes Heading Paragraph' style='position:absolute; left:375; top:172; width:300; height:28; visibility:'>Minutes Heading Paragraph</label>
    <input type='text' id=TextUser value='empty' style='position:absolute; left:12; top:855; width:613; height:60'></input>

    <label id=Label44 value='Enter the label\015\012you wish to use' style='position:absolute; left:22; top:813; width:150; height:43; visibility:'>Enter the label\015\012you wish to use</label>
    <input type='text' id=EducationText value='empty' style='position:absolute; left:619; top:856; width:481; height:60'></input>

    <label id=Label36 value='EDUCATION/TRAINING' style='position:absolute; left:711; top:822; width:256; height:28; visibility:'>EDUCATION/TRAINING</label>
    <input type='text' id=LabelUser value='empty' style='position:absolute; left:178; top:819; width:402; height:28'></input>
    <input type='text' id=AgendaTime value='empty' style='position:absolute; left:468; top:133; width:93; height:0'></input>

    <label id=Label38 value='Time' style='position:absolute; left:408; top:126; width:61; height:30; visibility:'>Time</label>
    <input type='text' id=AgendaLocation value='empty' style='position:absolute; left:661; top:133; width:411; height:0'></input>

    <label id=Label40 value='Location' style='position:absolute; left:568; top:126; width:90; height:30; visibility:'>Location</label>
    <input type='text' id=TextMeeting value='empty' style='position:absolute; left:270; top:133; width:130; height:0'></input>

    <label id=Label42 value='Meeting Date' style='position:absolute; left:132; top:126; width:135; height:33; visibility:'>Meeting Date</label>
    <label id=Label45 value='HPA/CoOrd/or Other Label' style='position:absolute; left:549; top:52; width:283; height:28; visibility:'>HPA/CoOrd/or Other Label</label>
    <label id=Label41 value='Name' style='position:absolute; left:955; top:52; width:70; height:28; visibility:'>Name</label>
    <input type='text' id=PreparerLabel value='empty' style='position:absolute; left:553; top:84; width:262; height:28'></input>


    <input type='button' id=CancelButton value='&Return' style='position:absolute; left:1015; top:42; width:97; height:40' onclick='CancelButton_Click();'></input>
    <input type='button' id=PreviewButton value='Preview (&Large Fonts)' style='position:absolute; left:735; top:42; width:280; height:40' onclick='PreviewButton_Click();'></input>
    <input type='button' id=PreviewButtonSmall value='Preview (&Small Fonts)' style='position:absolute; left:454; top:42; width:280; height:40' onclick='PreviewButtonSmall_Click();'></input>
    <label id=Label43 value='To view more data use the scroll bar--------' style='position:absolute; left:733; top:9; width:376; height:24; visibility:'>To view more data use the scroll bar--------</label>
    <input type='checkbox' id=CKCoINV value='False' style='position:absolute; left:4; top:46; width:25; height:10'></input>

    <label id=Label49 value='Check to Print Co-Investigators' style='position:absolute; left:22; top:37; width:337; height:28; visibility:'>Check to Print Co-Investigators</label>
    <input type='checkbox' id=CheckReportDate value='=True' style='position:absolute; left:6; top:13; width:85; height:19'></input>

    <label id=Label47 value='Print Report Date' style='position:absolute; left:22; top:6; width:198; height:28; visibility:'>Print Report Date</label>
    <input type='checkbox' id=chkItemNumber value='=True' style='position:absolute; left:6; top:72; width:85; height:19'></input>

    <label id=lbItemNumberPrint value='Print Item Numbers' style='position:absolute; left:22; top:66; width:214; height:28; visibility:'>Print Item Numbers</label>
    <input type='checkbox' id=CheckSites value='=False' style='position:absolute; left:252; top:13; width:85; height:19'></input>

    <label id=Label54 value='Print Sites' style='position:absolute; left:268; top:6; width:198; height:28; visibility:'>Print Sites</label>
  </body>
</html>
