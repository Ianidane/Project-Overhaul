<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function cmdFilterRecords_Click(){
//On //Error //GoTo //Err_cmdFilterRecords_Click;
//If //Forms![_IRBSearchNew].Filter //Like "Exists*" //Then;
        //'msgbox Me.Filter
        //MsgBox "You cannot Filter these records any further.";
        //Exit //Sub;
//Else;
    //SaveFilter = //Forms![_IRBSearchNew].Filter;
//End //If;
;
//Me.Visible = //False;
;
    var stDocName = '';
    var stLinkCriteria = '';
;
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
;
//Exit_cmdFilterRecords_Click:;
    //Exit //Sub;
;
//Err_cmdFilterRecords_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdFilterRecords_Click;
    ;
}
function Form_Close(){
//MyFilter = "None";
//Forms![_IRBSearchNew].Filter = "";
;
//Me.Filter = "";
//SaveFilter = " ";
//If //isFormLoaded("_irbsearchnew_qbf") //Then;
    //DoCmd.Close //acForm, "_irbsearchnew_qbf";
//End //If;
;
}
function Form_Current(){
//Me.lstReviewer.Requery;
;
//Me.lstCoInv.Requery;
//If //Me.NewRecord //And //Me.FilterOn = //True //Then;
    //MsgBox "There are no Studies meeting the filter criteria";
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
}
function Form_Open(Cancel){
//'msgbox Me.Filter
//If //HoldLocation = "***" //Then;
        //Me.SearchLocation = "All IRB's";
    //Else;
        //Me.SearchLocation = //HoldLocation;
    //End //If;
//Me.QueryName = //QueryString;
//Me.OrderBy = "";
//If //isFormLoaded("_irbsearchnew_qbf") //Then;
    //DoCmd.Close //acForm, "_irbsearchnew_qbf";
//End //If;
//SaveFilter = " ";
}
function Form_Timer(){
//On //Error //GoTo //Err_timer;
//Me.TimerInterval = //0;
//If //Me.RecordsetClone.RecordCount = //0 //Then;
    //MsgBox "There are no Studies for the Search Query or Filter Criteria.";
    //Me.Filter = "";
    //'Me.Requery
    //DoCmd.Close;
//End //If;
//exit_timer:;
//Exit //Sub;
//Err_timer:;
//MsgBox //Err.Description //& " " //& //Err.Number;
//GoTo //exit_timer;
;
}
function FullPrintButton_Click(){
//On //Error //GoTo //Err_FullPrintButton_Click;
//If //Me.FilterOn //Then;
     //MyFilter = //Me.Filter;
//Else;
    //MyFilter = "";
    //Forms![_IRBSearchNew].Filter = "";
//End //If;
    var stDocName = '';
   ;
    stDocName;
    //DoCmd.OpenReport //stDocName, //acViewPreview, //, //MyFilter;
    ;
    //Reports![SearchResults_FT_Details].OrderBy = "";
    //If //Forms![_IRBSearchNew].OrderBy = "" //Then;
        //Reports![SearchResults_FT_Details].OrderBy = "[IRB#]";
    //Else;
        //Reports![SearchResults_FT_Details].OrderBy = //Forms![_IRBSearchNew].OrderBy;
    //End //If;
    //Reports![SearchResults_FT_Details].OrderByOn = //True;
     //'DoCmd.RunCommand acCmdZoom75
//Exit_FullPrintButton_Click:;
    //Exit //Sub;
;
//Err_FullPrintButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_FullPrintButton_Click;
    ;
;
}
function Return_to_Menu_Click(){
//MyFilter = "";
//Me.Filter = "";
    ;
//DoCmd.Close;
}
function Switch_Click(){//'redisplays input form with this Study in edit mode
//If //isFormLoaded("_irbsearchnew_qbf") //Then;
    //DoCmd.Close //acForm, "_irbsearchnew_qbf";
//End //If;
;
//Forms![newmenu]![study] = //Me.ID_;
//holdcaption = "Study #  " //& //Forms![newmenu]![study];
//HoldLocation = //Me.IRB_Code;
//Forms![newmenu]![Location] = //Me.IRB_Code;
//DoCmd.Close;
location.href = "Form__irb input form.php";;
}
function PreviewSummary_Click(){
//On //Error //GoTo //Err_PreviewSummary_Click;
var stDocName = '';
//could initialize MyReport here
//If //YesNo("Do //you //wish //to //print //the //Full //Title //for //each //Study?") //Then;
    stDocName;
//Else;
    stDocName;
//End //If;
//'MsgBox Me.Caption
//If //Me.FilterOn //Then;
     //MyFilter = //Forms![_IRBSearchNew].Filter;
//Else;
    //MyFilter = "";
    //Forms![_IRBSearchNew].Filter = "";
//End //If;
    //SaveFilter = //MyFilter;
    //Me.HoldFilter = //SaveFilter;
    //DoCmd.OpenReport //stDocName, //acViewPreview, //, //MyFilter //' orig
    ;
    //If stDocName;
        //Set MyReport;
    //Else;
        //Set MyReport;
    //End //If;
    //MyReport.OrderBy = "";
    //'Reports![SearchResults_Summary].OrderBy = ""
    //If //Forms![_IRBSearchNew].OrderBy = "" //Then;
    //MyReport.OrderBy = "[IRB#]";
        //'Reports![SearchResults_Summary].OrderBy = "[expr1]"
    //Else;
        //MyReport.OrderBy = //Forms![_IRBSearchNew].OrderBy;
        //'Reports![SearchResults_Summary].OrderBy = Me.OrderBy
    //End //If;
    //MyReport.Filter = //SaveFilter;
    //MyReport.OrderByOn = //True //' this line was commented out on Lennox Miranda's
    //'Reports![SearchResults_Summary].OrderByOn = True
    //'DoCmd.RunCommand acCmdZoom75
//Exit_PreviewSummary_Click:;
    //Exit //Sub;
;
//Err_PreviewSummary_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_PreviewSummary_Click;
    ;
}
function FilterButton_Click(){
//On //Error //GoTo //Err_FilterButton_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
    ;
    ;
    //If //Me.Filter //Like "Exists*" //Then;
        //'msgbox Me.Filter
        //MsgBox "You cannot Filter these records any further.";
        //Exit //Sub;
    //End //If;
    //Me.Visible = //False;
    //Me.Filter = "";
    //MyFilter = "None";
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
  ;
//Exit_FilterButton_Click:;
    //Exit //Sub;
;
//Err_FilterButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_FilterButton_Click;
    ;
}
function cmdMore_Click(){
//On //Error //GoTo //Err_cmdMore_Click;
;
    var stDocName = '';
    var stLinkCriteria = '';
    stLinkCriteria = "[irb#] = " && "'" && Me.[IRB#] && "'";
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //, //, stLinkCriteria;
;
//Exit_cmdMore_Click:;
    //Exit //Sub;
;
//Err_cmdMore_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_cmdMore_Click;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>



    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>

    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>


    <input type='text' id=ID# value='empty' style='position:absolute; left:951; top:48; width:118; height:25'></input>

    <label id=Text15 value='IRB#:' style='position:absolute; left:975; top:12; width:57; height:24; visibility:'>IRB#:</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:144; top:133; width:396; height:96'></input>

    <label id=Text27 value='Protocol#/Title' style='position:absolute; left:6; top:133; width:145; height:43; visibility:'>Protocol#/Title</label>
    <input type='text' id=Approval Date  (original) value='empty' style='position:absolute; left:924; top:249; width:0; height:0'></input>

    <label id=Text31 value='Original Approval Date:' style='position:absolute; left:699; top:249; width:216; height:24; visibility:'>Original Approval Date:</label>
    <input type='text' id=Last Renewal Date value='empty' style='position:absolute; left:924; top:282; width:0; height:0'></input>

    <label id=Text33 value='Last IRB Renewal Date:' style='position:absolute; left:700; top:282; width:214; height:24; visibility:'>Last IRB Renewal Date:</label>
    <input type='text' id=# Patients enrolled value='empty' style='position:absolute; left:923; top:406; width:0; height:25'></input>

    <label id=Text37 value='Currently Enrolled:' style='position:absolute; left:750; top:406; width:165; height:24; visibility:'>Currently Enrolled:</label>
    <input type='button' id=Return to Menu value='&Return' style='position:absolute; left:948; top:621; width:135; height:48' onclick='Return to Menu_Click();'></input>
    <input type='text' id=InitEnroll value='empty' style='position:absolute; left:923; top:370; width:0; height:0'></input>

    <label id=Label470 value='Approved Patient Enrollment' style='position:absolute; left:661; top:370; width:250; height:24; visibility:'>Approved Patient Enrollment</label>
    <input type='button' id=Switch value='Switch to Edit Study' style='position:absolute; left:403; top:621; width:166; height:48' onclick='Switch_Click();'></input>
    <input type='text' id=Date Received value='empty' style='position:absolute; left:924; top:187; width:0; height:19'></input>

    <label id=Label475 value='Date Received:' style='position:absolute; left:771; top:187; width:142; height:24; visibility:'>Date Received:</label>
    <input type='text' id=First IRB Review value='empty' style='position:absolute; left:924; top:217; width:0; height:22'></input>

    <label id=Label476 value='First IRB Review:' style='position:absolute; left:757; top:217; width:157; height:24; visibility:'>First IRB Review:</label>
    <input type='text' id=TExpireDate value='empty' style='position:absolute; left:924; top:153; width:0; height:0'></input>

    <label id=Label492 value='Expiration Date' style='position:absolute; left:768; top:153; width:142; height:24; visibility:'>Expiration Date</label>
    <input type='text' id=MostRecentIRBMeeting value='empty' style='position:absolute; left:923; top:316; width:0; height:0'></input>

    <label id=Label497 value='Last seen by IRB:' style='position:absolute; left:753; top:316; width:160; height:24; visibility:'>Last seen by IRB:</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:33; top:39; width:159; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <input type='text' id=Study Active? value='empty' style='position:absolute; left:103; top:568; width:0; height:30'></input>

    <label id=Label454 value='Active?:' style='position:absolute; left:16; top:573; width:79; height:24; visibility:'>Active?:</label>
    <input type='text' id=StudyStatus value='empty' style='position:absolute; left:381; top:570; width:276; height:30'></input>

    <label id=Label441 value='Study Status:' style='position:absolute; left:252; top:570; width:123; height:24; visibility:'>Study Status:</label>
    <input type='button' id=PreviewSummary value='Results &Summary  Report ' style='position:absolute; left:760; top:621; width:187; height:48' onclick='PreviewSummary_Click();'></input>
    <label id=PIlabel value='P.I.' style='position:absolute; left:94; top:277; width:39; height:24; visibility:'>P.I.</label>
    <label id=Label447 value='Coordinator' style='position:absolute; left:24; top:312; width:106; height:24; visibility:'>Coordinator</label>
    <input type='text' id=IRB Code value='empty' style='position:absolute; left:639; top:48; width:301; height:0'></input>

    <label id=Label453 value='Study IRB' style='position:absolute; left:714; top:18; width:136; height:24; visibility:'>Study IRB</label>
    <input type='text' id=Renewal Cycle value='empty' style='position:absolute; left:925; top:100; width:34; height:0'></input>

    <label id=Label471 value='Renewal Cycle:' style='position:absolute; left:759; top:100; width:141; height:24; visibility:'>Renewal Cycle:</label>
    <label id=Label516 value='Months' style='position:absolute; left:969; top:100; width:70; height:24; visibility:'>Months</label>
    <input type='text' id=Sponsor value='empty' style='position:absolute; left:144; top:244; width:396; height:0'></input>

    <label id=Sponsor_Label value='Sponsor' style='position:absolute; left:55; top:244; width:78; height:24; visibility:'>Sponsor</label>
    <input type='text' id=PI value='empty' style='position:absolute; left:144; top:276; width:396; height:0'></input>
    <input type='button' id=FilterButton value='Filter Records' style='position:absolute; left:3; top:621; width:172; height:48' onclick='FilterButton_Click();'></input>
    <input type='text' id=SearchLocation value='empty' style='position:absolute; left:240; top:46; width:372; height:30'></input>

    <label id=Label500 value='Search IRB' style='position:absolute; left:310; top:10; width:177; height:28; visibility:'>Search IRB</label>
    <input type='text' id=QueryName value='empty' style='position:absolute; left:312; top:93; width:270; height:30'></input>
    <label id=Label515 value='Query Name' style='position:absolute; left:106; top:93; width:187; height:24; visibility:'>Query Name</label>
    <label id=Label488 value='On Agenda for:' style='position:absolute; left:780; top:444; width:138; height:24; visibility:'>On Agenda for:</label>
    <input type='text' id=CO value='empty' style='position:absolute; left:144; top:310; width:396; height:0'></input>
    <input type='button' id=cmdMore value='More Study Data.....' style='position:absolute; left:555; top:138; width:198; height:40' onclick='cmdMore_Click();'></input>
    <input type='button' id=cmdFilterRecords value='Additional Filter Fields' style='position:absolute; left:177; top:621; width:226; height:48' onclick='cmdFilterRecords_Click();'></input>
    <input type='button' id=FullPrintButton value='Results  Report \015\012(&Full Details)' style='position:absolute; left:570; top:621; width:192; height:48' onclick='FullPrintButton_Click();'></input>
    <input type='text' id=MaxOfMeetingDate value='empty' style='position:absolute; left:923; top:444; width:0; height:0'></input>
    <input type='text' id=ExemptCite value='empty' style='position:absolute; left:788; top:549; width:280; height:0'></input>

    <label id=Label528 value='ExemptCite:' style='position:absolute; left:661; top:549; width:109; height:24; visibility:'>ExemptCite:</label>
    <input type='text' id=ExpediteCite value='empty' style='position:absolute; left:788; top:585; width:280; height:0'></input>

    <label id=Label529 value='ExpediteCite:' style='position:absolute; left:661; top:585; width:121; height:24; visibility:'>ExpediteCite:</label>
    <select name='lstReviewer' Size='16' style='position:absolute; left:144; top:348; width:397; height:46'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_SelectCoInForSearchForm].[display name], [_SelectCoInForSearchForm].[StudyNo] FROM _SelectCoInForSearchForm WHERE ((([_SelectCoInForSearchForm].[StudyNo])=[Forms]![_IRBSearchNew]![ID#])) ORDER BY [_SelectCoInForSearchForm].[display name] WITH OWNERACCESS OPTION");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['display name']."'".">".$row['display name']."</Option>";
        }
    ?>
    </select>

    <label id=Label508 value='Reviewer(s)' style='position:absolute; left:18; top:348; width:115; height:24; visibility:'>Reviewer(s)</label>
    <input type='text' id=TextRemarks value='empty' style='position:absolute; left:144; top:504; width:397; height:55'></input>

    <label id=Label521 value='Remarks' style='position:absolute; left:54; top:504; width:79; height:24; visibility:'>Remarks</label>
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:40; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
    <select name='lstCoInv' Size='31' style='position:absolute; left:144; top:406; width:397; height:46'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [_SelectReviewerforInputForm].[Common Contact ID], [_SelectReviewerforInputForm].[Contact Type], [_SelectReviewerforInputForm].[Display Name] FROM _SelectReviewerforInputForm WHERE ((([_SelectReviewerforInputForm].[StudyNo])=[Forms]![_IRBSearchNew]![ID#])) ORDER BY [_SelectReviewerforInputForm].[Display Name] WITH OWNERACCESS OPTION");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Common Contact ID']."'".">".$row['Common Contact ID']."</Option>";
        }
    ?>
    </select>

    <label id=Label531 value='Co-        Investigator(s) ' style='position:absolute; left:6; top:402; width:130; height:49; visibility:'>Co-        Investigator(s) </label>
    <!--     <input type='text' id=HoldFilter value='empty' style='position:absolute; left:600; top:198; width:120; height:0'></input> -->

    <label id=Label533 value='Text532:' style='position:absolute; left:606; top:240; width:70; height:24; visibility:'>Text532:</label>
  </body>
</html>
