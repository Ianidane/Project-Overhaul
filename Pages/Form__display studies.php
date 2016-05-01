<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<?php include 'Serverinfo.php'; ?> 
<html>
  <head>
    <link href="ProIRB_PHP/images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function AgendaOnly_Click(){
//'OKswitch is true when an item on the subform is selected
//If //OKSwitch //Then;
    //OKSwitch = //False;
    //RenewalLetterSwitch = //True;
    //Call //PlaceRenewalonAgenda;
    //Forms![_display //studies]![_display //studies //subform].Form.RecordSource = "SelectALLupForRenewals";
    //Exit //Sub;
//Else;
    //MsgBox "You must first select a Study";
//End //If;
}
function CheckNotOnAgenda_Click(){
//If //Me.CheckNotOnAgenda //Then;
//Me.TextSHowOnly = "-1";
    ;
//Else;
    //Me.TextSHowOnly = "*";
//End //If;
//Forms![_display //studies]![_display //studies //subform].Form.RecordSource = "SelectALLupForRenewals";
}
function DatesCombo_AfterUpdate(){
//could initialize tempdate here
//Me.MeetingDate = //Me.DatesCombo;
//'trying this
;
//If //IsNull(Me.DatesCombo.ItemData(Me.DatesCombo.ListIndex //+ //1)) //Then;
    //MsgBox "You must add additional meeting dates before selecting the last one";
    //Exit //Sub;
//Else;
    tempdate;
    //Me.textholddate = tempdate;
  //End //If;
//Forms![_display //studies]![_display //studies //subform].Form.RecordSource = "SelectALLupForRenewals";
}
function Form_Activate(){
//If //HoldRowNumber //> //0 //Then;
    ;
    //Forms![_display //studies]![_display //studies //subform].Form.RecordSource = "SelectALLupForRenewals";
    ;
//End //If;
;
}
function Form_Close(){
//RenewalLetterSwitch = //False;
}
function Form_Open(Cancel){
//Me.ReturnButton.SetFocus;
//Me.AllowFilters = //False;
//Me.TextSHowOnly = "*";
//'OKswitch is true when an item on the subform is selected
//OKSwitch = //False;
//If //Forms![newmenu]![HoldReadOnly] //Or //AgendaDetail //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.AgendaOnly.Enabled = //False;
    //Me.LettersButton.Enabled = //False;
    //Me.DatesCombo.Locked = //False;
//End //If;
;
;
;
;
;
}
function LetterOnly_Click(){
//On //Error //GoTo //Err_Letteronly_Click;
//If //OKSwitch = //False //Then;
    //MsgBox "You must first select a Study";
    //Exit //Sub;
//End //If;
;
//RenewalLetterSwitch = //True;
    ;
    var stDocName = '';
    var stLinkCriteria = '';
    //searchirbs = //False;
    stLinkCriteria = "[IRB#]=" && "'" && HoldIRBNo && "'";
    stDocName;
    location.href = "Form_"+stDocName+".php";//, //acNormal, //, //stLinkCriteria, //, //acWindowNormal;
    //Me.Visible = //False;
//Exit_Letteronly_Click:;
    //Exit //Sub;
//Err_Letteronly_Click:;
    //MsgBox //Err.Description;
    //GoTo //Exit_Letteronly_Click;
;
}
function LettersButton_Click(){
//Call //AgendaAndLetter;
}
function PrintListing_Click(){
var HoldOrderBy = '';
HoldOrderBy;
//'msgbox HoldOrderBy
//On //Error //GoTo //Err_Print_Listing_Click;
;
;
    var stDocName = '';
;
    stDocName;
    //DoCmd.OpenReport //stDocName, //acPreview;
    ;
    //If HoldOrderBy;
        //Reports![PrintStudies //UpForRenewals].OrderBy = "ExpirationDate DESC, IRB#";
        ;
    //Else;
        //Reports![PrintStudies //UpForRenewals].OrderBy = HoldOrderBy;
        ;
    //End //If;
    //Reports![PrintStudies //UpForRenewals].OrderByOn = //True;
;
    ;
    ;
    ;
;
//Exit_Print_Listing_Click:;
    //Exit //Sub;
;
//Err_Print_Listing_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Print_Listing_Click;
}
function ReturnButton_Click(){
	window.history.back();
//DoCmd.Close;
}
function PlaceRenewalonAgenda(){
//On //Error //GoTo //error_Agenda;
//HoldMeetingDate = //Me.MeetingDate;
//DoAssignAgain1:;
location.href = "Form_ReassignAgendaRecord.php";//, //, //, //, //, //acDialog;
;
//If //ReassignSwitch = //True //Then;
        //Me.MeetingDate = //HoldMeetingDate;
        //If //CheckAgenda = //True //Then;
            //ReassignSwitch = //False;
            //GoTo //exit_Agenda;
        //Else;
            //DoCmd.SetWarnings //False;
            //DoCmd.OpenQuery "_CreateAgendaforRenewal";
            //MsgBox //("Study //#- " & HoldIRBNo & "  //Assigned //to //the--" //& //HoldMeetingDate //& "  Meeting"//);
            //DoCmd.SetWarnings //True;
            ;
        //End //If;
//Else;
    //DoCmd.CancelEvent;
//End //If;
;
//exit_Agenda:;
//Exit //Sub;
//error_Agenda:;
//MsgBox "Error trying to place Renewal on Agenda.  " //& //Err.Description;
//Resume //exit_Agenda;
}
function CheckAgenda(){//As //Boolean;
//On //Error //GoTo //Error_CheckAgenda;
//could initialize hope here
//could initialize dbs here
//Set dbs;
//could initialize GoForIt1 here
;
//sqlSTR = "SELECT [agenda records].[IRB#], [agenda records].MeetingDate, [agenda records].[IRB Action Category]" //_;
//& " FROM [agenda records]WHERE [agenda records].[IRB#]= " //& "'" //& //[Forms]![_display //studies]![_display //studies //subform]![studyNo] //& "'" //& " AND [agenda records].[Agenda Category] = 'Renewal'" //& " AND [agenda records].[IRB Action Category] is null ";
//'MsgBox SqlStr
//Set GoForIt1;
;
//If //GoForIt1.RecordCount //> //0 //Then;
    //GoForIt1.MoveFirst;
    //MsgBox "The Study is already on the agenda for Renewal on - " //& //GoForIt1![MeetingDate], //vbInformation;
    //HoldMeetingDate = //GoForIt1![MeetingDate] //'added 7/23
    //Me.MeetingDate = //HoldMeetingDate //' added 7/23
    //CheckAgenda = //True;
    //Exit //Function;
//Else;
    //CheckAgenda = //False;
//End //If;
//exit_CheckAgenda:;
//Exit //Function;
;
//Error_CheckAgenda:;
//MsgBox "Error Checking for Agenda date" //& //Err.Description //& "  " //& //Err.Number;
//Resume //exit_CheckAgenda;
;
}
function AgendaAndLetter(){
//could initialize INX here
//'OKswitch is true when an item on the subform is selected
//If //OKSwitch = //False //Then;
    //MsgBox "You must first select a Study";
    //Exit //Sub;
//End //If;
    //RenewalLetterSwitch = //True;
    //Call //PlaceRenewalonAgenda;
    //'Item was not placed on agenda if reassignswitch = true
      ;
    //OKSwitch = //False;
   ;
//On //Error //GoTo //Err_AgendaAndLetter;
;
//'MsgBox ([Forms]![_display studies]![_display studies subform]![StudyNo])
    //If //GetTemplateLocation("ContinuingReview") //Then;
        //'msgbox TemplateLocation
        //'msgbox TemplateName
        //If //TemplateName = "" //Then;
            //Exit //Sub;
        //Else;
            //DoCmd.SetWarnings //False;
            //Call //BuildCoInvestigatorString(HoldIRBNo);
            //'MsgBox RTrim$(([Forms]![NewMenu]![CoInvString]))
            ;
            ;
            //Me.ExpirationDays = //(DateDiff("d", //Date, //[Forms]![_display //studies]![_display //studies //subform]![ExpirationDate]));
             //'Me.ExpirationDays  = (DateDiff("d",
            //' MsgBox Me.MeetingDate
            //'MsgBox [Forms]![_display studies]![_display studies subform]![ExpirationDate]
           //' MsgBox "DAYS>>>  " & Me.ExpirationDays
            ;
            ;
            //DoCmd.OpenQuery //("ExportContinuingReviewWithCreateTable");
            //DoCmd.SetWarnings //True;
            //CustomLetter = //True;
        //End //If;
    //Else;
        //CustomLetter = //False;
    //End //If;
    ;
;
    //LetterName = "Continuing Review Notice";
    var stDocName = '';
    var stLinkCriteria = '';
    //searchirbs = //False;
    stLinkCriteria = "[IRB#]=" && "'" && HoldIRBNo && "'";
    stDocName;
    ;
    location.href = "Form_"+stDocName+".php";//, //acNormal, //, stLinkCriteria;
     //Forms![_display //studies]![_display //studies //subform].Form.RecordSource = "SelectALLupForRenewals";
//Exit_AgendaAndLetter:;
    //Exit //Sub;
//Err_AgendaAndLetter:;
    //MsgBox //Err.Description //& //Err.Source //& //Err.Number;
    //GoTo //Exit_AgendaAndLetter;
;
}

    </script>
  </head>
  	<style>
	table, th, td {
    	border: 1px solid black;
    	border-collapse: collapse;
		}
	</style>
  <body onLoad="Form_Open();Form_Active();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>


    
    


    <input type='text' id='empty' value='empty' style='position: absolute; left: 0; top: 0; width: 210; height: 96; visibility: hidden;'></input>
  
    <select name='empty' style='position: absolute; left: 41px; top: 266px; width: 102; height: 120; visibility: hidden;'></select>

    



    <label id=Label115 value='Studies Due for Renewal at the ' style='font-size: 25px; position: absolute; left: 7px; top: 89px; width: 315px; height: 34px;visibility:'>Studies Due for Renewal at the </label>
    <label id=Label122 value='---Report Received--' style='position: absolute; left: 693px; top: 184px; width: 223; height: 28;visibility:'>---Report Received--</label>
    <textarea type='text' id=IRB value='empty' style='position: absolute; left: 334; top: 24; width: 571px; height: 30'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRB Meeting Schedule1].[IRB Code] FROM [IRB Meeting Schedule1] ORDER BY [IRB Meeting Schedule1].[IRB Code]");
        while( $row = sqlsrv_fetch_array($sql)){
        echo $row['IRB Code'];
        }
    ?>
    </textarea>

    <label id=Label124 value='IRB' style='position: absolute; left: 295px; top: 29px; width: 36px; height: 22px;visibility:'>IRB</label>
    <input type='button' id=LettersButton value='Place on  Agenda and &Print Letter' style='position:absolute; left:472; top:528; width:460; height:45' onclick='LettersButton_Click();'></input>
    <input type='button' id=ReturnButton value='&Return' style='position:absolute; left:934; top:528; width:99; height:45' onclick='ReturnButton_Click();'></input>
    <label id=Label126 value='Meeting' style='font-size: 25px; position: absolute; left: 541px; top: 83px; width: 135; height: 45;visibility:'>Meeting</label>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='font-size:10px; position: absolute; left: 49px; top: 35px; width: 166; height: 25;visibility:'>by ProIRB Plus, Inc.</label>
    <input type='button' id=AgendaOnly value='Place on  &Agenda Only' style='position:absolute; left:199; top:528; width:274; height:45' onclick='AgendaOnly_Click();'></input>
    <!--     <input type='text' id=textholddate value='empty' style='position:absolute; left:288; top:96; width:96; height:48'></input> -->

    <label id=Label27 value='t12' style='position: absolute; left: 246; top: 96; width: 61; height: 24; visibility:; visibility: hidden;'>t12</label>
    <!--     <input type='text' id=MeetingDate value='empty' style='position:absolute; left:348; top:6; width:150; height:42'></input> -->
    <input type='button' id=PrintListing value='Renewal Listing' style='position:absolute; left:10; top:528; width:189; height:45' onclick='PrintListing_Click();'></input>
    <input type='checkbox' id=CheckNotOnAgenda value='False' style='position: absolute; left: 21px; top: 169px; width: 43px; height: 27px'></input>

  <label id=Label147 value='Check to show only those Studies not already on an Agenda for this date.' style='position: absolute; left: 67px; top: 170px; width: 247; height: 97;visibility:'>Check to show only those Studies not already on an Agenda for this date.</label>
    <!--     <input type='text' id=TextSHowOnly value='empty' style='position:absolute; left:54; top:78; width:90; height:18'></input> -->

  <label id=Label149 value='Text148:' style='position: absolute; left: 0; top: 78; width: 96; height: 28; visibility:; visibility: hidden;'>Text148:</label>
    <label id=Label108 value='ProIRB' style='font-size:28px; position: absolute; left: 2px; top: 2px; width: 130; height: 43;visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='font-size:10px; position: absolute; left: 91px; top: 2px; width: 18; height: 16;visibility:'>R</label>
    <!--     <input type='text' id=TxtHoldStudyNo value='empty' style='position:absolute; left:900; top:84; width:78; height:54'></input> -->

    <!--     <label id=Label151 value='Text150:' style='position:absolute; left:756; top:84; width:96; height:28; visibility:hidden'>Text150:</label> -->
    <!--     <input type='text' id=ExpirationDays value='empty' style='position:absolute; left:156; top:18; width:138; height:42'></input> -->

  <label id=Label153 value='Text152:' style='position: absolute; left: 12; top: 18; width: 96; height: 28; visibility:; visibility: hidden;'>Text152:</label>
  <table style='width: 1012px; position: absolute; left: 11px; top: 218px; height: 298px;'>
  
    <tr>
    <th width="8%">IRB#</th>
    <th width="10%">Protocol</th>
    <th width="11%">Expires</th>
    <th width="8%">On Agenda</th>
    <th width="8%">Exp</th>
    <th width="8%">Letter?</th> 
    <th width="7%">Sent</th>
    <th width="6%">Final</th>
    <th width="7%">Progress</th>
    <th width="9%">Rec'd</th>
    <th width="9%">Last Renewl</th>
    <th width="9%">Cycle</th>
    <th width="9%">P.I.</th>
    <th width="9%">Study</th>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
    </table>
  <select name='DatesCombo' size='3' style='position: absolute; left: 328px; top: 82px; width: 204; height: 126'>
  <?php
        $serverName = $SVNM;
        $connectionInfo = array( "Database"=>$DB, "UID"=>$UID, "PWD"=>$PWD, 'ReturnDatesAsStrings'=>true);
        $conn = sqlsrv_connect( $serverName, $connectionInfo);

        if( $conn ) {
        }
        else{
            echo "Connection could not be established.<br />";
            die( print_r( sqlsrv_errors(), true));
        }

        $sql = sqlsrv_query($conn, "SELECT [IRBMeetingDates1].MeetingDate, [IRBMeetingDates1].[key] FROM IRBMeetingDates1 WHERE ((([IRBMeetingDates1].MeetingDate)>DateAdd(m,-2,Getdate()))) ORDER BY [IRBMeetingDates1].MeetingDate");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['MeetingDate']."'".">".$row['MeetingDate']."</Option>";
        }
    ?>
  </select>
  </body>
</html>
