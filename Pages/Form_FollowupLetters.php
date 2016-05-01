<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function ButToday_Click(){
//'101019Me.UserDate.Value = Date
//Me.DatesCombo = //Date;
//Call //ReDisplay;
}
function Form_Close(){
//RenewalLetterSwitch = //False;
}
function Form_Current(){
//Call //ReDisplay;
//Forms![FollowupLetters].[followupletterssubform].Form.Requery;
//If //Demo //Then;
    //Me![followupletterssubform].Form![received].Locked = //True;
    //Me![followupletterssubform].Form![DateReceivedFollowUP].Locked = //True;
//Else;
    //If //UserEdit //Then;
        //Me![followupletterssubform].Form![received].Locked = //False;
        //Me![followupletterssubform].Form![DateReceivedFollowUP].Locked = //False;
    //Else;
        //Me![followupletterssubform].Form![received].Locked = //True;
        //Me![followupletterssubform].Form![DateReceivedFollowUP].Locked = //True;
    //End //If;
//End //If;
;
}
function Form_Load(){
//'101019Me.UserDate.Value = Date
//Me.DatesCombo = //Date;
 ;
}
function Form_Open(Cancel){
//Me.ReturnButton.SetFocus;
//OKSwitch = //False;
//Me.FUALL = //False;
;
//If //Forms![newmenu]![HoldReadOnly] //Or //AgendaDetail //Then;
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.followupletterssubform.Locked = //True;
    //Me.LettersButton.Enabled = //False;
    //Me.AllowDeletions = //False;
    ;
//End //If;
;
;
}
function LettersButton_Click(){
//'OKSwitch = False
//On //Error //GoTo //Err_LettersButton_Click:;
;
//If //Demo //Then;
    //DemoMsg;
    //Exit //Sub;
//End //If;
;
//If //OKSwitch = //False //Then;
    //MsgBox "You must first select a Study";
    //Exit //Sub;
//End //If;
;
    //FollowupLetterSwitch = //True;
    //OKSwitch = //False;
//'101019    HoldMeetingDate = Me.UserDate
    //HoldMeetingDate = //Me.DatesCombo //'101019
    //LetterName = "Follow-Up";
    var stDocName = '';
    var stLinkCriteria = '';
    //Forms!newmenu.study = //HoldIRBNo;
    ;
//'stLinkCriteria = "[IRB#]=" & "'" & HoldIRBNo & "'"
;
    stDocName;
       ;
    //If //GetTemplateLocation("FULetters") //Then;
        ;
        //If //TemplateName = "" //Then //Exit //Sub;
        //DoCmd.SetWarnings //False;
        //DoCmd.OpenQuery //("ExportFUWithCreateTable");
        //DoCmd.SetWarnings //True;
        //CustomLetter = //True;
    //Else;
        //CustomLetter = //False;
    //End //If;
   ;
    location.href = "Form_"+stDocName+".php";;
    ;
;
  //Me.Visible = //False;
;
//Exit_LettersButton_Click:;
    //Exit //Sub;
//Err_LettersButton_Click:;
    //MsgBox //Err.Description;
    //GoTo //Exit_LettersButton_Click;
}
function ReturnButton_Click(){
	window.history.back();
//DoCmd.Close;
}
function ShowOption_AfterUpdate(){
//'101019If Me.UserDate.Value = "12/31/2099" Then
//If //Me.DatesCombo.Value = "12/31/2099" //Then //'101019
//'101019        Me.UserDate.Value = Date
        //Me.DatesCombo = //Date;
//End //If;
//Call //ReDisplay;
}
function ReDisplay(){
//Me.[followupletterssubform].Form.FilterOn = //True;
//Select //Case //Me.ShowOption;
 //Case //1 //'Show Open
     //Me.FUALL = //False;
    //Forms![newmenu]![HoldFU] = //True;
    ;
    //Me.[followupletterssubform].Form.Filter = "((LettersSent.FollowUpDate)<=[Forms]![FollowupLetters]![userdate]) AND ((LettersSent.Received)=False)";
 //Case //2  //'Show Open and Closed
    //Forms![newmenu]![HoldFU] = //True;
    //Me.FUALL = //False;
    //Me.[followupletterssubform].Form.Filter = "((LettersSent.FollowUpDate)<=[Forms]![FollowupLetters]![userdate])";
 //Case //3 //'Show Past Due
    //Forms![newmenu]![HoldFU] = //True;
    //Me.FUALL = //False;
    //Me.[followupletterssubform].Form.Filter = "((LettersSent.DateRequiredBy) <=[Forms]![FollowupLetters]![userdate]) AND ((LettersSent.Received)=False)";
//Case //4  //'Show all letters sent in file
    //Forms![newmenu]![HoldFU] = //True;
    //Me.FUALL = //True;
//'101019    Me.UserDate.Value = "12/31/2099"
    //Me.DatesCombo = "12/31/2099";
    ;
    //Me.[followupletterssubform].Form.Filter = "((LettersSent.DateSent) <=[Forms]![FollowupLetters]![userdate])";
    //Me.[followupletterssubform].Form.OrderBy = "letterssent.[irb#], LettersSent.DateSent ";
    //Me.[followupletterssubform].Form.OrderByOn = //True;
    ;
    //End //Select;
//Me.[followupletterssubform].Form.Requery;
;
;
;
;
}
function UserDate_AfterUpdate(){
//'Me.DatesCombo = Me.UserDate
//'Call ReDisplay
//'End Sub
;
//Private //Sub //ReportButton_Click();
//On //Error //GoTo //Err_ReportButton_Click;
//MyFilter = //Forms![FollowupLetters].[followupletterssubform].Form.Filter;
;
;
    var stDocName = '';
;
    stDocName;
    //DoCmd.OpenReport //stDocName, //acPreview, //, //MyFilter;
   ;
//Exit_ReportButton_Click:;
    //Exit //Sub;
;
//Err_ReportButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_ReportButton_Click;
}

    </script>
  </head>
  	<style>
		table, th, td {
    		border: 1px solid black;
    		border-collapse: collapse;
		}
	</style>
  <body onLoad="Form_Load();Form_Open();Form_Current();" onUnload="Form_Close();">

	<input type='checkbox' id='Option176' value='False' style='position: absolute; left: 108px; top: 81px; width: 21px; height: 20px'></input>
    <input type='checkbox' id='Option178' value='False' style='position: absolute; left: 108px; top: 102px; width: 21px; height: 20px'></input>
    <input type='checkbox' id='Option180' value='False' style='position: absolute; left: 719px; top: 83px; width: 21px; height: 20px'></input>
    <input type='checkbox' id='Option182' value='False' style='position: absolute; left: 719px; top: 103px; width: 21px; height: 20px'></input>

  <label id='empty' value='empty' style='visibility:'></label>
  <select name='empty' style='position: absolute; left: 116px; top: 351px; width: 102; height: 120; visibility: hidden;'></select>
  <textarea type='text' id=Location value='empty' style='position:absolute; left:351; top:18; width:709; height:30'>
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
    <input type='button' id=LettersButton value='Send Follow Up Letter' style='position:absolute; left:432; top:564; width:465; height:45' onclick='LettersButton_Click();'></input>
    <input type='button' id=ReturnButton value='&Return' style='position:absolute; left:895; top:564; width:169; height:45' onclick='ReturnButton_Click();'></input>
    <input type='text' id=DatesCombo value='empty' style='position:absolute; left:466; top:190; width:180; height:40'></input>
    <label id=Label147 value='For Follow-Up\015\012On or Before' style='position: absolute; left: 463px; top: 157px; width: 231px; height: 26px;visibility:'>For Follow-Up\015\012On or Before</label>
    <label id=Label151 value='-----Completed-----' style='position: absolute; left: 878px; top: 257px; width: 157; height: 24;visibility:'>-----Completed-----</label>
    <label id=Label152 value='Follow-Up Manager' style='font-size: 28px; position: absolute; left: 69px; top: 16px; width: 283; height: 40;visibility:'>Follow-Up Manager</label>


    <label id=Label174 value='Letter Display Options' style='position: absolute; left: 442px; top: 61px; width: 238; height: 28;visibility:'>Letter Display Options</label>
    <input type='radio' style='position:absolute; left:33; top:128; width:0; height:0'></input>

    <label id=Label177 value='Response Required-OPEN' style='position: absolute; left: 138px; top: 86px; width: 288; height: 28;visibility:'>Response Required-OPEN</label>
    <input type='radio' style='position:absolute; left:33; top:164; width:0; height:0'></input>

    <label id=Label179 value='Response Required-Open + Closed' style='position: absolute; left: 137px; top: 109px; width: 281px; height: 28;visibility:'>Response Required-Open + Closed</label>
    <input type='radio' style='position:absolute; left:33; top:200; width:0; height:0'></input>

    <label id=Label181 value='Response Required-Past Due Only' style='position: absolute; left: 753px; top: 88px; width: 234px; height: 28;visibility:'>Response Required-Past Due Only</label>
    <input type='radio' style='position:absolute; left:33; top:242; width:0; height:0'></input>

    <label id=Label183 value='All Letters ' style='position: absolute; left: 754px; top: 108px; width: 271; height: 28;visibility:'>All Letters </label>
    <input type='button' id=ButToday value='Switch to Today' style='position:absolute; left:466; top:241; width:180; height:30' onclick='ButToday_Click();'></input>
    <input type='button' id=ReportButton value='Follow-up Report' style='position:absolute; left:43; top:564; width:388; height:45' onclick='ReportButton_Click();'></input>
    <!--     <input type='checkbox' id=FUALL value='False' style='position:absolute; left:438; top:78; width:138; height:12'></input> -->

    <label id=Label191 value='Check190' style='position: absolute; left: 461; top: 75; width: 109; height: 28; visibility:; visibility: hidden;'>Check190</label>
    <input type='button' id=cmdCalDate value='&Hire Date Calendar' style='position:absolute; left:648; top:192; width:33; height:33'></input>
    <table style='width: 1012px; position: absolute; left: 49px; top: 278px; height: 284px;'>
  
    <tr>
    <th width="8%">Study#</th>
    <th width="10%">F/I</th>
    <th width="11%">F/U Date</th>
    <th width="8%">Letter</th>
    <th width="8%">To</th>
    <th width="8%">Date Sent</th> 
    <th width="7%">Due By</th>
    <th width="6%">Recv'd</th>
    <th width="7%">Date</th>
    <th width="9%">LetterKey</th>
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
    </tr>
    </table>
    
</body>
</html>
