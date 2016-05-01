<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function CheckSpellButtonVote_Click(){
//could initialize doit here
doit;
;
}
function CloseButton_Click(){
//could initialize mybar here
//Set mybar;
//On //Error //GoTo //Err_CloseButton_Click;
 //Forms![_postmeetingform]![HoldVote] = //Me.TextVote;
 //If //IsNull(Me.TextVote) //Or //Me.TextVote = "" //Then;
    //Forms![_postmeetingform]!VoteButton.ForeColor = //8404992;
//Else;
    //Forms![_postmeetingform]!VoteButton.ForeColor = //vbCyan;
//End //If;
;
 ;
//mybar.Enabled = //True;
    //DoCmd.Close;
;
//Exit_CloseButton_Click:;
    //Exit //Sub;
;
//Err_CloseButton_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_CloseButton_Click;
    ;
}
function Form_Open(Cancel){
//could initialize mybar here
//Set mybar;
//mybar.Enabled = //False;
//Me.TextVote = //Forms![_postmeetingform]![HoldVote];
;
//If //Forms![newmenu]![HoldReadOnly] //Or //AgendaDetail //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
//End //If;
;
}
function CheckSpellingVote(){//As //Boolean;
;
;
//Me.TextVote.SetFocus;
//Me.TextVote.SelStart = //0;
//Me.TextVote.SelLength = //6556;
//RunCommand //acCmdSpelling;
//CheckSpellingVote = //True;
;
;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id=TextVote value='empty' style='position:absolute; left:18; top:48; width:636; height:222'></input>

    <label id=Label1 value='----Text of Voting Results -----------' style='position:absolute; left:1; top:12; width:645; height:30; visibility:'>----Text of Voting Results -----------</label>
    <input type='button' id=CloseButton value='&Close' style='position:absolute; left:555; top:282; width:102; height:45' onclick='CloseButton_Click();'></input>
    <input type='button' id=CheckSpellButtonVote value='Check &Spelling' style='position:absolute; left:18; top:282; width:181; height:45' onclick='CheckSpellButtonVote_Click();'></input>
  </body>
</html>
