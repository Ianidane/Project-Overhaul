<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function CheckSpellButton_Click(){
//could initialize doit here
doit;
//If doit //Then //Exit //Sub;
;
}
function CloseButton_Click(){
//could initialize mybar here
//Set mybar;
//On //Error //GoTo //Err_CloseButton_Click;
//If //IsNull(Forms![meetingdiscussion]!TextDiscussion) //Or //Forms![meetingdiscussion]!TextDiscussion = "" //Then;
    //Forms![_postmeetingform]!DiscussionButton.ForeColor = //8404992;
//Else;
    //Forms![_postmeetingform]!DiscussionButton.ForeColor = //vbCyan;
//End //If;
;
 //Forms![_postmeetingform]![holdDiscussion] = //Me.TextDiscussion;
;
     //DoCmd.Close;
;
//Exit_CloseButton_Click:;
//mybar.Enabled = //True;
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
//If //Not //IsNull(Forms![_postmeetingform]![holdDiscussion]) //Then;
;
   //Forms![meetingdiscussion]!TextDiscussion = //Forms![_postmeetingform]![holdDiscussion];
   //Forms![meetingdiscussion]!TextDiscussion.SetFocus;
    //Forms![meetingdiscussion]!TextDiscussion.SelStart = //0;
    //Forms![meetingdiscussion]!TextDiscussion.SelLength = //6556;
//End //If;
;
;
;
//If //Forms![newmenu]![HoldReadOnly] //Or //AgendaDetail //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
//End //If;
;
;
}
function CheckSpelling(){//As //Boolean;
;
//Me.TextDiscussion.SetFocus;
//Me.TextDiscussion.SelStart = //0;
//Me.TextDiscussion.SelLength = //6556;
//RunCommand //acCmdSpelling;
//CheckSpelling = //True;
;
;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <input type='text' id=TextDiscussion value='empty' style='position:absolute; left:18; top:48; width:636; height:222'></input>

    <label id=Label1 value='------------------Text of Discussion ---------------------------' style='position:absolute; left:120; top:12; width:408; height:30; visibility:'>------------------Text of Discussion ---------------------------</label>
    <input type='button' id=CloseButton value='&Close' style='position:absolute; left:552; top:282; width:102; height:45' onclick='CloseButton_Click();'></input>
    <input type='button' id=CheckSpellButton value='Check &Spelling' style='position:absolute; left:19; top:282; width:181; height:45' onclick='CheckSpellButton_Click();'></input>
  </body>
</html>
