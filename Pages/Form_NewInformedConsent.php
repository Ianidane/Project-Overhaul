<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Form_Activate(){
//'DoCmd.Restore
//Me.LetterDateCombo.Requery;
//Me.LetterDateCombo = //Me.LetterDateCombo.ItemData(0);
;
;
;
}
function Form_Close(){
//CancelUpdate = //False;
//If //isFormLoaded("_IRB //Input //Form") //Then;
    //Forms![_irb //input //form].Visible = //True;
//End //If;
}
function Form_Current(){
//On //Error //GoTo //err_consent;
//Me.LetterDateCombo = //Me.LetterDateCombo.ItemData(0);
//exit_consent:;
    //Exit //Sub;
//err_consent:;
    //MsgBox "Consent Form" //& //Err.Description, //vbCritical;
    //Resume //exit_consent;
}
function Form_Open(Cancel){
//DataSaved = //False;
//CancelUpdate = //False;
//Me![InformedConsentDetail //subform].Form.DatasheetFontWeight = //400;
//If //Forms![newmenu]![HoldReadOnly] //Then;
    //Call //adhToggleControl(Me, //True);
    //Me.Caption = //Me.Caption //& " --Read Only";
    //Me.SendCorresBut.Enabled = //False;
//End //If;
;
;
}
function ReturnBut_Click(){
//DoCmd.RunCommand //acCmdSaveRecord;
//DoCmd.Close;
}
function SendCorresBut_Click(){
//could initialize intX here
//could initialize db here
//could initialize rs here
//could initialize rs2 here
//Set db;
;
;
//DoCmd.RunCommand //acCmdSaveRecord;
//Me.ltextlettername.Caption = "Informed Consent Deficiencies";
//DoCmd.SetWarnings //False;
//DoCmd.OpenQuery "_DeleteCurrentLetterParagraphs";
//DoCmd.OpenQuery "AppendDeficienciesToCurrentLetter";
;
//' start code for custom letter here
;
//'build Currentletterdeficiencytable
//DoCmd.RunSQL "Delete * from CurrentLetterDeficiencyTable";
//DoCmd.OpenQuery "AppendDummyDef";
//Set rs2;
//Set rs;
//rs.MoveLast;
//rs.MoveFirst;
//DoCmd.SetWarnings //False;
//For intX;
    //Select //Case intX;
    //Case //1;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //'msgbox rs!Wording & "   " & Me.Requirement
        //rs2.Edit;
        //rs2("Req1") = //Me.Requirement;
        //rs2("Comment1") = //Me.Comment;
        //rs2.Update;
    //Case //2;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req2") = //Me.Requirement;
        //rs2("Comment2") = //Me.Comment;
        //rs2.Update;
    //Case //3;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req3") = //Me.Requirement;
        //rs2("Comment3") = //Me.Comment;
        //rs2.Update;
        ;
        ;
    //Case //4;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req4") = //Me.Requirement;
        //rs2("Comment4") = //Me.Comment;
        //rs2.Update;
        ;
    //Case //5;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req5") = //Me.Requirement;
        //rs2("Comment5") = //Me.Comment;
        //rs2.Update;
    //Case //6;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req6") = //Me.Requirement;
        //rs2("Comment6") = //Me.Comment;
        //rs2.Update;
    //Case //7;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req7") = //Me.Requirement;
        //rs2("Comment7") = //Me.Comment;
        //rs2.Update;
    //Case //8;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req8") = //Me.Requirement;
        //rs2("Comment8") = //Me.Comment;
        //rs2.Update;
    //Case //9;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req9") = //Me.Requirement;
        //rs2("Comment9") = //Me.Comment;
        //rs2.Update;
    //Case //10;
        //Me.Requirement = //rs!Wording;
        //Me.Comment = //rs!Comments;
        //rs2.Edit;
        //rs2("Req10") = //Me.Requirement;
        //rs2("Comment10") = //Me.Comment;
        //rs2.Update;
    //End //Select;
//If intX //<> //rs.RecordCount //Then //rs.MoveNext;
//Next intX;
;
//db.Close;
;
 //If //GetTemplateLocation("ConsentFormDeficiency") //Then;
        //'msgbox TemplateLocation
        //'msgbox TemplateName
        //If //TemplateName = "" //Then;
            //Exit //Sub;
        //Else;
            //DoCmd.SetWarnings //False;
            //DoCmd.OpenQuery //("ExportDefWithCreateTable");
            //DoCmd.SetWarnings //True;
            ;
            //CustomLetter = //True;
        //End //If;
    //Else;
        //'msgbox "Do PROIRB SAE Letter"
        //CustomLetter = //False;
    //End //If;
   ;
;
;
var stDocName = '';
var stLinkCriteria = '';
    //searchirbs = //False;
    //FromICChecklist = //True;
    //HoldIRBNo = //Me.[ID_];
    stLinkCriteria = "[IRB#]=" && "'" && HoldIRBNo && "'";
    //LetterName = "Informed Consent Deficiencies";
    stDocName;
    //DoCmd.SetWarnings //True;
   //DoCmd.Close;
   //'Me.Visible = False
    location.href = "Form_"+stDocName+".php";//, //acNormal, //, //stLinkCriteria, //, //acWindowNormal;
    ;
    ;
;
;
}

    </script>
  </head>
  <body onLoad="Form_Open();Form_Active();Form_Current();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>



    <input type='text' id=ID# value='empty' style='position:absolute; left:96; top:18; width:132; height:30'></input>

    <label id=Text15 value='Study #' style='position:absolute; left:6; top:19; width:75; height:24; visibility:'>Study #</label>
    <input type='text' id=ICNumber value='empty' style='position:absolute; left:156; top:144; width:103; height:0'></input>

    <label id=Label13 value='Internal IC\015\012Number' style='position:absolute; left:6; top:138; width:97; height:43; visibility:'>Internal IC\015\012Number</label>
    <input type='text' id=CPANumber value='empty' style='position:absolute; left:156; top:108; width:100; height:0'></input>

    <label id=Label8 value='CPA Reference' style='position:absolute; left:6; top:108; width:139; height:24; visibility:'>CPA Reference</label>
    <input type='text' id=Protocol Number & Title value='empty' style='position:absolute; left:804; top:36; width:280; height:162'></input>

    <label id=Text27 value='Protocol / Title:' style='position:absolute; left:864; top:6; width:144; height:24; visibility:'>Protocol / Title:</label>
    <input type='text' id=DateReceived value='empty' style='position:absolute; left:508; top:72; width:108; height:0'></input>

    <label id=Label2 value='Date Received:' style='position:absolute; left:361; top:72; width:142; height:24; visibility:'>Date Received:</label>
    <input type='text' id=DateLastEdited value='empty' style='position:absolute; left:508; top:108; width:108; height:0'></input>

    <label id=Label4 value='Date Last Edited:' style='position:absolute; left:346; top:108; width:157; height:24; visibility:'>Date Last Edited:</label>
    <input type='text' id=LastEditedByInit value='empty' style='position:absolute; left:508; top:142; width:108; height:0'></input>

    <label id=Label5 value='Edited by (Initials):' style='position:absolute; left:331; top:142; width:172; height:24; visibility:'>Edited by (Initials):</label>
    <label id=IClabel value='Informed Consent Checklist' style='position:absolute; left:359; top:6; width:391; height:40; visibility:'>Informed Consent Checklist</label>
    <select name='LetterDateCombo' Size='8' style='position:absolute; left:495; top:177; width:124; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT DISTINCTROW [CheckMaxICLetter].[irb#], [CheckMaxICLetter].[MaxOfDateSent], [CheckMaxICLetter].[LetterName] FROM CheckMaxICLetter WHERE ((([CheckMaxICLetter].[irb#])=[Forms]![NewInformedConsent]![ID#]) And (([CheckMaxICLetter].[LetterName])=\"Informed Consent Deficiencies\"))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['irb#']."'".">".$row['irb#']."</Option>";
        }
    ?>
    </select>

    <label id=Last Consent Letter sent on:_Label value='Last Consent Letter Sent On:' style='position:absolute; left:234; top:174; width:256; height:24; visibility:'>Last Consent Letter Sent On:</label>
    <input type='text' id=ApprovalDate value='empty' style='position:absolute; left:651; top:108; width:136; height:25'></input>

    <label id=ConsentApproval_Label value='       IRB \015\012Approval Date' style='position:absolute; left:651; top:60; width:136; height:40; visibility:'>       IRB \015\012Approval Date</label>
    <!--     <input type='text' id=TextParagraphNumber value='empty' style='position:absolute; left:648; top:144; width:258; height:30'></input> -->
    <!--     <label id=ltextlettername value='empty' style='visibility:hidden'></label> -->
    <!--     <input type='text' id=Text45 value='empty' style='position:absolute; left:634; top:132; width:102; height:42'></input> -->

    <label id=Label46 value='Text45:' style='position:absolute; left:490; top:132; width:61; height:24; visibility:'>Text45:</label>
    <!--     <input type='text' id=IRB Code value='empty' style='position:absolute; left:54; top:58; width:304; height:45'></input> -->

    <label id=Label453 value='IRB' style='position:absolute; left:6; top:58; width:40; height:24; visibility:'>IRB</label>
    <label id=Label39 value='Review the Informed Consent and check all that apply' style='position:absolute; left:240; top:208; width:561; height:28; visibility:'>Review the Informed Consent and check all that apply</label>





    <input type='button' id=SendCorresBut value='Send &Letter' style='position:absolute; left:684; top:0; width:204; height:0' onclick='SendCorresBut_Click();'></input>
    <input type='button' id=ReturnBut value='&Return to Study' style='position:absolute; left:888; top:0; width:204; height:0' onclick='ReturnBut_Click();'></input>
    <!--     <input type='text' id=Requirement value='empty' style='position:absolute; left:54; top:12; width:147; height:21'></input> -->

    <label id=Label70 value='Text69:' style='position:absolute; left:0; top:12; width:61; height:24; visibility:'>Text69:</label>
    <!--     <input type='text' id=Comment value='empty' style='position:absolute; left:349; top:15; width:142; height:30'></input> -->

    <label id=Label72 value='Text71:' style='position:absolute; left:205; top:15; width:61; height:24; visibility:'>Text71:</label>
  </body>
</html>
