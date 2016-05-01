<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function bFindProtocols_Click(){
//Call //FindProtocols;
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSAEsForDrugSearch";
}
function FindProtocols(){
//DoCmd.SetWarnings //False;
//strQuote = //Chr$(34);
//'******************** Code Start ************************
//could initialize frm here
//could initialize varItem here
var strSQL = '';
;
//HoldDrugsForPrinting = "";
//DoCmd.RunSQL "delete * from TempDrugSearch";
//DoCmd.RunSQL "delete * from TempDistinctStudies";
  ;
    //Set //ctl = //Forms!DrugSearch!lbDrugDescription;
   ;
    //Me.TBNumberOfItemsSelected = //ctl.ItemsSelected.Count;
    //'If Me.TBNumberOfItemsSelected = 0 Then
    ;
        //For //Each varItem //In //ctl.ItemsSelected;
            //Me.TBHoldDrug = //ctl.ItemData(varItem);
            //HoldDrugsForPrinting = //HoldDrugsForPrinting //& //Me.TBHoldDrug //& " AND ";
            //DoCmd.OpenQuery "qryTempDrugSearch";
        //Next varItem;
    //'Else
        ;
    //'End If
;
//'If HoldDrugsForPrinting = "" Then Exit Function
//If //ctl.ItemsSelected.Count = //0 //Then;
    //HoldDrugsForPrinting = "No Drugs/Devices Selected";
//Else;
    //HoldDrugsForPrinting = //Left$(HoldDrugsForPrinting, //Len(HoldDrugsForPrinting) //- //5);
//End //If;
;
}
function cmdClearDate_Click(){
//Me.txMeetingDateQry = "***";
//DoCmd.SetWarnings //False;
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSAEsForDrugSearch_dummy";
;
//DoCmd.RunSQL "delete * from TempDrugSearch";
//Call //lbDrugDescription_AfterUpdate;
;
}
function cmdSpecificMeeting_Click(){
//could initialize frm here
//could initialize varItem here
var strSQL = '';
//could initialize currentrow here
;
;
    ;
//If //IsNull(Me.txMeetingDateQry) //Or //Me.txMeetingDateQry = "" //Or //Me.txMeetingDateQry = "***" //Then;
    //MsgBox "You must enter a Meeting Date";
    //Exit //Sub;
//End //If;
;
;
//Set //ctl = //Forms!DrugSearch!lbDrugDescription;
;
//For //intCurrentRow = //0 //To //ctl.ListCount //- //1;
        //If //ctl.Selected(intCurrentRow) //Then;
            //ctl.Selected(intCurrentRow) = //False;
                ;
        //End //If;
//Next //intCurrentRow;
;
;
//Me.Frame39 = //3;
//DoCmd.SetWarnings //False;
   ;
//Me.TBActive = "***";
//DoCmd.RunSQL "delete * from TempDrugSearch";
//DoCmd.RunSQL "delete * from TempDistinctStudies";
//Call //FindProtocols;
//DoCmd.OpenQuery "qryTempDrugSearch_List_All";
//DoCmd.OpenQuery "qryDistinctTempDrugSearch_ListAll";
;
//'DoCmd.OpenQuery "qryFirstForDrugSearchListAll"
//'DoCmd.OpenQuery "qryTempDrugSearch_List_All"
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSaesfordrugsearch_dummy";
//'DoCmd.OpenQuery "qryDistinctTempDrugSearch"
;
//Me.TBHoldDrugsForPrinting1 = "All SAEs ALL Drugs/Devices " //& " For Meeting Date--" //& " " //& //Me.txMeetingDateQry;
;
;
;
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSAEsForDrugSearch";
;
//'DoCmd.RunSQL "delete * from TempDrugSearch"
//DoCmd.SetWarnings //True;
}
function Form_Close(){
//DoCmd.SetWarnings //True;
}
function Form_Open(Cancel){
//DoCmd.SetWarnings //False;
//HoldDrugsForPrinting = "";
//DoCmd.RunSQL "delete * from TempDrugSearch";
//DoCmd.RunSQL "delete * from TempDistinctStudies";
//DoCmd.SetWarnings //True;
;
;
//Me.txMeetingDateQry = "";
//Me.TBActive = "Open";
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSaesfordrugsearch_dummy";
;
//HoldDrugsForPrinting = "";
//HoldDrugsForPrinting1 = "OPEN Studies utilizing...";
//Me.TBHoldDrugsForPrinting1 = " ";
//Me.Frame39 = //1;
//Me.txMeetingDateQry = "***";
//Me.txMeetingDateQry.SetFocus;
;
}
function Frame39_AfterUpdate(){
//DoCmd.SetWarnings //False;
//If //Me.Frame39 = //1 //Then;
    //Me.TBActive = "OPEN";
    //HoldDrugsForPrinting1 = "OPEN Studies utilizing...";
//End //If;
//If //Me.Frame39 = //2 //Then;
    //Me.TBActive = "Closed";
    //HoldDrugsForPrinting1 = "Closed Studies utilizing...";
//End //If;
//If //Me.Frame39 = //3 //Then;
    //Me.TBActive = "***";
    //HoldDrugsForPrinting1 = "ALL Studies utilizing...";
//End //If;
;
//Call //FindProtocols;
//Me.TBHoldDrugsForPrinting1 = //HoldDrugsForPrinting1 //& //HoldDrugsForPrinting;
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSaesfordrugsearch_dummy";
 //DoCmd.OpenQuery "qryDistinctTempDrugSearch";
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSAEsForDrugSearch";
;
;
}
function lbDrugDescription_AfterUpdate(){
//DoCmd.SetWarnings //True;
//Call //FindProtocols;
//If //Me.Frame39 = //1 //Then;
    //Me.TBActive = "OPEN";
    //HoldDrugsForPrinting1 = "OPEN Studies utilizing...";
//End //If;
//If //Me.Frame39 = //2 //Then;
    //Me.TBActive = "Closed";
    //HoldDrugsForPrinting1 = "Closed Studies utilizing...";
//End //If;
//If //Me.Frame39 = //3 //Then;
    //Me.TBActive = "***";
    //HoldDrugsForPrinting1 = "ALL Studies utilizing...";
//End //If;
;
;
//Me.TBHoldDrugsForPrinting1 = //HoldDrugsForPrinting1 //& //HoldDrugsForPrinting;
//Me.TBHoldDrugsForPrinting1 = //Me.TBHoldDrugsForPrinting1 //& " For Meeting Date--" //& " " //& //Me.txMeetingDateQry;
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSaesfordrugsearch_dummy";
//DoCmd.OpenQuery "qryDistinctTempDrugSearch";
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSAEsForDrugSearch";
//'DoCmd.RunSQL "delete * from TempDrugSearch"
;
}
function Command49_Click(){
//On //Error //GoTo //Err_Command49_Click;
;
    var stDialStr = '';
    //could initialize PrevCtl here
    //Const //ERR_OBJNOTEXIST = //2467;
    //Const //ERR_OBJNOTSET = //91;
;
    //Set PrevCtl;
    ;
    //If //TypeOf PrevCtl //Is //TextBox //Then;
      stDialStr;
    //ElseIf //TypeOf PrevCtl //Is //ListBox //Then;
      stDialStr;
    //ElseIf //TypeOf PrevCtl //Is //ComboBox //Then;
      stDialStr;
    //Else;
      stDialStr;
    //End //If;
    ;
    //Application.Run "utility.wlib_AutoDial"//, stDialStr;
;
//Exit_Command49_Click:;
    //Exit //Sub;
;
//Err_Command49_Click:;
    //If //(Err = //ERR_OBJNOTEXIST) //Or //(Err = //ERR_OBJNOTSET) //Then;
      //Resume //Next;
    //End //If;
    //MsgBox //Err.Description;
    //Resume //Exit_Command49_Click;
    ;
}
function Preview_Click(){
//could initialize MyReport here
;
;
//MyFilter = //[_SelectSAEsForDrugSearch //subform].Form.OrderBy;
//'MsgBox MyFilter
//DoCmd.OpenReport "_SelectSAEsForDrugSearch"//, //acViewPreview;
//Set MyReport;
//MyReport.OrderBy = //MyFilter;
        //'Reports![SearchResults_Summary].OrderBy = "[expr1]"
 //'   Else
   //'     MyReport.OrderBy = Me.OrderBy
        //'Reports![SearchResults_Summary].OrderBy = Me.OrderBy
  //'  End If
//MyReport.OrderByOn = //True;
}
function Return_to_Menu_Click(){
//Forms![DrugSearch]![_SelectSAEsForDrugSearch //subform].Form.RecordSource = "_SelectSaesfordrugsearch_dummy";
//DoCmd.Close;
}
function Command57_Click(){
//On //Error //GoTo //Err_Command57_Click;
;
;
    //Screen.PreviousControl.SetFocus;
    //DoCmd.FindNext;
;
//Exit_Command57_Click:;
    //Exit //Sub;
;
//Err_Command57_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command57_Click;
    ;
}
function txMeetingDateQry_AfterUpdate(){
//If //IsNull(Me.txMeetingDateQry) //Or //Me.txMeetingDateQry = "" //Or //Me.txMeetingDateQry = " " //Then;
    //Me.txMeetingDateQry = "***";
//End //If;
    ;
//Call //lbDrugDescription_AfterUpdate;
}

    </script>
  </head>
  <body onLoad="Form_Open();" onUnload="Form_Close();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'></input>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>




    <select name='lbDrugDescription' style='position:absolute; left:153; top:81; width:450; height:126'>    </select>

    <label id=Label0 value='Select Drug(s)/Device' style='position:absolute; left:297; top:48; width:225; height:28; visibility:'>Select Drug(s)/Device</label>

    <!--     <input type='text' id=TBNumberOfItemsSelected value='empty' style='position:absolute; left:138; top:96; width:114; height:42'></input> -->

    <label id=Label28 value='Text27:' style='position:absolute; left:36; top:96; width:64; height:24; visibility:'>Text27:</label>
    <!--     <input type='text' id=TBHoldDrug value='empty' style='position:absolute; left:168; top:138; width:132; height:30'></input> -->

    <label id=Label30 value='Text29:' style='position:absolute; left:72; top:138; width:64; height:24; visibility:'>Text29:</label>
    <label id=Label32 value='Display Studies and their SAEs (if any) utilizing selected Drugs/Device' style='position:absolute; left:148; top:3; width:996; height:37; visibility:'>Display Studies and their SAEs (if any) utilizing selected Drugs/Device</label>
    <!--     <input type='text' id=TBActive value='empty' style='position:absolute; left:1230; top:36; width:96; height:0'></input> -->

    <label id=Label36 value='Text35:' style='position:absolute; left:1122; top:36; width:64; height:24; visibility:'>Text35:</label>


    <label id=Label40 value='     Which SAEs to Show  ' style='position:absolute; left:697; top:66; width:235; height:24; visibility:'>     Which SAEs to Show  </label>
    <input type='radio' style='position:absolute; left:714; top:103; width:0; height:0'></input>

    <label id=Label43 value='For OPEN Studies Only' style='position:absolute; left:737; top:100; width:186; height:24; visibility:'>For OPEN Studies Only</label>
    <input type='radio' style='position:absolute; left:714; top:136; width:0; height:0'></input>

    <label id=Label45 value='For CLOSED Studies Only' style='position:absolute; left:737; top:133; width:204; height:24; visibility:'>For CLOSED Studies Only</label>
    <input type='radio' style='position:absolute; left:714; top:169; width:0; height:0'></input>

    <label id=Label47 value='For ALL Studies' style='position:absolute; left:737; top:166; width:133; height:24; visibility:'>For ALL Studies</label>
    <input type='text' id=TBHoldDrugsForPrinting1 value='empty' style='position:absolute; left:259; top:235; width:1132; height:34'></input>
    <label id=Label53 value='Search Criteria = ' style='position:absolute; left:33; top:235; width:223; height:34; visibility:'>Search Criteria = </label>


    <input type='button' id=Preview value='Print Preview' style='position:absolute; left:1144; top:21; width:190; height:48' onclick='Preview_Click();'></input>
    <input type='button' id=Return to Menu value='&Return' style='position:absolute; left:1333; top:21; width:135; height:48' onclick='Return to Menu_Click();'></input>
    <input type='button' id=cmdSpecificMeeting value='List All SAEs forAll Studies on a Specific Meeting' style='position:absolute; left:24; top:19; width:189; height:72' onclick='cmdSpecificMeeting_Click();'></input>
    <input type='text' id=txMeetingDateQry value='empty' style='position:absolute; left:426; top:33; width:126; height:0'></input>

    <label id=Label56 value='Enter Meeting Date' style='position:absolute; left:240; top:33; width:174; height:24; visibility:'>Enter Meeting Date</label>
    <input type='button' id=cmdClearDate value='Clear Date' style='position:absolute; left:570; top:19; width:168; height:48' onclick='cmdClearDate_Click();'></input>
    <label id=Label58 value='(dd/mm/yy)' style='position:absolute; left:441; top:12; width:91; height:24; visibility:'>(dd/mm/yy)</label>
  </body>
</html>
