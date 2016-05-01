<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<?php include 'Serverinfo.php'; ?>
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Form_Load(){
var testlocation = '';
//could initialize dbs here
//could initialize doc here
//could initialize prp here
//could initialize rst here
//'MsgBox HoldListIndex
//Set dbs;
//BackendLocation = //GetBackendLocation();
    //' tries to open the only linked table "Demo Users" have permission on
    //'to determine if the link was broken.
    //On //Error //Resume //Next;
    //Set rst;
    //If //Err = //0 //And //Not //BackendLocation = "startup" //Then  //'Link is OK
        //rst.Close;
        //If //VersionCheck //Then;
            //DoCmd.Close //acForm, "refresh links"//, //acSaveYes;
            location.href = "Form_FlashScreen.php";;
            //Exit //Sub;
        //Else;
            //MsgBox "Quitting PRO_IRB";
            //dbs.Close;
            //Application.Quit;
        //End //If;
    //Else    //'Link has been broken now goto proirb_backup.mdb and get location of proirb_be
        //If //BackendLocation = "startup" //Then;
            //'BackendLocation = "C:"
//AskDrive:;
            //BackendLocation = //InputBox("Enter //the //Drive //letter //where //the //Folder //PRO_IRB //is //assigned //to:");
            //If //BackendLocation = "" //Then;
                //If //YesNo("Not //providing //a //Drive //letter //will //close //ProIRB. //Is //this //what //you //want?") //Then;
                    //Application.Quit;
                //Else;
                        //GoTo //AskDrive;
                //End //If;
             //Else;
                //If //Len(BackendLocation) //> //1 //Then;
                    //MsgBox "Please enter only the 1 character Drive letter";
                    //GoTo //AskDrive;
                //Else;
                    //BackendLocation = //BackendLocation //& ":\PRO_IRB";
                //End //If;
            //End //If;
            ;
        //End //If;
            //' Try to relink the tables; if it fails, shut down.
        //If //RelinkTables(BackendLocation) = //False //Then;
           //MsgBox //("Relinking //Failed //pleased //call //your //System //Admministrator"), //vbCritical;
           //CloseCurrentDatabase;
           //Exit //Sub;
        //Else;
            ;
            //MsgBox //("Sucessfully //established //new //links //to:  " & BackendLocation & "//\" //& "ProIRB_be.mdb" //& "  CLick OK then restart PRO_IRB"//), //vbCritical;
            //Call //UpdateDatabaseLocation;
           //' msgbox BackendLocation
            //Application.Quit;
        //End //If;
    //End //If;
    ;
//'err_find:
 //'  msgbox "PRO_IRB error in Refresh Links--Write down message and call PRO_IRB Technical Support" & Err.Description & "  " & Err.Number
  //'resuem exit_refreshlinks
}
function ReferenceCheck(){
//could initialize ref here
;
          //' Enumerate through References collection.
          //For //Each ref //In //References;
          //' Check IsBroken property.
    //On //Error //GoTo //ErrorHandler:;
             //If //ref.IsBroken = //True //Then;
             //' Print the name of the reference and the fact that it is
             //' broken.
                 //MsgBox "  " //& //ref.Name //& " is broken";
             //End //If;
          //Next ref;
;
        //Exit //Sub;
;
//ErrorHandler:;
;
        //' If the error received is 48 it's an error from the broken
        //' reference. Print the reference the fact it's broken and path to
        //' the Debug window.
        //If //Err.Number = //48 //Then;
           //MsgBox "Reference broken for: " //& //ref.FullPath //& " ---Call PRO_IRB Support";
        //Else;
            //MsgBox //Err.Description //& "   " //& //Err.Number;
        //End //If;
    //Application.Quit;
}
function Form_Open(Cancel){
//Call //ReferenceCheck;
}
function VersionCheck(){//As //Boolean;
//On //Error //GoTo //err_VersionCheck;
    //could initialize db here
    //could initialize BAckendVersion here
    //could initialize rs here
    //could initialize tdf here
    //Set db;
    //Set rs;
    BAckendVersion;
    ;
    //If //db.Properties("AppTitle") //Like "*" //& BAckendVersion //& "*" //Then;
        //VersionCheck = //True;
        //rs.Close;
        //Set rs;
        //Exit //Function;
    //Else;
        //MsgBox "Versions = Back-end Version = " //& //BackendLocation //& "  " //& BAckendVersion //_;
        //& "  " //& "Front-End version = " //& //db.Properties("AppTitle") //& "  Found and Don't match.  Make certain you have downloaded the latest version and your mapping is correct.";
        //VersionCheck = //False;
    //End //If;
    ;
    ;
//exit_versioncheck:;
    //Exit //Function;
//err_VersionCheck:;
    //MsgBox "Error trying to determine Version Numbers--Call PRO_IRB";
    //MsgBox "Versions = Back-end Version = " //& //BackendLocation //& "  " //& BAckendVersion //_;
        //& "  " //& "Front-End version = " //& //db.Properties("AppTitle") //& " Error and can'r match. Make certain you have downloaded the latest version and your mapping is correct.";
        ;
    //Application.Quit;
}
function UpdateDatabaseLocation(){
//On //Error //GoTo //Error_Update;
//strQuote = //Chr$(34);
;
;
//Me.THoldBackendLocation = //BackendLocation;
//DBEngine.SystemDB = "c:\cc\proirb_dev\security.mdw";
    ;
    //Set //wspace = //CreateWorkspace("", "MASTER"//, "3837"//);
    //could initialize db here
    var SqlStatement = '';
    //could initialize rs here
    //could initialize tdf here
    //Set db;
     //Set rs;
    //'Set rs = db.OpenRecordset("Select DatabaseLocation from NetworkLocation")
    //rs.Edit;
    //rs("databaselocation") = //BackendLocation;
    //rs("backupfolderpath") = //BackendLocation;
    //rs.Update;
    //rs.Close;
    //Set rs;
    //db.Close;
    //Set db;
    //Exit //Sub;
//Error_Update:;
    //MsgBox //Err.Description;
    //Application.Quit;
}
function Command4_Click(){
//On //Error //GoTo //Err_Command4_Click;
;
;
    //Screen.PreviousControl.SetFocus;
    //DoCmd.FindNext;
;
//Exit_Command4_Click:;
    //Exit //Sub;
;
//Err_Command4_Click:;
    //MsgBox //Err.Description;
    //Resume //Exit_Command4_Click;
    ;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>

    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'></input>


    <label id=Label0 value='Special Form for refreshing links' style='position:absolute; left:60; top:48; width:582; height:180; visibility:'>Special Form for refreshing links</label>
    <input type='text' id=THoldBackendLocation value='empty' style='position:absolute; left:154; top:238; width:313; height:30'></input>

    <label id=Label2 value='Text1:' style='position:absolute; left:10; top:238; width:52; height:24; visibility:'>Text1:</label>
    <input type='button' id=Command4 value='Command4' style='position:absolute; left:504; top:96; width:57; height:57' onclick='Command4_Click();'></input>
  </body>
</html>
