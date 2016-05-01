<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 
<html>
  <head>
    <link href="images/NewPro.css" rel="stylesheet" type="text/css">
    <meta http-equiv="Content-Type" content="text/html;">
    <title>NewPro</title>
    <script type="text/JavaScript">
function Registration(){//As //Boolean;
//'executed in version update 5.00.0006 only
//On //Error //GoTo //err_Registration;
   ;
;
    ;
        ;
            //If //IsNull(DMax("[NumberOfRegisteredUsers]", "Copyright"//)) //Then;
                //DoCmd.SetWarnings //False;
                //DoCmd.OpenQuery "Reg_SetRegisteredUsersToZero";
                //RegisteredUsers = //0;
            //Else;
                //RegisteredUsers = //DMax("[NumberOfRegisteredUsers]", "Copyright"//);
            //End //If;
            ;
            //If //RegisteredUsers //< //MaxUsersAllowed //Then //' proceed with registration
                //If //dir("C:\CC\PROIRB_dev\tempregdrf3jg.dll") = "" //Then;
                    //MsgBox "Can't find Registration File- Write down this error message and contact ProIRB at 727 490-4251 ";
                    //Application.Quit;
                //Else;
                    //Name "C:\CC\PROIRB_dev\tempregdrf3jg.dll" //As "C:\regdrf3jg.dll";
                //End //If;
               //DoCmd.SetWarnings //False;
               //DoCmd.OpenQuery "Reg_AddOneToNumberOfRegisteredUsers";
               //'check for good update
               //NewRegisteredUsers = //DMax("[NumberOfRegisteredUsers]", "Copyright"//);
               //If //Not //NewRegisteredUsers = //RegisteredUsers //+ //1 //Then //'bad update
                    //If //dir("C:\regdrf3jg.dll") = "" //Then;
                        //GoTo //err_Registration   //'renaming failed also
                    //Else;
                            //Kill //("C:\regdrf3jg.dll");
                            //GoTo //err_Registration;
                    //End //If;
               //End //If;
            //Else;
               //Registration = //False  //'too many registered users
               //MsgBox //("Install //aborted! //You //already //have " & RegisteredUsers & " //users //installed " _
               & "//which //is //the //number //of //users //licenses //you //have //purchased. //If //you //wish //to //purchase " _
               & "//additional //licenses //please //contact //ProIRB //at //727 //490-4251 ")
              '
              Kill ("//C:\CC\ProIRB_Dev\tempregdrf3jg.dll");
               //Application.Quit;
            //End //If;
        ;
            //Registration = //True;
            //DoCmd.SetWarnings //True;
            //MsgBox "Registration Successful for --" //& //NewRegisteredUsers //& " Users";
//exit_Registration:;
//DoCmd.SetWarnings //True;
    //Exit //Function;
//err_Registration:;
        //MsgBox "Error in Registration Process-write down this error message and contact ProIRB 727-490-4251 " //& //Err.Description;
      ;
        //Application.Quit;
   ;
//Exit //Function;
}
function CompareMaxusersToNumberofUsers(){
//MaxUsersAllowed = //DMax("[maxusers]", "copyright"//);
    //CountOfUserNames = //DCount("[ID]", "users"//);
    ;
    //If //MaxUsersAllowed = //CountOfUserNames //Then;
        //Exit //Sub;
    //Else;
        //MsgBox //("Please //write //down //the //following //data //and //call //ProIRB //technical //support //727-490-4251.      " _
        & "//MaxUsers = " & MaxUsersAllowed & "     //User //Count= " & CountOfUserNames & "   //Registration //cannot //proceed.");
        //Application.Quit;
    //End //If;
    ;
}
function CancelButton_Click(){
//'If Not Switch Then
;
//If //LoginLogoutSwitch //Then;
    ;
    //If //YesNo("Do //you //want //to //exit //ProIRB?") //Then;
        //GoTo //exit_app;
    //Else;
        //MsgBox "You are still Logged in as  " //& //UserLastName, //vbInformation;
        //LoginLogoutSwitch = //False;
        //DoCmd.Close;
        //Exit //Sub;
    //End //If;
//End //If;
//exit_app:;
//Application.Quit;
    ;
}
function Form_Load(){
//On //Error //GoTo //err_form_users_load;
//'MsgBox SysCmd(acSysCmdAccessVer)
//strQuote = //Chr$(34);
//'commented out for open version
//'DoCmd.Close acForm, "FlashScreen"
//If //CurrentUser = "master" //Then;
    //'msgbox ("Remember to put this false statement back in before final complile and distribution")
    //could initialize PROIRB here
    //Set PROIRB;
    //If //PROIRBSetAccessProperty(PROIRB, "AllowBypassKey"//, //dbBoolean, //False) //Then;
    //End //If;
//End //If;
//Exit //Sub;
;
//'obj =currentdbs, strname = the name of the property in quotes,
        //'inttype = data type constant see createproperty help, varsetting = the setting you want
//err_form_users_load:;
//MsgBox "PRO_IRB error in User From Load-Write dowm error message and call PRO_IRB Tecnical Service-" //& //Err.desc //& "  " //& //Err.Number;
//Application.Quit;
}
function Form_Open(Cancel){
//DefaultValue = //4;
//Me.Password = " ";
//DoCmd.Close //acForm, "FlashScreen";
//DoCmd.Close //acForm, "frm_display_generated_number";
}
function Login_Click(){
//On //Error //GoTo //error_log;
  ;
//Normal_Login_Starts_Here:;
;
//If //Me.CurrentUserName = "Admin-Trial" //Then;
    //GoTo //CheckAllow;
//Else;
    //Call //CompareMaxusersToNumberofUsers;
    ;
    //If //dir("C:\CC\ProIRB_Dev\tempregdrf3jg.dll") = "" //Then         //'if this file is present then this is the first time run since the latest install
        //If //dir("C:\regdrf3jg.dll") = "" //Then    //'C:\regdrf3jg.dll must be present otherwise this is illegal
            //MsgBox "Error-ProIRB appears to be unlicensed on this machine.  Please call ProIRB Tecnical Support 727-490-4251.";
            //Application.Quit;
        //Else;
            //GoTo //CheckAllow //' Registered copy program can proceed
        //End //If;
    //Else;
            //'not first run after install
        //If //dir("C:\regdrf3jg.dll") = "" //Then;
            //If //Registration() //Then //GoTo //CheckAllow //'Now register 'this should be the normal first run after install
        //Else;
            //Kill //("C:\regdrf3jg.dll")   //'was once registered on this machine caN REINSTALL
            //DoCmd.SetWarnings //False;
            //DoCmd.OpenQuery "Reg_SubTractOneFromRegisteredUsers" //'subtract 1 FROM REGISTERED USERS
            //If //Registration() //Then //GoTo //CheckAllow //'Now register 'this should be the normal first run after install
        //End //If;
        ;
    //End //If;
//End //If;
        ;
        ;
//CheckAllow:;
//If //Me.Password = "allow" //Then;
    //GoTo //AllowBypass;
//Else;
    //If //Me.Password = "close" //Then;
    //'sets the allowbypass to false each time program is started
        //could initialize PROIRBC here
        //could initialize wspace here
        //DBEngine.SystemDB = "C:\CC\PROIRB_Dev\security.mdw";
        //Set wspace;
    //'msgbox CurrentDb.Name
        //Set PROIRBC;
        //If //PROIRBSetAccessProperty(PROIRBC, "AllowBypassKey"//, //dbBoolean, //False) //Then;
        //End //If;
        //PROIRBC.Close;
        //wspace.Close;
        //MsgBox "Closed";
        //Application.Quit;
    //End //If;
;
    //SecurityDate = //Me.CurrentUserName.Column(4);
    //If //Me.CurrentUserName = "Admin-Trial" //Then;
        //If //SecurityDate //< //Date //Then;
            //'MsgBox "MSACCESS has encountered a memory error at page XoXoFFFF.  The application will close."
            //MsgBox "Trial version has expired. Contact service@Proirb.com";
            //Application.Quit;
        //End //If;
    //End //If;
    //If //DMax("[Field_AMANET_Flag]", "copyright"//) //Then;
        //If //SecurityDate //< //Date //Then;
            //MsgBox "MSACCESS has encountered a memory error at page XoXoFFFF.  The application will close.";
            //Application.Quit;
        //End //If;
    //End //If;
    ;
    ;
    ;
    //If //Me.NewPwd.Visible //Then;
            //If //Me.NewPwd = //Me.NewPwd2 //Then;
                //DoCmd.SetWarnings //False;
                //DoCmd.OpenQuery "UpdatePassword";
                //DoCmd.SetWarnings //True;
                //GoTo //OK;
            //Else;
                //MsgBox "The 2 new Passwords are not the same.  Reenter both";
                //Me.NewPwd = " ";
                //Me.NewPwd2 = " ";
                //Me.NewPwd.SetFocus;
                //Exit //Sub;
            //End //If;
    //End //If;
    //If //IsNull(Me.Password) //Or //Me.Password //<> //Me.CurrentUserName.Column(1) //Then;
            //GoTo //NotValid;
    //Else;
            //GoTo //OK;
    //End //If;
    ;
;
   ;
//NotValid:;
    //MsgBox "Invalid User or Password"//, //vbInformation;
    //Me.Password.SetFocus;
    //Me.Password = " ";
    //Exit //Sub;
;
//OK:;
    //UserEdit = //Me.CurrentUserName.Column(2);
    //If //Not //Me.CurrentUserName.Column(5) //Then;
        //SAEUser = //False;
          //Else;
        //SAEUser = //True;
    //End //If;
    ;
   //ExpertUser = //Me.CurrentUserName.Column(7);
   //UserLastName = //Me.CurrentUserName;
            ;
//CheckUserEdit:;
//If //isFormLoaded("NewMenu") //Then;
    //Forms![newmenu]![TCurrentUser] = //UserLastName;
    //If //UserEdit //Then;
        //Forms![newmenu]![HoldReadOnly] = //False;
    //Else;
        //Forms![newmenu]![HoldReadOnly] = //True;
    //End //If;
//End //If;
//If //Not //UserEdit //And //Not //SAEUser //Then //MsgBox "User has Read Only permissions";
//If //Not //UserEdit //And //SAEUser //Then //MsgBox "User has Read Only and SAE permissions";
        ;
    ;
//If //Me.CurrentUserName //Like "*Trial*" //Then;
        //Demo = //True;
//Else;
        //Demo = //False;
//End //If;
;
//Call //TextStreamTest  //'This will end the program if dYNO.ini DATE IS <TODAY
//'If RegisterAFile("C:\CC\PROIRB_DEV\MSADO15.DLL") Then
 //'   MsgBox "WAS ABLE TO REGISTER"
//'Else
 //'  MsgBox "Error but      WILL CONTINUE"
//'End If
//CyberIRB_Flag = //False;
//If //DMax("[reservetext3]", "copyright"//) = "CyberIRB" //Then //CyberIRB_Flag = //True;
;
//DoCmd.Close;
location.href = "Form_NewMenu.php";;
//If //CyberIRB_Flag //Then;
    //gclsCyberIRB.CyberProtocolsProcessing //' bridge
//End //If;
//Exit //Sub;
//End //If;
    ;
//AllowBypass:;
//'This procedure allows me to open the code when the user used the password of "allow"
//If //CurrentUser = "master" //Then;
    //could initialize PROIRB here
    //Set PROIRB;
    //If //PROIRBSetAccessProperty(PROIRB, "AllowBypassKey"//, //dbBoolean, //True) //Then;
    //End //If;
    //MsgBox "Bypass key enabled.  Close PRO_IRB and restart";
    //DoCmd.Close;
    //Application.Quit;
//End //If;
;
//exit_log:;
//Exit //Sub;
//error_log:;
//MsgBox //Err.Description;
//Resume //exit_log;
}
function NewPassword_Click(){
//If //Me.Password //<> //Me.CurrentUserName.Column(1) //Then;
        //GoTo //CantChange;
    //Else;
        //Me.NewPwdLabel.Visible = //True;
        //Me.NewPwd.Visible = //True;
        //Me.NewPwd2Label.Visible = //True;
        //Me.NewPwd2.Visible = //True;
        //MsgBox "Enter New Password Twice then Click Login";
        //Me.NewPwd.SetFocus;
        //Exit //Sub;
//End //If;
;
;
;
//CantChange:;
    //MsgBox "You must first enter the correct password";
    //Exit //Sub;
;
}
function Password_DblClick(Cancel){
//If //CurrentUser = "master" //Then;
//Me.Password = "3837";
//Login_Click;
//End //If;
}
function Password_GotFocus(){
//Me.Password = "";
;
}
function TextStreamTest(){
//On //Error //GoTo //err_read;
//'Exit Sub
//If //dir("C:\CC\PROIRB_dev\dyno.ini") //<> "" //Then;
  ;
    var DateInFile = '';
   ;
    //Open "C:\CC\PROIRB_dev\dyno.ini" //For //Input //As //#1    //' Open file for input.
    //Input //#1, DateInFile //' Read date
    ;
    //'MsgBox DateInFile
    //'MaxDate = DateAdd("d", 30, DateInFile)
    //'MsgBox MaxDate
    //If //DateValue(DateInFile) //< //Now //Then;
        //MsgBox "MSACCESS has encountered a memory error at page XoXoFFFF.  The application will close.";
        //Close //#1    //' Close file.
        //Application.Quit;
    //End //If;
    ;
//Else;
    //'MsgBox "No reg.ini was found"
    //Exit //Sub;
//End //If;
;
;
//Exit //Sub;
//err_read:;
    //MsgBox //Err.Description, //Err.Number;
    //Exit //Sub;
}

    </script>
  </head>
  <body onLoad="Form_Load();Form_Open();">


    <label id='empty' value='empty' style='visibility:'></label>




    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>




    <input type='text' id=Password value='empty' style='position:absolute; left:534; top:216; width:132; height:28'>

    <label id=Password Label value='Password' style='position:absolute; left:387; top:210; width:136; height:34; visibility:'>Password</label>

    <label id=Label5 value='Security Level \015\012Access Control' style='position:absolute; left:420; top:24; width:255; height:127; visibility:'>Security Level \015\012Access Control</label>
    <input type='button' id=Login value='&Login' style='position:absolute; left:526; top:373; width:108; height:42' onclick='Login_Click();'>
    <select name='CurrentUserName' style='position:absolute; left:534; top:156; width:207; height:31'>

    <label id=CurrentUserName Label value='User Name' style='position:absolute; left:366; top:156; width:148; height:34; visibility:'>User Name</label>
    <input type='button' id=CancelButton value='&Exit' style='position:absolute; left:633; top:373; width:108; height:42' onclick='CancelButton_Click();'>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:51; top:39; width:160; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <!--     <input type='text' id=NewPwd value='empty' style='position:absolute; left:534; top:264; width:132; height:28'> -->

    <!--     <label id=NewPwdLabel value='Enter New Password' style='position:absolute; left:384; top:252; width:144; height:60; visibility:hidden'>Enter New Password</label> -->
    <!--     <input type='text' id=NewPwd2 value='empty' style='position:absolute; left:534; top:324; width:132; height:28'> -->

    <!--     <label id=NewPwd2Label value='Enter Again ' style='position:absolute; left:378; top:327; width:144; height:28; visibility:hidden'>Enter Again </label> -->
    <input type='button' id=NewPassword value='&Change Password' style='position:absolute; left:60; top:372; width:228; height:0' onclick='NewPassword_Click();'>
    <!--     <select name='Label24' Size='7' style='position:absolute; left:64; top:316; width:79; height:25'> -->
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:39; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
  </body>
</html>
