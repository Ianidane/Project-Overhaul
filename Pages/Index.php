<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html>
  <head>
  <body>
    <label id='empty' value='empty' style='visibility:'></label>

    <input type='radio' style='position:absolute; left:0; top:0; width:0; height:0'>
    <input type='checkbox' id='empty' value='False' style='position:absolute; left:0; top:0; width:0; height:0'>


    <input type='text' id='empty' value='empty' style='position:absolute; left:0; top:0; width:0; height:0'>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>
    <select name='empty' style='position:absolute; left:0; top:0; width:0; height:0'>    </select>


    <button name='empty' type='submit' style='position:absolute; left:0; top:0; width:0; height:0'></button>




    <input type='text' id=Password value='empty' style='position:absolute; left:534; top:216; width:132; height:28'>

    <label id=Password Label value='Password' style='position:absolute; left:387; top:210; width:136; height:34; visibility:'>Password</label>

    <label id=Label5 value='Security Level \015\012Access Control' style='position:absolute; left:420; top:24; width:255; height:127; visibility:'>Security Level \015\012Access Control</label>
    <input type='button' id=Login value='&Login' style='position:absolute; left:526; top:373; width:108; height:42' onclick='Login_Click();'>
    <select name='CurrentUserName' style='position:absolute; left:534; top:156; width:207; height:31'>    </select>

    <label id=CurrentUserName Label value='User Name' style='position:absolute; left:366; top:156; width:148; height:34; visibility:'>User Name</label>
    <input type='button' id=CancelButton value='&Exit' style='position:absolute; left:633; top:373; width:108; height:42' onclick='CancelButton_Click();'>
    <label id=Label173 value='by ProIRB Plus, Inc.' style='position:absolute; left:51; top:39; width:160; height:25; visibility:'>by ProIRB Plus, Inc.</label>
    <!--     <input type='text' id=NewPwd value='empty' style='position:absolute; left:534; top:264; width:132; height:28'></input> -->

    <!--     <label id=NewPwdLabel value='Enter New Password' style='position:absolute; left:384; top:252; width:144; height:60; visibility:hidden'>Enter New Password</label> -->
    <!--     <input type='text' id=NewPwd2 value='empty' style='position:absolute; left:534; top:324; width:132; height:28'></input> -->

    <!--     <label id=NewPwd2Label value='Enter Again ' style='position:absolute; left:378; top:327; width:144; height:28; visibility:hidden'>Enter Again </label> -->
    <input type='button' id=NewPassword value='&Change Password' style='position:absolute; left:60; top:372; width:228; height:0' onclick='NewPassword_Click();'>
    <!--     <select name='Label24' Size='7' style='position:absolute; left:64; top:316; width:79; height:25'>
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

        $sql = sqlsrv_query($conn, "SELECT [IRB Meeting Schedule1].[Irb Code] FROM [IRB Meeting Schedule1] WHERE ((([IRB Meeting Schedule1].[Irb Code]) Like \"*Anthing I want*\"))");
        while( $row = sqlsrv_fetch_array($sql)){
        echo "<Option value='".$row['Irb Code']."'".">".$row['Irb Code']."</Option>";
        }
    ?>
    </select> -->
    <label id=Label108 value='ProIRB' style='position:absolute; left:0; top:0; width:130; height:39; visibility:'>ProIRB</label>
    <label id=Label48 value='R' style='position:absolute; left:121; top:6; width:18; height:16; visibility:'>R</label>
  </body>
</html>
