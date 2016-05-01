<?php
	$sstrTableName = $_POST['strTableName'];

  include("common-data.php");
  $ssqlGetRefNums = "SELECT uniquekey, ".$sstrTableName.".RefNum FROM ".$sstrTableName." INNER JOIN CIRB_CHOProtocol ON ".$sstrTableName.".RefNum = CIRB_CHOProtocol.RefNum";//join is to make sure we only fix ones that are actually in the protocol table
  if(!($sdbhGetRefNums = sqlsrv_connect($DBserver,$DBconnectionInfo))){
//      include_once("tkrTriggerTracker.php");
//      $stkrTriggerTracker = new tkrTriggerTracker;
//      $stkrTriggerTracker->RecordTrigger($sstrCustNum, $sstrLogonName, "BosInitdlgAssnRevrAvail.php", "Error", "ErrorError:sdbhGetRefNums couldn't connect to SQL Server");
//      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to SQL Server.";
  }
  if(!($sresGetRefNums=sqlsrv_query($sdbhGetRefNums, $ssqlGetRefNums))){
//      include_once("tkrTriggerTracker.php");
//      $stkrTriggerTracker = new tkrTriggerTracker;
//      $stkrTriggerTracker->RecordTrigger($sstrCustNum,$sstrLogonName,"BosInitdlgAssnRevrAvail.php","Error","ErrorError:ssqlGetRefNums is bad",$ssqlGetRefNums,substr($ssqlGetRefNums,50),substr($ssqlGetRefNums,100),substr($ssqlGetRefNums,150),substr($ssqlGetRefNums,200));
//      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to admin information at this time.";
  }
  while($srecGetRefNums=sqlsrv_fetch_array($sresGetRefNums)){
//    echo "<span onclick=GetMemFields(".$srecGetRefNums['RefNum'].")>".$srecGetRefNums['RefNum']."</span><br>";
    echo " <span id='".$srecGetRefNums['uniquekey']."' style='color:blue'>";
    echo $srecGetRefNums['uniquekey']." ";
    echo $srecGetRefNums['RefNum'];
//    echo " <input type=button value='all sync' onclick=GetMemFields('".$srecGetRefNums['RefNum']."',false,false)> ";
//    echo " <input type=button value='all async' onclick=GetMemFields('".$srecGetRefNums['RefNum']."',true,false)> ";
//    echo " <input type=button value='1 at a time' onclick=GetMemFields('".$srecGetRefNums['RefNum']."',true,true)> ";

//    echo " <input type=button value='Run' onclick=GetMemFields('".$srecGetRefNums['RefNum']."')> ";
    echo " <input type=button value='Run' onclick=GetMemFields('".$srecGetRefNums['uniquekey']."')> ";
    echo "</span>";

    echo " <br>";
  }


?>