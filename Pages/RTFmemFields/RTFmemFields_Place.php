<?php
	$sstrTableName = $_POST['strTableName'];
	$snumUniquekey = $_POST['numUniquekey'];
  $sstrMemFieldName = $_POST['strMemFieldName'];
  $sstrValue = $_POST['strValue'];

  include("common-data.php");
  $ssqlGetMemFields = "UPDATE ".$sstrTableName." SET ".$sstrMemFieldName."='".str_replace("'","''",stripslashes($sstrValue))."' WHERE uniquekey = '".$snumUniquekey."'";
//echo $ssqlGetMemFields;
  if(!($sdbhGetMemFields = sqlsrv_connect($DBserver,$DBconnectionInfo))){
//      include_once("tkrTriggerTracker.php");
//      $stkrTriggerTracker = new tkrTriggerTracker;
//      $stkrTriggerTracker->RecordTrigger($sstrCustNum, $sstrLogonName, "BosInitdlgAssnRevrAvail.php", "Error", "ErrorError:sdbhGetMemFields couldn't connect to SQL Server");
//      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to SQL Server.";
  }
  if(!($sresGetMemFields=sqlsrv_query($sdbhGetMemFields, $ssqlGetMemFields))){
//      include_once("tkrTriggerTracker.php");
//      $stkrTriggerTracker = new tkrTriggerTracker;
//      $stkrTriggerTracker->RecordTrigger($sstrCustNum,$sstrLogonName,"BosInitdlgAssnRevrAvail.php","Error","ErrorError:ssqlGetMemFields is bad",$ssqlGetMemFields,substr($ssqlGetMemFields,50),substr($ssqlGetMemFields,100),substr($ssqlGetMemFields,150),substr($ssqlGetMemFields,200));
//      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to admin information at this time.";
  }

//  $srecGetMemFields=sqlsrv_fetch_array($sresGetMemFields);
echo $ssqlGetMemFields;

?>