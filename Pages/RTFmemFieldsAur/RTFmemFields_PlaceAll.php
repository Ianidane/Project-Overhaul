<?php

//140725 way after comment - the way this looks now is really bad so I commented out - its looking up values in CHOAmend and then replacing the same uniquekey in CHOInit with the values in CHOAmend
//if doing this then use one varaible for the table name 

////	$sstrTableName = $_POST['strTableName'];
////	$snumUniquekey = $_POST['numUniquekey'];
////  $sstrMemFieldName = $_POST['strMemFieldName'];
////  $sstrValue = $_POST['strValue'];
//
//  include("common-data.php");
//
////$snumUniquekey = 74;
////  $ssqlGetMemFields = "SELECT * FROM ChoInit WHERE uniquekey = ".$snumUniquekey;
//$snumUniquekey = 289;
//  $ssqlGetMemFields = "SELECT * FROM ChoAmend WHERE uniquekey = ".$snumUniquekey;
////echo $ssqlGetMemFields;
//  if(!($sdbhGetMemFields = sqlsrv_connect($DBserver,$DBconnectionInfo))){
////      include_once("tkrTriggerTracker.php");
////      $stkrTriggerTracker = new tkrTriggerTracker;
////      $stkrTriggerTracker->RecordTrigger($sstrCustNum, $sstrLogonName, "BosInitdlgAssnRevrAvail.php", "Error", "ErrorError:sdbhGetMemFields couldn't connect to SQL Server");
////      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to SQL Server.";
//  }
//  if(!($sresGetMemFields=sqlsrv_query($sdbhGetMemFields, $ssqlGetMemFields))){
////      include_once("tkrTriggerTracker.php");
////      $stkrTriggerTracker = new tkrTriggerTracker;
////      $stkrTriggerTracker->RecordTrigger($sstrCustNum,$sstrLogonName,"BosInitdlgAssnRevrAvail.php","Error","ErrorError:ssqlGetMemFields is bad",$ssqlGetMemFields,substr($ssqlGetMemFields,50),substr($ssqlGetMemFields,100),substr($ssqlGetMemFields,150),substr($ssqlGetMemFields,200));
////      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to admin information at this time.";
//  }
//	//count the no. of  columns in the table
//	$fcount = sqlsrv_num_fields($sresGetMemFields);
//	$tag = sqlsrv_field_metadata($sresGetMemFields);
//
//  while($srecGetMemFields=sqlsrv_fetch_array($sresGetMemFields)){
//		for($i=0; $i< $fcount; $i++){
//      if(substr($tag[$i]['Name'],0,3) == 'mem'){
//
//        echo "<span id='".$tag[$i]['Name']."'>";
//        $text=$srecGetMemFields[$i];
//        $text=str_replace("%u2013","-",$text);
//
////        $text=str_replace("%u","",$text);
////        $text=str_replace("%u","",$text);
////        $text=str_replace("%u","",$text);
//
//        $ssqlUpdate = "UPDATE CHOInit SET .$tag[$i]['Name'].='".$text."' WHERE uniquekey = '".$snumUniquekey."'";
//      //echo $ssqlUpdate;
//        if(!($sdbhUpdate = sqlsrv_connect($DBserver,$DBconnectionInfo))){
//      //      include_once("tkrTriggerTracker.php");
//      //      $stkrTriggerTracker = new tkrTriggerTracker;
//      //      $stkrTriggerTracker->RecordTrigger($sstrCustNum, $sstrLogonName, "BosInitdlgAssnRevrAvail.php", "Error", "ErrorError:sdbhUpdate couldn't connect to SQL Server");
//      //      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to SQL Server.";
//        }
//        if(!($sresUpdate=sqlsrv_query($sdbhUpdate, $ssqlUpdate))){
//      //      include_once("tkrTriggerTracker.php");
//      //      $stkrTriggerTracker = new tkrTriggerTracker;
//      //      $stkrTriggerTracker->RecordTrigger($sstrCustNum,$sstrLogonName,"BosInitdlgAssnRevrAvail.php","Error","ErrorError:ssqlUpdate is bad",$ssqlUpdate,substr($ssqlUpdate,50),substr($ssqlUpdate,100),substr($ssqlUpdate,150),substr($ssqlUpdate,200));
//      //      return "An error occured trying to get admin information.  Operation did not complete successfully.  Couldn't connect to admin information at this time.";
//        }
//
////disabled        $srecUpdate=sqlsrv_fetch_array($sresUpdate);
//      }
//		}
//  }





?>