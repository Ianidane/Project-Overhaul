<?php
	$sstrTableName = $_POST['strTableName'];
//	$sstrRefNum = $_POST['strRefNum'];
	$snumUniquekey = $_POST['numUniquekey'];
//  $sbol1AtATime = $_POST['bol1AtATime'];

  include("common-data.php");
  $ssqlGetMemFields = "SELECT * FROM ".$sstrTableName." WHERE uniquekey = '".$snumUniquekey."'";
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
	//count the no. of  columns in the table
	$fcount = sqlsrv_num_fields($sresGetMemFields);
	$tag = sqlsrv_field_metadata($sresGetMemFields);
  
  while($srecGetMemFields=sqlsrv_fetch_array($sresGetMemFields)){

////just memInclExcl
//        echo "<span id='memInclExcl'>";
//        echo $tag[$i]['Name']." ";
//        echo "<input type=button value='Investigate' onclick=Investigate('memInclExcl')> ";
//        echo "<input type=button value='Investigate no RTF' onclick=InvestigateNoRTF('memInclExcl')> ";
////        echo "<input type=button value='Save' onclick=Place('memInclExcl',escape(document.getElementById('valuememInclExcl').innerHTML))>";
//        echo "<input type=button value='Save' onclick=Place('memInclExcl',escape(tinymce.get('valuememInclExcl').getContent()))>";
//        echo "</span>";
//        echo "<br>";
//        $text=$srecGetMemFields[$i];
//        echo "<textarea rows=10 cols=75 id='valuememInclExcl'>".$text."</textarea><br>";




		for($i=0; $i< $fcount; $i++){
      if(substr($tag[$i]['Name'],0,3) == 'mem'){
        echo "<span id='".$tag[$i]['Name']."'>";
        echo $tag[$i]['Name']." ";
//140725        echo "<input type=button value='Investigate' onclick=Investigate('".$tag[$i]['Name']."')> ";
//140725        echo "<input type=button value='Investigate no RTF' onclick=InvestigateNoRTF('".$tag[$i]['Name']."')> ";
//        echo "<input type=button value='Save' onclick=Place('".$tag[$i]['Name']."',escape(document.getElementById('value".$tag[$i]['Name']."').innerHTML))>";
        echo "<input type=button value='Save' onclick=Place('".$tag[$i]['Name']."',escape(tinymce.get('value".$tag[$i]['Name']."').getContent()))>";
        echo "</span>";
        echo "<br>";

//        echo "<textarea rows=10 cols=75>".$srecGetMemFields[$i]."</textarea><br>";
        
    $text=$srecGetMemFields[$i];
    
    
////    $text=str_replace("\n\n",  "z000"        ,$text);
//    $text=str_replace("\n",    "z111"        ,$text);
//    $text=filter_var($text,FILTER_SANITIZE_SPECIAL_CHARS, FILTER_FLAG_STRIP_LOW);
////    $text=str_replace("z000",  "<br/><br/>"  ,$text);
//    $text=str_replace("z111",  "<br/>"       ,$text);
//
////    $text=decodeURIComponent($text);
////    $text=encodeURIComponent($text);


    $text=nl2br($text);



//        $text = htmlentities($text,ENT_SUBSTITUTE, "UTF-8");
//        $text = htmlentities($text,ENT_IGNORE, "UTF-8");
//        $text = htmlspecialchars($text,ENT_IGNORE, "UTF-8");
//        $text = htmlspecialchars($text,ENT_SUBSTITUTE, "UTF-8");

        echo "<textarea rows=10 cols=75 id='value".$tag[$i]['Name']."'>".$text."</textarea><br>";

        echo "<span>".$text."</span><br>";

        //without nl2br
        //echo "<textarea rows=10 cols=75>".$srecGetMemFields[$i]."</textarea><br>";
        echo "<span>".$srecGetMemFields[$i]."</span><br>";

        echo "<br>-------------------------------------------------------------------------------";
        echo "<br><br>";
      }
		}




  }
?>