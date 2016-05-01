<html>
  <head>
    <title></title>
    <script>
    </script>
  </head>
  <body>
<?php
    $sstrTableName = $_GET['strTableName'];
    $snumUniquekey = $_GET['numUniquekey'];
    $sstrMemFieldName = $_GET['strMemFieldName'];
    echo $sstrTableName." ".$snumUniquekey." ".$sstrMemFieldName."<br>";

    include("common-data.php");
    $ssqlGetMemFields = "SELECT $sstrMemFieldName FROM $sstrTableName WHERE uniquekey = '$snumUniquekey'";
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

    if($srecGetMemFields=sqlsrv_fetch_array($sresGetMemFields)){
      echo "<textarea rows=10 cols=75>".$srecGetMemFields[$sstrMemFieldName]."</textarea><br>";
      echo "<span>".$srecGetMemFields[$sstrMemFieldName]."</span><br>";
      echo "<span>".unicode_to_entities(utf8_to_unicode($srecGetMemFields[$sstrMemFieldName]))."</span>";
    }


//    function utf8_to_unicode( $str ) {
//        $unicode = array();
//        $values = array();
//        $lookingFor = 1;
//        for ($i = 0; $i < strlen( $str ); $i++ ) {
//            $thisValue = ord( $str[ $i ] );
//            if ( $thisValue < 128 ) $unicode[] = $thisValue;
//            else {
//                if ( count( $values ) == 0 ) $lookingFor = ( $thisValue < 224 ) ? 2 : 3;
//                $values[] = $thisValue;
//                if ( count( $values ) == $lookingFor ) {
//                    $number = ( $lookingFor == 3 ) ?
//                        ( ( $values[0] % 16 ) * 4096 ) + ( ( $values[1] % 64 ) * 64 ) + ( $values[2] % 64 ):
//                    	( ( $values[0] % 32 ) * 64 ) + ( $values[1] % 64 );
//
//                    $unicode[] = $number;
//                    $values = array();
//                    $lookingFor = 1;
//                } // if
//            } // if
//        } // for
//        return $unicode;
//    } // utf8_to_unicode
//
//    function unicode_to_entities( $unicode ) {
//      $entities = '';
//      foreach( $unicode as $value ) $entities .= '&#' . $value . ';';
//      return $entities;
//    }

?>
  </body>
</html>