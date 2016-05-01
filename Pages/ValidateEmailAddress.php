<?php

//if(isset($_POST['strEmailAddr'])){
  $sstrEmailAddr =  "testing'apostrophe@proirb.com";

//150322  $sarrEmailAddrs = explode(',', str_replace(";",",",$sstrEmailAddr));
  $sarrEmailAddrs = explode(',', $sstrEmailAddr);//150322 moved this replace to the SendEmail dialog
  foreach($sarrEmailAddrs as $sstrEachEmailAddr) {
    if(!filter_var($sstrEachEmailAddr, FILTER_VALIDATE_EMAIL)){
      echo $sstrEachEmailAddr;
      break;
    }
    else
      echo "success";
  }
//}

?>