<html>
  <head>
    <title></title>
    <script type="text/javascript" src="tinymce/tinymce.min.js"></script>
    <script type="text/javascript" src="tinymce/CHO_Init1.js"></script>
    <script>
      function GetRefNums(){
        aHC_GetRefNums=GetXmlHttpObject();
        if (aHC_GetRefNums==null){
          alert ("Your browser does not support XMLHTTP! Currently known acceptable browsers are IE5+, Firefox, Chrome, Opera, Safari.");
          return;
        }
        aHC_GetRefNums.onreadystatechange=GetRefNums_stateChanged;
        aHC_GetRefNums.open("Post","RTFmemFields_GetRefNums.php",true);
        aHC_GetRefNums.setRequestHeader("Content-type","application/x-www-form-urlencoded");
        var parameters = "strTableName="+document.getElementById('txtTableName').value;
        aHC_GetRefNums.setRequestHeader("Content-length",parameters.length);
        aHC_GetRefNums.setRequestHeader("Connection","close");
        aHC_GetRefNums.send(parameters);
      }
      var aHC_GetRefNums;
      function GetRefNums_stateChanged(){
        if (aHC_GetRefNums.readyState==4){
          divGetRefNums.innerHTML = aHC_GetRefNums.responseText;
        }
      }

      var mnumUniquekey = "not set";
      function GetMemFields(pnumUniquekey){
        if(mnumUniquekey!="not set") {
          document.getElementById(mnumUniquekey).style.color = "red";
          document.getElementById(mnumUniquekey).style.backgroundColor = "white";
        }
        mnumUniquekey = pnumUniquekey;
        aHC_GetMemFields=GetXmlHttpObject();
        if (aHC_GetMemFields==null){
          alert ("Your browser does not support XMLHTTP! Currently known acceptable browsers are IE5+, Firefox, Chrome, Opera, Safari.");
          return;
        }
        aHC_GetMemFields.onreadystatechange=GetMemFields_stateChanged;
        aHC_GetMemFields.open("Post","RTFmemFields_GetMemFields.php", true);
        aHC_GetMemFields.setRequestHeader("Content-type","application/x-www-form-urlencoded");
        var parameters = "strTableName="+document.getElementById('txtTableName').value;
        parameters = parameters + "&numUniquekey="+pnumUniquekey;
//        parameters = parameters + "&bol1AtATime="+pbol1AtATime;
        aHC_GetMemFields.setRequestHeader("Content-length",parameters.length);
        aHC_GetMemFields.setRequestHeader("Connection","close");
        aHC_GetMemFields.send(parameters);
      }
      var aHC_GetMemFields;
      function GetMemFields_stateChanged(){
        if (aHC_GetMemFields.readyState==4){
          divGetMemFields.innerHTML = aHC_GetMemFields.responseText;

          tinymce.init({
            selector:'textarea',
            plugins:'spellchecker paste advlist lists link image charmap print preview anchor searchreplace visualblocks code fullscreen insertdatetime media table',
            fix_list_elements : true,
            toolbar:'bold italic underline | sub sup | bullist numlist | indent | paste cut',
            forced_root_block : false,
            menubar : false,
            statusbar : false,
          //  external_plugins: { "nanospell": "plugins/TinyMCESpellChecker/plugin.js" },
          // nanospell_server:"php",
            browser_spellcheck : true,
            paste_preprocess: function(plugin, args) {
              args.content = args.content.replace(/\uF0B7/g,"<li>"); //copying from PDF after javascript escape bullet showed up as %uF0B7
            }
//            ,
//            setup : function(ed) {
//              ed.on('blur', function(e){
//                tinymce.triggerSave();
//                document.getElementById(ed.id).onchange();
//              });
//
//              ed.on('keydown', function(e){
//                tinymce.triggerSave();
//                document.getElementById(ed.id).onkeydown();
//              });
//
//              ed.on('keyup', function(e){
//                tinymce.triggerSave();
//                document.getElementById(ed.id).onkeyup();
//              });
//
//              ed.on('focus', function(e){
//                tinymce.triggerSave();
//                document.getElementById(ed.id).onfocus();
//              });
//            }
          });

//          document.getElementById(mnumUniquekey).style.color = "yellow";
          document.getElementById(mnumUniquekey).style.backgroundColor = "yellow";
        }
      }

      function Investigate(pstrMemFieldName){
        		var sarrArgIn = [];
            sarrArgIn[pstrMemFieldName] = pstrMemFieldName;

            var sobjArguments = new Object();
            sobjArguments.arrArgIn = sarrArgIn;
            sdlgShowDialog = window.showModalDialog("Investigate.php", sobjArguments, "dialogWidth:500;dialogHeight:500;status:off;scroll:off;menubar:off;help:off");

//            sxmlAction = mxmlAction;
//            sxmlReply = mxmlReply;
//
////alert(sdlgShowDialog.xmlDlgReply);
//
//						if(sdlgShowDialog!=null){
//							sstrParms = "xmlDlgReply="+sdlgShowDialog.xmlDlgReply; //091026
//						}
//						else{
//							sxmlDlgReply = "<\?xml version=\"1.0\"\?>";
//							sxmlDlgReply = sxmlDlgReply + "<DlgReply>";
//							sxmlDlgReply = sxmlDlgReply + "<ButtonClicked>";
//              sxmlDlgReply = sxmlDlgReply + "btnCanel"; //091028 trying btnCancel - I think it makes sense
//							sxmlDlgReply = sxmlDlgReply + "</ButtonClicked>";
//							sxmlDlgReply = sxmlDlgReply + "</DlgReply>";
//							sstrParms = "xmlDlgReply="+sxmlDlgReply;
//						}
//
//            //101025
//            sstrParms = sstrParms + "&xmlClassVars=<\?xml version=\"1.0\"\?><xmlClassVars>" + sxmlReply + "</xmlClassVars>";
//
//            //091213aftercomment - below are parameters that are always searched for and passed back after dialog
//						sstrTriggerCall=sxmlAction.attributes.getNamedItem("TriggerCall");
//						if(sstrTriggerCall != null){
//							sstrFunctionCall=sxmlAction.attributes.getNamedItem("strFunctionCall");
//							if(sstrFunctionCall != null){
//								sstrParms = sstrParms + "&strFunctionCall=" + sstrFunctionCall.value;
//							}
//							sstrTrigger=sxmlAction.attributes.getNamedItem("Trigger");
//							if(sstrTrigger != null){
//								sstrParms = sstrParms + "&strTrigger=" + sstrTrigger.value;
//							}
//							sstrProjectName=sxmlAction.attributes.getNamedItem("strProjectName");
//							if(sstrProjectName != null){
//								sstrParms = sstrParms + "&strProjectName=" + sstrProjectName.value;
//							}
//							slonStartWithActionNumber=sxmlAction.attributes.getNamedItem("lonStartWithActionNumber");
//							if(slonStartWithActionNumber != null){
//								sstrParms = sstrParms + "&lonStartWithActionNumber=" + slonStartWithActionNumber.value;
//							}
//							//uniquekey 091026
//							slonUniquekey=sxmlAction.attributes.getNamedItem("lonUniquekey");
//							if(slonUniquekey != null){
//								sstrParms = sstrParms + "&lonUniquekey=" + slonUniquekey.value;
//							}
//              //091212strRefNum
//              sstrRefNum=sxmlAction.attributes.getNamedItem("strRefNum");
//              if(sstrRefNum != null){
//                sstrParms =sstrParms + "&strRefNum=" + sstrRefNum.value;
//              }
//              //091212strCustNum
//              sstrCustNum=sxmlAction.attributes.getNamedItem("strCustNum");
//              if(sstrCustNum != null){
//                sstrParms =sstrParms + "&strCustNum=" + sstrCustNum.value;
//              }
//              //091212strLogonName
//              sstrLogonName=sxmlAction.attributes.getNamedItem("strLogonName");
//              if(sstrLogonName != null){
//                sstrParms =sstrParms + "&strLogonName=" + sstrLogonName.value;
//              }
//              //091212bolAuthAdminCoor
//              sbolAuthAdminCoor=sxmlAction.attributes.getNamedItem("bolAuthAdminCoor");
//              if(sbolAuthAdminCoor != null){
//                sstrParms =sstrParms + "&bolAuthAdminCoor=" + sbolAuthAdminCoor.value;
//              }
//              //091213strMainPageMode
//              sstrMainPageMode=sxmlAction.attributes.getNamedItem("strMainPageMode");
//              if(sstrMainPageMode != null){
//                sstrParms =sstrParms + "&strMainPageMode=" + sstrMainPageMode.value;
//              }
//
////alert("(tkrHandleReturn(515))"+sstrTriggerCall.value + ":"+ sstrParms);
////140419							tkrHR_TriggerCall(sstrTriggerCall.value, sstrParms);
//							tkrHR_TriggerCall(sstrTriggerCall.value, sstrParms, pbolLayerDlg);//140419
////alert("(517)tkrHandleReturn");
//						}
//
//            //091213
//            else{
////140419              HandleReturn(sdlgShowDialog.xmlDlgReply);
//              HandleReturn(sdlgShowDialog.xmlDlgReply, pbolLayerDlg);//140419
//            }


          document.getElementById(pstrMemFieldName).style.color = "red";
      }

      function InvestigateNoRTF(pstrMemFieldName){
//        		var sarrArgIn = [];
//            sarrArgIn[pstrMemFieldName] = pstrMemFieldName;
//
//            var sobjArguments = new Object();
//            sobjArguments.arrArgIn = sarrArgIn;
//            sdlgShowDialog = window.showModalDialog("InvestigateNoRTF.php?strMemFieldName="+pstrMemFieldName, sobjArguments, "dialogWidth:500;dialogHeight:500;status:off;scroll:off;menubar:off;help:off");
            sdlgShowDialog = window.showModalDialog("InvestigateNoRTF.php?strTableName="+document.getElementById('txtTableName').value+"&numUniquekey="+mnumUniquekey+"&strMemFieldName="+pstrMemFieldName, "", "dialogWidth:250;dialogHeight:250");

//            sxmlAction = mxmlAction;
//            sxmlReply = mxmlReply;
//
////alert(sdlgShowDialog.xmlDlgReply);
//
//						if(sdlgShowDialog!=null){
//							sstrParms = "xmlDlgReply="+sdlgShowDialog.xmlDlgReply; //091026
//						}
//						else{
//							sxmlDlgReply = "<\?xml version=\"1.0\"\?>";
//							sxmlDlgReply = sxmlDlgReply + "<DlgReply>";
//							sxmlDlgReply = sxmlDlgReply + "<ButtonClicked>";
//              sxmlDlgReply = sxmlDlgReply + "btnCanel"; //091028 trying btnCancel - I think it makes sense
//							sxmlDlgReply = sxmlDlgReply + "</ButtonClicked>";
//							sxmlDlgReply = sxmlDlgReply + "</DlgReply>";
//							sstrParms = "xmlDlgReply="+sxmlDlgReply;
//						}
//
//            //101025
//            sstrParms = sstrParms + "&xmlClassVars=<\?xml version=\"1.0\"\?><xmlClassVars>" + sxmlReply + "</xmlClassVars>";
//
//            //091213aftercomment - below are parameters that are always searched for and passed back after dialog
//						sstrTriggerCall=sxmlAction.attributes.getNamedItem("TriggerCall");
//						if(sstrTriggerCall != null){
//							sstrFunctionCall=sxmlAction.attributes.getNamedItem("strFunctionCall");
//							if(sstrFunctionCall != null){
//								sstrParms = sstrParms + "&strFunctionCall=" + sstrFunctionCall.value;
//							}
//							sstrTrigger=sxmlAction.attributes.getNamedItem("Trigger");
//							if(sstrTrigger != null){
//								sstrParms = sstrParms + "&strTrigger=" + sstrTrigger.value;
//							}
//							sstrProjectName=sxmlAction.attributes.getNamedItem("strProjectName");
//							if(sstrProjectName != null){
//								sstrParms = sstrParms + "&strProjectName=" + sstrProjectName.value;
//							}
//							slonStartWithActionNumber=sxmlAction.attributes.getNamedItem("lonStartWithActionNumber");
//							if(slonStartWithActionNumber != null){
//								sstrParms = sstrParms + "&lonStartWithActionNumber=" + slonStartWithActionNumber.value;
//							}
//							//uniquekey 091026
//							slonUniquekey=sxmlAction.attributes.getNamedItem("lonUniquekey");
//							if(slonUniquekey != null){
//								sstrParms = sstrParms + "&lonUniquekey=" + slonUniquekey.value;
//							}
//              //091212strRefNum
//              sstrRefNum=sxmlAction.attributes.getNamedItem("strRefNum");
//              if(sstrRefNum != null){
//                sstrParms =sstrParms + "&strRefNum=" + sstrRefNum.value;
//              }
//              //091212strCustNum
//              sstrCustNum=sxmlAction.attributes.getNamedItem("strCustNum");
//              if(sstrCustNum != null){
//                sstrParms =sstrParms + "&strCustNum=" + sstrCustNum.value;
//              }
//              //091212strLogonName
//              sstrLogonName=sxmlAction.attributes.getNamedItem("strLogonName");
//              if(sstrLogonName != null){
//                sstrParms =sstrParms + "&strLogonName=" + sstrLogonName.value;
//              }
//              //091212bolAuthAdminCoor
//              sbolAuthAdminCoor=sxmlAction.attributes.getNamedItem("bolAuthAdminCoor");
//              if(sbolAuthAdminCoor != null){
//                sstrParms =sstrParms + "&bolAuthAdminCoor=" + sbolAuthAdminCoor.value;
//              }
//              //091213strMainPageMode
//              sstrMainPageMode=sxmlAction.attributes.getNamedItem("strMainPageMode");
//              if(sstrMainPageMode != null){
//                sstrParms =sstrParms + "&strMainPageMode=" + sstrMainPageMode.value;
//              }
//
////alert("(tkrHandleReturn(515))"+sstrTriggerCall.value + ":"+ sstrParms);
////140419							tkrHR_TriggerCall(sstrTriggerCall.value, sstrParms);
//							tkrHR_TriggerCall(sstrTriggerCall.value, sstrParms, pbolLayerDlg);//140419
////alert("(517)tkrHandleReturn");
//						}
//
//            //091213
//            else{
////140419              HandleReturn(sdlgShowDialog.xmlDlgReply);
//              HandleReturn(sdlgShowDialog.xmlDlgReply, pbolLayerDlg);//140419
//            }


          document.getElementById(pstrMemFieldName).style.color = "red";
      }

      function Place(pstrMemFieldName,pstrValue){
//alert(pstrMemFieldName+":"+pstrValue)
//alert(document.getElementById('divGetMemFields').innerHTML);
        aHC_Place=GetXmlHttpObject();
        if (aHC_Place==null){
          alert ("Your browser does not support XMLHTTP! Currently known acceptable browsers are IE5+, Firefox, Chrome, Opera, Safari.");
          return;
        }
        aHC_Place.onreadystatechange=Place_stateChanged;
        aHC_Place.open("Post","RTFmemFields_Place.php",true);
        aHC_Place.setRequestHeader("Content-type","application/x-www-form-urlencoded");
        var parameters = "strTableName="+document.getElementById('txtTableName').value;
        parameters = parameters + "&numUniquekey="+mnumUniquekey;
        parameters = parameters + "&strMemFieldName="+pstrMemFieldName;
        parameters = parameters + "&strValue="+pstrValue;
        aHC_Place.setRequestHeader("Content-length",parameters.length);
        aHC_Place.setRequestHeader("Connection","close");
        aHC_Place.send(parameters);
      }
      function Place_stateChanged(){
        if (aHC_Place.readyState==4){
          alert(unescape(aHC_Place.responseText));
        }
      }

      function GetXmlHttpObject(){
        if (window.XMLHttpRequest){
          // code for IE7+, Firefox, Chrome, Opera, Safari
          return new XMLHttpRequest();
        }
        if (window.ActiveXObject){
          // code for IE6, IE5
          return new ActiveXObject("Microsoft.XMLHTTP");
        }
        return null;
      }
      function loadXMLString(txt){
        try{ //Internet Explorer
          xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
          xmlDoc.async="false";
          xmlDoc.loadXML(txt);
          return xmlDoc;
        }
        catch (e){
          try{ //Firefox, Mozilla, Opera, etc.
            parser=new DOMParser();
            xmlDoc=parser.parseFromString(txt,"text/xml");
            return xmlDoc;
          }
          catch(e) {alert(e.message)}
        }
        return null;
      }


      function PlaceAll(){
//alert(pstrMemFieldName+":"+pstrValue)
//alert(document.getElementById('divGetMemFields').innerHTML);
        aHC_PlaceAll=GetXmlHttpObject();
        if (aHC_PlaceAll==null){
          alert ("Your browser does not support XMLHTTP! Currently known acceptable browsers are IE5+, Firefox, Chrome, Opera, Safari.");
          return;
        }
        aHC_PlaceAll.onreadystatechange=PlaceAll_stateChanged;
        aHC_PlaceAll.open("Post","RTFmemFields_PlaceAll.php",true);
        aHC_PlaceAll.setRequestHeader("Content-type","application/x-www-form-urlencoded");
        var parameters = "strTableName="+document.getElementById('txtTableName').value;
        parameters = parameters + "&numUniquekey="+mnumUniquekey;
//        parameters = parameters + "&strMemFieldName="+pstrMemFieldName;
//        parameters = parameters + "&strValue="+pstrValue;
        aHC_PlaceAll.setRequestHeader("Content-length",parameters.length);
        aHC_PlaceAll.setRequestHeader("Connection","close");
        aHC_PlaceAll.send(parameters);
      }
      function PlaceAll_stateChanged(){
        if (aHC_PlaceAll.readyState==4){
          alert(unescape(aHC_PlaceAll.responseText));
        }
      }

      function GetXmlHttpObject(){
        if (window.XMLHttpRequest){
          // code for IE7+, Firefox, Chrome, Opera, Safari
          return new XMLHttpRequest();
        }
        if (window.ActiveXObject){
          // code for IE6, IE5
          return new ActiveXObject("Microsoft.XMLHTTP");
        }
        return null;
      }
      function loadXMLString(txt){
        try{ //Internet Explorer
          xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
          xmlDoc.async="false";
          xmlDoc.loadXML(txt);
          return xmlDoc;
        }
        catch (e){
          try{ //Firefox, Mozilla, Opera, etc.
            parser=new DOMParser();
            xmlDoc=parser.parseFromString(txt,"text/xml");
            return xmlDoc;
          }
          catch(e) {alert(e.message)}
        }
        return null;
      }

    </script>
  </head>
  <body>
    <table>
      <tr>
        <td>
          Backup uniquekey?
        </td>
        <td colspan="3">
          CHOTest - Table Name:
          <INPUT id=txtTableName value="CHOInit">
          <input type="button" value="Get RefNums" id="btnGetRefNums" onclick="GetRefNums()">
          <input type="button" value="Replace %u" id="btnReplace" onclick="PlaceAll()">
        </td>
      </tr>
      <tr>
        <td style="width:200px"><div id="divGetRefNums"></div></td>
        <td><div id="divGetMemFields"></div></td>
      </tr>
    </table>
  </body>
</html>