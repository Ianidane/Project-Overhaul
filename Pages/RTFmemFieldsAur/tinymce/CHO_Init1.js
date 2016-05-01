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
//  ,
//  setup : function(ed) {
//    ed.on('blur', function(e){
//      tinymce.triggerSave();
//      document.getElementById(ed.id).onchange();
//    });
//
//    ed.on('keydown', function(e){
//      tinymce.triggerSave();
//      document.getElementById(ed.id).onkeydown();
//    });
//
//    ed.on('keyup', function(e){
//      tinymce.triggerSave();
//      document.getElementById(ed.id).onkeyup();
//    });
//
//    ed.on('focus', function(e){
//      tinymce.triggerSave();
//      document.getElementById(ed.id).onfocus();
//    });
//  }
});


//*******************************old code 140507***********************************

//tinymce.init({
  //selector : 'textarea',
////  theme : 'advanced',

////130411 plugins : "spellchecker,paste,safari,inlinepopups",
 //plugins : 'spellchecker paste', //130816
////130816
////plugins : "paste,safari,inlinepopups,AtD",
////atd_button_url : "https://www.cyberirb.us/auroratest/tiny_mce/plugins/atd-tinymce/atdbuttontr.gif",
////atd_rpc_url : "https://www.cyberirb.us/auroratest/tiny_mce/plugins/atd-tinymce/server/proxy.php?url=",
////atd_rpc_id  : "",
////atd_css_url  : "https://www.cyberirb.us/auroratest/tiny_mce/plugins/atd-tinymce/css/content.css",
//// Theme options
//// 101205 removing charmap till can get workign
////theme_advanced_buttons1 : "bold,italic,underline,|,sub,sup,charmap,|,spellchecker,|,bullist,numlist,|,outdent,indent,|,undo,redo",
////130816theme_advanced_buttons1 : "bold,italic,underline,|,sub,sup,|,bullist,numlist,|,indent,AtD",
//toolbar : 'bold italic underline | sub sup | bullist numlist | indent',//130816
////theme_advanced_buttons2 : "",
////theme_advanced_buttons3 : "",
////spellchecker_languages : "+English=en",
////theme_advanced_toolbar_location : "top",
////theme_advanced_toolbar_align : "left",
////theme_advanced_statusbar_location : "bottom",
////theme_advanced_resizing : true
////,spellchecker_rpc_url : "tiny_mce/plugins/spellchecker/rpc.php"
//// 130411 ,spellchecker_rpc_url : "http://integranite.net/Test/tinymce/jscripts/tiny_mce/plugins/spellchecker/rpc.php"

//forced_root_block : false,
////,force_br_newlines : true
////,force_p_newlines : false
//,
    //setup : function(ed) {
////130228
////      ed.onChange.add(function(ed, e) {
////        document.getElementById(ed.id).fireEvent("onChange");
////      });
////      ed.onKeyDown.add(function(ed, e) {
////        document.getElementById(ed.id).fireEvent("onKeyDown");
////      });
////      ed.onKeyUp.add(function(ed, e) {
////        document.getElementById(ed.id).innerText = tinyMCE.get(ed.id).getContent();
////        document.getElementById(ed.id).fireEvent("onKeyUp");
////      });
//////      ed.onCut.add(function(ed, e) {
////////        var mobjPage = document.getElementById("COC04InitTitle");
////////        var svarname = "mobjPage.memProtocolTitle_oncut()";
////////        eval(svarname);
//////        document.getElementById(ed.id).fireEvent("onCut");
//////      });
////      ed.onPaste.add(function(ed, e) {
////        document.getElementById(ed.id).fireEvent("onPaste");
////      });
////      ed.onMouseUp.add(function(ed, e) {
////        document.getElementById(ed.id).innerText = tinyMCE.get(ed.id).getContent();
////        document.getElementById(ed.id).fireEvent("onMouseUp");
////      });

////130429
    //ed.onInit.add(function(ed) {
      //tinyMCE.dom.Event.add(ed.getWin(), "blur", function(e){
        //tinyMCE.triggerSave();
        //document.getElementById(ed.id).onchange();
      //})
    //});

    //ed.onKeyDown.add(function(ed, e){
      //tinyMCE.triggerSave();
      //document.getElementById(ed.id).onkeydown();
    //});

    //ed.onKeyUp.add(function(ed, e){
      //tinyMCE.triggerSave();
      //document.getElementById(ed.id).onkeyup();
    //});
  //}
//});
