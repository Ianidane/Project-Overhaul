using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Linq;
using System.Threading.Tasks;
using IronPython.Hosting;

//Git Test


//Alex's git test
namespace wfavbVBA2PHPone
{
    public class PageInfo
    {

        private string mstrPageName;
        public string Name
        {
            get { return mstrPageName; }
        }

        private List<string> mlstOpenFormNames = new List<string>();
        public List<string> OpenFormNames
        {
            get { return mlstOpenFormNames; }
        }

        Collection mcolControlsByName; //not used?
        Collection mcolAllControls; //used in WritePHPPage function below
        FunctionsInfo mclsFunctionsInfo;
        Collection mlstFunctions = new Collection();
        public PageInfo(string pstrFilePathAndName)
        {
            //            mstrPageName = Strings.LCase(RemovePath(pstrFilePathAndName));
            mstrPageName = RemovePath(pstrFilePathAndName);

            string sstrSource = null;
            sstrSource = OpenFile(pstrFilePathAndName);

            int slonProcessLocation = 0;


            //get control info
            ControlsInfo sclsControlsInfo = null;
            sclsControlsInfo = new ControlsInfo();
            ControlInfo sclsFormControl = null;
            //first (top) control should be Begin Form
            sclsFormControl = sclsControlsInfo.GetControlsInfo(sstrSource, ref slonProcessLocation);
            mcolControlsByName = sclsControlsInfo.ControlsByName;//not used?
            mcolAllControls = sclsControlsInfo.AllControls;


            //GetControlsInfo above should leave slonProcessLocation at the end of Control area of file and just before Function area of file


            //get function info
            mclsFunctionsInfo = new FunctionsInfo(sstrSource, ref slonProcessLocation);
            mlstFunctions = mclsFunctionsInfo.Functions; //used in WritePHPPage function below
            mlstOpenFormNames = mclsFunctionsInfo.OpenFormNames; //used in WritePHPPage function below
        }
        //git test 2
        private string OpenFile(string pstrFilePathAndName)
        {
            string functionReturnValue = null;
            string sstrFile = null;
            Scripting.FileSystemObject fs = new Scripting.FileSystemObject();
            Scripting.TextStream ts = null;

            functionReturnValue = "";
            if (fs.FileExists(pstrFilePathAndName))
            {
                ts = fs.OpenTextFile(pstrFilePathAndName);
                sstrFile = ts.ReadAll();
                ts.Close();
                functionReturnValue = sstrFile;
            }
            else
                MessageBox.Show("File Not Found");
            return functionReturnValue;
        }

        public void WritePHPPage(string pstrWriteToPHPFilesLocation, bool pbolBackupExistingPHPPages)
        {
            if (pbolBackupExistingPHPPages)
            {
                MessageBox.Show("Backing up Existing PHP Pages is not done yet");
                return;
            }

            string sstrPHP = "";
            sstrPHP += "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\" > " + System.Environment.NewLine;
            sstrPHP += "<?php include 'Serverinfo.php'; ?>" + System.Environment.NewLine;
            sstrPHP += "<html>" + System.Environment.NewLine;
            sstrPHP += "  <head>" + System.Environment.NewLine;
            sstrPHP += "    <link href=\"images/NewPro.css\" rel=\"stylesheet\" type=\"text/css\">" + System.Environment.NewLine;
            sstrPHP += "    <meta http-equiv=\"Content-Type\" content=\"text/html;\">" + System.Environment.NewLine;
            sstrPHP += "    <title>NewPro</title>" + System.Environment.NewLine;

            //function code
            sstrPHP += "    <script type=\"text/JavaScript\">" + System.Environment.NewLine;
            foreach (FunctionInfo sclsEachFunction in mlstFunctions)
            {
                //150218                sstrPHP += "      function " + sclsEachFunction.Name + "(){" + System.Environment.NewLine;
                //150218
                //                if (sclsEachFunction.NumOpenFormNames > 0) {
                //                    sstrPHP += "        location.href = \"Form_" + sclsEachFunction.OpenFormNames[0] + ".php\";" + System.Environment.NewLine;
                //                }
                sstrPHP += sclsEachFunction.FunctionTranslated;

                //				sstrPHP += "      }" + System.Environment.NewLine;
            }
            sstrPHP += "" + System.Environment.NewLine;
            sstrPHP += "    </script>" + System.Environment.NewLine;

            sstrPHP += "  </head>" + System.Environment.NewLine;
            sstrPHP += "  <body";
            //from http://bytes.com/topic/access/answers/196005-difference-between-form_open-form_load about VBA
            //When you first open a form, the following events occur in this order:Open Þ Load Þ Resize Þ Activate Þ Current
            //If you Then 're trying to decide whether to use the Open or Load event for your macro or event procedure, one significant difference is that the Open event can be canceled, but the Load event can't. For example, if you're dynamically building a record source for a form in an event procedure for the Form 's Open event, you can cancel opening the form if there are no records to display.
            //When you close a form, the following events occur in this order:Unload Þ Deactivate Þ Close
            if (mclsFunctionsInfo.HasFormLoad | mclsFunctionsInfo.HasFormOpen | mclsFunctionsInfo.HasFormActive | mclsFunctionsInfo.HasFormCurrent)
            {
                sstrPHP += " onLoad=\"";
                if (mclsFunctionsInfo.HasFormLoad)
                    sstrPHP += "Form_Load();";
                if (mclsFunctionsInfo.HasFormOpen)
                    sstrPHP += "Form_Open();";
                if (mclsFunctionsInfo.HasFormActive)
                    sstrPHP += "Form_Active();";
                if (mclsFunctionsInfo.HasFormCurrent)
                    sstrPHP += "Form_Current();";
                sstrPHP += "\"";
            }
            if (mclsFunctionsInfo.HasFormUnload | mclsFunctionsInfo.HasFormClose)
            {
                sstrPHP += " onUnload=\"";
                if (mclsFunctionsInfo.HasFormUnload)
                    sstrPHP += "Form_Unload();";
                if (mclsFunctionsInfo.HasFormClose)
                    sstrPHP += "Form_Close();";
                sstrPHP += "\"";
            }
            sstrPHP += ">" + System.Environment.NewLine;


            foreach (ControlInfo sclsControlInfo in mcolAllControls)
            {
                //need to make sure that Control has Name before making control       
                if (sclsControlInfo.Name != null)
                {
                    if (sclsControlInfo.BeginType == "CommandButton")
                    {
                        if (sclsControlInfo.Visible != null)
                        {
                            sstrPHP += "    <!-- ";
                        }
                        sstrPHP += "    <input type='button' id=" + DropQuotes(sclsControlInfo.Name);
                        sstrPHP += " value='" + DropQuotes(sclsControlInfo.Caption) + "'";
                        sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left / 10 + "; top:" + sclsControlInfo.Top / 10 + "; width:" + sclsControlInfo.Width / 10 + "; height:" + sclsControlInfo.Height / 10 + "'";
                        if (sclsControlInfo.OnClick == "\"[Event Procedure]\"")
                            sstrPHP += " onclick='" + DropQuotes(sclsControlInfo.Name) + "_Click();'";
                        sstrPHP += "></input>";
                        if (sclsControlInfo.Visible != null)
                        {
                            sstrPHP += " -->";
                        }
                    }
                }
                if (sclsControlInfo.BeginType == "TextBox")
                {
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += "    <!-- ";
                    }
                    if (sclsControlInfo.Name == null)
                    {
                        sstrPHP += "    <input type='text' id='empty'";
                    }
                    else
                    {
                        sstrPHP += "    <input type='text' id=" + DropQuotes(sclsControlInfo.Name);
                    }
                    if (sclsControlInfo.Caption == null)
                    {
                        sstrPHP += " value='empty'";
                    }
                    else
                    {
                        sstrPHP += " value='" + DropQuotes(sclsControlInfo.Caption) + "'";
                    }
                    if (sclsControlInfo.Left == null && sclsControlInfo.Top == null && sclsControlInfo.Width == null && sclsControlInfo.Height == null)
                    {
                        sstrPHP += " style=''></input>";
                    }
                    else
                    {
                        sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left / 10 + "; top:" + sclsControlInfo.Top / 10 + "; width:" + sclsControlInfo.Width / 10 + "; height:" + sclsControlInfo.Height / 10 + "'></input>";
                    }
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += " -->";
                    }
                }
                if (sclsControlInfo.BeginType == "Label")
                {
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += "    <!-- ";
                    }
                    if (sclsControlInfo.Name == null)
                    {
                        sstrPHP += "    <label id='empty'";
                    }
                    else
                    {
                        sstrPHP += "    <label id=" + DropQuotes(sclsControlInfo.Name);
                    }
                    if (sclsControlInfo.Caption == null)
                    {
                        sstrPHP += " value='empty'";
                    }
                    else
                    {
                        sstrPHP += " value='" + DropQuotes(sclsControlInfo.Caption) + "'";
                    }
                    if (sclsControlInfo.Left == 0 && sclsControlInfo.Top == 0 && sclsControlInfo.Width == 0 && sclsControlInfo.Height == 0)
                    {
                        if (sclsControlInfo.Caption == null)
                        {
                            sstrPHP += " style='visibility:" + sclsControlInfo.Visible + "'></label>";
                        }
                        else
                        {
                            sstrPHP += " style='visibility:" + sclsControlInfo.Visible + "'>" + DropQuotes(sclsControlInfo.Caption) + "</label>";
                        }
                    }
                    else
                    {
                        if (sclsControlInfo.Caption == null)
                        {
                            sstrPHP += " style='visibility:" + sclsControlInfo.Visible + "'></label>";
                        }
                        else
                        {
                            sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left / 10 + "; top:" + sclsControlInfo.Top / 10 + "; width:" + sclsControlInfo.Width / 10 + "; height:" + sclsControlInfo.Height / 10 + "; visibility:" + sclsControlInfo.Visible + "'>" + DropQuotes(sclsControlInfo.Caption) + "</label>";
                        }

                    }
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += " -->";
                    }
                }
                if (sclsControlInfo.BeginType == "CheckBox")
                {
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += "    <!-- ";
                    }
                    if (sclsControlInfo.Name == null)
                    {
                        sstrPHP += "    <input type='checkbox' id='empty'";
                    }
                    else
                    {
                        sstrPHP += "    <input type='checkbox' id=" + DropQuotes(sclsControlInfo.Name);
                    }

                    if (sclsControlInfo.DefaultValue == null)
                    {
                        sstrPHP += " value='False'";
                    }
                    else
                    {
                        sstrPHP += " value='" + DropQuotes(sclsControlInfo.DefaultValue) + "'";
                    }
                    if (sclsControlInfo.Left == null && sclsControlInfo.Top == null && sclsControlInfo.Width == null && sclsControlInfo.Height == null)
                    {
                        sstrPHP += " style=''></input>";
                    }
                    else
                    {
                        sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left / 10 + "; top:" + sclsControlInfo.Top / 10 + "; width:" + sclsControlInfo.Width / 10 + "; height:" + sclsControlInfo.Height / 10 + "'></input>";
                    }
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += " -->";
                    }
                }
                if (sclsControlInfo.BeginType == "OptionButton")
                {
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += "    <!-- ";
                    }
                    sstrPHP += "    <input type='radio'";

                    if (sclsControlInfo.Left == null && sclsControlInfo.Top == null && sclsControlInfo.Width == null && sclsControlInfo.Height == null)
                    {
                        sstrPHP += " style=''></input>";
                    }
                    else
                    {
                        sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left / 10 + "; top:" + sclsControlInfo.Top / 10 + "; width:" + sclsControlInfo.Width / 10 + "; height:" + sclsControlInfo.Height / 10 + "'></input>";
                    }
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += " -->";
                    }
                }

                if (sclsControlInfo.BeginType == "ToggleButton") // I quite possibly may have used the wrong html for this one, I used a regular button. Couldnt find a toggleable button in html closest I think is a radiobutton but that requires a label next to it
                {
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += "    <!-- ";
                    }
                    sstrPHP += "    <button name='";

                    if (sclsControlInfo.Name == null)
                    {
                        sstrPHP += "empty'";
                    }
                    else
                    {
                        sstrPHP += "" + DropQuotes(sclsControlInfo.Name) + "'";
                    }

                    sstrPHP += " type='submit'";

                    if (sclsControlInfo.Left == null && sclsControlInfo.Top == null && sclsControlInfo.Width == null && sclsControlInfo.Height == null)
                    {
                        sstrPHP += " style=''>";
                    }
                    else
                    {
                        sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left / 10 + "; top:" + sclsControlInfo.Top / 10 + "; width:" + sclsControlInfo.Width / 10 + "; height:" + sclsControlInfo.Height / 10 + "'";
                    }

                    if (sclsControlInfo.OnClick == "\"[Event Procedure]\"")
                        sstrPHP += " onclick='" + DropQuotes(sclsControlInfo.Name) + "_Click();'";
                    sstrPHP += ">";

                    sstrPHP += sclsControlInfo.Caption;
                    sstrPHP += "</button>";
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += " -->";
                    }
                }

                if (sclsControlInfo.BeginType == "ListBox" || sclsControlInfo.BeginType == "ComboBox")
                {
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += "    <!-- ";
                    }
                    sstrPHP += "    <select name='";

                    if (sclsControlInfo.Name == null)
                    {
                        sstrPHP += "empty'";
                    }
                    else
                    {
                        sstrPHP += "" + DropQuotes(sclsControlInfo.Name) + "'";
                    }

                    if (sclsControlInfo.TabIndex != 0)
                    {
                        sstrPHP += " Size='" + sclsControlInfo.TabIndex + "'";
                    }

                    if (sclsControlInfo.Left == null && sclsControlInfo.Top == null && sclsControlInfo.Width == null && sclsControlInfo.Height == null)
                    {
                        sstrPHP += " style=''";
                    }
                    else
                    {
                        sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left / 10 + "; top:" + sclsControlInfo.Top / 10 + "; width:" + sclsControlInfo.Width / 10 + "; height:" + sclsControlInfo.Height / 10 + "'";
                    }

                    if (sclsControlInfo.OnClick == "\"[Event Procedure]\"")
                        sstrPHP += " onclick='" + DropQuotes(sclsControlInfo.Name) + "_Click();'";
                    sstrPHP += ">";
                    if (sclsControlInfo.RowSource != null)
                    {
                        var py = Python.CreateEngine();                           
                        var paths = py.GetSearchPaths();                               
                        var scope = py.CreateScope();                                   
                        paths.Add(@"C:\Python27\Lib");                            
                        py.SetSearchPaths(paths);                                       
                                                       


                        string rowsrc = sclsControlInfo.RowSource;

                        //if (rowsrc.StartsWith("\"\\"))
                        //{
                        //    scope.SetVariable("name", sclsControlInfo.Name);
                        //    scope.SetVariable("querytype", "noquery");
                        //    scope.SetVariable("rowsource", rowsrc);
                        //    scope.SetVariable("filename", mstrPageName);
                        //    py.ExecuteFile("QueryFind.py", scope);
                        //}

                        //else if (rowsrc.StartsWith("\"_\""))
                        //{
                        //    scope.SetVariable("querytype", "find");
                        //    scope.SetVariable("rowsource", rowsrc);
                        //    scope.SetVariable("filename", mstrPageName);
                        //    py.ExecuteFile("QueryFind.py", scope);
                        //}

                        if (rowsrc.StartsWith("\"SELECT"))
                        {
                            scope.SetVariable("name", sclsControlInfo.Name);
                            scope.SetVariable("namelength", sclsControlInfo.Name.Length);
                            scope.SetVariable("querytype", "needprettyquery");
                            scope.SetVariable("rowsource", rowsrc);
                            scope.SetVariable("filename", mstrPageName);
                            py.ExecuteFile("C:/Users/Alex/Desktop/Work/Project-Overhaul/C#/wfavbVBA2PHPone.ConvertedToC#/QueryFind.py", scope);
                            py.CreateScriptSourceFromFile("C:/Users/Alex/Desktop/Work/Project-Overhaul/C#/wfavbVBA2PHPone.ConvertedToC#/QueryFind.py").Execute(scope);        //get information to use .getvariable
                            string result = scope.GetVariable("result");
                            string selectname = scope.GetVariable("selectname");

                            //Connect to DB
                            sstrPHP += System.Environment.NewLine;
                            sstrPHP += "    <?php" + System.Environment.NewLine;
                            sstrPHP += "        $serverName = $SVNM;" + System.Environment.NewLine;
                            sstrPHP += "        $connectionInfo = array( \"Database\"=>$DB, \"UID\"=>$UID, \"PWD\"=>$PWD);" + System.Environment.NewLine;
                            sstrPHP += "        $conn = sqlsrv_connect( $serverName, $connectionInfo);" + System.Environment.NewLine + System.Environment.NewLine;
                            sstrPHP += "        if( $conn ) {" + System.Environment.NewLine + "        }" + System.Environment.NewLine;
                            sstrPHP += "        else{" + System.Environment.NewLine;
                            sstrPHP += "            echo \"Connection could not be established.<br />\";" + System.Environment.NewLine;
                            sstrPHP += "            die( print_r( sqlsrv_errors(), true));" + System.Environment.NewLine + "        }" + System.Environment.NewLine + System.Environment.NewLine;

                            //Specific query
                            sstrPHP += "        $sql = sqlsrv_query($conn, " + "\"" + result + "\"" + ");" + System.Environment.NewLine;
                            sstrPHP += "        while( $row = sqlsrv_fetch_array($sql)){" + System.Environment.NewLine;
                            sstrPHP += "        echo \"<Option value='\".$row['" + selectname + "'].\"'\".\">\".$row['" + selectname + "'].\"</Option>\";" + System.Environment.NewLine + "        }" + System.Environment.NewLine + "    ?>" + System.Environment.NewLine;

                        }

                    }

                    sstrPHP += "    </select>";
                    if (sclsControlInfo.Visible != null)
                    {
                        sstrPHP += " -->";
                    }

                   
                }
                

                //Needed since it is above with the button write code?
                //if (sclsControlInfo.OnClick == "\"[Event Procedure]\"")
                //    sstrPHP += " onclick='" + DropQuotes(sclsControlInfo.Name) + "_Click();'>";
                sstrPHP += System.Environment.NewLine;
            }


            sstrPHP += "  </body>" + System.Environment.NewLine;
            sstrPHP += "</html>" + System.Environment.NewLine;

            WriteFile(pstrWriteToPHPFilesLocation + RemoveFileExtension(mstrPageName) + ".php", sstrPHP);
        }
        private void WriteFile(string pstrFilePathAndName, string pstrString)
        {
            Scripting.FileSystemObject fs = new Scripting.FileSystemObject();
            Scripting.TextStream ts = null;

            if (fs.FileExists(pstrFilePathAndName))
            {
                ts = fs.OpenTextFile(pstrFilePathAndName, Scripting.IOMode.ForWriting);
            }
            else
            {
                ts = fs.CreateTextFile(pstrFilePathAndName);
            }

            ts.Write(pstrString);
            ts.Close();
        }

        private string RemovePath(string pstrFileNameAndPath)
        {
            int sintLastSlash = 0;
            int sintInc = 0;
            for (sintInc = 0; sintInc < pstrFileNameAndPath.Length; sintInc++)
            {
                if (pstrFileNameAndPath.Substring(sintInc, 1) == "\\")
                    sintLastSlash = sintInc + 1;
            }
            return pstrFileNameAndPath.Substring(sintLastSlash);
        }
        private string RemoveFileExtension(string pstrFileName)
        {
            int sintInc = 0;
            for (sintInc = 0; sintInc <= pstrFileName.Length + 1; sintInc++)
            {
                if (pstrFileName.Substring(sintInc, 1) == ".")
                    break; // TODO: might not be correct. Was : Exit For
            }
            return pstrFileName.Substring(0, sintInc);
        }

        private string DropQuotes(string pstrString)
        {
            string sstrQuotesDropped = "";
            if (pstrString.Substring(0, 1) == "'")
            {
                if (pstrString.Substring(pstrString.Length - 1, 1) == "'")
                    return pstrString.Substring(1, pstrString.Length - 2);
            }
            else if (pstrString.Substring(0, 1) == "\"")
                if (pstrString.Substring(pstrString.Length - 1, 1) == "\"")
                    return pstrString.Substring(1, pstrString.Length - 2);
            return sstrQuotesDropped;
        }

    }
}
