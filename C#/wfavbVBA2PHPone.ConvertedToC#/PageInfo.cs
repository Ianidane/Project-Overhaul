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

namespace wfavbVBA2PHPone
{
	public class PageInfo
	{

		private string mstrPageName;
		public string Name {
			get { return mstrPageName; }
		}
		//Set(ByVal pstrName As String)
		//    mstrName = pstrName
		//End Set

		private List<string> mlstOpenFormNames = new List<string>();
		public List<string> OpenFormNames {
			get { return mlstOpenFormNames; }
		}
		//Set(ByVal pstrName As String)
		//    mstrName = pstrName
		//End Set

        Collection mcolControlsByName;//150215
        Collection mcolAllControls;
		FunctionsInfo mclsFunctionsInfo;
		Collection mlstFunctions = new Collection();
		public PageInfo(string pstrFilePathAndName)
		{
			mstrPageName = Strings.LCase(RemovePath(pstrFilePathAndName));

			string sstrSource = null;
			sstrSource = OpenFile(pstrFilePathAndName);

			int slonProcessLocation = 0;
//150215            long sslonProcessLocation = 0;
			//sstrSource string starts at location 1
//            slonProcessLocation = 1;
            slonProcessLocation = 0;
//150215            sslonProcessLocation = 1;

			//get control info
			ControlsInfo sclsControlsInfo = null;
			sclsControlsInfo = new ControlsInfo();
			ControlInfo sclsFormControl = null;
			//first (top) control should be Begin Form
//150203            sclsFormControl = sclsControlsInfo.GetControlsInfo(sstrSource, ref sslonProcessLocation);
            sclsFormControl = sclsControlsInfo.GetControlsInfo(sstrSource, ref slonProcessLocation);//150203
            mcolControlsByName = sclsControlsInfo.ControlsByName;
            mcolAllControls = sclsControlsInfo.AllControls;
			//Ian - break here 
			//inpect mlstAttributes and you can see attributes in a control
			//if you inspect sclsFormControl you can keep drilling down into SubControls to see the whole control part of the file 
			//I haven't had time yet - but I think the next step is to create a Collection that has all the Controls on the page using its Name as the index, not caring about where it is in the form
			//  what we are looking for is the button that is calling a certain OpenForm so we can put that control on the page and have it when clicked open the next page
			//        MsgBox("Ian break here")


			//GetControlsInfo above should leave slonProcessLocation at the end of Control area of file and just before Function area of file

			//get function info

			//I think we need lstOpenFormsNamesUnique - just names of forms to be processed
			//-change so only making unique right when adding to lstOpenFormsNameUnique if not doing tha talready
			//and a list of open forms with associalted function name
			//-for now skip controls
			//--goto functions - if only 1 OpenForm and its in Load then redirect to new page
			//---so need OpenForms by Function

			//150110        Dim sclsFunctionsInfo As FunctionsInfo
			//150110        sclsFunctionsInfo = New FunctionsInfo
			//150110        sclsFunctionsInfo.GetFunctionsInfo(sstrSource, slonProcessLocation)
			//151010        sclsFunctionsInfo = New FunctionsInfo(sstrSource, slonProcessLocation) '150110
			mclsFunctionsInfo = new FunctionsInfo(sstrSource, ref slonProcessLocation);
			//150110
			mlstFunctions = mclsFunctionsInfo.Functions;
			mlstOpenFormNames = mclsFunctionsInfo.OpenFormNames;
			//
			//        MsgBox("Ian break here to check out the mlstOpenFormsNames")
		}
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
            else//150215
                MessageBox.Show("File Not Found");//150215
			return functionReturnValue;
		}

		public void WritePHPPage(string pstrWriteToPHPFilesLocation, bool pbolBackupExistingPHPPages)
		{
			if (pbolBackupExistingPHPPages) {
				MessageBox.Show("Backing up Existing PHP Pages is not done yet");
				return;
			}

			string sstrPHP = null;
			sstrPHP = "";
			sstrPHP = sstrPHP + "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\" > " + System.Environment.NewLine;
			sstrPHP = sstrPHP + "<html>" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "  <head>" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "    <link href=\"images/AWH/AWHirb.css\" rel=\"stylesheet\" type=\"text/css\">" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "    <meta http-equiv=\"Content-Type\" content=\"text/html;\">" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "    <title>NewPro</title>" + System.Environment.NewLine;

			//styles to place buttons
			sstrPHP = sstrPHP + "    <style type=\"text/css\">" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "    </style>" + System.Environment.NewLine;

			//function code
			sstrPHP = sstrPHP + "    <script type=\"text/JavaScript\">" + System.Environment.NewLine;
			FunctionInfo sclsFunction = null;
			foreach (FunctionInfo sclsFunction_loopVariable in mlstFunctions) {
				sclsFunction = sclsFunction_loopVariable;
				sstrPHP = sstrPHP + "      function " + sclsFunction.Name + "(){" + System.Environment.NewLine;
				if (sclsFunction.NumOpenFormNames > 0) {
					sstrPHP = sstrPHP + "        location.href = \"" + sclsFunction.OpenFormNames[0] + ".php\";" + System.Environment.NewLine;
				}
				sstrPHP = sstrPHP + "      }" + System.Environment.NewLine;
			}
			sstrPHP = sstrPHP + "" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "    </script>" + System.Environment.NewLine;

			sstrPHP = sstrPHP + "  </head>" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "  <body ";
			//from http://bytes.com/topic/access/answers/196005-difference-between-form_open-form_load about VBA
			//When you first open a form, the following events occur in this order:Open Þ Load Þ Resize Þ Activate Þ Current
			//If you Then 're trying to decide whether to use the Open or Load event for your macro or event procedure, one significant difference is that the Open event can be canceled, but the Load event can't. For example, if you're dynamically building a record source for a form in an event procedure for the Form 's Open event, you can cancel opening the form if there are no records to display.
			//When you close a form, the following events occur in this order:Unload Þ Deactivate Þ Close
			if (mclsFunctionsInfo.HasFormLoad | mclsFunctionsInfo.HasFormOpen | mclsFunctionsInfo.HasFormActive | mclsFunctionsInfo.HasFormCurrent) {
				sstrPHP += "onLoad=\"";
				if (mclsFunctionsInfo.HasFormLoad)
					sstrPHP += "Form_Load();";
				if (mclsFunctionsInfo.HasFormOpen)
					sstrPHP += "Form_Open();";
                if (mclsFunctionsInfo.HasFormActive)
                    sstrPHP += "Form_Active();";
                if (mclsFunctionsInfo.HasFormCurrent)
                    sstrPHP += "Form_Current();";
				sstrPHP = sstrPHP + "\"";
			}
			if (mclsFunctionsInfo.HasFormUnload | mclsFunctionsInfo.HasFormClose) {
				sstrPHP += "onUnload=\"";
                if (mclsFunctionsInfo.HasFormUnload)
                    sstrPHP += "Form_Unload();";
                if (mclsFunctionsInfo.HasFormClose)
                    sstrPHP += "Form_Close();";
                sstrPHP += "\"";
            }
			sstrPHP = sstrPHP + ">" + System.Environment.NewLine;



            foreach (ControlInfo sclsControlInfo in mcolAllControls)
            {
                //need to make sure that Control has Name before making control
//MessageBox.Show(sclsControlInfo.Name);
                if (sclsControlInfo.Name != null)
                    if (sclsControlInfo.BeginType == "CommandButton")
                    {
                        sstrPHP += "    <input type='button' id=" + DropQuotes(sclsControlInfo.Name);
                        sstrPHP += " value='" + DropQuotes(sclsControlInfo.Caption) + "'";
                        sstrPHP += " style='position:absolute; left:" + sclsControlInfo.Left/10 + "; top:" + sclsControlInfo.Top/10 + "; width:" + sclsControlInfo.Width/10 + "; height:" + sclsControlInfo.Height/10 + "'";
                        if (sclsControlInfo.OnClick == "\"[Event Procedure]\"")
                            sstrPHP += " onclick='" + DropQuotes(sclsControlInfo.Name) + "_Click();'";
                        sstrPHP += ">" + System.Environment.NewLine;
                    }
            }

            sstrPHP = sstrPHP + "" + System.Environment.NewLine;
            sstrPHP = sstrPHP + "" + System.Environment.NewLine;




			sstrPHP = sstrPHP + "  </body>" + System.Environment.NewLine;
			sstrPHP = sstrPHP + "</html>" + System.Environment.NewLine;

			WriteFile(pstrWriteToPHPFilesLocation + RemoveFileExtension(mstrPageName) + ".php", sstrPHP);
		}
		private void WriteFile(string pstrFilePathAndName, string pstrString)
		{
			Scripting.FileSystemObject fs = new Scripting.FileSystemObject();
			Scripting.TextStream ts = null;

			if (fs.FileExists(pstrFilePathAndName)) {
				ts = fs.OpenTextFile(pstrFilePathAndName, Scripting.IOMode.ForWriting);
			} else {
				ts = fs.CreateTextFile(pstrFilePathAndName);
			}

			ts.Write(pstrString);
			ts.Close();
		}

		private string RemovePath(string pstrFileNameAndPath)
		{
			int sintLastSlash = 0;
			int sintInc = 0;
//150215			for (sintInc = 1; sintInc <= pstrFileNameAndPath.Length + 1; sintInc++) {
			for (sintInc = 0; sintInc < pstrFileNameAndPath.Length; sintInc++) {
				if (pstrFileNameAndPath.Substring(sintInc, 1) == "\\")
					sintLastSlash = sintInc + 1;
			}
//150215            return Strings.Right(pstrFileNameAndPath, pstrFileNameAndPath.Length - sintLastSlash);
            return pstrFileNameAndPath.Substring(sintLastSlash);
		}
		private string RemoveFileExtension(string pstrFileName)
		{
			int sintInc = 0;
//			for (sintInc = 1; sintInc <= pstrFileName.Length + 1; sintInc++) {
			for (sintInc = 0; sintInc <= pstrFileName.Length + 1; sintInc++) {
				if (pstrFileName.Substring(sintInc, 1) == ".")
					break; // TODO: might not be correct. Was : Exit For
			}
//150215            return Strings.Left(pstrFileName, sintInc - 1);
            return pstrFileName.Substring(0, sintInc);
		}

        private string DropQuotes(string pstrString)
        {
            string sstrQuotesDropped = "";
            if (pstrString.Substring(0, 1) == "'")
            {
                if (pstrString.Substring(pstrString.Length-1, 1) == "'")
                    return pstrString.Substring(1, pstrString.Length - 2);
            }
            else if (pstrString.Substring(0, 1) == "\"")
                if (pstrString.Substring(pstrString.Length-1, 1) == "\"")
                    return pstrString.Substring(1, pstrString.Length - 2);
            return sstrQuotesDropped;
        }
	}
}
