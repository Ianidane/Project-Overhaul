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
	public class FunctionInfo
	{
		private string mstrFunctionName;
		public string Name {
			get { return mstrFunctionName; }
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

		private long mlonNumOpenFormNames;
		public long NumOpenFormNames {
			get { return mlonNumOpenFormNames; }
		}
		//Set(ByVal pstrName As String)
		//    mstrName = pstrName
		//End Set

		public FunctionInfo()
		{
			mlonNumOpenFormNames = 0;
		}

		public bool PopulateFunctionInfoWithNextFunction(string pstrSource, ref int plonProcessLocation)
		{
			bool sbolFoundFunction = false;
			sbolFoundFunction = false;
			//note - you'll notice most "Do Until"s will have "Or plonProcessLocation >= Len(pstrSource)" in the condition for stopping the looping.  This is to be sure if we unexpectedly find the end of the file that we don't error.  Probably would be a better idea to actually throw an error if that happened rather than just falling out of the function
			//find Sub or Function
//150215			while (!(Strings.Mid(pstrSource, plonProcessLocation, 3) == "Sub" | Strings.Mid(pstrSource, plonProcessLocation, 8) == "Function" | plonProcessLocation >= Strings.Len(pstrSource))) {
			while (!(pstrSource.Substring(plonProcessLocation, 3) == "Sub" | pstrSource.Substring(plonProcessLocation, 8) == "Function" | plonProcessLocation >= pstrSource.Length)) {
				plonProcessLocation = plonProcessLocation + 1;
			}
			string EndFunctionType = null;
			//used to find the end of the Sub or Function
			EndFunctionType = "";
            if (pstrSource.Substring(plonProcessLocation, 3) == "Sub")
            {
                EndFunctionType = "End Sub";
                plonProcessLocation = plonProcessLocation + 3;//jump over the word "Sub"
            }
            else if (pstrSource.Substring(plonProcessLocation, 8) == "Function")
            {
                EndFunctionType = "End Function";
                plonProcessLocation = plonProcessLocation + 8;//jump over the word Function
            }
            else
                MessageBox.Show("failed to find function type");
			plonProcessLocation = plonProcessLocation + 1;//step over space after Sub or Function

			//set start FunctionName
			int slonStartFunctionNamePos = 0;
			slonStartFunctionNamePos = plonProcessLocation;
			//find end of function name - pretty sure that function always has ( after it
//150215			while (!(Strings.Mid(pstrSource, plonProcessLocation, 1) == "(" | plonProcessLocation >= Strings.Len(pstrSource))) {
			while (!(pstrSource.Substring(plonProcessLocation, 1) == "(" | plonProcessLocation >= pstrSource.Length)) {
				plonProcessLocation = plonProcessLocation + 1;
			}
			//set length of FunctionName
			int slonFunctionNameLength = 0;
			slonFunctionNameLength = plonProcessLocation - slonStartFunctionNamePos;

			//FunctionName
//150215            mstrFunctionName = Strings.Mid(pstrSource, slonStartFunctionNamePos, slonFunctionNameLength);
            mstrFunctionName = pstrSource.Substring(slonStartFunctionNamePos, slonFunctionNameLength);
			//        MsgBox("Function Name: " & mstrFunctionName)
			sbolFoundFunction = true;

			//loop through rest of the function and find needed items

//150215			while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len(EndFunctionType)) == EndFunctionType | plonProcessLocation >= Strings.Len(pstrSource))) {
			while (!(pstrSource.Substring(plonProcessLocation, EndFunctionType.Length) == EndFunctionType | plonProcessLocation >= pstrSource.Length)) {
				//if its a comment then skip it
//150215				while (!(Strings.Mid(pstrSource, plonProcessLocation, 1) != "'")) {
				while (!(pstrSource.Substring(plonProcessLocation, 1) != "'")) {
//150215					while (!(Strings.Mid(pstrSource, plonProcessLocation, 2) == Constants.vbCrLf | plonProcessLocation >= Strings.Len(pstrSource))) {
                    while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length))
                    {
						plonProcessLocation = plonProcessLocation + 1;
					}
					plonProcessLocation = plonProcessLocation + 1;//skip over vbcrlf 
				}

//150215                if (Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("DoCmd.OpenForm")) == "DoCmd.OpenForm"){
//                if (pstrSource.Substring(plonProcessLocation, 14) == "DoCmd.OpenForm"){
                if (plonProcessLocation + 14 < pstrSource.Length)
                if (pstrSource.Substring(plonProcessLocation, 14) == "DoCmd.OpenForm"){
                    //**found DoCmd.OpenForm
//150215                    plonProcessLocation = plonProcessLocation + Strings.Len("DoCmd.OpenForm");
                    plonProcessLocation = plonProcessLocation + 14;
                    plonProcessLocation = plonProcessLocation + 1;
                    //step over space after DoCmd.OpenForm
                    //get OpenForm being name
                    string sstrClosingChar = null;
                    //used to find the end of the OpenForm name
                    sstrClosingChar = "";
//150215                    if (Strings.Mid(pstrSource, plonProcessLocation, 1) == "'")
                    if (pstrSource.Substring(plonProcessLocation, 1) == "'")
                    {
                        sstrClosingChar = "'";
                    }
//150215                    else if (Strings.Mid(pstrSource, plonProcessLocation, 1) == "\"")
                    else if (pstrSource.Substring(plonProcessLocation, 1) == "\"")
                    {
                        sstrClosingChar = "\"";
                    }
                    plonProcessLocation = plonProcessLocation + 1;//step over open quote
                    int slonStartOpenFormNamePos = 0;
                    slonStartOpenFormNamePos = plonProcessLocation;
                    //find end of OpenForm name
//150215                    while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len(sstrClosingChar)) == sstrClosingChar | plonProcessLocation >= Strings.Len(pstrSource)))
                    while (!(pstrSource.Substring(plonProcessLocation, sstrClosingChar.Length) == sstrClosingChar | plonProcessLocation >= pstrSource.Length))
                    {
                        plonProcessLocation = plonProcessLocation + 1;
                    }
                    //set length of OpenFormName
                    int slonOpenFormNameLength = 0;
                    slonOpenFormNameLength = plonProcessLocation - slonStartOpenFormNamePos;

                    //was storing unique OpenFormNames but changing to doing that in FunctionsInfo to a variable name OpenFormNamesUnique
                    //If Not mlstOpenFormNames.Contains(LCase(Mid(pstrSource, slonStartOpenFormNamePos, slonOpenFormNameLength))) Then
                    //    mlstOpenFormNames.Add(LCase(Mid(pstrSource, slonStartOpenFormNamePos, slonOpenFormNameLength)))
                    //End If
//150215                    mlstOpenFormNames.Add(Strings.Mid(pstrSource, slonStartOpenFormNamePos, slonOpenFormNameLength));
                    mlstOpenFormNames.Add(pstrSource.Substring(slonStartOpenFormNamePos, slonOpenFormNameLength));
                    mlonNumOpenFormNames = mlonNumOpenFormNames + 1;

                    //break;



                    //want to get SafeSubString working for switch below https://sharpsnippets.wordpress.com/2014/01/25/safe-substring/

				//this is where we build arrays of info we will need later
				//switch (true) {
					//case Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("DoCmd.OpenForm")) == "DoCmd.OpenForm":
						//**found DoCmd.OpenForm
						//plonProcessLocation = plonProcessLocation + Strings.Len("DoCmd.OpenForm");
						//plonProcessLocation = plonProcessLocation + 1;
						//step over space after DoCmd.OpenForm
						//get OpenForm being name
						//string sstrClosingChar = null;
						//used to find the end of the OpenForm name
						//sstrClosingChar = "";
						//if (Strings.Mid(pstrSource, plonProcessLocation, 1) == "'") {
						//	sstrClosingChar = "'";
						//} else if (Strings.Mid(pstrSource, plonProcessLocation, 1) == "\"") {
						//	sstrClosingChar = "\"";
						//}
						//plonProcessLocation = plonProcessLocation + 1;
						//step over open quote
						//int slonStartOpenFormNamePos = 0;
						//slonStartOpenFormNamePos = plonProcessLocation;
						//find end of OpenForm name
						//while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len(sstrClosingChar)) == sstrClosingChar | plonProcessLocation >= Strings.Len(pstrSource))) {
						//	plonProcessLocation = plonProcessLocation + 1;
						//}
						//set length of OpenFormName
						//int slonOpenFormNameLength = 0;
						//slonOpenFormNameLength = plonProcessLocation - slonStartOpenFormNamePos;

						//was storing unique OpenFormNames but changing to doing that in FunctionsInfo to a variable name OpenFormNamesUnique
						//If Not mlstOpenFormNames.Contains(LCase(Mid(pstrSource, slonStartOpenFormNamePos, slonOpenFormNameLength))) Then
						//    mlstOpenFormNames.Add(LCase(Mid(pstrSource, slonStartOpenFormNamePos, slonOpenFormNameLength)))
						//End If
						//mlstOpenFormNames.Add(Strings.Mid(pstrSource, slonStartOpenFormNamePos, slonOpenFormNameLength));
						//mlonNumOpenFormNames = mlonNumOpenFormNames + 1;

						//break;
					//                    MsgBox("OpenFormName: " & LCase(Mid(pstrSource, slonStartOpenFormNamePos, slonOpenFormNameLength)))
				}
				plonProcessLocation = plonProcessLocation + 1;
				//next character
			}
//150215            plonProcessLocation = plonProcessLocation + Strings.Len(EndFunctionType);
            plonProcessLocation = plonProcessLocation + EndFunctionType.Length;
			//step over End Sub or End Function
			return sbolFoundFunction;
		}
	}
}
