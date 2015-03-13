//using Microsoft.VisualBasic;
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

//can't get SafeSubstring working
//namespace wfavbVBA2PHPone
//{
//        public static string SafeSubstring(this string value, int startIndex, int length)
//        {
//            return new string((value ?? string.Empty).Skip(startIndex).Take(length).ToArray());
//        }
//}

namespace wfavbVBA2PHPone
{
    public class FunctionInfo
    {
        private string mstrFunctionName;
        public string Name
        {
            get { return mstrFunctionName; }
        }

        private List<string> mlstOpenFormNames = new List<string>();
        public List<string> OpenFormNames
        {
            get { return mlstOpenFormNames; }
        }

        private long mlonNumOpenFormNames;
        public long NumOpenFormNames
        {
            get { return mlonNumOpenFormNames; }
        }

        private string mstrFunctionTranslated = "";
        public string FunctionTranslated
        {
            get { return mstrFunctionTranslated; }
        }


        public FunctionInfo()
        {
            mlonNumOpenFormNames = 0;
        }

        public bool PopulateFunctionInfoWithNextFunction(string pstrSource, ref int plonProcessLocation)
        {
            bool sbolFoundFunction = false;

            if (!(plonProcessLocation + 8 < pstrSource.Length)) //would prefer SafeSubString
                return sbolFoundFunction;

            //note - you'll notice most "Do Until"s will have "Or plonProcessLocation >= Len(pstrSource)" in the condition for stopping the looping.  This is to be sure if we unexpectedly find the end of the file that we don't error.  Probably would be a better idea to actually throw an error if that happened rather than just falling out of the function
            //find Sub or Function
            while (!(pstrSource.Substring(plonProcessLocation, 3) == "Sub" | pstrSource.Substring(plonProcessLocation, 8) == "Function" | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
            //used to find the end of the Sub or Function
            string EndFunctionType = "";
            if (pstrSource.Substring(plonProcessLocation, 3) == "Sub")
            {
                EndFunctionType = "End Sub";
                plonProcessLocation += 3;//jump over the word "Sub"
            }
            else if (pstrSource.Substring(plonProcessLocation, 8) == "Function")
            {
                EndFunctionType = "End Function";
                plonProcessLocation += 8;//jump over the word Function
            }
            else MessageBox.Show("failed to find function type");
            plonProcessLocation++;//step over space after Sub or Function

            //set start FunctionName
            int slonStartFunctionNamePos = plonProcessLocation;
            //find end of function name - pretty sure that function always has ( after it
            while (!(pstrSource.Substring(plonProcessLocation, 1) == "(" | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
            //set length of FunctionName
            int slonFunctionNameLength = plonProcessLocation - slonStartFunctionNamePos;

            //FunctionName
            mstrFunctionName = pstrSource.Substring(slonStartFunctionNamePos, slonFunctionNameLength);
            sbolFoundFunction = true;

            plonProcessLocation++; //step over "("

            mstrFunctionTranslated += "function " + mstrFunctionName + "(";
            if (mstrFunctionName == "QuickStudyNumber_KeyDown")
    MessageBox.Show("hit");
            string sstrComma = "";
            while (!(pstrSource.Substring(plonProcessLocation, 1) == ")"))
            {
                int sintStartParmPos = plonProcessLocation;
                while (!(pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 1) == "," | pstrSource.Substring(plonProcessLocation, 1) == ")" | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++; //find space, "," or ) - end of Parm Name
                mstrFunctionTranslated += sstrComma + pstrSource.Substring(sintStartParmPos, plonProcessLocation - sintStartParmPos);
                sstrComma = ", ";
                if (pstrSource.Substring(plonProcessLocation, 4) == " As ")
                {
                    plonProcessLocation += 4;
                    while (!(pstrSource.Substring(plonProcessLocation, 1) == "," | pstrSource.Substring(plonProcessLocation, 1) == ")" | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++; //find , or ) - end of Parm Type
                    if (pstrSource.Substring(plonProcessLocation, 1) == ",") plonProcessLocation++; //step over ","
                }
                else
                    MessageBox.Show("Parameter without Type");
            }
            plonProcessLocation++; //step over ")"
            mstrFunctionTranslated += "){" + System.Environment.NewLine;

            //loop through rest of the function and find needed items
            while (!(pstrSource.Substring(plonProcessLocation, EndFunctionType.Length) == EndFunctionType | plonProcessLocation >= pstrSource.Length))
            {
                //150218 handle comments in TranslateNextWord
                //				//if its a comment then skip it
                //				while (!(pstrSource.Substring(plonProcessLocation, 1) != "'")) {
                //                    while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                //					plonProcessLocation++;//skip over vbcrlf 
                //				}

                //if (plonProcessLocation + 14 < pstrSource.Length) //would like SafeSubString working - but till it does need to make sure that there is enough string left to check 14 characters
                //if (pstrSource.Substring(plonProcessLocation, 14) == "DoCmd.OpenForm"){
                //    //**found DoCmd.OpenForm
                //    plonProcessLocation = plonProcessLocation + 14; //step over "DoCmd.OpenForm"
                //    plonProcessLocation = plonProcessLocation + 1; //step over space after DoCmd.OpenForm
                //    //get OpenForm being name
                //    //used to find the end of the OpenForm name
                //    string sstrClosingChar = "";
                //    if (pstrSource.Substring(plonProcessLocation, 1) == "'")
                //    {
                //        sstrClosingChar = "'";
                //    }
                //    else if (pstrSource.Substring(plonProcessLocation, 1) == "\"")
                //    {
                //        sstrClosingChar = "\"";
                //    }
                //    plonProcessLocation = plonProcessLocation + 1;//step over open quote
                //    int slonStartOpenFormNamePos = plonProcessLocation;
                //    //find end of OpenForm name
                //    while (!(pstrSource.Substring(plonProcessLocation, sstrClosingChar.Length) == sstrClosingChar | plonProcessLocation >= pstrSource.Length))
                //    {
                //        plonProcessLocation = plonProcessLocation + 1;
                //    }
                //    //set length of OpenFormName
                //    int slonOpenFormNameLength = plonProcessLocation - slonStartOpenFormNamePos;

                //    mlstOpenFormNames.Add(pstrSource.Substring(slonStartOpenFormNamePos, slonOpenFormNameLength));
                //    mlonNumOpenFormNames = mlonNumOpenFormNames + 1;



                //I'm think this is where the bulk of the coding need to be done
                //This if I think need to be replaced with a check for each word of code in sequence and translate it from VBA to PHP



                //}
                //plonProcessLocation = plonProcessLocation + 1; //next character



                mstrFunctionTranslated += TranslateNextWord(pstrSource, ref plonProcessLocation);

                //skip over any spaces or CrLf - also put any space or CrLf into mstrFunctionTranslated
                while ((pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine))
                {
                    if (pstrSource.Substring(plonProcessLocation, 1) == " ")
                    {
                        mstrFunctionTranslated += " ";
                        plonProcessLocation++;
                    }
                    else if (pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine)
                    {
                        mstrFunctionTranslated += System.Environment.NewLine;
                        plonProcessLocation += 2;
                    }
                }
            }


            mstrFunctionTranslated += "}" + System.Environment.NewLine; //close function


            plonProcessLocation = plonProcessLocation + EndFunctionType.Length; //step over End Sub or End Function
            return sbolFoundFunction;
        }

        public string TranslateNextWord(string pstrSource, ref int plonProcessLocation)
        {
            string sstrTranslatedWord = "";

//            //find next word
//            if (!(plonProcessLocation + 14 < pstrSource.Length)) //would prefer SafeSubString - 14 is longest Reserve word so far
//                return "";

            //skip any spaces or CrLf
            while (pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine)
            {
                if (pstrSource.Substring(plonProcessLocation, 1) == " ")
                    plonProcessLocation++;
                else if (pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine)
                    plonProcessLocation += 2;
            }
            int slonWordStartPos = plonProcessLocation;

            //find end of word being space or CrLf
            while (!(pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;

            string sstrWord = pstrSource.Substring(slonWordStartPos, plonProcessLocation - slonWordStartPos);

            //Reserve words first - and if not one of those then variable or function
            //I'm thinking sort by longer to shorter so doens't find a shorter reserve within a longer one
            //want to get SafeSubString working for switch below https://sharpsnippets.wordpress.com/2014/01/25/safe-substring/
            switch (sstrWord)
            {
                case "'": //insert // before ' and then run to end of line
                    int sintStartCommentPos = plonProcessLocation;
                    while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++; //find CrLf
                    plonProcessLocation += 2; //skip over vbcrlf 
                    sstrTranslatedWord += "//'" + pstrSource.Substring(sintStartCommentPos, plonProcessLocation - sintStartCommentPos);
                    break;

                case "Dim":

                    break;

                case "DoCmd.OpenForm":
                    //**found DoCmd.OpenForm
                    //                    plonProcessLocation += 14; //step over "DoCmd.OpenForm"
                    plonProcessLocation++; //step over space after DoCmd.OpenForm
                    string sstrClosingChar = ""; //used to find the end of the OpenForm name
                    if (pstrSource.Substring(plonProcessLocation, 1) == "'") sstrClosingChar = "'";
                    else if (pstrSource.Substring(plonProcessLocation, 1) == "\"") sstrClosingChar = "\"";
                    plonProcessLocation++;//step over open quote
                    int slonStartOpenFormNamePos = plonProcessLocation;
                    //find end of OpenForm name
                    while (!(pstrSource.Substring(plonProcessLocation, sstrClosingChar.Length) == sstrClosingChar | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                    mlstOpenFormNames.Add(pstrSource.Substring(slonStartOpenFormNamePos, plonProcessLocation - slonStartOpenFormNamePos));
                    mlonNumOpenFormNames++;
                    sstrTranslatedWord += "        location.href = \"Form_" + pstrSource.Substring(slonStartOpenFormNamePos, plonProcessLocation - slonStartOpenFormNamePos) + ".php\";" + System.Environment.NewLine;
                    plonProcessLocation++; //step over sstrClosingChar

                    break;

                default:
                    //must be Variable or Function
                    //NEED VARIABLE AND FUNCTION CODE HERE
                    break;
            }
            return sstrTranslatedWord;
        }

    }
}
