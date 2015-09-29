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
        private Dictionary<string, List<string>> mdicSubVariableValues4OpenForms = new Dictionary<string,List<string>>();//150315

        private List<string> mlstSubVariables = new List<string>();//150314

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
            string sstrComma = "";
            while (!(pstrSource.Substring(plonProcessLocation, 1) == ")"))
            {
                while (pstrSource.Substring(plonProcessLocation, 1) == " ") plonProcessLocation++;
                if (pstrSource.Substring(plonProcessLocation, 9) == "Optional " | pstrSource.Substring(plonProcessLocation, 11) == System.Environment.NewLine + "Optional") plonProcessLocation += 9;
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
            mstrFunctionTranslated += "){";
            if (pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine) mstrFunctionTranslated += System.Environment.NewLine;

            //loop through rest of the function and find needed items
            while (!(pstrSource.Substring(plonProcessLocation, EndFunctionType.Length) == EndFunctionType | plonProcessLocation >= pstrSource.Length))
            {
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
                        mstrFunctionTranslated += ";" + System.Environment.NewLine;
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
                {
                    mstrFunctionTranslated += " ";
                    plonProcessLocation++;
                }
                else if (pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine)
                {
                    mstrFunctionTranslated += ";" + System.Environment.NewLine;
                    plonProcessLocation += 2;
                }
            }
            int slonWordStartPos = plonProcessLocation;

            //find end of word being space or CrLf
            while (!(pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;

            string sstrWord = pstrSource.Substring(slonWordStartPos, plonProcessLocation - slonWordStartPos);

            if (sstrWord.Substring(0,1) == "'"){ //insert // before ' and then run to end of line - this can't be in the switch below because its checking the first character of sstrWord only
                int sintStartCommentPos = plonProcessLocation;
                while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++; //find CrLf
                plonProcessLocation += 2; //skip over vbcrlf 
                sstrTranslatedWord += "//" +sstrWord + pstrSource.Substring(sintStartCommentPos, plonProcessLocation - sintStartCommentPos);
            }
            else{
                //Reserve words first - and if not one of those then variable or function
                //I'm thinking sort by longer to shorter so doens't find a shorter reserve within a longer one
                //want to get SafeSubString working for switch below https://sharpsnippets.wordpress.com/2014/01/25/safe-substring/
                switch (sstrWord)
                {
    //                case "'": //insert // before ' and then run to end of line
    //                    int sintStartCommentPos = plonProcessLocation;
    //                    while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++; //find CrLf
    //                    plonProcessLocation += 2; //skip over vbcrlf 
    //                    sstrTranslatedWord += "//'" + pstrSource.Substring(sintStartCommentPos, plonProcessLocation - sintStartCommentPos);
    //                    break;

                    case "=": //this should never be called for an = that is within an if, do while, do until... - only for actual assignments, not evaluations
                        sstrTranslatedWord += sstrWord;

                        break;

                    case "Dim":
                        plonProcessLocation++; //step over space after Dim
                        int slonStartOpenVarNamePos = plonProcessLocation;
                        //find end of VarName
                        while (!(pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                        string sstrVarName = pstrSource.Substring(slonStartOpenVarNamePos, plonProcessLocation - slonStartOpenVarNamePos);
                        mlstSubVariables.Add(sstrVarName);
                        //step over 'As'
                        plonProcessLocation++;
                        while (!(pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                        plonProcessLocation++;
                        int slonStartOpenVarTypePos = plonProcessLocation;
                        //find end of VarType
                        while (!(pstrSource.Substring(plonProcessLocation, 1) == " " | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                        string sstrVarType = pstrSource.Substring(slonStartOpenVarTypePos, plonProcessLocation - slonStartOpenVarTypePos);
                        if (sstrVarType == "String")
                        {
                            sstrTranslatedWord += "var " + sstrVarName + " = " + "'';" + System.Environment.NewLine;
                        }
                        else if (sstrVarType == "Long")
                        {
                            sstrTranslatedWord += "var " + sstrVarName + " = " + "0;" + System.Environment.NewLine;
                        }
                        else
                        {
                            sstrTranslatedWord += "//could initialize " + sstrVarName + " here" + System.Environment.NewLine;
                        }
                        while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                        plonProcessLocation += 2; //skip over vbcrlf 
                        break;

                    case "DoCmd.OpenForm":
                        //**found DoCmd.OpenForm
                        //                    plonProcessLocation += 14; //step over "DoCmd.OpenForm"
                        plonProcessLocation++; //step over space after DoCmd.OpenForm
                        if (pstrSource.Substring(plonProcessLocation, 1) == "'" | pstrSource.Substring(plonProcessLocation, 1) == "\"")
                        {
                            //OpenFormName starts with quote so its a constant
                            string sstrClosingChar = ""; //used to find the end of the OpenForm name
                            if (pstrSource.Substring(plonProcessLocation, 1) == "'") sstrClosingChar = "'";
                            else if (pstrSource.Substring(plonProcessLocation, 1) == "\"") sstrClosingChar = "\"";
                            plonProcessLocation++;//step over open quote
                            int slonStartOpenFormNamePos = plonProcessLocation;
                            //find end of OpenForm name
                            while (!(pstrSource.Substring(plonProcessLocation, sstrClosingChar.Length) == sstrClosingChar | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                            mlstOpenFormNames.Add(pstrSource.Substring(slonStartOpenFormNamePos, plonProcessLocation - slonStartOpenFormNamePos));
                            mlonNumOpenFormNames++;
                            sstrTranslatedWord += "location.href = \"Form_" + pstrSource.Substring(slonStartOpenFormNamePos, plonProcessLocation - slonStartOpenFormNamePos) + ".php\";";
                            plonProcessLocation++; //step over sstrClosingChar
                        }
                        else
                        {
                            //OpenformName does not start with quote so its a variable
                            int slonStartOpenFormNamePos = plonProcessLocation;
                            //find end of variable name - guessing either , or crlf
                            while (!(pstrSource.Substring(plonProcessLocation, 1) == "," | pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                            //Below commented out by Alex, because it was throwing an error after 35 pages. Cant figure out why, will come back later.
                            //foreach(string sstrEachOpenform in mdicSubVariableValues4OpenForms[pstrSource.Substring(slonStartOpenFormNamePos, plonProcessLocation - slonStartOpenFormNamePos)])
                            //{
                            //    mlstOpenFormNames.Add(sstrEachOpenform);
                            //    mlonNumOpenFormNames++;
                            //}
                            sstrTranslatedWord += "location.href = \"Form_\"+" + pstrSource.Substring(slonStartOpenFormNamePos, plonProcessLocation - slonStartOpenFormNamePos) + "+\".php\";";
                        }

                        break;


                    default:
                        //quote (constant) - (example of problem: Form_NewMenu.txt Exit_Application_Click())
                        if (sstrWord.Substring(0, 1) == "\"" | sstrWord.Substring(0, 1) == "'") //find close quote and pass that as is
                        {
                            plonProcessLocation -= sstrWord.Length; //need to start at begining of word to find close quote so backing up to do that
                            int sintStartQuote = plonProcessLocation;
                            string sstrCloseQuoteType;
                            if (pstrSource.Substring(plonProcessLocation, 1) == "\"") sstrCloseQuoteType = "\"";
                            else if (pstrSource.Substring(plonProcessLocation, 1) == "'") sstrCloseQuoteType = "'";
                            else
                            {
                                MessageBox.Show("unknown Close Quote Type");
                                sstrCloseQuoteType = "";
                            }
                            plonProcessLocation++; //jump over first quote
                            while (!(pstrSource.Substring(plonProcessLocation, 1) == sstrCloseQuoteType)) plonProcessLocation++;
                            plonProcessLocation++; //include close quote
                            sstrTranslatedWord += pstrSource.Substring(sintStartQuote, plonProcessLocation-sintStartQuote);
                        }

                        //Sub/Function level variables
                        else if (mlstSubVariables.Contains(sstrWord))
                        {
                            sstrTranslatedWord += sstrWord;
                            //OpenForms that use variables need the value of those to add to mlstOpenFormNames

                            //value should be after = - going to use slonProcLoc so as not to mess with plonProcessLocation to do this
                            int slonProcLoc = plonProcessLocation;
                            if (pstrSource.Substring(slonProcLoc, 3) == " = ")
                            {
                                slonProcLoc += 3;
                                //if first character is quote then we are going to save it
                                string sstrCloseQuoteType = "";
                                if (pstrSource.Substring(slonProcLoc, 1) == "\"") sstrCloseQuoteType = "\"";
                                else if (pstrSource.Substring(slonProcLoc, 1) == "'") sstrCloseQuoteType = "'";
                                if (sstrCloseQuoteType != "")
                                {
                                    slonProcLoc++; //step over quote
                                    int sintStartValue = slonProcLoc;
                                    while (!(pstrSource.Substring(slonProcLoc, 1) == sstrCloseQuoteType)) slonProcLoc++;
                                    string sstrValue = pstrSource.Substring(sintStartValue, slonProcLoc - sintStartValue);
                                    slonProcLoc++; //step over quote
                                    if (pstrSource.Substring(slonProcLoc, 2) == System.Environment.NewLine)
                                    {
                                        if (mdicSubVariableValues4OpenForms.ContainsKey(sstrWord))
                                        {
                                            List<string> slstVarValues = mdicSubVariableValues4OpenForms[sstrWord];
                                            slstVarValues.Add(sstrValue);
                                        }
                                        else
                                        {
                                            List<string> slstVarValues = new List<string>();
                                            slstVarValues.Add(sstrValue);
                                            mdicSubVariableValues4OpenForms.Add(sstrWord, slstVarValues);
                                        }
                                    }

                                    //temporary to translate everything after the last quote
                                    int slonVarInfoStartPos = plonProcessLocation;
                                    while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                                    string sstrVarInfo = pstrSource.Substring(slonVarInfoStartPos, plonProcessLocation - slonVarInfoStartPos);
                                    sstrTranslatedWord += sstrVarInfo;
                                  
                                
                                }
                                else //translate everything after the equal sign
                                {
                                    
                                    int slonVarInfoStartPos = plonProcessLocation;
                                    while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length)) plonProcessLocation++;
                                    string sstrVarInfo = pstrSource.Substring(slonVarInfoStartPos, plonProcessLocation - slonVarInfoStartPos);
                                    sstrTranslatedWord += sstrVarInfo;

                                }
                            }

                      
                        }


                        //need else if() for module level variables
                        //need else if() for global level varaibles
                        //need else if() for module level functions
                        //need else if() for global level funcitons

                        else
                        {
                            //write out each word here not processed with // infront to show in output still needs translating
                            sstrTranslatedWord += "//" + sstrWord;
                        }

                        break;
                }
            }
            return sstrTranslatedWord;
        }

    }
}
