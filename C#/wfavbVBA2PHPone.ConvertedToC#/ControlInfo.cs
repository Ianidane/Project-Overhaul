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
	public struct Attribute
	{
		public string Name;
		public string Value;
	}
}
namespace wfavbVBA2PHPone
{

	public class ControlInfo
	{
		private string mstrBeginType;
        public string BeginType
        {
            get { return mstrBeginType; }
        }
        
		private List<ControlInfo> mclsSubControls = new List<ControlInfo>(); //pretty sure not used yet - not sure if will be needed
		public void AddSubControl(ControlInfo pclsSubControl)
		{
			mclsSubControls.Add(pclsSubControl);
		}

		private int mlonControlLinesStartPos; //pretty sure not used yet - not sure if will be needed

        private string mstrName;
        public string Name
        {
            get { return mstrName; }
        }
        private string mstrOnClick;
        public string OnClick
        {
            get { return mstrOnClick; }
        }
        private string mstrCaption;
        public string Caption
        {
            get { return mstrCaption; }
        }
        private int mintLeft;
        public int Left
        {
            get { return mintLeft; }
        }
        private int mintTop;
        public int Top
        {
            get { return mintTop; }
        }
        private int mintWidth;
        public int Width
        {
            get { return mintWidth; }
        }
        private int mintHeight;
        public int Height
        {
            get { return mintHeight; }
        }
        private string mstrVisible;
        public string Visible
        {
            get { return mstrVisible; }
        }
        private string mstrDefaultValue;
        public string DefaultValue
        {
            get { return mstrDefaultValue; }
        }
        private int mintSpecialEffect;
        public int SpecialEffect
        {
            get { return mintSpecialEffect; }
        }

        //Ian,Alex - might need to add more properties here from Controls as needed - (was hoping mlstAttributes would give all attributes like these as needed but can't get working)
            //Ian,Alex will need to add code to set these properties below where commented

		private List<Attribute> mlstAttributes = new List<Attribute>(); //was going to have PageInfo.cs access mlstAttributes to get to Name, Caption, Left, Top... - but can't get to work like collection in VB so instead making the properties Name, Caption, Left, Top... and setting those
		private string mstrControlLines;

		//don't think plonNestDepth is really needed - pcolControlsByName just parameter so far, code to set not done yet
        //pcolControlsByName is being populated, but so far not used - I thought it would be needed
		public ControlInfo(string pstrSource, ref int plonProcessLocation, ref int plonNestDepth, ref Collection plstAllControls, ref Collection pcolControlsByName)
		{
            plstAllControls.Add(this);

			//note - you'll notice most "Do Until"s will have "Or plonProcessLocation >= Len(pstrSource)" in the condition for stopping the looping.  This is to be sure if we unexpectedly find the end of the file that we don't error.  Probably would be a better idea to actually throw an error if that happened rather than just falling out of the function
			while (!(pstrSource.Substring(plonProcessLocation, 5) == "Begin" | plonProcessLocation >= pstrSource.Length)) {
				plonProcessLocation = plonProcessLocation + 1;
			}

			plonProcessLocation = plonProcessLocation + 5;//step over "Begin"
			plonNestDepth = plonNestDepth + 1;

            if(pstrSource.Substring(plonProcessLocation, 1) == " ") //if there is a begin type then there will be a space
                plonProcessLocation++;

			//find Begin type (if it has one)
            if(pstrSource.Substring(plonProcessLocation, 2) != System.Environment.NewLine)
            {
				//has Begin type

				//set start Begin type
				int slonStartBeginTypePos = 0;
				slonStartBeginTypePos = plonProcessLocation;
				//find end of Begin type - pretty sure that function always is vbCrLf
                while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length))
                {
					plonProcessLocation = plonProcessLocation + 1;
				}
				//set length of Begin type
				int slonBeginTypeLength = 0;
				slonBeginTypeLength = plonProcessLocation - slonStartBeginTypePos;

				//Begin type
                mstrBeginType = pstrSource.Substring(slonStartBeginTypePos, slonBeginTypeLength);
			}
			plonProcessLocation = plonProcessLocation + 2;//step over vbCRLf
			mlonControlLinesStartPos = plonProcessLocation;

			while (!(pstrSource.Substring(plonProcessLocation,3) == "End")) { //run loop until finding "End"
				while (!(pstrSource.Substring(plonProcessLocation, 1) != " " | plonProcessLocation >= pstrSource.Length)) {
					plonProcessLocation = plonProcessLocation + 1;
				}

				if (pstrSource.Substring(plonProcessLocation, 5) == "Begin") {
					ControlInfo sclsControlInfo = null;
					sclsControlInfo = new ControlInfo(pstrSource, ref plonProcessLocation, ref plonNestDepth, ref plstAllControls,ref pcolControlsByName);
					mclsSubControls.Add(sclsControlInfo);
				} else if (pstrSource.Substring(plonProcessLocation, 3) == "End") {
					//do nothing - will kick out of Do Until loop
				} else {
					//its an Attribute
					Attribute sstcAttr = default(Attribute);
					sstcAttr = new Attribute();

					//AttrName
					int slonStartAttrNamePos = 0;
					slonStartAttrNamePos = plonProcessLocation;
					while (!(pstrSource.Substring(plonProcessLocation, 1) == " " | plonProcessLocation >= pstrSource.Length)) { //find space - there is always a space after attr name
						plonProcessLocation = plonProcessLocation + 1;
					}
                    sstcAttr.Name = pstrSource.Substring(slonStartAttrNamePos, plonProcessLocation - slonStartAttrNamePos);

					while (!(pstrSource.Substring(plonProcessLocation, 1) == "=" | plonProcessLocation >= pstrSource.Length)) {
						plonProcessLocation = plonProcessLocation + 1;
					}
					plonProcessLocation = plonProcessLocation + 1; //step over "="

					//AttrValue
					int slonStartAttrValuePos = 0;
					slonStartAttrValuePos = plonProcessLocation;
                    while (!(pstrSource.Substring(plonProcessLocation, 2) == System.Environment.NewLine | plonProcessLocation >= pstrSource.Length))
                    {
						plonProcessLocation = plonProcessLocation + 1;
					}
                    sstcAttr.Value = pstrSource.Substring(slonStartAttrValuePos, plonProcessLocation - slonStartAttrValuePos);

                    if (sstcAttr.Name == "Name") mstrName = sstcAttr.Value; //150216
                    if (sstcAttr.Name == "OnClick") mstrOnClick = sstcAttr.Value; 
                    if (sstcAttr.Name == "Caption") mstrCaption = sstcAttr.Value;
                    if (sstcAttr.Name == "Visible") mstrVisible = sstcAttr.Value;
                    if (sstcAttr.Name == "DefaultValue") mstrDefaultValue = sstcAttr.Value; 
                    if (sstcAttr.Name == "Left") mintLeft = Convert.ToInt32(sstcAttr.Value);
                    if (sstcAttr.Name == "Top") mintTop = Convert.ToInt32(sstcAttr.Value);
                    if (sstcAttr.Name == "Width") mintWidth = Convert.ToInt32(sstcAttr.Value);
                    if (sstcAttr.Name == "Height") mintHeight = Convert.ToInt32(sstcAttr.Value);
                    if (sstcAttr.Name == "SpecialEffect") mintSpecialEffect = Convert.ToInt32(sstcAttr.Value);
                    //Ian,Alex - this is where you'd set any new properties you've added, after adding above where comment indicates

                    //Translate vauge attribute values to correct HTML values
                    if (mstrVisible == " NotDefault")
                    {
                        mstrVisible = "hidden";
                    }

					if (sstcAttr.Name == "Name") {
						pcolControlsByName.Add(this, sstcAttr.Value); //not used yet - I thought it would be needed but so far no
					}


					//there are Attributes (mostly bits, like GUID, PictureDate,...) that we don't care about - form is GUID = BEGIN .... END
					if (sstcAttr.Value == " Begin") {
						while (!(pstrSource.Substring(plonProcessLocation, 3) == "End" | plonProcessLocation >= pstrSource.Length)) {
							plonProcessLocation = plonProcessLocation + 1;
						}
						plonProcessLocation = plonProcessLocation + 3; //step over "End"
					} else {
                        mlstAttributes.Add(sstcAttr);
					}
                    plonProcessLocation = plonProcessLocation + 2; //step over vbCrLf
				}
			}

			plonNestDepth = plonNestDepth - 1;
            mstrControlLines = pstrSource.Substring(mlonControlLinesStartPos, plonProcessLocation - mlonControlLinesStartPos);
            plonProcessLocation = plonProcessLocation + 3; //step over "End"
			plonProcessLocation = plonProcessLocation + 2;//step over vbCrLf
		}

	}
}
