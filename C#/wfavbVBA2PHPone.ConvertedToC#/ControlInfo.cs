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
		//Public ReadOnly Property BeginType() As String
		//    Get
		//        Return mstrBeginType
		//    End Get
		//    'Set(ByVal pstrBeginType As String)
		//    '    mstrBeginType = pstrBeginType
		//    'End Set
		//End Property

		private List<ControlInfo> mclsSubControls = new List<ControlInfo>();
		//Public ReadOnly Property SubControls() As List(Of ControlInfo)
		//    Get
		//        Return mclsSubControls
		//    End Get
		//    'Set(ByVal pclsSubControl As Long)
		//    '    mlonNestDepth = plonNestDepth
		//    'End Set
		//End Property
		public void AddSubControl(ControlInfo pclsSubControl)
		{
			mclsSubControls.Add(pclsSubControl);
		}

		private int mlonControlLinesStartPos;
		//Public ReadOnly Property ControlLinesStartPos() As Long
		//    Get
		//        Return mlonControlLinesStartPos
		//    End Get
		//    'Set(ByVal plonControlLinesStartPos As Long)
		//    '    mlonControlLinesStartPos = plonControlLinesStartPos
		//    'End Set
		//End Property


		private List<Attribute> mlstAttributes = new List<Attribute>();
		private string mstrControlLines;
        private string pstrSource;
        private long plonProcessLocation;
        private long slonNestDepth;
        private Collection slstAllControls;
        private Collection scolControlsByName;
		//Public Property ControlLines() As String
		//    Get
		//        Return mstrControlLines
		//    End Get
		//    Set(ByVal pstrControlLines As String)
		//        mstrControlLines = pstrControlLines
		//    End Set
		//End Property

		//150108    Public Sub New(pstrSource As String, ByRef plonProcessLocation As Long, ByRef plonNestDepth As Long, ByRef plstAllControls As Collection)
		//don't think plonNestDepth is really needed - pcolControlsByName just parameter so far, code to set not done yet
		public ControlInfo(string pstrSource, ref int plonProcessLocation, ref int plonNestDepth, ref Collection plstAllControls, ref Collection pcolControlsByName)
		{
			//note - you'll notice most "Do Until"s will have "Or plonProcessLocation >= Len(pstrSource)" in the condition for stopping the looping.  This is to be sure if we unexpectedly find the end of the file that we don't error.  Probably would be a better idea to actually throw an error if that happened rather than just falling out of the function
			while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("Begin")) == "Begin" | plonProcessLocation >= Strings.Len(pstrSource))) {
				plonProcessLocation = plonProcessLocation + 1;
			}

			plonProcessLocation = plonProcessLocation + Strings.Len("Begin");
			//step over Begin
			plonNestDepth = plonNestDepth + 1;

			//find Begin type (if it has one)
			if (Strings.Mid(pstrSource, plonProcessLocation, 2) != Constants.vbCrLf) {
				//has Begin type
				plonProcessLocation = plonProcessLocation + 1;
				//step over space after Begin

				//set start Begin type
				int slonStartBeginTypePos = 0;
				slonStartBeginTypePos = plonProcessLocation;
				//find end of Begin type - pretty sure that function always is vbCrLf
				while (!(Strings.Mid(pstrSource, plonProcessLocation, 2) == Constants.vbCrLf | plonProcessLocation >= Strings.Len(pstrSource))) {
					plonProcessLocation = plonProcessLocation + 1;
				}
				//set length of Begin type
				int slonBeginTypeLength = 0;
				slonBeginTypeLength = plonProcessLocation - slonStartBeginTypePos;

				//Begin type
				mstrBeginType = Strings.Mid(pstrSource, slonStartBeginTypePos, slonBeginTypeLength);
			}
			plonProcessLocation = plonProcessLocation + 2;
			//step over vbCRLf
			mlonControlLinesStartPos = plonProcessLocation;

			while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("End")) == "End")) {
				while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len(" ")) != " " | plonProcessLocation >= Strings.Len(pstrSource))) {
					plonProcessLocation = plonProcessLocation + 1;
				}

				if (Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("Begin")) == "Begin") {
					ControlInfo sclsControlInfo = null;
					sclsControlInfo = new ControlInfo(pstrSource, ref plonProcessLocation, ref plonNestDepth, ref plstAllControls,ref pcolControlsByName);
					mclsSubControls.Add(sclsControlInfo);
				} else if (Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("End")) == "End") {
					//do nothing - will kick out of Do Until loop
				} else {
					//its an Attribute
					Attribute sstcAttr = default(Attribute);
					sstcAttr = new Attribute();

					//AttrName
					int slonStartAttrNamePos = 0;
					slonStartAttrNamePos = plonProcessLocation;
					while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len(" ")) == " " | plonProcessLocation >= Strings.Len(pstrSource))) {
						plonProcessLocation = plonProcessLocation + 1;
					}
					sstcAttr.Name = Strings.Mid(pstrSource, slonStartAttrNamePos, plonProcessLocation - slonStartAttrNamePos);

					while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("=")) == "=" | plonProcessLocation >= Strings.Len(pstrSource))) {
						plonProcessLocation = plonProcessLocation + 1;
					}
					plonProcessLocation = plonProcessLocation + Strings.Len("=");
					//step over =

					//AttrValue
					int slonStartAttrValuePos = 0;
					slonStartAttrValuePos = plonProcessLocation;
					while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len(Constants.vbCrLf)) == Constants.vbCrLf | plonProcessLocation >= Strings.Len(pstrSource))) {
						plonProcessLocation = plonProcessLocation + 1;
					}
					sstcAttr.Value = Strings.Mid(pstrSource, slonStartAttrValuePos, plonProcessLocation - slonStartAttrValuePos);


					if (sstcAttr.Name == "Name") {
						pcolControlsByName.Add(this, sstcAttr.Value);
					}


					//there are Attributes (mostly bits, like GUID, PictureDate,...) that we don't care about - form is GUID = BEGIN .... END
					if (sstcAttr.Value == " Begin") {
						while (!(Strings.Mid(pstrSource, plonProcessLocation, Strings.Len("End")) == "End" | plonProcessLocation >= Strings.Len(pstrSource))) {
							plonProcessLocation = plonProcessLocation + 1;
						}
						plonProcessLocation = plonProcessLocation + Strings.Len("End");
						//step over End
					} else {
						mlstAttributes.Add(sstcAttr);
					}
					plonProcessLocation = plonProcessLocation + Strings.Len(Constants.vbCrLf);
					//step over vbCrLf
				}
			}

			plonNestDepth = plonNestDepth - 1;
			mstrControlLines = Strings.Mid(pstrSource, mlonControlLinesStartPos, plonProcessLocation - mlonControlLinesStartPos);
			plonProcessLocation = plonProcessLocation + Strings.Len("End");
			//step over End
			plonProcessLocation = plonProcessLocation + Strings.Len(Constants.vbCrLf);
			//step over vbCrLf
		}

        public ControlInfo(string pstrSource, ref long plonProcessLocation, ref long slonNestDepth, ref Collection slstAllControls, Collection scolControlsByName)
        {
            // TODO: Complete member initialization
            this.pstrSource = pstrSource;
            this.plonProcessLocation = plonProcessLocation;
            this.slonNestDepth = slonNestDepth;
            this.slstAllControls = slstAllControls;
            this.scolControlsByName = scolControlsByName;
        }
	}
}
