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
	public class FunctionsInfo
	{
		private Collection mlstFunctions = new Collection();
		public Collection Functions {
			get { return mlstFunctions; }
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

		private bool mbolHasFormLoad;
		public bool HasFormLoad {
			get { return mbolHasFormLoad; }
		}
		//Set(ByVal pstrName As String)
		//    mstrName = pstrName
		//End Set
		private bool mbolHasFormOpen;
		public bool HasFormOpen {
			get { return mbolHasFormOpen; }
		}
        private bool mbolHasActive;
        public bool HasFormActive
        {
            get { return mbolHasActive; }
        }
        private bool mbolHasCurrent;
        public bool HasFormCurrent
        {
            get { return mbolHasCurrent; }
        }
		//Set(ByVal pstrName As String)
		//    mstrName = pstrName
		//End Set
		private bool mbolHasFormUnload;
		public bool HasFormUnload {
			get { return mbolHasFormUnload; }
		}
		//Set(ByVal pstrName As String)
		//    mstrName = pstrName
		//End Set
        private bool mbolHasFormClose;
        public bool HasFormClose
        {
            get { return mbolHasFormClose; }
        }

		//150110    Public Function GetFunctionsInfo(pstrSource As String, ByRef plonProcessLocation As Long) As List(Of FunctionInfo)
		public FunctionsInfo(string pstrSource, ref int plonProcessLocation)
		{
			mlonNumOpenFormNames = 0;
			mbolHasFormLoad = false;
			mbolHasFormOpen = false;
            mbolHasActive = false;
            mbolHasCurrent = false;
            mbolHasFormClose = false;
			while (!(plonProcessLocation >= Strings.Len(pstrSource))) {
				//get function info
				FunctionInfo sclsFunctionInfo = null;
				sclsFunctionInfo = new FunctionInfo();
				bool sbolFoundFunction = false;
				sbolFoundFunction = sclsFunctionInfo.PopulateFunctionInfoWithNextFunction(pstrSource, ref plonProcessLocation);
				if (sbolFoundFunction) {
//150215                    mlstFunctions.Add(sclsFunctionInfo, Strings.LCase(sclsFunctionInfo.Name));
                    mlstFunctions.Add(sclsFunctionInfo, sclsFunctionInfo.Name.ToLower());
					foreach (string OpenFormName_loopVariable in sclsFunctionInfo.OpenFormNames) {
						OpenFormName = OpenFormName_loopVariable;
						//moving making Unique just to Forms1 file
						//If Not mlstOpenFormNames.Contains(OpenFormName) Then
						//    mlstOpenFormNames.Add(OpenFormName)
						//End If
						mlstOpenFormNames.Add(OpenFormName);
						mlonNumOpenFormNames = mlonNumOpenFormNames + 1;
					}
					//from http://bytes.com/topic/access/answers/196005-difference-between-form_open-form_load about VBA
					//When you first open a form, the following events occur in this order:Open Þ Load Þ Resize Þ Activate Þ Current
					//If you Then 're trying to decide whether to use the Open or Load event for your macro or event procedure, one significant difference is that the Open event can be canceled, but the Load event can't. For example, if you're dynamically building a record source for a form in an event procedure for the Form 's Open event, you can cancel opening the form if there are no records to display.
					//When you close a form, the following events occur in this order:Unload Þ Deactivate Þ Close
					if (Strings.LCase(sclsFunctionInfo.Name) == "form_open") {
						mbolHasFormOpen = true;
					} else if (sclsFunctionInfo.Name.ToLower() == "form_load") {
						mbolHasFormLoad = true;
					} else if (sclsFunctionInfo.Name.ToLower() == "form_resize") {
						MessageBox.Show("config form_resize - update PageInfo.vb:WritePHPPage as well");
					} else if (sclsFunctionInfo.Name.ToLower() == "form_activate") {
						mbolHasActive = true;
					} else if (sclsFunctionInfo.Name.ToLower() == "form_current") {
						mbolHasCurrent = true;
					} else if (sclsFunctionInfo.Name.ToLower() == "form_unLoad") {
						mbolHasFormUnload = true;
					} else if (sclsFunctionInfo.Name.ToLower() == "form_deactivate") {
						MessageBox.Show("config form_deactivate - update PageInfo.vb:WritePHPPage as well");
					} else if (sclsFunctionInfo.Name.ToLower() == "form_close") {
						mbolHasFormClose = true;
					}
				} else {
					break; // TODO: might not be correct. Was : Exit Do
				}
			}
			return;
		}

        public string OpenFormName { get; set; }
    }
}
