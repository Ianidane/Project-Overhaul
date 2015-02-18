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
	public class ControlsInfo
	{
		public ControlsInfo()
		{
		}

        //150215
        private Collection mcolControlsByName = new Collection();
        public Collection ControlsByName
        {
            get { return mcolControlsByName; }
        }

        //150215
        private Collection mcolAllControls = new Collection();
        public Collection AllControls
        {
            get { return mcolAllControls; }
        }

//150203		public ControlInfo GetControlsInfo(string pstrSource, ref long plonProcessLocation)
		public ControlInfo GetControlsInfo(string pstrSource, ref int plonProcessLocation)
		{
//150203            long slonNestDepth = 0;
            int slonNestDepth = 0; //150203
			slonNestDepth = 0;

//150216			Collection slstAllControls = new Collection();
//150215			Collection scolControlsByName = new Collection();

			ControlInfo sclsControlInfo = null;
//150215            sclsControlInfo = new ControlInfo(pstrSource, ref plonProcessLocation, ref slonNestDepth, ref slstAllControls, ref scolControlsByName); //not sure slstAllControls is being used
            sclsControlInfo = new ControlInfo(pstrSource, ref plonProcessLocation, ref slonNestDepth, ref mcolAllControls, ref mcolControlsByName); //not sure slstAllControls is being used

			return sclsControlInfo;
		}
	}
}
