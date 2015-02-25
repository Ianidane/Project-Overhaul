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

        private Collection mcolControlsByName = new Collection();
        public Collection ControlsByName
        {
            get { return mcolControlsByName; }
        }

        private Collection mcolAllControls = new Collection();
        public Collection AllControls
        {
            get { return mcolAllControls; }
        }

        public ControlInfo GetControlsInfo(string pstrSource, ref int plonProcessLocation)
        {
            int slonNestDepth = 0;
            slonNestDepth = 0;

            ControlInfo sclsControlInfo = null;
            sclsControlInfo = new ControlInfo(pstrSource, ref plonProcessLocation, ref slonNestDepth, ref mcolAllControls, ref mcolControlsByName); //not sure slstAllControls is being used

            return sclsControlInfo;
        }
    }
}
