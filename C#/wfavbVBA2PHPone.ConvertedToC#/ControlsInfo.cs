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

		public ControlInfo GetControlsInfo(string pstrSource, ref long plonProcessLocation)
		{
			long slonNestDepth = 0;
			slonNestDepth = 0;

			Collection slstAllControls = new Collection();
			Collection scolControlsByName = new Collection();

			ControlInfo sclsControlInfo = null;
			sclsControlInfo = new ControlInfo(pstrSource, ref plonProcessLocation, ref slonNestDepth, ref slstAllControls, scolControlsByName);

			return sclsControlInfo;
		}
	}
}
