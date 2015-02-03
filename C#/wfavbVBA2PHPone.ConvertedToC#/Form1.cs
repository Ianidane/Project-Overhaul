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
	public partial class Form1
	{
		//there are a bunch of line numbers in these functions - these are just hold overs from the VB6 program
        private void Button1_Click(object sender, EventArgs e)
        {
            List<PageInfo> slstAllPages = new List<PageInfo>();
            List<string> slstAllPageNames = new List<string>();
            List<string> slstUnprocessPageNames = new List<string>();

            PageInfo sclsPageInfo = new PageInfo("C:\\Users\\Ian\\Documents\\VBA2PHP\\Form_refresh Links.txt");
            bool sbolBackupExistingPHPPages = false;
            sbolBackupExistingPHPPages = false;
            sclsPageInfo.WritePHPPage("C:\\Users\\Ian\\Documents\\ProIRB_PHP\\", sbolBackupExistingPHPPages);

            slstAllPages.Add(sclsPageInfo);
            slstAllPageNames.Add(sclsPageInfo.Name);
            foreach (string sstrPagesOpenForm_loopVariable in sclsPageInfo.OpenFormNames)
            {
                sstrPagesOpenForm = sstrPagesOpenForm_loopVariable;
                if (!slstAllPageNames.Contains(sstrPagesOpenForm, StringComparer.OrdinalIgnoreCase))
                {
                    slstUnprocessPageNames.Add(sstrPagesOpenForm);
                }
            }

            //enable code below to process more than just first page
            //only process 5 pages for now
            int slonTempPagesProcess = 0;
            while (slstUnprocessPageNames.Count > 0 & slonTempPagesProcess < 5)
            {
                sclsPageInfo = new PageInfo("C:\\Users\\Ian\\Documents\\ProIRB_Text\\Form_" + slstUnprocessPageNames[0] + ".txt");
                sclsPageInfo.WritePHPPage("C:\\Users\\Ian\\Documents\\ProIRB_PHP\\", sbolBackupExistingPHPPages);

                slstUnprocessPageNames.Remove(slstUnprocessPageNames[0]);

                Interaction.MsgBox("done");
            }
        }
		public Form1()
		{
			InitializeComponent();
		}


        public string sstrPagesOpenForm { get; set; }
    }
}

