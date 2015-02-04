using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace wfavbVBA2PHPone
{
    public partial class Form1 : Form
    {
        private void button1_Click(object sender, EventArgs e)
        {
            List<PageInfo> slstAllPages = new List<PageInfo>();
            List<string> slstAllPageNames = new List<string>();
            List<string> slstUnprocessPageNames = new List<string>();

            PageInfo sclsPageInfo = new PageInfo("C:\\VBA2PHP\\Form_refresh Links.txt");
            bool sbolBackupExistingPHPPages = false;
            sbolBackupExistingPHPPages = false;
            sclsPageInfo.WritePHPPage("C:\\ProIRB_PHP\\", sbolBackupExistingPHPPages);

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
                sclsPageInfo = new PageInfo("C:\\ProIRB_Text\\Form_" + slstUnprocessPageNames[0] + ".txt");
                sclsPageInfo.WritePHPPage("C:\\ProIRB_PHP\\", sbolBackupExistingPHPPages);

                slstUnprocessPageNames.Remove(slstUnprocessPageNames[0]);

//150203                Interaction.MsgBox("done");
                MessageBox.Show("done");
            }
        }
        public Form1()
        {
            InitializeComponent();
        }

        public string sstrPagesOpenForm { get; set; }

    }
}
