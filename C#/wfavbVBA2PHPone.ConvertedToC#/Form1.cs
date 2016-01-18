using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronPython.Hosting;

namespace wfavbVBA2PHPone
{

    public partial class Form1 : Form
    {

        private void button1_Click(object sender, EventArgs e)
        {
            List<PageInfo> slstAllPages = new List<PageInfo>();
            List<string> slstAllPageNames = new List<string>();
            List<string> slstUnprocessPageNames = new List<string>();

            PageInfo sclsPageInfo = new PageInfo("C:\\VBA2PHP\\ProIRB_Text\\Form_refresh Links.txt");
            bool sbolBackupExistingPHPPages = false;
            sbolBackupExistingPHPPages = false;
            sclsPageInfo.WritePHPPage("C:\\VBA2PHP\\ProIRB_PHP\\", sbolBackupExistingPHPPages);

            slstAllPages.Add(sclsPageInfo);
            slstAllPageNames.Add(sclsPageInfo.Name);
            //            foreach (string sstrPagesOpenForm_loopVariable in sclsPageInfo.OpenFormNames)
            foreach (string sstrEachPagesOpenForm in sclsPageInfo.OpenFormNames)
            {
                //                sstrPagesOpenForm = sstrPagesOpenForm_loopVariable;
                //                if (!slstAllPageNames.Contains(sstrPagesOpenForm, StringComparer.OrdinalIgnoreCase))
                //                if (!slstAllPageNames.Contains(sstrEachPagesOpenForm, StringComparer.OrdinalIgnoreCase))
                //                if (!slstAllPageNames.Contains("Form_" + sstrEachPagesOpenForm, StringComparer.OrdinalIgnoreCase))
                if (!slstAllPageNames.Contains("Form_" + sstrEachPagesOpenForm + ".txt", StringComparer.OrdinalIgnoreCase))
                {
                    //                    slstUnprocessPageNames.Add(sstrPagesOpenForm);
                    //                    slstUnprocessPageNames.Add(sstrEachPagesOpenForm);
                    //                    slstUnprocessPageNames.Add("Form_" + sstrEachPagesOpenForm);
                    slstUnprocessPageNames.Add("Form_" + sstrEachPagesOpenForm + ".txt");
                }
            }

            //only process 20 pages for now
            int slonTempPagesProcess = 0;
            while (slstUnprocessPageNames.Count > 0 & slonTempPagesProcess < 9999)
            {
                //                sclsPageInfo = new PageInfo("C:\\VBA2PHP\\ProIRB_Text\\Form_" + slstUnprocessPageNames[0] + ".txt");
                //                sclsPageInfo = new PageInfo("C:\\VBA2PHP\\ProIRB_Text\\" + slstUnprocessPageNames[0] + ".txt");
                sclsPageInfo = new PageInfo("C:\\VBA2PHP\\ProIRB_Text\\" + slstUnprocessPageNames[0]);
                sclsPageInfo.WritePHPPage("C:\\VBA2PHP\\ProIRB_PHP\\", sbolBackupExistingPHPPages);

                slstUnprocessPageNames.Remove(slstUnprocessPageNames[0]);

                slstAllPages.Add(sclsPageInfo);
                slstAllPageNames.Add(sclsPageInfo.Name);
                //                foreach (string sstrPagesOpenForm_loopVariable in sclsPageInfo.OpenFormNames)
                foreach (string sstrEachPagesOpenForm in sclsPageInfo.OpenFormNames)
                {
                    //                    sstrPagesOpenForm = sstrPagesOpenForm_loopVariable;
                    //                    if (!slstAllPageNames.Contains(sstrPagesOpenForm, StringComparer.OrdinalIgnoreCase))
                    //150218                    if (!slstAllPageNames.Contains(sstrEachPagesOpenForm, StringComparer.OrdinalIgnoreCase))
                    //                    if (!slstAllPageNames.Contains("Form_" + sstrEachPagesOpenForm, StringComparer.OrdinalIgnoreCase))
                    if (!slstAllPageNames.Contains("Form_" + sstrEachPagesOpenForm + ".txt", StringComparer.OrdinalIgnoreCase))
                    {
                        //                        slstUnprocessPageNames.Add(sstrPagesOpenForm);
                        //                        slstUnprocessPageNames.Add(sstrEachPagesOpenForm);
                        if (!slstUnprocessPageNames.Contains("Form_" + sstrEachPagesOpenForm + ".txt", StringComparer.OrdinalIgnoreCase))
                        {
                            if (sstrEachPagesOpenForm == "")
                                //MessageBox.Show("skipping blank filename - need to get local variables working");
                                Console.Write("skipping blank filename - need to get local variables working");
                            else
                                slstUnprocessPageNames.Add("Form_" + sstrEachPagesOpenForm + ".txt");
                        }
                    }
                }
                slonTempPagesProcess = slonTempPagesProcess + 1;
            }


            //All this python stuff is a test/example
            var py = Python.CreateEngine();                                 //create python engine
            var paths = py.GetSearchPaths();                                //used to inlude python libraries
            var scope = py.CreateScope();                                   //allows us to exchange information between C# and Python script
            paths.Add(@"C:\Python27\Lib");                                  //add to include python libraries
            py.SetSearchPaths(paths);                                       //includes them
            scope.SetVariable("input1", "Yaaay");                           //sets a variable in the scope to "yaaay" this variable can then be used in a python script
            py.ExecuteFile("test.py", scope);                               //runs test.py using the variables in the scope like "yaaay"
            py.CreateScriptSourceFromFile("test.py").Execute(scope);        //get information to use .getvariable
            var result = scope.GetVariable("input2");                       //get specific variable


            MessageBox.Show("done");
        }
        public Form1()
        {
            InitializeComponent();
        }

        //        public string sstrPagesOpenForm { get; set; }

    }
}
