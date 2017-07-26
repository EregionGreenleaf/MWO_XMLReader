using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MWO_XMLReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ConfigFile.Initialize();
            List<MechStats> mechList = Worker.LoadQuirks(ConfigFile.DIR_QUIRK);
            Worker.PrintExcel(mechList);
        }
    }
}
