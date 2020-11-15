using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excell_okuma_islemleri
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OpenFileDialog op;connect c = new connect();
        private void button1_Click(object sender, EventArgs e)
        {
            op = new OpenFileDialog();
            if (op.ShowDialog()==DialogResult.OK)
            {
                c.xlsxad = op.FileName.ToString();
                c.EBSconnecitonExcellsayfaadi();
                c.excelldata("select * from ["+c.sayfaadi+"]",dataGridView1);
            }
        }
    }
}
