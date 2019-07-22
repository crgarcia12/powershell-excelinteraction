using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelUpdater
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            txtFilePath.Text = @"C:\gitrepos\github\crgarcia12\powershell-excelinteraction\workbook.xlsx";
        }

        private void btnOpenAndTransform_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelUpdater updater = new ExcelUpdater();
                updater.UpdateExcel(txtFilePath.Text);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
