﻿using System;
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
                btnOpenAndTransform.Enabled = false;

                int initialRow;
                int copyUntilMaxPlusRows;

                Int32.TryParse(txtInitialRowNr.Text, out initialRow);
                Int32.TryParse(txtCopyUntil.Text, out copyUntilMaxPlusRows);

                initialRow = initialRow == 0 ? 1 : initialRow; // Default InitialRow = 1
                copyUntilMaxPlusRows = copyUntilMaxPlusRows < 0 ? -1 : copyUntilMaxPlusRows; // Default copyUntilMaxPlusRows = -1

                ExcelUpdater updater = new ExcelUpdater();
                updater.ProcessExcelFile(txtFilePath.Text, initialRow, copyUntilMaxPlusRows);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnOpenAndTransform.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            int initialRow;
            int copyUntilMaxPlusRows;

            Int32.TryParse(txtInitialRowNr.Text, out initialRow);
            Int32.TryParse(txtCopyUntil.Text, out copyUntilMaxPlusRows);

            initialRow = initialRow == 0 ? 1 : initialRow; // Default InitialRow = 1
            copyUntilMaxPlusRows = copyUntilMaxPlusRows < 0 ? -1 : copyUntilMaxPlusRows; // Default copyUntilMaxPlusRows = -1

            ExcelUpdater updater = new ExcelUpdater();
            updater.Play(txtFilePath.Text, initialRow, copyUntilMaxPlusRows);
        }
    }
}
