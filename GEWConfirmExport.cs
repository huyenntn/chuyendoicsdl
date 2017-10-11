using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AVDApplication
{
    public partial class GEWConfirmExport : Form
    {
        public GEWConfirmExport()
        {
            InitializeComponent();
        }
        DataTable dtSourceExport;
        private string sheetName = default(string);

        public GEWConfirmExport(DataTable dtExport, bool isExportTran, string nameSheet)
        {
            InitializeComponent();
            Utilities utils = new Utilities();
            sheetName = nameSheet;
            dtSourceExport = dtExport;

            if (isExportTran)
            {
                rdFre.Checked = false;
                rdTran.Checked = true;
                rdFre.Enabled = false;
                rdTran.Enabled = false;

                //// Export to CSV file
                //SaveFileDialog dialog = new SaveFileDialog();
                //dialog.Filter = "CSV file (*.csv)|*.csv";
                //dialog.Title = "Save file CSV convert.";

                //Utilities utilities = new Utilities();

                //if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                //{
                //    utilities.SaveFileCSV(dialog.FileName, dtExport);
                //}
            }
            else
            {
                //dtSourceExport = dtExport;
                rdFre.Checked = true;
                rdTran.Checked = false;
                rdFre.Enabled = false;
                rdTran.Enabled = false;

                // Export to Excel file
                // TBD
            }
        }

        private void rdbuttonExcel_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnGEWOK_Click(object sender, EventArgs e)
        {
            if (rdTran.Checked)
            {
                if (dtSourceExport != null)
                {
                    bool isExportSuccess = Utilities.exportDataToExcel(null, dtSourceExport, sheetName);
                    if (isExportSuccess)
                    {
                        MessageBox.Show("File export has been exported successful.", "Export GEW message.", MessageBoxButtons.OK);
                    }
                    else
                    {
                        MessageBox.Show("File export has been exported fail.", "Export GEW message.", MessageBoxButtons.OK);
                    }
                }
            }
            else
            {
                if (rdFre.Checked)
                {

                    /*Utilities util = new Utilities();
                    dtSourceExport = util.FormatTableGEW(dtSourceExport);
                    dtSourceExport = util.FreqencyChangeColumnName(dtSourceExport);

                    bool isExportSuccess = Utilities.exportDataToExcel(null, dtSourceExport, sheetName);
                    if (isExportSuccess)
                    {
                        MessageBox.Show("File export has been exported successful.", "Export R&S message.", MessageBoxButtons.OK);
                    }
                    else
                    {
                        MessageBox.Show("File export has been exported fail.", "Export R&S message.", MessageBoxButtons.OK);
                    }*/
                    if (dtSourceExport != null)
                    {
                        bool isExportSuccess = Utilities.exportDataToExcel(null, dtSourceExport, sheetName);
                        if (isExportSuccess)
                        {
                            MessageBox.Show("File export has been exported successful.", "Export GEW message.", MessageBoxButtons.OK);
                        }
                        else
                        {
                            MessageBox.Show("File export has been exported fail.", "Export GEW message.", MessageBoxButtons.OK);
                        }
                    }
                }
            }
        }

        private void btnGEWCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void rdFre_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void rdTran_CheckedChanged(object sender, EventArgs e)
        {

        }

    }
}
