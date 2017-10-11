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
    public partial class ConfirmExport : Form
    {
        public ConfirmExport()
        {
            InitializeComponent();
        }
        DataTable dtSourceExport;
        private string sheetName = default(string);
        public ConfirmExport(DataTable dtExport, bool isCSV, string nameSheet)
        {
            InitializeComponent();
            Utilities utils = new Utilities();
            sheetName = nameSheet;
            dtSourceExport = dtExport;

            if (isCSV)
            {
                rdbuttonExcel.Checked = false;
                rdCSV.Checked = true;
                rdCSV.Enabled = false;
                rdbuttonExcel.Enabled = false;

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
                rdbuttonExcel.Checked = true;
                rdCSV.Checked = false;
                rdCSV.Enabled = false;
                rdbuttonExcel.Enabled = false;

                // Export to Excel file
                // TBD
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (rdbuttonExcel.Checked)
            {
                if (dtSourceExport != null)
                {
                    // Export excel
                    //SaveFileDialog fileSave = new SaveFileDialog();
                    //fileSave.Filter = "Excel file (*.xls)|*.xls";
                    //fileSave.Title = "Save file Excel for suppress list.";
                    //if (fileSave.ShowDialog() == DialogResult.OK)
                    //{
                    //    ExportTableToExcel.exportToExcelSheetName(dtSourceExport, fileSave.FileName, sheetName);
                    //    MessageBox.Show("File export has been exported successful.", "Export Excel message.", MessageBoxButtons.OK);
                    //}

                    bool isExportSuccess = Utilities.exportDataToExcel(null, dtSourceExport, sheetName);
                    if (isExportSuccess)
                    {
                        MessageBox.Show("File export has been exported successful.", "Export R&S message.", MessageBoxButtons.OK);
                    }
                    else
                    {
                        MessageBox.Show("File export has been exported fail.", "Export R&S message.", MessageBoxButtons.OK);
                    }
                }
            }
            else
            {
                if (rdCSV.Checked)
                {
                    // Export CSV
                    Utilities util = new Utilities();
                    dtSourceExport = util.FormatTableRS(dtSourceExport);
                    dtSourceExport = util.ChangeColumnName(dtSourceExport);

                    bool isExportSuccess = Utilities.exportDataToExcel(null, dtSourceExport, sheetName);
                    if (isExportSuccess)
                    {
                        MessageBox.Show("File export has been exported successful.", "Export R&S message.", MessageBoxButtons.OK);
                    }
                    else
                    {
                        MessageBox.Show("File export has been exported fail.", "Export R&S message.", MessageBoxButtons.OK);
                    }

                    //SaveFileDialog fileSave = new SaveFileDialog();
                    //fileSave.Filter = "Excel file (*.xls)|*.xls";
                    //fileSave.Title = "Save file for Rohde and Swcharz.";
                    //if (fileSave.ShowDialog() == DialogResult.OK)
                    //{
                    //    //util.SaveFileCSV(fileSave.FileName, dtSourceExport);
                    //    dtSourceExport = util.FormatTableRS(dtSourceExport);
                    //    ExportTableToExcel.exportToExcelSheetName(dtSourceExport, fileSave.FileName, sheetName);
                    //    MessageBox.Show("File export has been exported successful.", "Export R&S message.", MessageBoxButtons.OK);
                    //}
                }
            }
        }

        private void ConfirmExport_Load(object sender, EventArgs e)
        {

        }

        private void rdbuttonExcel_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
