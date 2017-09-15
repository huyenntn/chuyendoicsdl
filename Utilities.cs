using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AVDApplication
{
    public class Utilities
    {

        public ArrayList GetColumnName(DataTable tb)
        {
            ArrayList arr = new ArrayList();
            foreach (var col in tb.Columns)
            {
                arr.Add(col.ToString());
            }
            return arr;
        }


        static public bool exportDataToExcel(string tieude, DataTable dt, string sheetName)
        {
            bool result = false;
            //khoi tao cac doi tuong Com Excel de lam viec
            Excel.Application xlApp;
            Excel.Worksheet xlSheet;
            Excel.Workbook xlBook;
            //doi tuong Trống để thêm  vào xlApp sau đó lưu lại sau
            object missValue = System.Reflection.Missing.Value;
            //khoi tao doi tuong Com Excel moi
            xlApp = new Excel.Application();
            xlBook = xlApp.Workbooks.Add(missValue);
            //su dung Sheet dau tien de thao tac
            xlSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(1);
            xlSheet.Name = sheetName;
            //không cho hiện ứng dụng Excel lên để tránh gây đơ máy
            xlApp.Visible = false;
            int socot = dt.Columns.Count;
            int sohang = dt.Rows.Count;
            int i, j;

            SaveFileDialog f = new SaveFileDialog();
            f.Filter = "Excel file (*.xls)|*.xls";
            if (f.ShowDialog() == DialogResult.OK)
            {


                ////set thuoc tinh cho tieu de
                xlSheet.get_Range("A1", Convert.ToChar(socot + 65) + "1");
                //Excel.Range caption = xlSheet.get_Range("A1", Convert.ToChar(socot + 65) + "1");
                //caption.Select();
                //caption.FormulaR1C1 = tieude;
                ////căn lề cho tiêu đề
                //caption.HorizontalAlignment = Excel.Constants.xlCenter;
                //caption.Font.Bold = true;
                //caption.VerticalAlignment = Excel.Constants.xlCenter;
                //caption.Font.Size = 15;
                ////màu nền cho tiêu đề
                //caption.Interior.ColorIndex = 20;
                //caption.RowHeight = 30;
                //set thuoc tinh cho cac header
                Excel.Range header = xlSheet.get_Range("A1", Convert.ToChar(socot + 65) + "1");
                header.Select();

                header.HorizontalAlignment = Excel.Constants.xlLeft;
                header.Font.Bold = true;
                header.Font.Size = 10;
                //điền tiêu đề cho các cột trong file excel
                for (i = 0; i < socot; i++)
                    xlSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                //dien cot stt
                //xlSheet.Cells[2, 1] = "STT";
                //for (i = 0; i < sohang; i++)
                //    xlSheet.Cells[i + 3, 1] = i + 1;
                //dien du lieu vao sheet


                for (i = 0; i < sohang; i++)
                    for (j = 0; j < socot; j++)
                    {
                        xlSheet.Cells[i + 2, j + 1] = dt.Rows[i][j];

                    }
                //autofit độ rộng cho các cột
                for (i = 0; i < sohang; i++)
                    ((Excel.Range)xlSheet.Cells[1, i + 1]).EntireColumn.AutoFit();

                //save file
                xlBook.SaveAs(f.FileName, Excel.XlFileFormat.xlWorkbookNormal, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
                xlBook.Close(true, missValue, missValue);
                xlApp.Quit();

                // release cac doi tuong COM
                //releaseObject(xlSheet);
                //releaseObject(xlBook);
                //releaseObject(xlApp);
                result = true;
            }
            return result;
        }

        static public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public DataTable ChangeColumnName(DataTable dtExport)
        {
            ArrayList arrNewColumnName = new ArrayList();
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.TRANSNAME);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.FREQUENCY);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.CHANNELOFS);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.SERVICE);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.SIGNATURE);

            arrNewColumnName.Add(Constants.TableExport.RSTABLE.CALLSIGN);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.LICENSEE);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.TELEPHONE);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.COUNTRY);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.ZIPCODE);

            arrNewColumnName.Add(Constants.TableExport.RSTABLE.CITY);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.STREET);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.LONGITUDE);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.LATITUDE);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.DIRECTION);

            arrNewColumnName.Add(Constants.TableExport.RSTABLE.DISTANCE);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.OFFSET);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.BANDWIDTH);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.MODULATION);
            arrNewColumnName.Add(Constants.TableExport.RSTABLE.MOD_UNIT);



            if (dtExport != null && dtExport.Rows.Count > 0 && dtExport.Columns != null && dtExport.Columns.Count > 0)
            {
                for (int i = 0; i < dtExport.Columns.Count; i++)
                {
                    dtExport.Columns[i].ColumnName = arrNewColumnName[i].ToString(); ;
                }
            }
            return dtExport;
        }

        public DataTable FormatTableRS(DataTable dtExport)
        {
            if (dtExport != null && dtExport.Rows != null && dtExport.Rows.Count > 0)
            {
                for (int ro = 0; ro <dtExport.Rows.Count; ro++)
                {
                    dtExport.Rows[ro][Constants.TableExport.TEN_KHACH_HANG] = "'" + dtExport.Rows[ro][Constants.TableExport.TEN_KHACH_HANG];
                    dtExport.Rows[ro][Constants.TableExport.DICH_VU] = "'" + dtExport.Rows[ro][Constants.TableExport.DICH_VU];

                    dtExport.Rows[ro][Constants.TableExport.KY_HIEU] = "'" + dtExport.Rows[ro][Constants.TableExport.KY_HIEU];
                    dtExport.Rows[ro][Constants.TableExport.HO_HIEU] = "'" + dtExport.Rows[ro][Constants.TableExport.HO_HIEU];

                    dtExport.Rows[ro][Constants.TableExport.GPNo] = "'" + dtExport.Rows[ro][Constants.TableExport.GPNo];
                    dtExport.Rows[ro][Constants.TableExport.DIEN_THOAI] = "'" + dtExport.Rows[ro][Constants.TableExport.DIEN_THOAI];

                    dtExport.Rows[ro][Constants.TableExport.TEN_MA_DAT_NUOC] = "'" + dtExport.Rows[ro][Constants.TableExport.TEN_MA_DAT_NUOC];
                    dtExport.Rows[ro][Constants.TableExport.ZIP_CODE] = "'" + dtExport.Rows[ro][Constants.TableExport.ZIP_CODE];

                    dtExport.Rows[ro][Constants.TableExport.TINH_THANH] = "'" + dtExport.Rows[ro][Constants.TableExport.TINH_THANH];
                    dtExport.Rows[ro][Constants.TableExport.DUONG_PHO] = "'" + dtExport.Rows[ro][Constants.TableExport.DUONG_PHO];

                    dtExport.Rows[ro][Constants.TableExport.KINH_DO] = "'" + dtExport.Rows[ro][Constants.TableExport.KINH_DO];
                    dtExport.Rows[ro][Constants.TableExport.VI_DO] = "'" + dtExport.Rows[ro][Constants.TableExport.VI_DO];

                    dtExport.Rows[ro][Constants.TableExport.DON_VI_DIEU_CHE] = "'" + dtExport.Rows[ro][Constants.TableExport.DON_VI_DIEU_CHE];
                }
            }
            return dtExport;
        }

        public void SaveFileCSV(string filePath, DataTable dtExport)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(filePath, true, Encoding.Unicode);
            try
            {
               
                if(dtExport != null&& dtExport.Rows != null && dtExport.Rows.Count > 0)
                {
                    for (int i = 0; i < dtExport.Rows.Count; i++)
                    {
                        StringBuilder sbLine = new StringBuilder();
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.TEN_KHACH_HANG]);
                        sbLine.Append(",");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.TAN_SO]);
                        sbLine.Append(",");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.OFFSET_FREQ]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.DICH_VU]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.KY_HIEU]);
                        sbLine.Append(",");

                        // 5 fields 2
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.HO_HIEU]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.GPNo]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.DIEN_THOAI]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.TEN_MA_DAT_NUOC]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.ZIP_CODE]);
                        sbLine.Append(",");

                        // 5 fields 3
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.TINH_THANH]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.DUONG_PHO]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.KINH_DO]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.VI_DO]);
                        sbLine.Append(",");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.HUONG_DAI_PHAT]);
                        sbLine.Append(",");

                        // 5 fields 4
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.KHOANG_CACH_DAI_PHAT]);
                        sbLine.Append(",");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.MIN_DO_LECH_FREQ]);
                        sbLine.Append(",");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.BANG_THONG]);
                        sbLine.Append(",");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.MIN_DIEU_CHE]);
                        sbLine.Append(",");
                        sbLine.Append("'");
                        sbLine.Append(dtExport.Rows[i][Constants.TableExport.DON_VI_DIEU_CHE]);

                        file.WriteLine(sbLine.ToString());
                    }
                }
            }
            catch (Exception)
            {
                file.Close();
                if (MessageBox.Show("Has unexpected error, please contact to administrator."
                                , "Error when export file.", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                {
                    //btnBrowse.Focus();
                }
            }
            finally
            {
              
            }
        }
        /// <summary>
        /// Get all data from file excel
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public DataSet GetAllDataFromFileExcel(string path, string sheetName)
        {
            String sConnectionString = path;
            sConnectionString = MergeNewConnect(sConnectionString);
            OleDbConnection objConn = null;
            DataSet objDataset1 = new DataSet();
            
                objConn = new OleDbConnection(sConnectionString);

                objConn.Open();

                OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [" + sheetName + "$]", objConn);//[Sheet1$]"

                OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();

                objAdapter1.SelectCommand = objCmdSelect;

                objAdapter1.Fill(objDataset1);
            
           

            return objDataset1;
        }

        public bool MachTDMB(string value)
        {
            bool match;
            string pattern = "T-DMB";

            Regex regex = new Regex(pattern);

            match = regex.Match(value).Success ? true : false;
            return match;
        }

        private bool IsNumeric(string value)
        {
            //bool variable to hold the return value

            bool match;

            //regula expression to match numeric values

            string pattern =
                "(^[-+]?\\d+(,?\\d*)*\\.?\\d*([Ee][-+]\\d*)?$)|(^[-+]?\\d?(,?\\d*)*\\.\\d+([Ee][-+]\\d*)?$)";

            //generate new Regulsr Exoression eith the pattern and a couple RegExOptions

            Regex regEx = new Regex(pattern,
                                    RegexOptions.Compiled | RegexOptions.IgnoreCase |
                                    RegexOptions.IgnorePatternWhitespace);

            //tereny expresson to see if we have a match or not

            match = regEx.Match(value).Success ? true : false;

            //return the match value (true or false)

            return match;

        }


        public bool IsKinhdo(string value)
        {
            bool match = true;

            string pattern = "[0-9][0-9][0-9]°[0-9][0-9]'[0-9][0-9][0-9.][0-9]\"E";

            Regex regEx = new Regex(pattern);

            match = regEx.Match(value).Success ? true : false;

          return match;
          }

        public bool IsVido(string value)
        {
            bool match = true;

            string pattern = "[0-9][0-9]°[0-9][0-9]'[0-9][0-9][0-9.][0-9]\"N";

            Regex regEx = new Regex(pattern);

            match = regEx.Match(value).Success ? true : false;

            return match;
        }

        public bool IsKinhdoVido(string value)
        {
            bool match = true;

            value = value.Replace("(", "");
            value = value.Replace(")", "");
            value = value.Replace(";", "");

            value = value.Replace(".", ",");

            value = value.Trim();

            value = value.Replace(" ", "");

            string pattern = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9]\"";
            Regex regEx = new Regex(pattern);
            match = regEx.Match(value).Success ? true : false;

            bool temp = false;
            string pattern2 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9],[0-9][0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9],[0-9][0-9]\"";
            regEx = new Regex(pattern2);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern3 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9],[0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9],[0-9]\"";
            regEx = new Regex(pattern3);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern4 = "[0-9][0-9][0-9]E/[0-9][0-9]N";
            regEx = new Regex(pattern4);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;


            string pattern5 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9],[0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9]\"";
            regEx = new Regex(pattern5);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern6 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9],[0-9]\"";
            regEx = new Regex(pattern6);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;


            string pattern7 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9]\"/[0-9][0-9]N[0-9]'[0-9][0-9]\"";
            regEx = new Regex(pattern7);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern8 = "[0-9][0-9][0-9]E[0-9]'[0-9][0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9]\"";
            regEx = new Regex(pattern8);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;


            string pattern9 = "[0-9][0-9][0-9]E[0-9]'[0-9][0-9]\"/[0-9][0-9]N[0-9]'[0-9][0-9]\"";
            regEx = new Regex(pattern9);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern10 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9]\"";
            regEx = new Regex(pattern10);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern11 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9]\"";
            regEx = new Regex(pattern11);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;


            string pattern12 = "[0-9][0-9][0-9]E[0-9]'[0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9]\"";
            regEx = new Regex(pattern12);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;


           
            regEx = new Regex(pattern10);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

           
            regEx = new Regex(pattern11);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;


            string pattern15 = "[0-9][0-9][0-9]E[0-9]'[0-9],[0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9],[0-9]\"";
            regEx = new Regex(pattern15);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern16 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9],[0-9]\"";
            regEx = new Regex(pattern16);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern17 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9],[0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9]\"";
            regEx = new Regex(pattern17);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern18 = "[0-9][0-9][0-9]E[0-9][0-9]'[0-9][0-9],[0-9][0-9]\"/[0-9][0-9]N[0-9][0-9]'[0-9][0-9],[0-9][0-9]\"";
            regEx = new Regex(pattern18);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;

            string pattern19 = "E/N";
            regEx = new Regex(pattern19);
            temp = regEx.Match(value).Success ? true : false;

            match = match || temp;


            return match;
        }

        public string CorrectKinhdoVido(string value, ref bool hasError)
        {
            bool match = true;

            value = value.Replace("(", "");
            value = value.Replace(")", "");
            //value = value.Replace(";", "");
            value = value.Replace(" ", "");

            string valueCorrect = "";
            hasError = false;

            if (!String.IsNullOrEmpty(value.Trim()))
            {
                try
                {
                    // Split kinhdo vido
                    string[] tempS = value.Split(';');
                    ArrayList listTemp = new ArrayList();

                    if (tempS.Length > 0)
                    {
                        for (int i = 0; i < tempS.Length; i++)
                        {
                            if (!String.IsNullOrEmpty(tempS[i]))
                            {
                                listTemp.Add(tempS[i]);
                                break;
                            }
                        }
                    }
                    if (listTemp.Count > 0)
                        value = listTemp[0].ToString();

                    string[] kinhdovido = value.Split('/');

                    if (kinhdovido.Length == 2)
                    {
                        string kinhdo = kinhdovido[0].Trim().ToUpper();
                        kinhdo = kinhdo.Replace("'", "|");
                        kinhdo = kinhdo.Replace(",", ".");

                        #region Kinhdo
                        int indexE = 0, indexPhay = 0, index2Phay = 0;

                        for (int i = 0; i < kinhdo.Length; i++)
                        {
                            if (kinhdo[i] == 'E')
                            {
                                // Set indexE
                                indexE = i;
                            }
                            if (kinhdo[i] == '|')
                            {
                                indexPhay = i;
                            }
                            if (kinhdo[i] == '\"')
                            {
                                index2Phay = i;
                            }
                        }
                        // Get kinh do hour
                        string kinhdoHour = kinhdo.Substring(0, indexE).Trim();
                        if (!String.IsNullOrEmpty(kinhdoHour))
                        {
                            double dHTemp = Convert.ToDouble(kinhdoHour);
                            int intKdHour = Convert.ToInt32(dHTemp);
                            if (intKdHour < 10)
                            {
                                // set 00F
                                kinhdoHour = "00" + intKdHour.ToString();
                            }
                            else
                            {
                                if (intKdHour < 100)
                                {
                                    // Set 0FF
                                    kinhdoHour = "0" + intKdHour.ToString();
                                }
                                else
                                {
                                    // So no mama
                                    if (intKdHour < 1000)
                                    {
                                        kinhdoHour = intKdHour.ToString();
                                    }
                                    else
                                    {
                                        // cut
                                        kinhdoHour = intKdHour.ToString().Substring(0, 3);
                                    }
                                }
                            }
                        }
                        else
                        {
                            kinhdoHour = "000";
                        }

                        // Get kinh do phut
                        string kinhdoPhut = "00";
                        string kinhdoGiay = "00";
                        if (indexPhay >= indexE && indexPhay > 0)
                        {
                            kinhdoPhut = kinhdo.Substring(indexE + 1, indexPhay - indexE - 1).Trim();
                            if (!String.IsNullOrEmpty(kinhdoPhut))
                            {
                                double dPTemp = Convert.ToDouble(kinhdoPhut);
                                int intKdPhut = Convert.ToInt32(dPTemp);
                                if (intKdPhut < 10)
                                {
                                    // set 00F
                                    kinhdoPhut = "0" + intKdPhut.ToString();
                                }
                                else
                                {
                                    if (intKdPhut < 100)
                                    {
                                        // Set 0FF
                                        kinhdoPhut = intKdPhut.ToString();
                                    }
                                    else
                                    {
                                        // cut
                                        kinhdoPhut = intKdPhut.ToString().Substring(0, 2);
                                    }
                                }
                            }
                            else
                            {
                                kinhdoPhut = "00";
                            }
                        }
                        else
                        {
                            if (indexPhay == 0)
                            {
                                if (index2Phay > 0)
                                {
                                    // Do not phay but had 2phay
                                    string kinhphutgiay = kinhdo.Substring(indexE + 1, kinhdo.Length - indexE - 1);
                                    kinhdoPhut = kinhphutgiay.Substring(0, 2);
                                    kinhphutgiay = kinhphutgiay.Replace("\"", "");
                                    kinhdoGiay = kinhphutgiay.Substring(2, kinhphutgiay.Length - 2);
                                }
                                else
                                {
                                    // Do not had phay and do not had 2phay
                                    int tep = kinhdo.Length - indexE - 1;

                                    if (tep > 0)
                                    {
                                        // Do not phay but had 2phay
                                        string kinhphutgiay = kinhdo.Substring(indexE + 1, kinhdo.Length - indexE - 1);
                                        kinhdoPhut = kinhphutgiay.Substring(0, 2);
                                        kinhphutgiay = kinhphutgiay.Replace("\"", "");
                                        kinhdoGiay = kinhphutgiay.Substring(3, kinhphutgiay.Length - 2);
                                    }
                                    else
                                    {
                                        kinhdoPhut = "00";
                                        kinhdoGiay = "00";
                                    }
                                }
                            }
                        }

                        // Get kinh do giay
                        //string kinhdoGiay = "00";
                        if (index2Phay >= indexPhay && index2Phay > 0)
                        {
                            if (indexPhay > 0)
                            {

                                kinhdoGiay = kinhdo.Substring(indexPhay + 1, index2Phay - indexPhay - 1).Trim();
                                if (!String.IsNullOrEmpty(kinhdoGiay))
                                {
                                    double dPTemp = Convert.ToDouble(kinhdoGiay);
                                    int intKdGiay = Convert.ToInt32(dPTemp);
                                    if (intKdGiay < 10)
                                    {
                                        // set 00F
                                        kinhdoGiay = "0" + intKdGiay.ToString();
                                    }
                                    else
                                    {
                                        if (intKdGiay < 100)
                                        {
                                            // Set 0FF
                                            kinhdoGiay = intKdGiay.ToString();
                                        }
                                        else
                                        {
                                            // cut
                                            kinhdoGiay = intKdGiay.ToString().Substring(0, 2);
                                        }
                                    }
                                }
                                else
                                {
                                    kinhdoGiay = "00";
                                }
                            }
                            else
                            {
                                // Khong co "phay"
                                string phutgiay = kinhdo.Substring(indexE + 1, kinhdo.Length - indexE - 1);
                                kinhdoGiay = phutgiay.Replace("\"", "").Substring(phutgiay.Length - 2 - 1, 2);
                                //string phut = phutgiay.Replace(giay, "");

                            }
                        }
                        else
                        {
                            kinhdoGiay = "00";
                        }

                        #endregion

                        #region Getvido
                        string vido = kinhdovido[1].Trim().ToUpper();
                        vido = vido.Replace("'", "|");
                        vido = vido.Replace(",", ".");

                        int indexVE = 0, indexVPhay = 0, indexV2Phay = 0;

                        for (int i = 0; i < vido.Length; i++)
                        {
                            if (vido[i] == 'N')
                            {
                                // Set indexE
                                indexVE = i;
                            }
                            if (vido[i] == '|')
                            {
                                indexVPhay = i;
                            }
                            if (vido[i] == '\"')
                            {
                                indexV2Phay = i;
                            }
                        }
                        // Get vi do hour
                        string vidoHour = vido.Substring(0, indexVE).Trim();
                        if (!String.IsNullOrEmpty(vidoHour))
                        {
                            double dHTemp = Convert.ToDouble(vidoHour);
                            int intVdHour = Convert.ToInt32(dHTemp);
                            if (intVdHour < 10)
                            {
                                // set 00F
                                vidoHour = "0" + intVdHour.ToString();
                            }
                            else
                            {
                                if (intVdHour < 100)
                                {
                                    // Set 0FF
                                    vidoHour = intVdHour.ToString();
                                }
                                else
                                {
                                    // cut
                                    vidoHour = intVdHour.ToString().Substring(0, 2);
                                }
                            }
                        }
                        else
                        {
                            vidoHour = "00";
                        }

                        // Get vi do phut
                        string vidoPhut = "00";
                        string vidoGiay = "00";
                        if (indexVPhay >= indexVE && indexVPhay > 0)
                        {
                            vidoPhut = vido.Substring(indexVE + 1, indexVPhay - indexVE - 1).Trim();
                            if (!String.IsNullOrEmpty(vidoPhut))
                            {
                                double dPTemp = Convert.ToDouble(vidoPhut);
                                int intVdPhut = Convert.ToInt32(dPTemp);
                                if (intVdPhut < 10)
                                {
                                    // set 00F
                                    vidoPhut = "0" + intVdPhut.ToString();
                                }
                                else
                                {
                                    if (intVdPhut < 100)
                                    {
                                        // Set 0FF
                                        vidoPhut = intVdPhut.ToString();
                                    }
                                    else
                                    {
                                        // cut
                                        vidoPhut = intVdPhut.ToString().Substring(0, 2);
                                    }
                                }
                            }
                            else
                            {
                                vidoPhut = "00";
                            }
                        }
                        else
                        {
                            if (indexVPhay == 0)
                            {
                                if (indexV2Phay > 0)
                                {
                                    // Do not phay but had 2phay
                                    string viphutgiay = vido.Substring(indexVE + 1, vido.Length - indexVE - 1);
                                    vidoPhut = viphutgiay.Substring(0, 2);
                                    viphutgiay = viphutgiay.Replace("\"", "");
                                    vidoGiay = viphutgiay.Substring(2, viphutgiay.Length - 2);
                                }
                                else
                                {
                                    // Do not had phay and do not had 2phay
                                    int tep = vido.Length - indexVE - 1;

                                    if (tep > 0)
                                    {
                                        // Do not phay but had 2phay
                                        string viphutgiay = kinhdo.Substring(indexVE + 1, vido.Length - indexVE - 1);
                                        vidoPhut = viphutgiay.Substring(0, 2);
                                        viphutgiay = viphutgiay.Replace("\"", "");
                                        vidoGiay = viphutgiay.Substring(3, viphutgiay.Length - 2);
                                    }
                                    else
                                    {
                                        vidoPhut = "00";
                                        vidoGiay = "00";
                                    }
                                }
                            }
                        }

                        // Get vi do giay
                        //string vidoGiay = "00";
                        if (indexV2Phay >= indexVPhay && indexV2Phay > 0)
                        {
                            if (indexVPhay > 0)
                            {
                                vidoGiay = vido.Substring(indexVPhay + 1, indexV2Phay - indexVPhay - 1).Trim();
                                if (!String.IsNullOrEmpty(vidoGiay))
                                {
                                    double dPTemp = Convert.ToDouble(vidoGiay);
                                    int intVdGiay = Convert.ToInt32(dPTemp);
                                    if (intVdGiay < 10)
                                    {
                                        // set 00F
                                        vidoGiay = "0" + intVdGiay.ToString();
                                    }
                                    else
                                    {
                                        if (intVdGiay < 100)
                                        {
                                            // Set 0FF
                                            vidoGiay = intVdGiay.ToString();
                                        }
                                        else
                                        {
                                            // cut
                                            vidoGiay = intVdGiay.ToString().Substring(0, 2);
                                        }
                                    }
                                }
                                else
                                {
                                    vidoGiay = "00";
                                }
                            }
                            else
                            {
                                // Khong co "phay"
                                string phutgiay = vido.Substring(indexVE + 1, vido.Length - indexVE - 1);
                                vidoGiay = phutgiay.Replace("\"", "").Substring(phutgiay.Length - 2 - 1, 2);
                                //string phut = phutgiay.Replace(giay, "");
                            }
                        }
                        else
                        {
                            vidoGiay = "00";
                        }
                        #endregion

                        StringBuilder temp = new StringBuilder();
                        temp.Append("(");

                        temp.Append(kinhdoHour);
                        temp.Append("E");
                        temp.Append(kinhdoPhut);
                        temp.Append("'");
                        temp.Append(kinhdoGiay);
                        temp.Append("\"");

                        temp.Append("/");

                        temp.Append(vidoHour);
                        temp.Append("N");
                        temp.Append(vidoPhut);
                        temp.Append("'");
                        temp.Append(vidoGiay);
                        temp.Append("\"");

                        temp.Append(")");

                        // Set value
                        valueCorrect = temp.ToString();
                    }
                    else
                    {
                        // Has error
                        hasError = true;
                        valueCorrect = value;
                    }
                }
                catch (Exception)
                {
                    // Has error
                    hasError = true;
                    valueCorrect = value;
                }
            }
            //else
            //{
            //    valueCorrect = "(000E00'00\"/00N00'00\")";
            //}

            return valueCorrect;
        }

        public Dictionary<string, string> CreateFreqAndStepTCI(List<string> listForStep)
        {
            Dictionary<string, string> dicFreqAndStep = new Dictionary<string, string>();

            dicFreqAndStep.Add(listForStep[0], Constants.FreqAndStep.Step.STEP_25);
            dicFreqAndStep.Add(listForStep[1], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[2], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[3], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[4], Constants.FreqAndStep.Step.STEP_6_25 + ";" + Constants.FreqAndStep.Step.STEP_100);


            string step12_5and25 = Constants.FreqAndStep.Step.STEP_5+';'+Constants.FreqAndStep.Step.STEP_12_5 + ";" + Constants.FreqAndStep.Step.STEP_25;
            dicFreqAndStep.Add(listForStep[5], step12_5and25);
            dicFreqAndStep.Add(listForStep[6], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[7], step12_5and25);

            dicFreqAndStep.Add(listForStep[8], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[9], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[10], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[11], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[12], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[13], Constants.FreqAndStep.Step.STEP_100);

            return dicFreqAndStep;
        }

        public Dictionary<string, string> CreateFreqAndStepRS(List<string> listForStep)
        {
            Dictionary<string, string> dicFreqAndStep = new Dictionary<string, string>();

            dicFreqAndStep.Add(listForStep[0], Constants.FreqAndStep.Step.STEP_3);
            dicFreqAndStep.Add(listForStep[1], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[2], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[3], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[4], Constants.FreqAndStep.Step.STEP_100);

            string step12_5and25 = Constants.FreqAndStep.Step.STEP_12_5 + ";" + Constants.FreqAndStep.Step.STEP_25;
            dicFreqAndStep.Add(listForStep[5], Constants.FreqAndStep.Step.STEP_5);
            dicFreqAndStep.Add(listForStep[6], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[7], Constants.FreqAndStep.Step.STEP_5);

            dicFreqAndStep.Add(listForStep[8], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[9], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[10], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[11], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[12], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStep.Add(listForStep[13], Constants.FreqAndStep.Step.STEP_100);

            return dicFreqAndStep;
        }

        public Dictionary<string, string> CreateFreqAndStepGEW(List<string> listForStepGEW)
        {
            Dictionary<string, string> dicFreqAndStepGEW = new Dictionary<string, string>();

            dicFreqAndStepGEW.Add(listForStepGEW[0], Constants.FreqAndStep.Step.STEP_1 + ";" + Constants.FreqAndStep.Step.STEP_5);
            dicFreqAndStepGEW.Add(listForStepGEW[1], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[2], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[3], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[4], Constants.FreqAndStep.Step.STEP_12_5 + ";" + Constants.FreqAndStep.Step.STEP_25);

            string step5and12_5and25 = Constants.FreqAndStep.Step.STEP_5 + ';' + Constants.FreqAndStep.Step.STEP_12_5 + ";" + Constants.FreqAndStep.Step.STEP_25;
            dicFreqAndStepGEW.Add(listForStepGEW[5], step5and12_5and25);
            dicFreqAndStepGEW.Add(listForStepGEW[6], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[7], step5and12_5and25);

            dicFreqAndStepGEW.Add(listForStepGEW[8], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[9], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[10], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[11], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[12], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[13], Constants.FreqAndStep.Step.STEP_100);
            dicFreqAndStepGEW.Add(listForStepGEW[14], Constants.FreqAndStep.Step.STEP_100);

            return dicFreqAndStepGEW;
        }

        public double FormatFrequency(string value)
        {
            string pattern = "^*khz";
            Regex regex = new Regex(pattern);
            double dValue = default(double);

            if(regex.Match(value.ToLower()).Success)
            {
                value.Replace(',', '.');
                value = value.ToLower();
                value = value.Replace("khz", "");
                value = value.Replace(',', '.');
                value = value.Replace(";", "");
                if(double.TryParse(value, out dValue))
                {
                    // OK
                    dValue = dValue*1000;
                }
            }
            else
            {
                value = value.Replace(',', '.');
                value = value.Replace(";", "");
                value = value.ToLower().Replace("mhz", "");
                if (double.TryParse(value, out dValue))
                {
                    // MHZ
                    dValue = dValue * 1000000;
                }
            }
            return dValue;
        }

        public ArrayList GetAllFrequencyByRange( ArrayList groupFreq, ArrayList allFreq, ref bool hasError)
        {
            for (int i = 0; i < groupFreq.Count; i++)
            {
                allFreq.Add(groupFreq[i]);

                //if(!allFreq.Contains(groupFreq[i]))
                //{
                //    allFreq.Add(groupFreq[i]);
                //}
                //else
                //{
                //    hasError = true;
                //    break;
                //}
            }
            return allFreq;
        }

        public string CorrectFrequencyByRange(string range, double step, string freq)
        {
            StringBuilder returnFreq = new StringBuilder();
          
            string[] listfreq = freq.Split(';');
           
            if (listfreq != null
                && listfreq.Length > 0)
            {
                //string[] listfreqRange;
                for (int cntFreq = 0; cntFreq < listfreq.Length; cntFreq++)
                {
                    string[] listfreqRange = listfreq[cntFreq].Trim().Split('-');
                    if (listfreqRange.Length > 1)
                    {
                        if (listfreqRange.Length == 2)
                        {
                            double dStart = FormatFrequency(listfreqRange[0].Trim());
                            double dStop = FormatFrequency(listfreqRange[1].Trim());

                            string[] allRange = range.Split('_');
                            double dAllRangeStart = FormatFrequency(allRange[0]);
                            double dAllRangeStop = FormatFrequency(allRange[1]);

                            // Check range freq with range total
                            if (dStart >= dAllRangeStart && dStop <= dAllRangeStop)
                            {
                                returnFreq.Append(listfreq[cntFreq]);
                                returnFreq.Append(";");
                            }
                            else
                            {
                                // Do not add frequency
                            }
                        }
                    }
                    else
                    {
                        // Single value
                        double dFrequeceValue = FormatFrequency(listfreq[cntFreq].Trim());
                        string[] allRange = range.Split('_');
                        double dAllRangeStart = FormatFrequency(allRange[0]);
                        double dAllRangeStop = FormatFrequency(allRange[1]);
                        if (dFrequeceValue >= dAllRangeStart && dFrequeceValue <= dAllRangeStop)
                        {
                            returnFreq.Append(listfreq[cntFreq]);
                            returnFreq.Append(";");
                        }
                        else
                        {
                            // Do not add frequency
                        }
                        
                    }
                }
            }

            return returnFreq.ToString();
        }


        public Dictionary<int, ArrayList> GetFrequencyByRange(string range, double step, string freq, int indexOfRow, ref bool hasError)
        {
            Dictionary<int, ArrayList> groupfreqbyRange = new Dictionary<int, ArrayList>();
            ArrayList arrFreq = new ArrayList();

            string[] listfreq = freq.Split(';');
            bool hasContainKhz = false;
            bool isRangeMhz = true;
            ArrayList arrKhz = new ArrayList();

            if(listfreq != null
                && listfreq.Length > 0)
            {
                //string[] listfreqRange;
                for(int cntFreq = 0; cntFreq < listfreq.Length; cntFreq++)
                {
                    string[] listfreqRange = listfreq[cntFreq].Trim().Split('-');
                    if (listfreqRange.Length > 1)
                    {
                        if (listfreqRange.Length == 2)
                        {
                            double dStart = FormatFrequency(listfreqRange[0].Trim());
                            double dStop = FormatFrequency(listfreqRange[1].Trim());

                            string[] allRange = range.Split('_');
                            double dAllRangeStart = FormatFrequency(allRange[0]);
                            double dAllRangeStop = FormatFrequency(allRange[1]);

                            // Check range freq with range total
                            if (dStart >= dAllRangeStart && dStop <= dAllRangeStop)
                            {
                                // Has range
                                double dfreqRange = dStart;
                                string pattern = "^*khz";
                                Regex regex = new Regex(pattern);
                                if (regex.Match(listfreqRange[0].Trim()).Success
                                    || regex.Match(listfreqRange[1].Trim()).Success)
                                {
                                    isRangeMhz = false;
                                }

                                if (!arrFreq.Contains(dfreqRange))
                                {
                                    while (dfreqRange <= dStop)
                                    {
                                        if (!arrFreq.Contains(dfreqRange))
                                        {
                                            arrFreq.Add(dfreqRange);
                                            dfreqRange = dfreqRange + step;
                                        }
                                        //else
                                        //{
                                        //    // Has error
                                        //    hasError = true;
                                        //    break;
                                        //}
                                    }
                                }
                                //else
                                //{
                                //    // Has error
                                //    hasError = true;
                                //    break;
                                //}
                            }
                            else
                            {
                                // Vuot qua range tong
                                hasError = true;
                                break;
                            }
                        }
                        //else
                        //{
                        //    // Not valid
                        //    hasError = true;
                        //    break;
                        //}
                    }
                    else
                    {
                        // Single value
                        double dFrequeceValue = FormatFrequency(listfreq[cntFreq].Trim());
                        if (dFrequeceValue > 0)
                        {
                            string[] allRange = range.Split('_');
                            double dAllRangeStart = FormatFrequency(allRange[0]);
                            double dAllRangeStop = FormatFrequency(allRange[1]);
                            if (dFrequeceValue >= dAllRangeStart && dFrequeceValue <= dAllRangeStop)
                            {
                                if (!arrFreq.Contains(FormatFrequency(listfreq[cntFreq].Trim())))
                                {
                                    string value = listfreq[cntFreq].Trim().ToLower();
                                    string pattern = "^*khz";
                                    Regex regex = new Regex(pattern);
                                    if (regex.Match(value).Success)
                                    {
                                        hasContainKhz = true;
                                        arrKhz.Add(FormatFrequency(value));
                                    }
                                    if (!String.IsNullOrEmpty(value))
                                    {
                                        arrFreq.Add(FormatFrequency(value));
                                    }
                                }
                            }
                            else
                            {
                                // Has error
                                hasError = true;
                                break;
                            }
                        }
                        //else
                        //{
                        //    // Has error
                        //    hasError = true;
                        //    break;
                        //}
                        //groupfreqbyRange.Add(indexOfRow,arrFreq);
                    }
                }
            }

            if(isRangeMhz && hasContainKhz
                && arrFreq.Count > 0 && arrKhz.Count > 0)
            {
                // Remove all khz
                for (int i = 0; i < arrKhz.Count; i++)
                {
                    if(arrFreq.Contains(arrKhz[i]))
                    {
                        arrFreq.Remove(arrKhz[i]);
                    }
                }
            }
            groupfreqbyRange.Add(indexOfRow, arrFreq);

            return groupfreqbyRange;
        }

        //private bool CheckFrequency(string freq, double startFreq, double stopFreq )

        public bool IsValidateField(string fieldName, string fieldValue, bool isTCI)
        {
            bool hasValidate = true;
            switch (fieldName)
            {
                case (Constants.TableExport.BANG_THONG):
                    // Is Numeric
                    // Check field value is Numeric,isn't it.
                    hasValidate = IsNumeric(fieldValue);
                    break;

                // Khong su dung
                //case (Constants.TableExport.BRAND_UU_TIEN):

                //    break;

                case (Constants.TableExport.DICH_VU):
                    // Not NULL and <=2 characters
                    if(!String.IsNullOrEmpty(fieldValue))
                    {
                        if(fieldValue.Length > 2)
                        {
                            hasValidate = false;
                        }
                    }
                    else
                    {
                        hasValidate = false;
                    }
                    break;

                case (Constants.TableExport.DIEN_THOAI):
                    // <= 20 characters
                    if(!String.IsNullOrEmpty(fieldValue))
                    {
                        if (fieldValue.Length > 20)
                            hasValidate = false;
                    }
                    break;

                    // Khong su dung
                //case (Constants.TableExport.DO_LECH_F):
                //    break;
                //case (Constants.TableExport.DO_RONG_KENH):
                //    break;

                case (Constants.TableExport.DON_VI_DIEU_CHE):
                    // <= 3 characters
                    if (!String.IsNullOrEmpty(fieldValue))
                    {
                        if (fieldValue.Length > 3)
                            hasValidate = false;
                    }
                    break;

                case (Constants.TableExport.DUONG_PHO):
                    // <= 50 characters
                    if (!String.IsNullOrEmpty(fieldValue))
                    {
                        if (fieldValue.Length > 50)
                            hasValidate = false;
                    }
                    break;
                // Khong su dung
                case (Constants.TableExport.GPNo):
                    if(!isTCI)
                    {
                        // <= 32 characters
                        if (!String.IsNullOrEmpty(fieldValue))
                        {
                            if (fieldValue.Length > 32)
                                hasValidate = false;
                        }
                    }
                    break;
                case (Constants.TableExport.HO_HIEU):
                    if (!isTCI)
                    {
                        // <= 32 characters
                        if (!String.IsNullOrEmpty(fieldValue))
                        {
                            if (fieldValue.Length > 32)
                                hasValidate = false;
                        }
                    }
                    break;

                case (Constants.TableExport.HUONG_DAI_PHAT):
                    // Is Numeric
                    // Check field value is Numeric,isn't it.
                    hasValidate = IsNumeric(fieldValue);
                    break;

                //case (Constants.TableExport.ID):
                //    break;

                case (Constants.TableExport.KHOANG_CACH_DAI_PHAT):
                    // Is Numeric
                    // Check field value is Numeric,isn't it.
                    hasValidate = IsNumeric(fieldValue);
                    break;

                case (Constants.TableExport.KINH_DO):
                    // xxx°xx'xxxx"E
                    hasValidate = IsKinhdo(fieldValue);
                    break;
                case (Constants.TableExport.KINHDO_VIDO):
                    // xxExx'xx"/xxNxx'xx"
                    hasValidate = IsKinhdoVido(fieldValue);
                    break;
                case (Constants.TableExport.KY_HIEU):
                    // <= 20 characters
                    if (!String.IsNullOrEmpty(fieldValue))
                    {
                        if (fieldValue.Length > 20)
                            hasValidate = false;
                    }
                    break;
                //case (Constants.TableExport.MAU_GIAY_PHEP):
                //    break;
                case (Constants.TableExport.MIN_DIEU_CHE):
                    // Is Numeric
                    // Check field value is Numeric,isn't it.
                    hasValidate = IsNumeric(fieldValue);
                    break;
                case (Constants.TableExport.MIN_DO_LECH_FREQ):
                    // Is Numeric
                    // Check field value is Numeric,isn't it.
                    hasValidate = IsNumeric(fieldValue);
                    break;
                case (Constants.TableExport.OFFSET_FREQ):
                    // Is Numeric
                    // Check field value is Numeric,isn't it.
                    hasValidate = IsNumeric(fieldValue);
                    break;
                //case (Constants.TableExport.SO_KENH):
                    //break;
                //case (Constants.TableExport.SO_THAM_CHIEU):
                //    break;
                case (Constants.TableExport.TAN_SO):
                    break;
                case (Constants.TableExport.TEN_MA_DAT_NUOC):
                    break;
                case (Constants.TableExport.TEN_KHACH_HANG):
                    // <= 25 characters
                    if (!String.IsNullOrEmpty(fieldValue))
                    {
                        if (fieldValue.Length > 25)
                            hasValidate = false;
                    }
                    break;
                case (Constants.TableExport.TEN_MAY):
                    break;
                case (Constants.TableExport.TINH_THANH):
                    break;
                case (Constants.TableExport.VI_DO):
                    // xx°xx'xxxx"N
                    hasValidate = IsVido(fieldValue);
                    break;
                case (Constants.TableExport.ZIP_CODE):
                    break;

            }
            return hasValidate;
        }

        /// <summary>
        /// Merg link
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private string MergeNewConnect(string path)
        {
            StringBuilder newConn = new StringBuilder();
            newConn.Append("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
            newConn.Append(path);
            newConn.Append(@";Extended Properties=Excel 8.0;");
            return newConn.ToString();
        }

        public DataTable GetTemplateTableRS()
        {
            DataTable dtExcel = new DataTable();
            dtExcel.TableName = "TableRSCSV";
            //dtExcel.Columns.Add(Constants.TableExport.ID);

            dtExcel.Columns.Add(Constants.TableExport.TEN_KHACH_HANG);
            dtExcel.Columns.Add(Constants.TableExport.TAN_SO);
            dtExcel.Columns.Add(Constants.TableExport.OFFSET_FREQ);
            dtExcel.Columns.Add(Constants.TableExport.DICH_VU);
            dtExcel.Columns.Add(Constants.TableExport.KY_HIEU);

            dtExcel.Columns.Add(Constants.TableExport.HO_HIEU);
            dtExcel.Columns.Add(Constants.TableExport.GPNo);
            dtExcel.Columns.Add(Constants.TableExport.DIEN_THOAI);
            dtExcel.Columns.Add(Constants.TableExport.TEN_MA_DAT_NUOC);
            dtExcel.Columns.Add(Constants.TableExport.ZIP_CODE);

            dtExcel.Columns.Add(Constants.TableExport.TINH_THANH);
            dtExcel.Columns.Add(Constants.TableExport.DUONG_PHO);
            dtExcel.Columns.Add(Constants.TableExport.KINH_DO);
            dtExcel.Columns.Add(Constants.TableExport.VI_DO);
            dtExcel.Columns.Add(Constants.TableExport.HUONG_DAI_PHAT);

            dtExcel.Columns.Add(Constants.TableExport.KHOANG_CACH_DAI_PHAT);
            dtExcel.Columns.Add(Constants.TableExport.MIN_DO_LECH_FREQ);
            dtExcel.Columns.Add(Constants.TableExport.BANG_THONG);
            dtExcel.Columns.Add(Constants.TableExport.MIN_DIEU_CHE);
            dtExcel.Columns.Add(Constants.TableExport.DON_VI_DIEU_CHE);

            return dtExcel;

        }

        public DataTable GetTemplateTable()
        {
            DataTable dtExcel = new DataTable();
            dtExcel.TableName = "TableDes";
            dtExcel.Columns.Add(Constants.TableExport.ID);
            dtExcel.Columns.Add(Constants.TableExport.GPNo);
            dtExcel.Columns.Add(Constants.TableExport.MAU_GIAY_PHEP);
            dtExcel.Columns.Add(Constants.TableExport.SO_THAM_CHIEU);
            dtExcel.Columns.Add(Constants.TableExport.DO_LECH_F);
            dtExcel.Columns.Add(Constants.TableExport.TAN_SO);
            dtExcel.Columns.Add(Constants.TableExport.BRAND_UU_TIEN);
            dtExcel.Columns.Add(Constants.TableExport.DO_RONG_KENH);
            dtExcel.Columns.Add(Constants.TableExport.SO_KENH);
            dtExcel.Columns.Add(Constants.TableExport.TEN_KHACH_HANG);
            dtExcel.Columns.Add(Constants.TableExport.HO_HIEU);
            dtExcel.Columns.Add(Constants.TableExport.VI_DO);
            dtExcel.Columns.Add(Constants.TableExport.KINH_DO);
            dtExcel.Columns.Add(Constants.TableExport.TEN_MAY);

            return dtExcel;

        }
        public DataTable GetTemplateTableDFScan()
        {
            DataTable dtExcel = new DataTable();
            dtExcel.TableName = "TableDFS";
            dtExcel.Columns.Add(Constants.TableExport.ID);
            dtExcel.Columns.Add(Constants.TableExport.TAN_SO);
            //dtExcel.Columns.Count == 2
            return dtExcel;

        }


        // Save random value
        public Dictionary<double, double> randomLongtitudeDict = null;

        public Dictionary<double, double> randomLatitudeDict = null;
        /// <summary>
        /// Format vi do
        /// </summary>
        /// <param name="latitude"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public double FormatLatitude(string latitude, string type)
        {
            double latitudeFormat = default(double);

            if (type != Constants.ValueConstant.RANDOM)
            {
                // (105E50'57.35" /20N56'52.77" );
                string strDo = latitude.Substring(0, 2);

                // minute and minus
                string strMinuteAndMinus = latitude.Substring(3);

                // Split minute and minus
                string[] splitArr = strMinuteAndMinus.Split(Convert.ToChar("'"));

                // minute
                string strMinute = splitArr[0];

                // minus
                string strMinus = splitArr[1].Replace("\"", "");

                strMinus = strMinus.Replace(',', '.');

                latitudeFormat = Convert.ToDouble(strDo) + Convert.ToDouble(strMinute) / 60 +
                                 Convert.ToDouble(strMinus) / 3600;
            }
            else
            {
                //Random ran = new Random();
                //int minute = ran.Next(0, 59);
                //int minus = ran.Next(0, 59);
                //latitudeFormat = Constants.ValueConstant.LATITUDE + Convert.ToDouble(minute)/60 +
                //                 Convert.ToDouble(minus)/3600;

                Random ranMinute = new Random();
                double temp = default(double);
                while (true)
                {
                    temp = Constants.ValueConstant.LATITUDE + ranMinute.NextDouble() / 1;
                    if (randomLatitudeDict == null ||
                        !randomLatitudeDict.ContainsKey(temp))
                    {
                        if (randomLatitudeDict == null)
                        {
                            randomLatitudeDict = new Dictionary<double, double>();
                            randomLatitudeDict.Add(temp, temp);
                        }
                        else
                            randomLatitudeDict.Add(temp, temp);

                        latitudeFormat = temp;
                        break;
                    }
                }
            }

            return latitudeFormat;
        }

        public string FormatLatitudeRS(string latitude, string type)
        {
            string latitudeFormat = default(string);

            if (type != Constants.ValueConstant.RANDOM)
            {
                // (105E50'57.35" /20N56'52.77" );
                string strDo = latitude.Substring(0, 2);

                // minute and minus
                string strMinuteAndMinus = latitude.Substring(3);

                // Split minute and minus
                string[] splitArr = strMinuteAndMinus.Split(Convert.ToChar("'"));

                // minute
                string strMinute = splitArr[0];

                // minus
                string strMinus = splitArr[1].Replace("\"", "");

                strMinus = strMinus.Replace(',', '.');

                latitudeFormat = strDo + Constants.ValueConstant.HOURVALUE + strMinute + "'" + strMinus + "''" + "N";
            }
            else
            {               
                Random ranMinute = new Random();
                double temp = default(double);
                while (true)
                {
                    temp = Constants.ValueConstant.LATITUDE + ranMinute.NextDouble() / 1;
                    if (randomLatitudeDict == null ||
                        !randomLatitudeDict.ContainsKey(temp))
                    {
                        if (randomLatitudeDict == null)
                        {
                            randomLatitudeDict = new Dictionary<double, double>();
                            randomLatitudeDict.Add(temp, temp);
                        }
                        else
                            randomLatitudeDict.Add(temp, temp);

                        latitudeFormat = ConvertKinhdo_Vido(temp, false);
                        break;
                    }
                }
            }

            return latitudeFormat;
        }

        private string ConvertKinhdo_Vido(double kinhdovido, bool isKinhdo)
        {
            string returnValue = default(string);
            if (isKinhdo)
            {
                // Format kinhdo
                string[] kinhdovidoSplit = kinhdovido.ToString().Split('.');
                string hour = default(string);
                string minus = default(string);
                string second = default(string);

                if (kinhdovidoSplit.Length > 0)
                {
                    hour = kinhdovidoSplit[0];
                    string minussecondtemp = "0." + kinhdovidoSplit[1];
                    double minustemp = (Convert.ToDouble(minussecondtemp) * 60);
                    string[] minussecondSplit = minustemp.ToString().Split('.');

                    if (minussecondSplit.Length > 0)
                    {
                        minus = minussecondSplit[0];

                        string secondtemp = "0." + minussecondSplit[1];
                        //double secondValuetemp = Convert.ToDouble(minussecondSplit[1]) / 60;
                        double test = Math.Round(Convert.ToDouble(secondtemp) * 60);
                        second = (Math.Round(Convert.ToDouble(secondtemp)*60)).ToString();
                    }

                    returnValue = hour + Constants.ValueConstant.HOURVALUE + minus + "'" + second + "''" + "E";
                }
            }
            else
            {
                // Format vido
                string[] kinhdovidoSplit = kinhdovido.ToString().Split('.');
                string hour = default(string);
                string minus = default(string);
                string second = default(string);

                if (kinhdovidoSplit.Length > 0)
                {
                    hour = kinhdovidoSplit[0];
                    string minussecondtemp = "0." + kinhdovidoSplit[1];
                    double minustemp = (Convert.ToDouble(minussecondtemp) * 60);
                    string[] minussecondSplit = minustemp.ToString().Split('.');

                    if (minussecondSplit.Length > 0)
                    {
                        minus = minussecondSplit[0];

                        string secondtemp = "0." + minussecondSplit[1];
                        //double secondValuetemp = Convert.ToDouble(minussecondSplit[1]) / 60;

                        second = (Math.Round(Convert.ToDouble(secondtemp) * 60)).ToString();
                    }

                    returnValue = hour + Constants.ValueConstant.HOURVALUE + minus + "'" + second + "''" + "N";
                }                
            }

            return returnValue;
        }

        /// <summary>
        /// Format kinh do
        /// </summary>
        /// <param name="longtitude"></param>
        /// <returns></returns>
        public double FormatLongtitude(string longtitude, string type)
        {
            double longtitudeFormat = default(double);

            if (type != Constants.ValueConstant.RANDOM)
            {
                // (105E50'57.35" /20N56'52.77" );
                string strDo = longtitude.Substring(0, 3);

                // minute and minus
                string strMinuteAndMinus = longtitude.Substring(4);

                // Split minute and minus
                string[] splitArr = strMinuteAndMinus.Split(Convert.ToChar("'"));

                // minute
                string strMinute = splitArr[0];

                // minus
                string strMinus = splitArr[1].Replace("\"", "");

                strMinus = strMinus.Replace(',', '.');

                longtitudeFormat = Convert.ToDouble(strDo) + Convert.ToDouble(strMinute) / 60 + Convert.ToDouble(strMinus) / 3600;
            }
            else
            {
                Random ranMinute = new Random();
                double temp = default(double);
                while (true)
                {
                    temp = Constants.ValueConstant.LONGTIDUDE + ranMinute.NextDouble() / 1;
                    if (randomLongtitudeDict == null ||
                        !randomLongtitudeDict.ContainsKey(temp))
                    {
                        if (randomLongtitudeDict == null)
                        {
                            randomLongtitudeDict = new Dictionary<double, double>();
                            randomLongtitudeDict.Add(temp, temp);
                        }
                        else
                            randomLongtitudeDict.Add(temp, temp);

                        longtitudeFormat = temp;
                        break;
                    }
                }
            }

            return longtitudeFormat;
        }


        public string FormatLongtitudeRS(string longtitude, string type)
        {
            string longtitudeFormat = default(string);

            if (type != Constants.ValueConstant.RANDOM)
            {
                // (105E50'57.35" /20N56'52.77" );
                string strDo = longtitude.Substring(0, 3);

                // minute and minus
                string strMinuteAndMinus = longtitude.Substring(4);

                // Split minute and minus
                string[] splitArr = strMinuteAndMinus.Split(Convert.ToChar("'"));

                // minute
                string strMinute = splitArr[0];

                // minus
                string strMinus = splitArr[1].Replace("\"", "");

                strMinus = strMinus.Replace(',', '.');

                longtitudeFormat = strDo + Constants.ValueConstant.HOURVALUE + strMinute + "'" + strMinus + "''" + "E";
            }
            else
            {
                Random ranMinute = new Random();
                double temp = default(double);
                while (true)
                {
                    temp = Constants.ValueConstant.LONGTIDUDE + ranMinute.NextDouble() / 1;
                    if (randomLongtitudeDict == null ||
                        !randomLongtitudeDict.ContainsKey(temp))
                    {
                        if (randomLongtitudeDict == null)
                        {
                            randomLongtitudeDict = new Dictionary<double, double>();
                            randomLongtitudeDict.Add(temp, temp);
                        }
                        else
                            randomLongtitudeDict.Add(temp, temp);

                        longtitudeFormat = ConvertKinhdo_Vido(temp, true);
                        break;
                    }
                }
            }

            return longtitudeFormat;
        }
        /// <summary>
        /// Get kinh do and vi do
        /// </summary>
        /// <param name="kinhdovido"></param>
        /// <returns></returns>
        public string[] GetKinhdoAndVido(string kinhdovido)
        {
            string[] kinhvidoArr = new string[2];

            //string[] temp = kinhdovido.Split(';');

            //if (temp.Length >= 1 && temp.Length <= 2)
            //{
            //    if (!string.IsNullOrEmpty(temp[0]))
            //    {
            //        kinhdovido = temp[0];
            //    }
            //    else
            //    {
            //        kinhdovido = temp[1];
            //    }

            //}
            //else
            //{
            //    kinhdovido = temp[0];
            //}

            // Split kinhdovido
            // (105E50'57.35" /20N56'52.77" );
            // Replace character no need
            kinhdovido = kinhdovido.Replace("(", "");
            kinhdovido = kinhdovido.Replace(")", "");
            kinhdovido = kinhdovido.Replace(";", "");

            string[] arrTemp = kinhdovido.Trim().Split('/');

            if (arrTemp.Length <= 2
                && arrTemp.Length > 0
                && arrTemp[0].Length > 3
                && arrTemp[1].Length > 2)
            {
                kinhvidoArr[0] = arrTemp[0];
                kinhvidoArr[1] = arrTemp[1];
            }
            else
            {
                // Set random longtitude and latitude
                kinhvidoArr[0] = Constants.ValueConstant.RANDOM;
                kinhvidoArr[1] = Constants.ValueConstant.RANDOM;
            }
            return kinhvidoArr;
        }

    }
}
