using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace AVDApplication
{
    public partial class MainMenu : Form
    {
        public MainMenu()
        {
            InitializeComponent();
        }

        private void MainMenu_Load(object sender, EventArgs e)
        {
            
            //cmbChooseFreq.DataSource 
            List<string> freqRange = new List<string>();
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_HF_9_30);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_47_50);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_54_68);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_87_108);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_HKHONG_108_138);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_138_174);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_174_230);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_400_463);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_CDMA_806_890);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_EGDSM_890_960);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_GSM_1800_1900);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_3G_2100_2170);
            freqRange.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_3G_2620_2680);

            List<string> freqRangeGEW = new List<string>();
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_HF_9_30);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_47_50);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_54_68);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_PT_87_108);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_HK_108_137);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_174_230);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_470_790);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_790_890);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_890_960);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_1710_1785);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_1805_1880);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_1920_1980);
            freqRangeGEW.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_2110_2170);

            cmbChooseFreq.DataSource = freqRange;
            cmbRSChooseFreq.DataSource = freqRange;
            cmbGEChooseFreq.DataSource = freqRangeGEW;

            List<string> freqRangeForStep = new List<string>();
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_HF_9_30);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_FM_47_50);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_FM_54_68);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_FM_87_108);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_HKHONG_108_138);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_VHF_138_174);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_VHF_174_230);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_UHF_400_463);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_UHF_470_806);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_CDMA_806_890);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_EGDSM_890_960);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_GSM_1800_1900);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_3G_2100_2170);
            freqRangeForStep.Add(Constants.FreqAndStep.Frequency.FREQ_3G_2620_2680);

            List<string> freqRangeForStepGEW = new List<string>();
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_HF_9_30);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTKD_47_50);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTKD_54_68);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_PT_87_108);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_HK_108_137);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_DR_137_174);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TH_174_230);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_DR_400_470);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TH_470_790);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTDD_790_890);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTDD_890_960);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTDD_1710_1785);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTDD_1805_1880);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTDD_1920_1980);
            freqRangeForStepGEW.Add(Constants.FreqAndStep.FrequencyGEW.FREQ_TTDD_2110_2170);

            

            Utilities utilities = new Utilities();
            Dictionary<string, string> dicStep = utilities.CreateFreqAndStepTCI(freqRange);
            if (cmbChooseFreq.SelectedItem != null)
            {
                string[] listStep = dicStep[cmbChooseFreq.SelectedItem.ToString()].Split(';');

                cmbRSStep.DataSource = listStep;
            }

            Dictionary<string, string> dicStepRS = utilities.CreateFreqAndStepRS(freqRange);
            if (cmbRSChooseFreq.SelectedItem != null)
            {
                string[] listStepRS = dicStepRS[cmbRSChooseFreq.SelectedItem.ToString()].Split(';');

                cmbRSStep.DataSource = listStepRS;
            }

            Dictionary<string, string> dicStepGEW = utilities.CreateFreqAndStepGEW(freqRangeGEW);
            if (cmbGEChooseFreq.SelectedItem != null)
            {
                string[] listStepGEW = dicStepGEW[cmbGEChooseFreq.SelectedItem.ToString()].Split(';');

                cmbGEStep.DataSource = listStepGEW;
                
            }
            btnCheckError.Enabled = false;
            btnFormat.Enabled = false;
            btnRSCheckError.Enabled = false;
            btnRSShow.Enabled = false;
            btnGECheckError.Enabled = false;
            btnGECorrectError.Enabled = false;
            btnTranFormat.Enabled = false;
            btnFrequencies.Enabled = false;
            buttonGExport.Enabled = false;
            btnGEShow.Enabled = false;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel file (*.xls)|*.xls";
            dialog.Title = "Open file Excel convert.";

            // Clean table dtSource
            dtTCISource = null;

            Utilities utilities = new Utilities();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string sheetName = this.GetSheetName(dialog.FileName);
                //string sheetName = "sheet1";
                DataSet dsExcel = utilities.GetAllDataFromFileExcel(dialog.FileName, sheetName);

                if (dsExcel != null
                && dsExcel.Tables != null
                && dsExcel.Tables.Count > 0
                && dsExcel.Tables[0].Rows.Count > 0)
                {
                    //dgRSDetailInformation.DataSource = null;
                    if (dgDetailInformation.DataSource != null)
                    {
                        dgDetailInformation.DataSource = null;
                        dgDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    else
                    {
                        dgDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    if (allListRange != null && allListRange.Count > 0)
                    {
                        allListRange.Clear();
                    }

                    // dgDetailInformation.DataSource = dsExcel.Tables[0];

                    // if (allListRange != null && allListRange.Count > 0)
                    // {
                    //     allListRange.Clear();
                    // }
                }

                // Enable button
                imgTCI.Visible = false;
                btnCheckError.Enabled = true;
                btnFormat.Enabled = false;
                btnCorrectError.Enabled = false;
                btnShow.Enabled = false;
                button9.Enabled = false;
                btnExport.Enabled = false;
            }
        }

        /// <summary>
        /// Get frequence, check has multi and get them.
        /// </summary>
        /// <param name="strFrequence"></param>
        /// <returns></returns>
        private ArrayList GetFrequence(string strFrequence)
        {
            ArrayList arrAllFrequence = new ArrayList();

            string[] allFrequence = null;

            // Split strFrequence
            if (!String.IsNullOrEmpty(strFrequence))
            {
                allFrequence = strFrequence.Trim().Split(';');
                if (allFrequence != null
                    && allFrequence.Length > 0)
                {
                    for (int i = 0; i < allFrequence.Length; i++)
                    {
                        allFrequence[i] = allFrequence[i].ToLower();
                        if (allFrequence[i].Contains("mhz"))
                        {
                            arrAllFrequence.Add(allFrequence[i]);
                        }
                    }
                }
            }

            // return
            return arrAllFrequence;
        }


        /// <summary>
        /// Get kinh do and vi do
        /// </summary>
        /// <param name="kinhdovido"></param>
        /// <returns></returns>
        private string[] GetKinhdoAndVido(string kinhdovido)
        {
            string[] kinhvidoArr = new string[2];

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


        private DataTable GetAllDataToMerge(DataSet dsExcel)
        {
            Utilities utilities = new Utilities();

            DataTable dtExcel = utilities.GetTemplateTable();

            if (dsExcel != null
                && dsExcel.Tables != null
                && dsExcel.Tables.Count > 0
                && dsExcel.Tables[0].Rows.Count > 0)
            {
                // ID of line
                int intId = 1;

                for (int i = 0; i < dsExcel.Tables[0].Rows.Count; i++)
                {
                    DataRow row = dtExcel.NewRow();
                    // Comment ID
                    //row[Constants.TableExport.ID] = dsExcel.Tables[0].Rows[i]["STT"];


                    if (!String.IsNullOrEmpty(dsExcel.Tables[0].Rows[i][Constants.TableExport.GPNo].ToString()))
                    {
                        row[Constants.TableExport.GPNo] =
                            dsExcel.Tables[0].Rows[i][Constants.TableExport.GPNo].ToString().Trim().Replace(";", "");
                    }
                    else
                    {
                        row[Constants.TableExport.GPNo] = Constants.ValueConstant.SPACE;
                    }
                    if (!String.IsNullOrEmpty(dsExcel.Tables[0].Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                    {
                        row[Constants.TableExport.MAU_GIAY_PHEP] =
                            dsExcel.Tables[0].Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim().Replace(
                                ";", "");
                    }
                    else
                    {
                        row[Constants.TableExport.MAU_GIAY_PHEP] = Constants.ValueConstant.SPACE;
                    }

                    // Two columns insert space
                    row[Constants.TableExport.SO_THAM_CHIEU] = Constants.ValueConstant.SPACE;
                    row[Constants.TableExport.DO_LECH_F] = Constants.ValueConstant.SPACE;

                    // Frequence
                    // Get multi frequence
                    ArrayList arrFrequece =
                        this.GetFrequence(dsExcel.Tables[0].Rows[i][Constants.TableExport.TAN_SO].ToString().Trim());
                    //row[Constants.TableExport.TAN_SO] = this.FormatFrequence(dsExcel.Tables[0].Rows[i][Constants.TableExport.TAN_SO].ToString());


                    // Three columns insert space
                    row[Constants.TableExport.BRAND_UU_TIEN] = Constants.ValueConstant.SPACE;
                    row[Constants.TableExport.DO_RONG_KENH] = Constants.ValueConstant.SPACE;
                    row[Constants.TableExport.SO_KENH] = Constants.ValueConstant.SPACE;

                    // Five columns normal
                    if (!String.IsNullOrEmpty(dsExcel.Tables[0].Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                    {
                        row[Constants.TableExport.TEN_KHACH_HANG] =
                            dsExcel.Tables[0].Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(
                                ";", "");
                    }
                    else
                    {
                        row[Constants.TableExport.TEN_KHACH_HANG] = Constants.ValueConstant.SPACE;
                    }

                    if (!String.IsNullOrEmpty(dsExcel.Tables[0].Rows[i][Constants.TableExport.HO_HIEU].ToString()))
                    {
                        row[Constants.TableExport.HO_HIEU] =
                            dsExcel.Tables[0].Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Replace(";", "");
                    }
                    else
                    {
                        row[Constants.TableExport.HO_HIEU] = Constants.ValueConstant.SPACE;
                    }
                    // Longtitude and latitude
                    string[] kinhdoVidoArr =
                        this.GetKinhdoAndVido(dsExcel.Tables[0].Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
                    if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
                    {
                        // Get normal
                        row[Constants.TableExport.VI_DO] = this.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);

                    }
                    else
                    {
                        // Call random value;
                        row[Constants.TableExport.VI_DO] = this.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
                    }

                    if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
                    {
                        row[Constants.TableExport.KINH_DO] = this.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
                    }
                    else
                    {
                        row[Constants.TableExport.KINH_DO] = this.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
                    }
                    // Ten may
                    if (!String.IsNullOrEmpty(dsExcel.Tables[0].Rows[i][Constants.TableExport.TEN_MAY].ToString()))
                    {
                        row[Constants.TableExport.TEN_MAY] =
                            dsExcel.Tables[0].Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Replace(";", "");
                    }
                    else
                    {
                        row[Constants.TableExport.TEN_MAY] = Constants.ValueConstant.SPACE;
                    }

                    // Check multi frequence
                    if (arrFrequece.Count > 1)
                    {
                        for (int fre = 0; fre < arrFrequece.Count; fre++)
                        {
                            //row[Constants.TableExport.TAN_SO] = this.FormatFrequence(arrFrequece[fre].ToString());
                            if (fre > 0)
                            {
                                DataRow dtRowTemp = dtExcel.NewRow();
                                dtRowTemp.ItemArray = row.ItemArray;
                                dtRowTemp[Constants.TableExport.TAN_SO] = this.FormatFrequence(arrFrequece[fre].ToString());
                                dtRowTemp[Constants.TableExport.ID] = intId;

                                dtExcel.Rows.Add(dtRowTemp);

                                // Increase ID
                                intId++;
                            }
                            else
                            {
                                // ID
                                row[Constants.TableExport.ID] = intId;
                                row[Constants.TableExport.TAN_SO] = this.FormatFrequence(arrFrequece[fre].ToString());
                                dtExcel.Rows.Add(row);

                                // Increase Id
                                intId++;
                            }
                        }
                    }
                    else
                    {
                        // ID
                        row[Constants.TableExport.ID] = intId;
                        row[Constants.TableExport.TAN_SO] = this.FormatFrequence(dsExcel.Tables[0].Rows[i][Constants.TableExport.TAN_SO].ToString());
                        dtExcel.Rows.Add(row);

                        // Increase id
                        intId++;
                    }

                }
            }
            return dtExcel;
        }

        private double FormatFrequence(string frequence)
        {
            double formatFrequence = default(double);

            // To lower
            frequence = frequence.ToLower();

            // Replace MHZ
            frequence = frequence.Replace("mhz", "");

            // Replace ;
            frequence = frequence.Replace(";", "");

            // Replace "," --> "."
            frequence = frequence.Replace(",", ".");

            bool temp = double.TryParse(frequence, out formatFrequence);

            return formatFrequence * 1000000;
        }

        // Save random value
        private Dictionary<double, double> randomLongtitudeDict = null;

        private Dictionary<double, double> randomLatitudeDict = null;

        /// <summary>
        /// Format vi do
        /// </summary>
        /// <param name="latitude"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private double FormatLatitude(string latitude, string type)
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


        /// <summary>
        /// Format kinh do
        /// </summary>
        /// <param name="longtitude"></param>
        /// <returns></returns>
        private double FormatLongtitude(string longtitude, string type)
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

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void tabRandS_Click(object sender, EventArgs e)
        {

        }

        public Dictionary<int, ArrayList> allListRange;

        private bool CheckErrorRS()
        {

            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();

            Utilities utilities = new Utilities();
            DataTable dtSource = null;
            if (dtRSSource != null)
            {
                dtSource = dtRSSource;
            }
            else
            {
                dtSource = (DataTable)dgRSDetailInformation.DataSource;
            }
            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbRSChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbRSStep.SelectedItem) * 1000;

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnRSCorrectError.Enabled = true;
                    }
                    else
                    {
                        allFreq = utilities.GetAllFrequencyByRange(arrFreqByRow[i], allFreq, ref hasError);

                        IsValidate = IsValidate && !hasError;

                        if (hasError)
                        {
                            // Tan so bi loi
                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                            DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnRSCorrectError.Enabled = true;
                        }
                    }

                    // Check customer
                    #region Ten khach hang
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                    //&& dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Length <= 25)
                    {
                        // Good
                    }
                    else
                    {
                        // Had error
                        hasError = true;
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            "Ten khach hang error";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnRSCorrectError.Enabled = true;

                    }
                    #endregion
                    #region kinhdo_vido
                    // check kinh do vi do
                    //hasError = utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    //IsValidate = IsValidate && !hasError;

                    //if (hasError)
                    //{
                    //    // Kinh do vi do bi loi
                    //    dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                    //    DataGridViewRow row = dgRSDetailInformation.Rows[i];
                    //    row.DefaultCellStyle.BackColor = Color.Yellow;
                    //    btnRSCorrectError.Enabled = true;
                    //}

                    // check kinh do vi do
                    //hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    //IsValidate = IsValidate && !hasError;

                    //if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString()))
                    //{
                    //    if (hasError)
                    //    {
                    //        // Kinh do vi do bi loi
                    //        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                    //        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                    //        row.DefaultCellStyle.BackColor = Color.Yellow;
                    //        btnRSCorrectError.Enabled = true;
                    //    }
                    //}
                    //else
                    //{
                    //    // Do not have error
                    //    IsValidate = true;
                    //}

                    #endregion

                    // check HO HIEU
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Length > 32)
                    {
                        IsValidate = false;

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ToolTipText = "Error";
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnRSCorrectError.Enabled = true;
                    }


                    // check So GP
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.GPNo].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Length > 32)
                    {
                        IsValidate = false;

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ToolTipText = "Error";
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnRSCorrectError.Enabled = true;
                    }
                    //Mau giay phep
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()) )
                    {
                        string maugiayphep = (dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString());
                        if (maugiayphep == Constants.ValueConstant.DAI_TAU)
                        {
                            // Had error
                            IsValidate = false;
                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText =
                                "Ten khach hang error";
                            DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnRSCorrectError.Enabled = true;
                        }
 
                    }

                }
            }
            btnRSCorrectError.Enabled = true;
            return IsValidate;
        }

        private bool CheckErrorTCI()
        {
            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();
            //test save logfile
            List<string> list = new List<string>();
            listTCIExport = new List<string>();

            Utilities utilities = new Utilities();
            DataTable dtSource = null;

            //ArrayList test = utilities.GetColumnName(dtSource);
            if (dtTCISource != null)
            {
                dtSource = dtTCISource;
            }
            else
            {
                dtSource = (DataTable)dgDetailInformation.DataSource;
            }

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate by row
                    // if has error
                    // Set error into datagrid
                    StringBuilder stbuilderRow = new StringBuilder();
                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    //allListRange = new Dictionary<int, ArrayList>();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnCorrectError.Enabled = true;
                    }
                    else
                    {
                        allFreq = utilities.GetAllFrequencyByRange(arrFreqByRow[i], allFreq, ref hasError);

                        IsValidate = IsValidate && !hasError;

                        if (hasError)
                        {
                            // Tan so bi loi
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnCorrectError.Enabled = true;
                        }
                    }

                    // check kinh do vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    IsValidate = IsValidate && !hasError;

                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString()))
                    {
                        if (hasError)
                        {
                            // Kinh do vi do bi loi
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnCorrectError.Enabled = true;
                        }
                    }
                    else
                    {
                        // Do not have error
                        IsValidate = true;
                    }


                    // check ten may
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Length > 50)
                    {
                        IsValidate = false;

                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ToolTipText = "Error";
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnCorrectError.Enabled = true;
                    }

                    // Check dai tan 470 - 806
                    string valueCombobox = cmbChooseFreq.SelectedItem.ToString();
                    if (valueCombobox == Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806)
                    {
                        string maugiayphep = dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().Trim();
                        if (maugiayphep != Constants.ValueConstant.THTS && maugiayphep != Constants.ValueConstant.THTT)
                        {
                            IsValidate = false;

                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ToolTipText = "Error";
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnCorrectError.Enabled = true;
                        }
                    }
                    if (valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174 || valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470)
                    {
                        string maugiayphep = dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().Trim();
                        if (maugiayphep == Constants.ValueConstant.DAI_TAU)
                        {
                            IsValidate = false;

                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ToolTipText = "Error";
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnCorrectError.Enabled = true;
                        }
                    }

                }
            }
            btnCorrectError.Enabled = true;
            return IsValidate;
        }

        private bool CheckErrorGEW()
        {
            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();
            //test save logfile
            List<string> list = new List<string>();
            listGEWExport = new List<string>();

            Utilities utilities = new Utilities();
            DataTable dtSource = null;

            //ArrayList test = utilities.GetColumnName(dtSource);
            if (dtGESource != null)
            {
                dtSource = dtGESource;
            }
            else
            {
                dtSource = (DataTable)dgGEDetailInformation.DataSource;
            }

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate by row
                    // if has error
                    // Set error into datagrid
                    StringBuilder stbuilderRow = new StringBuilder();
                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbGEChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbGEStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    //allListRange = new Dictionary<int, ArrayList>();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnGECorrectError.Enabled = true;
                    }
                    else
                    {
                        allFreq = utilities.GetAllFrequencyByRange(arrFreqByRow[i], allFreq, ref hasError);

                        IsValidate = IsValidate && !hasError;

                        if (hasError)
                        {
                            // Tan so bi loi
                            dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                            DataGridViewRow row = dgGEDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnGECorrectError.Enabled = true;
                        }
                    }

                    // check kinh do vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    IsValidate = IsValidate && !hasError;

                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString()))
                    {
                        if (hasError)
                        {
                            // Kinh do vi do bi loi
                            dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                            DataGridViewRow row = dgGEDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnGECorrectError.Enabled = true;
                        }
                    }
                    else
                    {
                        // Do not have error
                        IsValidate = true;
                    }

                    // Check customer
                    #region Ten khach hang
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                    //&& dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Length <= 25)
                    {
                        // Good
                    }
                    else
                    {
                        // Had error
                        hasError = true;
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            "Ten khach hang error";
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnGECorrectError.Enabled = true;

                    }
                    #endregion
                    // check ten may
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Length > 50)
                    {
                        IsValidate = false;

                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ToolTipText = "Error";
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnGECorrectError.Enabled = true;
                    }

                    // Check dai tan 470 - 790
                    //string valueCombobox = cmbGEChooseFreq.SelectedItem.ToString();
                    //if (valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_470_790)
                    //{
                    //    string maugiayphep = dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().Trim();
                    //    if (maugiayphep != Constants.ValueConstant.THTS && maugiayphep != Constants.ValueConstant.THTT)
                    //    {
                    //        IsValidate = false;

                    //        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ToolTipText = "Error";
                    //        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = "Error";
                    //        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                    //        row.DefaultCellStyle.BackColor = Color.Yellow;
                    //        btnGECorrectError.Enabled = true;
                    //    }
                    //}
                    string valueCombobox = cmbGEChooseFreq.SelectedItem.ToString();
                    if (valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174 || valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470)
                    {
                        if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                        {
                            string maugiayphep = dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString();
                            if (maugiayphep == Constants.ValueConstant.DAI_TAU)
                            {
                                IsValidate = false;
                                dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ToolTipText = "Error";
                                dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText =
                                    "Test thu ErrorText";
                                DataGridViewRow row = dgGEDetailInformation.Rows[i];
                                row.DefaultCellStyle.BackColor = Color.Yellow;
                                btnGECorrectError.Enabled = true;
                            }
                        }
                        
                    }

                }
            }
            btnGECorrectError.Enabled = true;
            return IsValidate;
        }
        /// <summary>
        /// Check error of TCI
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 


        private void btnCheckError_Click(object sender, EventArgs e)
        {
            bool isValidate = CheckErrorTCI();

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Test data
            string inputtest = "000°00'0000\"E";
            string inputtest2 = "000°00'000\"E";


            Utilities utilities = new Utilities();

            bool isboo = utilities.IsKinhdo(inputtest);
            isboo = utilities.IsKinhdo((inputtest2));
        }

        private void btnCorrectError_Click(object sender, EventArgs e)
        {
            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();

            Utilities utilities = new Utilities();
            DataTable dtSource = (DataTable)dgDetailInformation.DataSource;

            bool isMustReBindDataSource = false;

            bool hasUnExpectedError = false;

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate by row
                    // if has error
                    // Set error into datagrid

                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    //allListRange = new Dictionary<int, ArrayList>();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    //// Add arraylist with no error
                    //if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    //{
                    //    allListRange.Add(i, arrFreqByRow[i]);
                    //}

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].Value =
                            utilities.CorrectFrequencyByRange(strStart, dStep, strFreq);
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = string.Empty;
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                        btnCorrectError.Enabled = true;
                    }
                    #region KinhdoVido
                    // check kinh do vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    IsValidate = IsValidate && !hasError;

                    //if (hasError)
                    //{
                    // Kinh do vi do bi loi
                    dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].Value =
                        utilities.CorrectKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString(),
                                                    ref hasError);

                    if (!hasError)
                    {
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            string.Empty;
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                    else
                    {
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            "Unexpected Error.";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnFormat.Enabled = false;
                        hasUnExpectedError = true;
                    }
                    //}
                    //else
                    //{
                    //    dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = string.Empty;
                    //    DataGridViewRow row = dgDetailInformation.Rows[i];
                    //    row.DefaultCellStyle.BackColor = Color.White;
                    //    //btnCorrectError.Enabled = true;
                    //}
                    #endregion
                    // check ten may
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Length > 50)
                    {
                        IsValidate = false;

                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].Value =
                            dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Substring(0, 50);
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ToolTipText = string.Empty;
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ErrorText = string.Empty;
                        //    "Test thu ErrorText";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                        btnCorrectError.Enabled = true;
                    }

                    // Remove row khong phai PTTH
                    string valueCombobox = cmbChooseFreq.SelectedItem.ToString();
                    if (valueCombobox == Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806)
                    {
                        if (dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value != null)
                        {
                            string maugiayphep =
                                dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().
                                    Trim();
                            if (maugiayphep != Constants.ValueConstant.THTS &&
                                maugiayphep != Constants.ValueConstant.THTT)
                            {
                                // Remove row
                                //dgDetailInformation.Rows.RemoveAt(i);
                                dtSource.Rows.RemoveAt(i);
                                isMustReBindDataSource = true;
                            }
                        }
                    }
                    if (valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174 || valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470)
                    {
                        if (dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value != null)
                        {
                            string maugiayphep =
                                dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().
                                    Trim();
                            if (maugiayphep == Constants.ValueConstant.DAI_TAU)
                            {
                                dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value = string.Empty;
                                // Remove row
                                //dgDetailInformation.Rows.RemoveAt(i);
                                dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = string.Empty;
                                //    "Test thu ErrorText";
                                DataGridViewRow row = dgDetailInformation.Rows[i];
                                row.DefaultCellStyle.BackColor = Color.White;
                                btnCorrectError.Enabled = true;
                            }
                        }
                    }

                }
            }
            if (hasUnExpectedError)
            {
                btnFormat.Enabled = false;
                button9.Enabled = false;
            }
            else
            {
                btnFormat.Enabled = true;
                button9.Enabled = true;
                dtTCISource = (DataTable)dgDetailInformation.DataSource;
            }

            if (isMustReBindDataSource)
                dgDetailInformation.DataSource = dtSource;
            button9.Enabled = true;
            btnFormat.Enabled = true;
            dtTCISource = (DataTable)dgDetailInformation.DataSource;
        }

        private void cmbChooseFreq_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> freqRangeForStep = new List<string>();
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_HF_9_30);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_47_50);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_54_68);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_87_108);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_HKHONG_108_138);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_138_174);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_174_230);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_400_463);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_CDMA_806_890);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_EGDSM_890_960);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_GSM_1800_1900);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_3G_2100_2170);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_3G_2620_2680);

            Utilities utilities = new Utilities();
            Dictionary<string, string> dicStep = utilities.CreateFreqAndStepTCI(freqRangeForStep);
            if (cmbChooseFreq.SelectedItem != null)
            {
                string[] listStep = dicStep[cmbChooseFreq.SelectedItem.ToString()].Split(';');

                cmbFreqStep.DataSource = listStep;
            }
        }

        private void btnFormat_Click(object sender, EventArgs e)
        {
            OutFormatBO objFormat = new OutFormatBO();

            dgDetailInformation.DataSource = objFormat.GetTCITableOutput((DataTable)dgDetailInformation.DataSource,
                                                                         allListRange);
        }

        private void btnFormat_Click_1(object sender, EventArgs e)
        {
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            bool isValidate = CheckErrorTCI();
            Utilities util = new Utilities();

            if (isValidate)
            {
                // Check value of combobox
                string valueCombobox = cmbChooseFreq.SelectedItem.ToString();
                double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString()) * 1000;
                //DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;
                DataTable dtGrid = null;
                Utilities utilities = new Utilities();
                string strStart = cmbRSChooseFreq.SelectedItem.ToString();

                if (dtTCISource != null)
                    dtGrid = dtTCISource;
                else
                {
                    dtGrid = (DataTable)dgDetailInformation.DataSource;
                }

                switch (valueCombobox)
                {
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_47_50):
                        #region FREQ_FM_47_50
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 100000;
                                    double freqLower = freqBase - 100000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_54_68):
                        #region FREQ_FM_54_68
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 100000;
                                    double freqLower = freqBase - 100000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_87_108):
                        #region FREQ_FM_87_108
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 100000;
                                    double freqLower = freqBase - 100000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_138_174):
                        #region FREQ_VHF_138_174
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 5000;
                                    double freqLower = freqBase - 5000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_174_230):
                        #region FREQ_VHF_174_230
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (util.MachTDMB(dtGrid.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim()))
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);
                                            double freqBegin = freqBase - 868000;
                                            double freqEnd = freqBase + 868000;

                                            while (freqBegin <= freqEnd)
                                            {
                                                if (!arrFreq.Contains(freqBegin.ToString()))
                                                {
                                                    arrFreqTemp.Add(freqBegin);
                                                    hasChange = true;
                                                }
                                                freqBegin = freqBegin + 100000; // Frequency plus Step
                                            }
                                        }
                                        if (hasChange)
                                        {
                                            // Clear old data
                                            arrFreq.Clear();

                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {

                                        if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() ==
                                            Constants.ValueConstant.THTS)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            for (int j = 0; j < arrFreq.Count; j++)
                                            {
                                                double freqBase = Convert.ToDouble(arrFreq[j]);
                                                double freqBegin = freqBase - 4000000;
                                                double freqEnd = freqBase + 4000000;

                                                while (freqBegin <= freqEnd)
                                                {
                                                    if (!arrFreq.Contains(freqBegin.ToString()))
                                                    {
                                                        arrFreqTemp.Add(freqBegin);
                                                        hasChange = true;
                                                    }
                                                    freqBegin = freqBegin + 100000; // Frequency plus Step
                                                }
                                            }
                                            if (hasChange)
                                            {
                                                // Clear old data
                                                arrFreq.Clear();

                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (arrFreq != null && arrFreq.Count == 2)
                                            {
                                                //arrFreqTemp = arrFreq;
                                                // Create freq base
                                                // fBase = ((f1 + f2)-1)/2;
                                                double freqBase = ((Convert.ToDouble(arrFreq[0]) +
                                                                    Convert.ToDouble(arrFreq[1])) - 1000000) / 2;

                                                double freqBegin = freqBase - 4000000;
                                                double freqEnd = freqBase + 4000000;

                                                while (freqBegin <= freqEnd)
                                                {
                                                    if (!arrFreq.Contains(freqBegin.ToString()))
                                                    {
                                                        arrFreqTemp.Add(freqBegin);
                                                        hasChange = true;
                                                    }
                                                    // Get step
                                                    freqBegin = freqBegin + dStep; // Frequency plus Step
                                                }
                                            }

                                            if (hasChange)
                                            {
                                                // Clear old data
                                                arrFreq.Clear();
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        //
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_400_463):
                        #region FREQ_UHF_400_463
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 5000;
                                    double freqLower = freqBase - 5000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806):
                        #region UHF_470_806
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);
                                            double freqBegin = freqBase - 4000000;
                                            double freqEnd = freqBase + 4000000;

                                            while (freqBegin <= freqEnd)
                                            {
                                                if (!arrFreq.Contains(freqBegin.ToString()))
                                                {
                                                    arrFreqTemp.Add(freqBegin);
                                                    hasChange = true;
                                                }
                                                freqBegin = freqBegin + 100000; // Frequency plus Step
                                            }
                                        }
                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (arrFreq != null && arrFreq.Count == 2)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            // Create freq base
                                            // fBase = ((f1 + f2)-1)/2;
                                            double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;

                                            double freqBegin = freqBase - 4000000;
                                            double freqEnd = freqBase + 4000000;

                                            while (freqBegin <= freqEnd)
                                            {
                                                if (!arrFreq.Contains(freqBegin.ToString()))
                                                {
                                                    arrFreqTemp.Add(freqBegin);
                                                    hasChange = true;
                                                }
                                                // Get step
                                                freqBegin = freqBegin + dStep; // Frequency plus Step
                                            }
                                        }

                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    default:
                        //// Action default
                        //dgDetailInformation.DataSource = objFormat.GetTCITableOutput((DataTable)dgDetailInformation.DataSource,
                        //                                                     allListRange);
                        break;

                }

                // Common action// Action default

                dgDetailInformation.DataSource = null;

                dgDetailInformation.DataSource = objFormat.GetTCITableOutput(dtGrid, allListRange);

                //   dgDetailInformation.DataSource = objFormat.GetTCITableBeforeFormat((DataTable)dgDetailInformation.DataSource,
                //          allListRange);                                                   

                btnShow.Enabled = true;
                btnExport.Enabled = true;
                btnImport.Enabled = true;
                button9.Enabled = true;
                btnFormat.Enabled = true;
                btnCheckError.Enabled = false;
                btnCorrectError.Enabled = false;

                dgDetailInformation.ReadOnly = true;

            }
            else
            {
                // Check error
                btnCheckError.Enabled = true;
                btnImport.Enabled = true;
                btnFormat.Enabled = false;
                btnShow.Enabled = false;
                btnExport.Enabled = false;

            }

        }

        private Form2 form1;
        private List<string> listTCIExport = null;
        private List<string> listGEWExport = null;


        private void btnShow_Click(object sender, EventArgs e)
        {
            //OutFormatBO objFormat = new OutFormatBO();
            double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString());
            if (dgDetailInformation != null && dgDetailInformation.DataSource != null)
            {
                List<string> list = new List<string>();
                listTCIExport = new List<string>();
                DataTable tbTCIInfo = (DataTable)dgDetailInformation.DataSource;
                if (tbTCIInfo != null && tbTCIInfo.Rows.Count > 0)
                {
                    for (int i = 0; i < tbTCIInfo.Rows.Count; i++)
                    {
                        if (tbTCIInfo.Columns.Count > 2)
                        {

                            StringBuilder stbuilderRow = new StringBuilder();
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.ID]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.GPNo]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.MAU_GIAY_PHEP]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.SO_THAM_CHIEU]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.DO_LECH_F]);
                            stbuilderRow.Append(";");

                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.TAN_SO]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.BRAND_UU_TIEN]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.DO_RONG_KENH]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.SO_KENH]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.TEN_KHACH_HANG]);
                            stbuilderRow.Append(";");

                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.HO_HIEU]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.VI_DO]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.KINH_DO]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.TEN_MAY]);

                            list.Add(stbuilderRow.ToString());
                        }
                        else
                        {
                            StringBuilder stbuilderRow = new StringBuilder();
                            stbuilderRow.Append(Convert.ToDouble(tbTCIInfo.Rows[i][Constants.TableExport.TAN_SO]) / 1000000);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(dStep);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append("0");
                            stbuilderRow.Append(";");
                            stbuilderRow.Append("0");
                            stbuilderRow.Append(";");
                            stbuilderRow.Append("0");
                            stbuilderRow.Append(";");
                            stbuilderRow.Append("1");
                            stbuilderRow.Append(";");
                            stbuilderRow.Append("1");
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.ID]);
                            list.Add(stbuilderRow.ToString());
                        }
                    }
                }

                foreach (var openForm in Application.OpenForms)
                {
                    if (openForm.Equals(form1))
                    {

                    }
                    else
                    {
                        form1 = new Form2(list);
                        //form1.Show();
                    }
                }
                form1.Show();
                listTCIExport = list;
                // Enable button show
                btnExport.Enabled = true;
                btnShow.Enabled = true;
                btnImport.Enabled = true;

            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            //   bool isValidate = true;

            // if (isMustCheckRS)
            //     isValidate = CheckErrorTCI();

            //  if (isValidate)
            //   {
            // Check value of combobox
            string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
            //double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString());
            //  DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;


            double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString());
            List<string> list = new List<string>();
            List<string> listSplit = new List<string>();
            listTCIExport = new List<string>();
            DataTable tbTCIInfo = (DataTable)dgDetailInformation.DataSource;
            if (tbTCIInfo != null && tbTCIInfo.Rows.Count > 0)
            {
                for (int i = 0; i < tbTCIInfo.Rows.Count; i++)
                {
                    if (tbTCIInfo.Columns.Count > 2)
                    {

                        StringBuilder stbuilderRow = new StringBuilder();
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.ID]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.GPNo]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.MAU_GIAY_PHEP]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.SO_THAM_CHIEU]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.DO_LECH_F]);
                        stbuilderRow.Append(";");

                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.TAN_SO]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.BRAND_UU_TIEN]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.DO_RONG_KENH]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.SO_KENH]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.TEN_KHACH_HANG]);
                        stbuilderRow.Append(";");

                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.HO_HIEU]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.VI_DO]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.KINH_DO]);
                        stbuilderRow.Append(";");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.TEN_MAY]);

                        list.Add(stbuilderRow.ToString());
                    }
                    else
                    {
                        StringBuilder stbuilderRow = new StringBuilder();
                        stbuilderRow.Append(Convert.ToDouble(tbTCIInfo.Rows[i][Constants.TableExport.TAN_SO]) / 1000000);
                        stbuilderRow.Append(",");
                        stbuilderRow.Append(dStep);
                        stbuilderRow.Append(",");
                        stbuilderRow.Append("0");
                        stbuilderRow.Append(",");
                        stbuilderRow.Append("0");
                        stbuilderRow.Append(",");
                        stbuilderRow.Append("0");
                        stbuilderRow.Append(",");
                        stbuilderRow.Append("1");
                        stbuilderRow.Append(",");
                        stbuilderRow.Append(tbTCIInfo.Rows[i][Constants.TableExport.ID]);
                        list.Add(stbuilderRow.ToString());
                    }
                }
            }
            List<List<string>> allList = SplitIntoChunks(list, 100);
            bool isFirst = true;
            string pathsv = "";
            var charsToRemove = new string[] { "lic", "dbl" };
            int numberfile = 1;
            bool isLic = true;
            foreach (List<string> splList in allList)
            {
                listTCIExport = splList;
                if (listTCIExport != null && listTCIExport.Count > 0)
                {
                    List<string> lines = listTCIExport;
                    string pathSave = default(string);
                    
                    if (isFirst)
                    {
                        // Save file
                        SaveFileDialog dialog = new SaveFileDialog();
                        dialog.Filter = "Export file (*.lic)|*.lic|Export file (*.dbl)|*.dbl";
                        dialog.Title = "Save file type.";



                        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            pathSave = dialog.FileName;
                            pathsv = dialog.FileName;
                        }

                        if (!String.IsNullOrEmpty(pathSave))
                        {
                            Console.WriteLine(pathSave);
                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(pathSave))
                            {
                                foreach (string row in lines)
                                {
                                    // Writer into file
                                    file.WriteLine(row);
                                }
                            }
                        }
                        listTCIExport.Clear();
                        lines.Clear();
                        listSplit.Clear();
                        isFirst = false;
                    }
                    else
                    {
                        
                        string pth = pathsv;
                        Console.WriteLine("pth: " + pth);
                        Console.WriteLine("pathsv: " + pathsv);
                        string extension;
                        extension = Path.GetExtension(pathsv);
                        foreach (var c in charsToRemove)
                        {
                            pth = pth.Replace(c, string.Empty);
                        }
                        if (extension == ".lic")
                        {
                            isLic = true;
                        }
                        else if (extension == ".dbl")
                        {
                            isLic = false;
                        }
                        if (isLic)
                        {
                            pth = pth + "(" + numberfile + ").lic";
                        }
                        else
                        {
                            pth = pth + "(" + numberfile + ").dbl";
                        }
                        numberfile++;

                        if (!String.IsNullOrEmpty(pth))
                        {
                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(pth))
                            {
                                foreach (string row in lines)
                                {
                                    // Writer into file
                                    file.WriteLine(row);
                                }
                            }
                        }
                        listTCIExport.Clear();
                        lines.Clear();
                        listSplit.Clear();
                    }

                }
            }
                

                        MessageBox.Show("Convert Sucessful", "Message box", MessageBoxButtons.OK);
                        btnImport.Enabled = true;
                        btnShow.Enabled = true;
                        btnCorrectError.Enabled = false;
                        btnFormat.Enabled = true;
                        button9.Enabled = true;
                        btnExport.Enabled = true;
                  

        }

        public static List<List<T>> SplitIntoChunks<T>(List<T> list, int chunkSize)
        {
            if (chunkSize <= 0)
            {
                throw new ArgumentException("chunkSize must be greater than 0.");
            }

            List<List<T>> retVal = new List<List<T>>();
            int index = 0;
            while (index < list.Count)
            {
                int count = list.Count - index > chunkSize ? chunkSize : list.Count - index;
                retVal.Add(list.GetRange(index, count));

                index += chunkSize;
            }

            return retVal;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmbFreqStep_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbRSChooseFreq_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> freqRangeForStep = new List<string>();
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_HF_9_30);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_47_50);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_54_68);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_87_108);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_HKHONG_108_138);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_138_174);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_174_230);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_400_463);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_CDMA_806_890);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_EGDSM_890_960);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_GSM_1800_1900);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_3G_2100_2170);
            freqRangeForStep.Add(Constants.FreqAndStep.FrequencyDisplay.FREQ_3G_2620_2680);

            Utilities utilities = new Utilities();
            Dictionary<string, string> dicStep = utilities.CreateFreqAndStepRS(freqRangeForStep);
            if (cmbRSChooseFreq.SelectedItem != null)
            {
                string[] listStep = dicStep[cmbRSChooseFreq.SelectedItem.ToString()].Split(';');

                cmbRSStep.DataSource = listStep;
            }
        }
        private DataTable dtRSSource = null;
        private DataTable dtTCISource = null;
        private DataTable dtGESource = null;

        private void btnRSCorrectError_Click(object sender, EventArgs e)
        {
            Utilities utilities = new Utilities();
            DataTable dtSource = (DataTable)dgRSDetailInformation.DataSource;
            bool hasError = false;
            bool IsValidate = true;
            bool IsValidateAll = true;

            bool hasUnExpectedError = false;

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate Frequency
                    string strStart = cmbRSChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbRSStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].Value =
                            utilities.CorrectFrequencyByRange(strStart, dStep, strFreq);
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = string.Empty;
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                        btnRSCorrectError.Enabled = true;

                        // Ep gia tri
                        IsValidate = true;
                        hasError = false;
                    }

                    //if (hasError)
                    //{
                    //    // Tan so bi loi
                    //    dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].Value =
                    //        utilities.CorrectFrequencyByRange(strStart, dStep, strFreq);
                    //    dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = string.Empty;
                    //    DataGridViewRow row = dgRSDetailInformation.Rows[i];
                    //    row.DefaultCellStyle.BackColor = Color.White;
                    //    //btnRSCorrectError.Enabled = true;
                    //    // Ep gia tri.
                    //    hasError = false;

                    //    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                    //                                                                        ref hasError);

                    //    // Add arraylist with no error
                    //    if (!allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    //    {
                    //        allListRange.Add(i, arrFreqByRow[i]);
                    //    }
                    //}

                    // Check customer
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                    {
                        if (dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Length > 25)
                        {
                            //string tenkhachhangCut =
                            //dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Substring(0, 25);

                            //dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].Value = tenkhachhangCut;
                            //dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            //    string.Empty;
                            //DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            //row.DefaultCellStyle.BackColor = Color.White;
                        }
                    }
                    else
                    {
                        // Had error
                        hasError = true;
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            "Ten khach hang blank";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnRSCorrectError.Enabled = true;

                    }
                    IsValidate = IsValidate && !hasError;

                    // check kinh do vi do
                    #region Kinhdo, vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());



                    //if (hasError)
                    //{
                    // Kinh do vi do bi loi
                    dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].Value =
                        utilities.CorrectKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString(),
                                                    ref hasError);

                    if (!hasError)
                    {
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            string.Empty;
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                    else
                    {
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            "Unexpected Error.";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnRSFormat.Enabled = false;
                        hasUnExpectedError = true;
                    }
                    IsValidate = IsValidate && !hasError;

                    #endregion

                    // check HO HIEU
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString()))
                    {
                        if (dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Length > 32)
                        {
                            string hohieuCut =
                               dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString().Substring(0, 32);

                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].Value = hohieuCut;
                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ErrorText =
                                string.Empty;
                            DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.White;
                        }
                    }
                    else
                    {

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].Value = string.Empty;
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ErrorText =
                            string.Empty;
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                    }

                    // check So GP
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.GPNo].ToString()))
                    {
                        if (dtSource.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Length > 32)
                        {
                            string gpNoCut =
                               dtSource.Rows[i][Constants.TableExport.GPNo].ToString().Substring(0, 32);

                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].Value = gpNoCut;
                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ErrorText =
                                string.Empty;
                            DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.White;
                        }
                    }
                    else
                    {

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].Value = string.Empty;
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ErrorText =
                            string.Empty;
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                    }

                    //Mau giay phep
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                    {
                        if (dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString() == Constants.ValueConstant.DAI_TAU)
                        {

                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value = String.Empty;
                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = string.Empty;
                            DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.White;
                        }
                    }
                    else
                    {

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value = string.Empty;
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText =
                            string.Empty;
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                    //IsValidate = IsValidate && !hasError;

                    //// If is not validate 
                    //// --> Disable button Export
                    //if (!IsValidate)
                    //    btnRSExport.Enabled = false;
                }

                // Check all row is OK
                IsValidateAll = IsValidateAll && IsValidate;
                if (IsValidateAll)
                {
                   
                    btnRSFormat.Enabled = true;
                    btnRSCorrectError.Enabled = false;
                    btnRSCheckError.Enabled = false;
                    button8.Enabled = true;
                   

                    //// Get record by frequence and save into datable to export file excel.
                    //this.GetFrequenceBeforeFormat();

                    // Save data source
                    dtRSSource = (DataTable)dgRSDetailInformation.DataSource;

                    // Has been corrected error --> must not be checked
                    isMustCheckRS = false;

                }
                else
                {
                    btnRSExport.Enabled = false;
                    btnRSFormat.Enabled = false;
                    btnRSCorrectError.Enabled = true;
                    button8.Enabled = false;
                }
            }
        }

        private bool isMustCheckRS = true;
        private bool isMustCheckGEW = true;
        private bool canExportCSV = false;
        private bool isExportTran = true;

        private string GetSheetName(string fileName)
        {
            string returnSheet = default(string);
            try
            {
                Excel.Application ExcelObj = new Excel.Application();
                Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                Excel.Sheets sheets = theWorkbook.Worksheets;
                //for (int i = 1; i < sheets.Count + 1; i++)
                //{
                //    Excel.Worksheet sheetA = (Excel.Worksheet)sheets[i];
                //    string s = sheetA.Name;
                //}
                Excel.Worksheet sheetA = (Excel.Worksheet)sheets[1];
                returnSheet = sheetA.Name;
            }
            catch
            {
                MessageBox.Show("Has unknow exception when read file Excel", "UnExpected error", MessageBoxButtons.OK);
                return "Error";
            }


            return returnSheet;


        }

        private void btnRSImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel file (*.xls)|*.xls";
            dialog.Title = "Open file Excel convert.";

            // Clean dtRSSource
            dtRSSource = null;

            Utilities utilities = new Utilities();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string sheetName = this.GetSheetName(dialog.FileName);

                DataSet dsExcel = utilities.GetAllDataFromFileExcel(dialog.FileName, sheetName);

                if (dsExcel != null
                && dsExcel.Tables != null
                && dsExcel.Tables.Count > 0
                && dsExcel.Tables[0].Rows.Count > 0)
                {
                    //dgRSDetailInformation.DataSource = null;
                    if (dgRSDetailInformation.DataSource != null)
                    {
                        dgRSDetailInformation.DataSource = null;
                        dgRSDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    else
                    {
                        dgRSDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    if (allListRange != null && allListRange.Count > 0)
                    {
                        allListRange.Clear();
                    }
                }

                // Enable button
                btnRSCheckError.Enabled = true;
                btnRSFormat.Enabled = false;
                btnRSCorrectError.Enabled = false;
                button8.Enabled = false;
                btnRSExport.Enabled = false;
                btnRSShow.Enabled = false;
                imgRS.Visible = false;
                // Reset can export value
                canExportCSV = false;

            }
        }

        private void btnRSCheckError_Click(object sender, EventArgs e)
        {
            bool isValidate = CheckErrorRS();
        }

        //// Define a variable for check export
        //private bool isRSExportCSV = false;

        private void btnRSFormat_Click(object sender, EventArgs e)
        {
            //Set can export CSV
            canExportCSV = true;

            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            bool isValidate = true;

            if (isMustCheckRS)
            {
                isValidate = CheckErrorRS();
            }

            if (isValidate)
            {
                // Check value of combobox
                string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
                double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString()) * 1000;
                DataTable dtGrid = null;
                Utilities utilities = new Utilities();
                string strStart = cmbRSChooseFreq.SelectedItem.ToString();
                // Convert NewStart to OldStart
                strStart = strStart.Replace("KHz", "");
                strStart = strStart.Replace("MHz", "");
                strStart = strStart.Replace(" ", "");
                strStart = strStart.Replace("-", "_");

                if (dtRSSource != null)
                    dtGrid = dtRSSource;
                else
                {
                    dtGrid = (DataTable)dgRSDetailInformation.DataSource;
                }
                for (int i = 0; i < dtGrid.Rows.Count; i++)
                {
                    string strFreq = dtGrid.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    bool hasError = false;

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }
                }

                switch (valueCombobox)
                {
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_47_50):
                        #region FREQ_FM_47_50
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_54_68):
                        #region FREQ_FM_54_68
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_87_108):
                        #region FREQ_FM_87_108
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_138_174):
                        #region FREQ_VHF_138_174
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_174_230):
                        #region FREQ_VHF_174_230
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    Utilities util = new Utilities();
                                    if (util.MachTDMB(dtGrid.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim()))
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            // Clear old data
                                            arrFreq.Clear();

                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            for (int j = 0; j < arrFreq.Count; j++)
                                            {
                                                double freqBase = Convert.ToDouble(arrFreq[j]);

                                            }
                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (arrFreq != null && arrFreq.Count == 2)
                                            {
                                                //arrFreqTemp = arrFreq;
                                                // Create freq base
                                                // fBase = ((f1 + f2)-1)/2;
                                                double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;


                                            }

                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                    }
                                    if (hasChange)
                                    {
                                        allListRange[i] = arrFreq;
                                    }
                                }
                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_400_463):
                        #region FREQ_UHF_400_463
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806):
                        #region FREQ_UHF_470_806
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (arrFreq != null && arrFreq.Count == 2)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            // Create freq base
                                            // fBase = ((f1 + f2)-1)/2;
                                            double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;


                                        }

                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    default:
                        //// Action default
                        //dgDetailInformation.DataSource = objFormat.GetTCITableOutput((DataTable)dgDetailInformation.DataSource,
                        //                                                     allListRange);
                        break;

                }

                // Common action// Action default
                dgRSDetailInformation.DataSource = null;
                dgRSDetailInformation.DataSource = objFormat.GetRSTableOutput(dtGrid, allListRange);

                btnRSShow.Enabled = true;
                btnRSExport.Enabled = true;
                button8.Enabled = true;
                btnRSImport.Enabled = true;
                btnRSFormat.Enabled = true;
                btnRSCheckError.Enabled = false;
                btnRSCorrectError.Enabled = false;
                dgRSDetailInformation.ReadOnly = true;

                // Set can export 
                canExportCSV = true;
                isMustCheckRS = false;

            }
            else
            {
                // Check error
                btnRSCheckError.Enabled = true;
                btnRSImport.Enabled = true;
                btnRSFormat.Enabled = false;
                btnRSShow.Enabled = false;

            }
        }

        private void GetFrequenceBeforeFormat()
        {
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            bool isValidate = true;

            if (isMustCheckRS)
                isValidate = CheckErrorRS();

            if (isValidate)
            {
                // Check value of combobox
                string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
                double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString()) * 1000;
                //DataTable dtGrid = (DataTable)dgRSDetailInformation.DataSource;
                DataTable dtGrid = null;
                if (dtRSSource != null)
                    dtGrid = dtRSSource;
                else
                {
                    dtGrid = (DataTable)dgRSDetailInformation.DataSource;
                }

                Utilities utilities = new Utilities();
                string strStart = cmbRSChooseFreq.SelectedItem.ToString();
                // Convert NewStart to OldStart
                strStart = strStart.Replace("KHz", "");
                strStart = strStart.Replace("MHz", "");
                strStart = strStart.Replace(" ", "");
                strStart = strStart.Replace("-", "_");

                for (int i = 0; i < dtGrid.Rows.Count; i++)
                {
                    string strFreq = dtGrid.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    bool hasError = false;

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }
                }

                switch (valueCombobox)
                {
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_47_50):
                        #region FREQ_FM_47_50
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 100000;
                                    double freqLower = freqBase - 100000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_54_68):
                        #region FREQ_FM_54_68
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 100000;
                                    double freqLower = freqBase - 100000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_87_108):
                        #region FREQ_FM_87_108
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 100000;
                                    double freqLower = freqBase - 100000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_138_174):
                        #region FREQ_VHF_138_174
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 5000;
                                    double freqLower = freqBase - 5000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_174_230):
                        #region FREQ_VHF_174_230
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    Utilities util = new Utilities();
                                    if (util.MachTDMB(dtGrid.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim()))
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);
                                            double freqBegin = freqBase - 868000;
                                            double freqEnd = freqBase + 868000;

                                            while (freqBegin <= freqEnd)
                                            {
                                                if (!arrFreq.Contains(freqBegin.ToString()))
                                                {
                                                    arrFreqTemp.Add(freqBegin);
                                                    hasChange = true;
                                                }
                                                freqBegin = freqBegin + 100000; // Frequency plus Step
                                            }
                                        }
                                        if (hasChange)
                                        {
                                            // Clear old data
                                            arrFreq.Clear();

                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            for (int j = 0; j < arrFreq.Count; j++)
                                            {
                                                double freqBase = Convert.ToDouble(arrFreq[j]);
                                                double freqBegin = freqBase - 4000000;
                                                double freqEnd = freqBase + 4000000;

                                                while (freqBegin <= freqEnd)
                                                {
                                                    if (!arrFreq.Contains(freqBegin.ToString()))
                                                    {
                                                        arrFreqTemp.Add(freqBegin);
                                                        hasChange = true;
                                                    }
                                                    freqBegin = freqBegin + 100000; // Frequency plus Step
                                                }
                                            }
                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (arrFreq != null && arrFreq.Count == 2)
                                            {
                                                //arrFreqTemp = arrFreq;
                                                // Create freq base
                                                // fBase = ((f1 + f2)-1)/2;
                                                double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;

                                                double freqBegin = freqBase - 4000000;
                                                double freqEnd = freqBase + 4000000;

                                                while (freqBegin <= freqEnd)
                                                {
                                                    if (!arrFreq.Contains(freqBegin.ToString()))
                                                    {
                                                        arrFreqTemp.Add(freqBegin);
                                                        hasChange = true;
                                                    }
                                                    // Get step
                                                    freqBegin = freqBegin + dStep; // Frequency plus Step
                                                }
                                            }

                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                    }
                                    if (hasChange)
                                    {
                                        allListRange[i] = arrFreq;
                                    }
                                }
                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_400_463):
                        #region FREQ_UHF_400_463
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);
                                    double freqUpper = freqBase + 5000;
                                    double freqLower = freqBase - 5000;

                                    if (!arrFreq.Contains(freqUpper.ToString()))
                                    {
                                        //arrFreq.Add(freqUpper.ToString());
                                        arrFreqTemp.Add(freqUpper.ToString());
                                        hasChange = true;
                                    }

                                    if (!arrFreq.Contains(freqLower.ToString()))
                                    {
                                        //arrFreq.Add(freqLower.ToString());
                                        arrFreqTemp.Add(freqLower.ToString());
                                        hasChange = true;
                                    }
                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806):
                        #region FREQ_UHF_470_806
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);
                                            double freqBegin = freqBase - 100000;//4000000;
                                            double freqEnd = freqBase + 100000;//4000000;

                                            while (freqBegin <= freqEnd)
                                            {
                                                if (!arrFreq.Contains(freqBegin.ToString()))
                                                {
                                                    arrFreqTemp.Add(freqBegin);
                                                    hasChange = true;
                                                }
                                                freqBegin = freqBegin + 100000; // Frequency plus Step
                                            }
                                        }
                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (arrFreq != null && arrFreq.Count == 2)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            // Create freq base
                                            // fBase = ((f1 + f2)-1)/2;
                                            double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;

                                            double freqBegin = freqBase - 4000000;
                                            double freqEnd = freqBase + 4000000;

                                            while (freqBegin <= freqEnd)
                                            {
                                                if (!arrFreq.Contains(freqBegin.ToString()))
                                                {
                                                    arrFreqTemp.Add(freqBegin);
                                                    hasChange = true;
                                                }
                                                // Get step
                                                freqBegin = freqBegin + dStep; // Frequency plus Step
                                            }
                                        }

                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    default:
                        //// Action default
                        //dgDetailInformation.DataSource = objFormat.GetTCITableOutput((DataTable)dgDetailInformation.DataSource,
                        //                                                     allListRange);
                        break;

                }

                // Common action// Action default
                dgRSDetailInformation.DataSource = null;
                dgRSDetailInformation.DataSource = objFormat.GetRSTableBeforeFormat(dtGrid, allListRange);

                btnRSShow.Enabled = true;
                btnRSExport.Enabled = true;
                button8.Enabled = true;
                btnRSImport.Enabled = true;
                btnRSFormat.Enabled = true;
                btnRSCheckError.Enabled = false;
                btnRSCorrectError.Enabled = false;

                dgRSDetailInformation.ReadOnly = true;
            }
            else
            {
                // Check error
                btnRSCheckError.Enabled = true;
                btnRSImport.Enabled = true;
                btnRSFormat.Enabled = false;
                btnRSShow.Enabled = false;

            }
        }

        private ConfirmExport confirmExport;
        private GEWConfirmExport GEWconfirmExport;
        //  private abtAbout AboutBox;
        private void btnRSExport_Click(object sender, EventArgs e)
        {
            if (dgRSDetailInformation != null && dgRSDetailInformation.DataSource != null)
            {
                //List<string> list = new List<string>();

                //Utilities utils = new Utilities();
                //DataTable dtRS = utils.GetTemplateTableRS();

                listTCIExport = new List<string>();
                DataTable tbRSInfo = (DataTable)dgRSDetailInformation.DataSource;
                string frequenceRange = cmbRSChooseFreq.SelectedItem.ToString();

                foreach (var openForm in Application.OpenForms)
                {
                    if (openForm.Equals(confirmExport))
                    {

                    }
                    else
                    {
                        confirmExport = new ConfirmExport(tbRSInfo, canExportCSV, frequenceRange);
                        //form1.Show();
                    }
                }
                confirmExport.Show();

                // Enable button show
                btnRSExport.Enabled = true;
                btnRSShow.Enabled = true;
                btnRSImport.Enabled = true;
                btnRSFormat.Enabled = true;
                button8.Enabled = true;

            }
        }


        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            this.GetFrequenceBeforeFormat();
            canExportCSV = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            bool isValidate = CheckErrorTCI();
            Utilities util = new Utilities();

            if (isValidate)
            {
                // Check value of combobox
                string valueCombobox = cmbChooseFreq.SelectedItem.ToString();
                double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString()) * 1000;
                // DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;
                DataTable dtGrid = null;
                Utilities utilities = new Utilities();
                string strStart = cmbRSChooseFreq.SelectedItem.ToString();

                if (dtTCISource != null)
                    dtGrid = dtTCISource;
                else
                {
                    dtGrid = (DataTable)dgDetailInformation.DataSource;
                }

                switch (valueCombobox)
                {
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_47_50):
                        #region FREQ_FM_47_50
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_54_68):
                        #region FREQ_FM_54_68
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_FM_87_108):
                        #region FREQ_FM_87_108
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_138_174):
                        #region FREQ_VHF_138_174
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_VHF_174_230):
                        #region FREQ_VHF_174_230
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (util.MachTDMB(dtGrid.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim()))
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            // Clear old data
                                            arrFreq.Clear();

                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {

                                        if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() ==
                                            Constants.ValueConstant.THTS)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            for (int j = 0; j < arrFreq.Count; j++)
                                            {
                                                double freqBase = Convert.ToDouble(arrFreq[j]);

                                            }
                                            if (hasChange)
                                            {
                                                // Clear old data
                                                arrFreq.Clear();

                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (arrFreq != null && arrFreq.Count == 2)
                                            {
                                                //arrFreqTemp = arrFreq;
                                                // Create freq base
                                                // fBase = ((f1 + f2)-1)/2;
                                                double freqBase = ((Convert.ToDouble(arrFreq[0]) +
                                                                    Convert.ToDouble(arrFreq[1])) - 1000000) / 2;

                                            }

                                            if (hasChange)
                                            {
                                                // Clear old data
                                                arrFreq.Clear();
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        //
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_400_463):
                        #region FREQ_UHF_400_463
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806):
                        #region UHF_470_806
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (arrFreq != null && arrFreq.Count == 2)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            // Create freq base
                                            // fBase = ((f1 + f2)-1)/2;
                                            double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;


                                        }

                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    default:
                        //// Action default
                        //dgDetailInformation.DataSource = objFormat.GetTCITableOutput((DataTable)dgDetailInformation.DataSource,
                        //                                                     allListRange);
                        break;

                }

                // Common action// Action default
                dgDetailInformation.DataSource = null;

                dgDetailInformation.DataSource = objFormat.GetTCITableOutput_DFSCAN(dtGrid, allListRange);

                btnShow.Enabled = true;
                btnExport.Enabled = true;
                btnImport.Enabled = true;
                btnFormat.Enabled = true;

                // btnFormat.Enabled = false;
                btnCheckError.Enabled = false;
                btnCorrectError.Enabled = false;

                dgDetailInformation.ReadOnly = true;


            }
            else
            {
                // Check error
                btnCheckError.Enabled = true;
                btnImport.Enabled = true;
                btnFormat.Enabled = false;
                btnShow.Enabled = false;
                btnExport.Enabled = false;

            }
        }

        private void cmbRSStep_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dgDetailInformation_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void changeTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Database update (*.dbs)|*.dbs";
            dialog.Title = "Open file to update database";
            dialog.ShowDialog();
        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {

            
        }



        private void test_Click_1(object sender, EventArgs e)
        {
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            //   bool isValidate = true;

            // if (isMustCheckRS)
            //     isValidate = CheckErrorTCI();

            //  if (isValidate)
            //   {
            // Check value of combobox
            string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
            double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString());
            //  DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;
            DataTable dtGrid = null;
            if (dtTCISource != null)
                dtGrid = dtTCISource;
            else
            {
                dtGrid = (DataTable)dgDetailInformation.DataSource;
            }

            // Common action// Action default
            dgDetailInformation.DataSource = null;

            dgDetailInformation.DataSource = objFormat.GetTCITableBeforeFormat(dtGrid, allListRange);
            btnFormat.Enabled = true;
            //  dgDetailInformation.ReadOnly = true;


            //  }  
            bool isValidate = CheckErrorTCI();
            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();

            Utilities utilities = new Utilities();
            DataTable dtSource = (DataTable)dgDetailInformation.DataSource;

            bool isMustReBindDataSource = false;

            bool hasUnExpectedError = false;

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate by row
                    // if has error
                    // Set error into datagrid

                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    //  double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    //allListRange = new Dictionary<int, ArrayList>();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    //// Add arraylist with no error
                    //if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    //{
                    //    allListRange.Add(i, arrFreqByRow[i]);
                    //}

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].Value =
                            utilities.CorrectFrequencyByRange(strStart, dStep, strFreq);
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = string.Empty;
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                        btnCorrectError.Enabled = true;
                    }
                    #region KinhdoVido
                    // check kinh do vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    IsValidate = IsValidate && !hasError;

                    //if (hasError)
                    //{
                    // Kinh do vi do bi loi
                    dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].Value =
                        utilities.CorrectKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString(),
                                                    ref hasError);

                    if (!hasError)
                    {
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            string.Empty;
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                    else
                    {
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            "Unexpected Error.";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnFormat.Enabled = false;
                        hasUnExpectedError = true;
                    }
                    //}
                    //else
                    //{
                    //    dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = string.Empty;
                    //    DataGridViewRow row = dgDetailInformation.Rows[i];
                    //    row.DefaultCellStyle.BackColor = Color.White;
                    //    //btnCorrectError.Enabled = true;
                    //}
                    #endregion
                    // check ten may
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Length > 50)
                    {
                        IsValidate = false;

                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].Value =
                            dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Substring(0, 50);
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ToolTipText = string.Empty;
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ErrorText = string.Empty;
                        //    "Test thu ErrorText";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                        btnCorrectError.Enabled = true;
                    }

                    // Remove row khong phai PTTH
                    //    string valueCombobox = cmbChooseFreq.SelectedItem.ToString();
                    if (valueCombobox == Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806)
                    {
                        if (dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value != null)
                        {
                            string maugiayphep =
                                dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().
                                    Trim();
                            if (maugiayphep != Constants.ValueConstant.THTS &&
                                maugiayphep != Constants.ValueConstant.THTT)
                            {
                                // Remove row
                                //dgDetailInformation.Rows.RemoveAt(i);
                                dtSource.Rows.RemoveAt(i);
                                isMustReBindDataSource = true;
                            }
                        }
                    }

                }
            }
            if (hasUnExpectedError)
            {
                btnFormat.Enabled = false;
            }
            else
            {
                btnFormat.Enabled = true;
                dtTCISource = (DataTable)dgDetailInformation.DataSource;
            }

            if (isMustReBindDataSource)
                dgDetailInformation.DataSource = dtSource;
            button9.Enabled = true;
            btnFormat.Enabled = true;
        }



        private void button10_Click(object sender, EventArgs e)
        {
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            bool isValidate = true;

            //if (isMustCheckRS)
            //       isValidate = CheckErrorRS();

            //    if (isValidate)
            //   {
            // Check value of combobox
            string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
            double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString());
            //DataTable dtGrid = (DataTable)dgRSDetailInformation.DataSource;
            DataTable dtGrid = null;
            if (dtRSSource != null)
                dtGrid = dtRSSource;
            else
            {
                dtGrid = (DataTable)dgRSDetailInformation.DataSource;
            }

            // Common action// Action default

            dgRSDetailInformation.DataSource = null;
            dgRSDetailInformation.DataSource = objFormat.GetTCITableBeforeFormat(dtGrid, allListRange);
            //   }

        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var testhelp = new abtAbout();
            testhelp.Show();
        }

        private void btnRSShow_Click(object sender, EventArgs e)
        {

        }

        private void tCIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnGEExport.SelectedIndex = 0;
        }

        private void rodToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnGEExport.SelectedIndex = 1;
        }

        private void button10_Click_1(object sender, EventArgs e)
        {

            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            //   bool isValidate = true;

            // if (isMustCheckRS)
            //     isValidate = CheckErrorTCI();

            //  if (isValidate)
            //   {
            // Check value of combobox
            // string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
            //double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString());
            //  DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;


            //  double dStep1 = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString());
            List<string> list = new List<string>();
            listTCIExport = new List<string>();
            DataTable tbTCIInfo = (DataTable)dgDetailInformation.DataSource;

            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();
            //test save logfile
            // List<string> list = new List<string>();
            //  listTCIExport = new List<string>();

            Utilities utilities = new Utilities();


            DataTable dtSource = (DataTable)dgDetailInformation.DataSource;


            //ArrayList test = utilities.GetColumnName(dtSource);

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 1; i < dtSource.Rows.Count; i++)
                {
                    // Validate by row
                    // if has error
                    // Set error into datagrid
                    StringBuilder stbuilderRow = new StringBuilder();
                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    //allListRange = new Dictionary<int, ArrayList>();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                                       ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnCorrectError.Enabled = true;

                        //     stbuilderRow.Append("loi cot tan so, dong thu ");
                        stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                        // stbuilderRow.Insert(26, i+1);
                        list.Add(stbuilderRow.ToString());
                    }
                    else
                    {
                        allFreq = utilities.GetAllFrequencyByRange(arrFreqByRow[i], allFreq, ref hasError);

                        IsValidate = IsValidate && !hasError;

                        if (hasError)
                        {
                            // Tan so bi loi
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnCorrectError.Enabled = true;

                            //   stbuilderRow.Append("loi cot tan so, dong thu ");
                            stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                            //stbuilderRow.Insert(26, i+1);
                            list.Add(stbuilderRow.ToString());
                        }
                    }

                    // check kinh do vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    IsValidate = IsValidate && !hasError;

                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString()))
                    {
                        if (hasError)
                        {
                            // Kinh do vi do bi loi
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnCorrectError.Enabled = true;

                            //    stbuilderRow.Append("loi cot kinh do/vi do, dong thu ");
                            stbuilderRow.AppendFormat("Error dong thu {0}, cot kinh do/vi do ", i + 2);
                            // stbuilderRow.Insert(33, i + 1);
                            list.Add(stbuilderRow.ToString());
                        }
                    }
                    else
                    {
                        // Do not have error
                        IsValidate = true;
                    }


                    // check ten may
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Length > 50)
                    {
                        IsValidate = false;

                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ToolTipText = "Error";
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ErrorText = "Test thu ErrorText";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnCorrectError.Enabled = true;

                        //  stbuilderRow.Append("loi cot ten may, dong thu ");
                        stbuilderRow.AppendFormat("Error dong thu {0}, cot ten may", i + 2);
                        //stbuilderRow.Insert(27, i + 1);
                        list.Add(stbuilderRow.ToString());
                    }

                    // Check dai tan 470 - 806
                    string valueCombobox = cmbChooseFreq.SelectedItem.ToString();
                    if (valueCombobox == Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806)
                    {
                        string maugiayphep = dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().Trim();
                        if (maugiayphep != Constants.ValueConstant.THTS && maugiayphep != Constants.ValueConstant.THTT)
                        {
                            IsValidate = false;

                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ToolTipText = "Error";
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            btnCorrectError.Enabled = true;

                            // stbuilderRow.Append("loi cot mau giay phep, dong thu ");
                            stbuilderRow.AppendFormat("Error dong thu {0}, cot mau giay phep", i + 2);
                            // stbuilderRow.Insert(33, i + 1);
                            list.Add(stbuilderRow.ToString());
                        }
                    }

                }
            }
            // return IsValidate;

            listTCIExport = list;

            // Get source
            if (listTCIExport != null && listTCIExport.Count > 0)
            {
                List<string> lines = listTCIExport;

                // Save file
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Export file (*.txt)|*.txt|Export file (*.doc)|*.doc";
                dialog.Title = "Save file type.";

                string pathSave = default(string);

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    pathSave = dialog.FileName;
                }

                if (!String.IsNullOrEmpty(pathSave))
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(pathSave))
                    {
                        foreach (string line in lines)
                        {
                            // Writer into file
                            file.WriteLine(line);
                        }
                    }

                    MessageBox.Show("Save logfile Sucessful", "Message box", MessageBoxButtons.OK);

                }

            }


        }

        private void button11_Click(object sender, EventArgs e)
        {
            allListRange = new Dictionary<int, ArrayList>();
            //   bool isValidate = true;

            // if (isMustCheckRS)
            //     isValidate = CheckErrorTCI();

            //  if (isValidate)
            //   {
            // Check value of combobox
            // string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
            //double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString());
            //  DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;


            //  double dStep1 = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString());
            List<string> list = new List<string>();
            listTCIExport = new List<string>();
            DataTable tbTCIInfo = (DataTable)dgRSDetailInformation.DataSource;


            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();

            Utilities utilities = new Utilities();
            DataTable dtSource = (DataTable)dgRSDetailInformation.DataSource;

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate Frequency
                    StringBuilder stbuilderRow = new StringBuilder();
                    bool hasError = false;
                    string strStart = cmbRSChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbRSStep.SelectedItem) * 1000;

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                        list.Add(stbuilderRow.ToString());

                    }
                    else
                    {
                        allFreq = utilities.GetAllFrequencyByRange(arrFreqByRow[i], allFreq, ref hasError);

                        IsValidate = IsValidate && !hasError;

                        if (hasError)
                        {
                            // Tan so bi loi
                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                            DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;

                            stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                            list.Add(stbuilderRow.ToString());
                        }
                    }

                    // Check customer
                    #region Ten khach hang
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                    //&& dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Length <= 25)
                    {
                        // Good
                    }
                    else
                    {
                        // Had error
                        hasError = true;
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            "Ten khach hang error";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot ten khach hang", i + 2);
                        list.Add(stbuilderRow.ToString());

                    }
                    #endregion
                    #region kinhdo_vido
                    // check kinh do vi do
                    //hasError = utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    //IsValidate = IsValidate && !hasError;

                    //if (hasError)
                    //{
                    //    // Kinh do vi do bi loi
                    //    dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                    //    DataGridViewRow row = dgRSDetailInformation.Rows[i];
                    //    row.DefaultCellStyle.BackColor = Color.Yellow;
                    //    btnRSCorrectError.Enabled = true;
                    //}

                    // check kinh do vi do
                    //hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    //IsValidate = IsValidate && !hasError;

                    //if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString()))
                    //{
                    //    if (hasError)
                    //    {
                    //        // Kinh do vi do bi loi
                    //        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                    //        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                    //        row.DefaultCellStyle.BackColor = Color.Yellow;
                    //        btnRSCorrectError.Enabled = true;
                    //    }
                    //}
                    //else
                    //{
                    //    // Do not have error
                    //    IsValidate = true;
                    //}

                    #endregion

                    // check HO HIEU
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Length > 32)
                    {
                        IsValidate = false;

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ToolTipText = "Error";
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot ho hieu", i + 2);
                        list.Add(stbuilderRow.ToString());
                    }


                    // check So GP
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.GPNo].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Length > 32)
                    {
                        IsValidate = false;

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ToolTipText = "Error";
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot so giap phep", i + 2);
                        list.Add(stbuilderRow.ToString());
                    }
                }
            }


            listTCIExport = list;

            // Get source
            if (listTCIExport != null && listTCIExport.Count > 0)
            {
                List<string> lines = listTCIExport;

                // Save file
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Export file (*.txt)|*.txt|Export file (*.doc)|*.doc";
                dialog.Title = "Save file type.";

                string pathSave = default(string);

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    pathSave = dialog.FileName;
                }

                if (!String.IsNullOrEmpty(pathSave))
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(pathSave))
                    {
                        foreach (string line in lines)
                        {
                            // Writer into file
                            file.WriteLine(line);
                        }
                    }

                    MessageBox.Show("Save logfile Sucessful", "Message box", MessageBoxButtons.OK);

                }

            }



        }

        private void saveTCILogfileToolStripMenuItem_Click(object sender, EventArgs e)
        {

            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            //   bool isValidate = true;

            // if (isMustCheckRS)
            //     isValidate = CheckErrorTCI();

            //  if (isValidate)
            //   {
            // Check value of combobox
            // string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
            //double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString());
            //  DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;


            //  double dStep1 = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString());
            List<string> list = new List<string>();
            listTCIExport = new List<string>();
            DataTable tbTCIInfo = (DataTable)dgDetailInformation.DataSource;

            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();
            //test save logfile
            // List<string> list = new List<string>();
            //  listTCIExport = new List<string>();

            Utilities utilities = new Utilities();


            DataTable dtSource = (DataTable)dgDetailInformation.DataSource;


            //ArrayList test = utilities.GetColumnName(dtSource);

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 1; i < dtSource.Rows.Count; i++)
                {
                    // Validate by row
                    // if has error
                    // Set error into datagrid
                    StringBuilder stbuilderRow = new StringBuilder();
                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbFreqStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    //allListRange = new Dictionary<int, ArrayList>();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                                       ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        
                        //     stbuilderRow.Append("loi cot tan so, dong thu ");
                        stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                        // stbuilderRow.Insert(26, i+1);
                        list.Add(stbuilderRow.ToString());
                    }
                    else
                    {
                        allFreq = utilities.GetAllFrequencyByRange(arrFreqByRow[i], allFreq, ref hasError);

                        IsValidate = IsValidate && !hasError;

                        if (hasError)
                        {
                            // Tan so bi loi
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            
                            //   stbuilderRow.Append("loi cot tan so, dong thu ");
                            stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                            //stbuilderRow.Insert(26, i+1);
                            list.Add(stbuilderRow.ToString());
                        }
                    }

                    // check kinh do vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    IsValidate = IsValidate && !hasError;

                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString()))
                    {
                        if (hasError)
                        {
                            // Kinh do vi do bi loi
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            
                            //    stbuilderRow.Append("loi cot kinh do/vi do, dong thu ");
                            stbuilderRow.AppendFormat("Error dong thu {0}, cot kinh do/vi do ", i + 2);
                            // stbuilderRow.Insert(33, i + 1);
                            list.Add(stbuilderRow.ToString());
                        }
                    }
                    else
                    {
                        // Do not have error
                        IsValidate = true;
                    }


                    // check ten may
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Length > 50)
                    {
                        IsValidate = false;

                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ToolTipText = "Error";
                        dgDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ErrorText = "Test thu ErrorText";
                        DataGridViewRow row = dgDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        
                        //  stbuilderRow.Append("loi cot ten may, dong thu ");
                        stbuilderRow.AppendFormat("Error dong thu {0}, cot ten may", i + 2);
                        //stbuilderRow.Insert(27, i + 1);
                        list.Add(stbuilderRow.ToString());
                    }

                    // Check dai tan 470 - 806
                    string valueCombobox = cmbChooseFreq.SelectedItem.ToString();
                    if (valueCombobox == Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806)
                    {
                        string maugiayphep = dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().Trim();
                        if (maugiayphep != Constants.ValueConstant.THTS && maugiayphep != Constants.ValueConstant.THTT)
                        {
                            IsValidate = false;

                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ToolTipText = "Error";
                            dgDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = "Error";
                            DataGridViewRow row = dgDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;
                            
                            // stbuilderRow.Append("loi cot mau giay phep, dong thu ");
                            stbuilderRow.AppendFormat("Error dong thu {0}, cot mau giay phep", i + 2);
                            // stbuilderRow.Insert(33, i + 1);
                            list.Add(stbuilderRow.ToString());
                        }
                    }

                }
            }
            // return IsValidate;

            listTCIExport = list;

            // Get source
            if (listTCIExport != null && listTCIExport.Count > 0)
            {
                List<string> lines = listTCIExport;

                // Save file
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Export file (*.txt)|*.txt|Export file (*.doc)|*.doc";
                dialog.Title = "Save file type.";

                string pathSave = default(string);

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    pathSave = dialog.FileName;
                }

                if (!String.IsNullOrEmpty(pathSave))
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(pathSave))
                    {
                        foreach (string line in lines)
                        {
                            // Writer into file
                            file.WriteLine(line);
                        }
                    }

                    MessageBox.Show("Save logfile Sucessful", "Message box", MessageBoxButtons.OK);

                }

            }

        }

        private void saveRSLogfileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            allListRange = new Dictionary<int, ArrayList>();
            //   bool isValidate = true;

            // if (isMustCheckRS)
            //     isValidate = CheckErrorTCI();

            //  if (isValidate)
            //   {
            // Check value of combobox
            // string valueCombobox = cmbRSChooseFreq.SelectedItem.ToString();
            //double dStep = Convert.ToDouble(cmbRSStep.SelectedItem.ToString());
            //  DataTable dtGrid = (DataTable)dgDetailInformation.DataSource;


            //  double dStep1 = Convert.ToDouble(cmbFreqStep.SelectedItem.ToString());
            List<string> list = new List<string>();
            listTCIExport = new List<string>();
            DataTable tbTCIInfo = (DataTable)dgRSDetailInformation.DataSource;


            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();

            Utilities utilities = new Utilities();
            DataTable dtSource = (DataTable)dgRSDetailInformation.DataSource;

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate Frequency
                    StringBuilder stbuilderRow = new StringBuilder();
                    bool hasError = false;
                    string strStart = cmbRSChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbRSStep.SelectedItem) * 1000;

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                        list.Add(stbuilderRow.ToString());

                    }
                    else
                    {
                        allFreq = utilities.GetAllFrequencyByRange(arrFreqByRow[i], allFreq, ref hasError);

                        IsValidate = IsValidate && !hasError;

                        if (hasError)
                        {
                            // Tan so bi loi
                            dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = "Error";
                            DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.Yellow;

                            stbuilderRow.AppendFormat("error dong thu {0}, cot tan so", i + 2);
                            list.Add(stbuilderRow.ToString());
                        }
                    }

                    // Check customer
                    #region Ten khach hang
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                    //&& dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Length <= 25)
                    {
                        // Good
                    }
                    else
                    {
                        // Had error
                        hasError = true;
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            "Ten khach hang error";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot ten khach hang", i + 2);
                        list.Add(stbuilderRow.ToString());

                    }
                    #endregion
                    #region kinhdo_vido
                    // check kinh do vi do
                    //hasError = utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    //IsValidate = IsValidate && !hasError;

                    //if (hasError)
                    //{
                    //    // Kinh do vi do bi loi
                    //    dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                    //    DataGridViewRow row = dgRSDetailInformation.Rows[i];
                    //    row.DefaultCellStyle.BackColor = Color.Yellow;
                    //    btnRSCorrectError.Enabled = true;
                    //}

                    // check kinh do vi do
                    //hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    //IsValidate = IsValidate && !hasError;

                    //if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString()))
                    //{
                    //    if (hasError)
                    //    {
                    //        // Kinh do vi do bi loi
                    //        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = "Error";
                    //        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                    //        row.DefaultCellStyle.BackColor = Color.Yellow;
                    //        btnRSCorrectError.Enabled = true;
                    //    }
                    //}
                    //else
                    //{
                    //    // Do not have error
                    //    IsValidate = true;
                    //}

                    #endregion

                    // check HO HIEU
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Length > 32)
                    {
                        IsValidate = false;

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ToolTipText = "Error";
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.HO_HIEU].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot ho hieu", i + 2);
                        list.Add(stbuilderRow.ToString());
                    }


                    // check So GP
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.GPNo].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Length > 32)
                    {
                        IsValidate = false;

                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ToolTipText = "Error";
                        dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.GPNo].ErrorText =
                            "Test thu ErrorText";
                        DataGridViewRow row = dgRSDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;

                        stbuilderRow.AppendFormat("error dong thu {0}, cot so giap phep", i + 2);
                        list.Add(stbuilderRow.ToString());
                    }
                }
            }


            listTCIExport = list;

            // Get source
            if (listTCIExport != null && listTCIExport.Count > 0)
            {
                List<string> lines = listTCIExport;

                // Save file
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Export file (*.txt)|*.txt|Export file (*.doc)|*.doc";
                dialog.Title = "Save file type.";

                string pathSave = default(string);

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    pathSave = dialog.FileName;
                }

                if (!String.IsNullOrEmpty(pathSave))
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(pathSave))
                    {
                        foreach (string line in lines)
                        {
                            // Writer into file
                            file.WriteLine(line);
                        }
                    }

                    MessageBox.Show("Save logfile Sucessful", "Message box", MessageBoxButtons.OK);

                }

            }
        }

        private void tCIToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Excel file (*.xls)|*.xls";
                dialog.Title = "Open file Excel convert.";

                // Clean table dtSource


                Utilities utilities = new Utilities();

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string sheetName = this.GetSheetName(dialog.FileName);
                    //string sheetName = "sheet1";
                    DataSet dsExcel = utilities.GetAllDataFromFileExcel(dialog.FileName, sheetName);

                    if (dsExcel != null
                    && dsExcel.Tables != null
                    && dsExcel.Tables.Count > 0
                    && dsExcel.Tables[0].Rows.Count > 0)
                    {
                        dgDetailInformation.DataSource = dsExcel.Tables[0];

                        if (allListRange != null && allListRange.Count > 0)
                        {
                            allListRange.Clear();
                        }
                    }

                    // Enable button
                    btnCheckError.Enabled = true;
                    btnFormat.Enabled = true;

                }
            }
        }

        private void rohdeSchwarzToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel file (*.xls)|*.xls";
            dialog.Title = "Open file Excel convert.";

            // Clean dtRSSource
            dtRSSource = null;

            Utilities utilities = new Utilities();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string sheetName = this.GetSheetName(dialog.FileName);

                DataSet dsExcel = utilities.GetAllDataFromFileExcel(dialog.FileName, sheetName);

                if (dsExcel != null
                && dsExcel.Tables != null
                && dsExcel.Tables.Count > 0
                && dsExcel.Tables[0].Rows.Count > 0)
                {
                    //dgRSDetailInformation.DataSource = null;
                    if (dgRSDetailInformation.DataSource != null)
                    {
                        dgRSDetailInformation.DataSource = null;
                        dgRSDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    else
                    {
                        dgRSDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    if (allListRange != null && allListRange.Count > 0)
                    {
                        allListRange.Clear();
                    }
                }

                // Enable button
                btnRSCheckError.Enabled = true;
                btnRSFormat.Enabled = false;

                // Reset can export value
                canExportCSV = false;

            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void cmbGStep_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnTranFormat_Click(object sender, EventArgs e)
        {
            //Set can export CSV
            isExportTran = true;
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();

            bool isValidate = CheckErrorGEW();

            if (isValidate)
            {
                // Check value of combobox
                string valueCombobox = cmbGEChooseFreq.SelectedItem.ToString();
                double dStep = Convert.ToDouble(cmbGEStep.SelectedItem.ToString()) * 1000;
                DataTable dtGrid = null;
                Utilities utilities = new Utilities();
                string strStart = cmbGEChooseFreq.SelectedItem.ToString();
                // Convert NewStart to OldStart
                strStart = strStart.Replace("KHz", "");
                strStart = strStart.Replace("MHz", "");
                strStart = strStart.Replace(" ", "");
                strStart = strStart.Replace("-", "_");

                if (dtGESource != null)
                    dtGrid = dtGESource;
                else
                {
                    dtGrid = (DataTable)dgGEDetailInformation.DataSource;
                }
                for (int i = 0; i < dtGrid.Rows.Count; i++)
                {
                    string strFreq = dtGrid.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    bool hasError = false;

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }
                }

                switch (valueCombobox)
                {
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_47_50):
                        #region FREQ_TTKD_47_50
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_54_68):
                        #region FREQ_TTKD_54_68
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_PT_87_108):
                        #region FREQ_PT_87_108
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174):
                        #region FREQ_DR_137_174
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_174_230):
                        #region FREQ_TH_174_230
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    Utilities util = new Utilities();
                                    if (util.MachTDMB(dtGrid.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim()))
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            // Clear old data
                                            arrFreq.Clear();

                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            for (int j = 0; j < arrFreq.Count; j++)
                                            {
                                                double freqBase = Convert.ToDouble(arrFreq[j]);

                                            }
                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (arrFreq != null && arrFreq.Count == 2)
                                            {
                                                //arrFreqTemp = arrFreq;
                                                // Create freq base
                                                // fBase = ((f1 + f2)-1)/2;
                                                double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;


                                            }

                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                    }
                                    if (hasChange)
                                    {
                                        allListRange[i] = arrFreq;
                                    }
                                }
                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470):
                        #region FREQ_DR_400_470
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_470_790):
                        #region FREQ_TH_470_790
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (arrFreq != null && arrFreq.Count == 2)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            // Create freq base
                                            // fBase = ((f1 + f2)-1)/2;
                                            double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;


                                        }

                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    default:
                        //// Action default
                        //dgDetailInformation.DataSource = objFormat.GetTCITableOutput((DataTable)dgDetailInformation.DataSource,
                        //                                                     allListRange);
                        break;

                }

                // Common action// Action default
                dgGEDetailInformation.DataSource = null;
                dgGEDetailInformation.DataSource = objFormat.GetGETranmisterTableOutput(dtGrid, allListRange);

                btnGEShow.Enabled = true;
                buttonGExport.Enabled = true;
                btnTranFormat.Enabled = true;
                btnGEImport.Enabled = true;
                btnFrequencies.Enabled = true;
                btnGECheckError.Enabled = false;
                btnGECorrectError.Enabled = false;
                dgGEDetailInformation.ReadOnly = true;

                // Set can export 
                canExportCSV = true;

            }
            else
            {
                // Check error
                btnGECheckError.Enabled = true;
                btnGEImport.Enabled = true;
                btnFrequencies.Enabled = false;
                btnTranFormat.Enabled = false;
                btnGEShow.Enabled = false;

            }
        }

        private void cmbGEChooseFreq_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> freqRangeForStepGE = new List<string>();
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_HF_9_30);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_47_50);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_54_68);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_PT_87_108);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_HK_108_137);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_174_230);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_470_790);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_790_890);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_890_960);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_1710_1785);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_1805_1880);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_1920_1980);
            freqRangeForStepGE.Add(Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTDD_2110_2170);

            Utilities utilities = new Utilities();
            Dictionary<string, string> dicStepGE = utilities.CreateFreqAndStepGEW(freqRangeForStepGE);
            if (cmbGEChooseFreq.SelectedItem != null)
            {
                string[] listStepGE = dicStepGE[cmbGEChooseFreq.SelectedItem.ToString()].Split(';');

                cmbGEStep.DataSource = listStepGE;
            }
        }

        private void btnGEImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel file (*.xls)|*.xls";
            dialog.Title = "Open file Excel convert.";

            // Clean table dtSource
            dtGESource = null;

            Utilities utilities = new Utilities();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string sheetName = this.GetSheetName(dialog.FileName);
                //string sheetName = "sheet1";
                DataSet dsExcel = utilities.GetAllDataFromFileExcel(dialog.FileName, sheetName);

                if (dsExcel != null
                && dsExcel.Tables != null
                && dsExcel.Tables.Count > 0
                && dsExcel.Tables[0].Rows.Count > 0)
                {
                    //dgRSDetailInformation.DataSource = null;
                    if (dgGEDetailInformation.DataSource != null)
                    {
                        dgGEDetailInformation.DataSource = null;
                        dgGEDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    else
                    {
                        dgGEDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    if (allListRange != null && allListRange.Count > 0)
                    {
                        allListRange.Clear();
                    }

                    // dgDetailInformation.DataSource = dsExcel.Tables[0];

                    // if (allListRange != null && allListRange.Count > 0)
                    // {
                    //     allListRange.Clear();
                    // }
                }

                // Enable button
                btnGECheckError.Enabled = true;
                imgGEW.Visible = false;
            }
        }

        private void btnGECheckError_Click(object sender, EventArgs e)
        {
            bool isValidate = CheckErrorGEW();
        }

        private void btnGECorrectError_Click(object sender, EventArgs e)
        {
            bool IsValidate = true;
            allListRange = new Dictionary<int, ArrayList>();

            Utilities utilities = new Utilities();
            DataTable dtSource = (DataTable)dgGEDetailInformation.DataSource;

            bool isMustReBindDataSource = false;

            bool hasUnExpectedError = false;

            if (dtSource != null && dtSource.Rows.Count > 0)
            {
                ArrayList allFreq = new ArrayList();
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    // Validate by row
                    // if has error
                    // Set error into datagrid

                    // Validate Frequency
                    bool hasError = false;
                    string strStart = cmbGEChooseFreq.SelectedItem.ToString();
                    // Convert NewStart to OldStart
                    strStart = strStart.Replace("KHz", "");
                    strStart = strStart.Replace("MHz", "");
                    strStart = strStart.Replace(" ", "");
                    strStart = strStart.Replace("-", "_");

                    double dStep = Convert.ToDouble(cmbGEStep.SelectedItem) * 1000;

                    ////Test
                    //strStart = "800Mhz_10000mhz";

                    //dStep = 100000;
                    //ArrayList allFreq = new ArrayList();

                    string strFreq = dtSource.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    //allListRange = new Dictionary<int, ArrayList>();

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    //// Add arraylist with no error
                    //if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    //{
                    //    allListRange.Add(i, arrFreqByRow[i]);
                    //}

                    // 
                    IsValidate = IsValidate && !hasError;

                    if (hasError)
                    {
                        // Tan so bi loi
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].Value =
                            utilities.CorrectFrequencyByRange(strStart, dStep, strFreq);
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TAN_SO].ErrorText = string.Empty;
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                        btnGECorrectError.Enabled = true;
                    }
                    #region KinhdoVido
                    // check kinh do vi do
                    hasError = !utilities.IsKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString());

                    IsValidate = IsValidate && !hasError;

                    //if (hasError)
                    //{
                    // Kinh do vi do bi loi
                    dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].Value =
                        utilities.CorrectKinhdoVido(dtSource.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString(),
                                                    ref hasError);

                    if (!hasError)
                    {
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            string.Empty;
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                    else
                    {
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText =
                            "Unexpected Error.";
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnTranFormat.Enabled = false;
                        hasUnExpectedError = true;
                    }
                    //}
                    //else
                    //{
                    //    dgDetailInformation.Rows[i].Cells[Constants.TableExport.KINHDO_VIDO].ErrorText = string.Empty;
                    //    DataGridViewRow row = dgDetailInformation.Rows[i];
                    //    row.DefaultCellStyle.BackColor = Color.White;
                    //    //btnCorrectError.Enabled = true;
                    //}
                    #endregion

                    // Check customer
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                    {
                        if (dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Length > 25)
                        {
                            //string tenkhachhangCut =
                            //dtSource.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Substring(0, 25);

                            //dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].Value = tenkhachhangCut;
                            //dgRSDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            //    string.Empty;
                            //DataGridViewRow row = dgRSDetailInformation.Rows[i];
                            //row.DefaultCellStyle.BackColor = Color.White;
                        }
                    }
                    else
                    {
                        // Had error
                        hasError = true;
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_KHACH_HANG].ErrorText =
                            "Ten khach hang blank";
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                        btnRSCorrectError.Enabled = true;

                    }
                    IsValidate = IsValidate && !hasError;

                    // check ten may
                    if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString()) &&
                        dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Length > 50)
                    {
                        IsValidate = false;

                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].Value =
                            dtSource.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Substring(0, 50);
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ToolTipText = string.Empty;
                        dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.TEN_MAY].ErrorText = string.Empty;
                        //    "Test thu ErrorText";
                        DataGridViewRow row = dgGEDetailInformation.Rows[i];
                        row.DefaultCellStyle.BackColor = Color.White;
                        btnGECorrectError.Enabled = true;
                    }

                    // Remove row khong phai PTTH
                    string valueCombobox = cmbGEChooseFreq.SelectedItem.ToString();
                    if (valueCombobox == Constants.FreqAndStep.FrequencyDisplay.FREQ_UHF_470_806)
                    {
                        if (dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value != null)
                        {
                            string maugiayphep =
                                dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value.ToString().
                                    Trim();
                            if (maugiayphep != Constants.ValueConstant.THTS &&
                                maugiayphep != Constants.ValueConstant.THTT)
                            {
                                // Remove row
                                //dgDetailInformation.Rows.RemoveAt(i);
                                dtSource.Rows.RemoveAt(i);
                                isMustReBindDataSource = true;
                            }
                        }
                    }

                    if (valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174 || valueCombobox == Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470)
                    {
                        if (!String.IsNullOrEmpty(dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                        {
                            if (dtSource.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString() == Constants.ValueConstant.DAI_TAU)
                            {

                                dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value = String.Empty;
                                dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText = string.Empty;
                                DataGridViewRow row = dgGEDetailInformation.Rows[i];
                                row.DefaultCellStyle.BackColor = Color.White;
                            }
                        }
                        else
                        {

                            dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].Value = string.Empty;
                            dgGEDetailInformation.Rows[i].Cells[Constants.TableExport.MAU_GIAY_PHEP].ErrorText =
                                string.Empty;
                            DataGridViewRow row = dgGEDetailInformation.Rows[i];
                            row.DefaultCellStyle.BackColor = Color.White;
                        }
                    }
                }
            }
            if (hasUnExpectedError)
            {
                btnTranFormat.Enabled = false;
                btnFrequencies.Enabled = false;
            }
            else
            {
                btnTranFormat.Enabled = true;
                btnFrequencies.Enabled = true;
                dtGESource = (DataTable)dgGEDetailInformation.DataSource;
            }

            if (isMustReBindDataSource)
                dgGEDetailInformation.DataSource = dtSource;
            btnTranFormat.Enabled = true;
            btnFrequencies.Enabled = true;
            dtGESource = (DataTable)dgGEDetailInformation.DataSource;
        }

        private void btnFrequencies_Click(object sender, EventArgs e)
        {
            //Set can export CSV
            isExportTran = false;
            OutFormatBO objFormat = new OutFormatBO();

            allListRange = new Dictionary<int, ArrayList>();
            bool isValidate = true;

            if (isMustCheckRS)
            {
                isValidate = CheckErrorRS();
            }

            if (isValidate)
            {
                // Check value of combobox
                string valueCombobox = cmbGEChooseFreq.SelectedItem.ToString();
                double dStep = Convert.ToDouble(cmbGEStep.SelectedItem.ToString()) * 1000;
                DataTable dtGrid = null;
                Utilities utilities = new Utilities();
                string strStart = cmbGEChooseFreq.SelectedItem.ToString();
                // Convert NewStart to OldStart
                strStart = strStart.Replace("KHz", "");
                strStart = strStart.Replace("MHz", "");
                strStart = strStart.Replace(" ", "");
                strStart = strStart.Replace("-", "_");

                if (dtGESource != null)
                    dtGrid = dtGESource;
                else
                {
                    dtGrid = (DataTable)dgGEDetailInformation.DataSource;
                }
                for (int i = 0; i < dtGrid.Rows.Count; i++)
                {
                    string strFreq = dtGrid.Rows[i][Constants.TableExport.TAN_SO].ToString();

                    bool hasError = false;

                    Dictionary<int, ArrayList> arrFreqByRow = utilities.GetFrequencyByRange(strStart, dStep, strFreq, i,
                                                                                            ref hasError);

                    // Add arraylist with no error
                    if (!hasError && !allListRange.ContainsKey(i) && arrFreqByRow[i] != null && arrFreqByRow[i].Count > 0)
                    {
                        allListRange.Add(i, arrFreqByRow[i]);
                    }
                }

                switch (valueCombobox)
                {
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_47_50):
                        #region FREQ_TTKD_47_50
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TTKD_54_68):
                        #region FREQ_TTKD_54_68
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_PT_87_108):
                        #region FREQ_PT_87_108
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_137_174):
                        #region FREQ_DR_137_174
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_174_230):
                        #region FREQ_TH_174_230
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    Utilities util = new Utilities();
                                    if (util.MachTDMB(dtGrid.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim()))
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            // Clear old data
                                            arrFreq.Clear();

                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            for (int j = 0; j < arrFreq.Count; j++)
                                            {
                                                double freqBase = Convert.ToDouble(arrFreq[j]);

                                            }
                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (arrFreq != null && arrFreq.Count == 2)
                                            {
                                                //arrFreqTemp = arrFreq;
                                                // Create freq base
                                                // fBase = ((f1 + f2)-1)/2;
                                                double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;


                                            }

                                            if (hasChange)
                                            {
                                                foreach (var list in arrFreqTemp)
                                                {
                                                    if (!arrFreq.Contains(list))
                                                        arrFreq.Add(list);
                                                }
                                            }
                                        }
                                    }
                                    if (hasChange)
                                    {
                                        allListRange[i] = arrFreq;
                                    }
                                }
                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_DR_400_470):
                        #region FREQ_DR_400_470
                        // FM range
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq upper and lower
                                ArrayList arrFreq = allListRange[i];

                                ArrayList arrFreqTemp = new ArrayList();
                                bool hasChange = false;
                                for (int j = 0; j < arrFreq.Count; j++)
                                {
                                    double freqBase = Convert.ToDouble(arrFreq[j]);

                                }
                                if (hasChange)
                                {
                                    foreach (var list in arrFreqTemp)
                                    {
                                        if (!arrFreq.Contains(list))
                                            arrFreq.Add(list);
                                    }
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    case (Constants.FreqAndStep.FrequencyGEWDisplay.FREQ_TH_470_790):
                        #region FREQ_TH_470_790
                        // Analog, Digital TV
                        // Get allListRage
                        if (allListRange != null
                            && allListRange.Count > 0)
                        {
                            for (int i = 0; i < allListRange.Count; i++)
                            {
                                // Create freq Begin and End frequency
                                ArrayList arrFreq = allListRange[i];
                                bool hasChange = false;
                                if (dtGrid != null
                                    && dtGrid.Rows != null
                                    && dtGrid.Rows.Count > 0)
                                {
                                    ArrayList arrFreqTemp = new ArrayList();
                                    if (dtGrid.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim() == Constants.ValueConstant.THTS)
                                    {
                                        //arrFreqTemp = arrFreq;
                                        for (int j = 0; j < arrFreq.Count; j++)
                                        {
                                            double freqBase = Convert.ToDouble(arrFreq[j]);

                                        }
                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (arrFreq != null && arrFreq.Count == 2)
                                        {
                                            //arrFreqTemp = arrFreq;
                                            // Create freq base
                                            // fBase = ((f1 + f2)-1)/2;
                                            double freqBase = ((Convert.ToDouble(arrFreq[0]) + Convert.ToDouble(arrFreq[1])) - 1000000) / 2;


                                        }

                                        if (hasChange)
                                        {
                                            foreach (var list in arrFreqTemp)
                                            {
                                                if (!arrFreq.Contains(list))
                                                    arrFreq.Add(list);
                                            }
                                        }
                                    }
                                }
                                if (hasChange)
                                {
                                    allListRange[i] = arrFreq;
                                }

                            }
                        }
                        break;
                        #endregion
                    default:
                        //// Action default
                        //dgDetailInformation.DataSource = objFormat.GetTCITableOutput((DataTable)dgDetailInformation.DataSource,
                        //                                                     allListRange);
                        break;

                }

                // Common action// Action default
                dgGEDetailInformation.DataSource = null;
                dgGEDetailInformation.DataSource = objFormat.GetGEFrequencyTableOutput(dtGrid, allListRange);

                btnGEShow.Enabled = true;
                buttonGExport.Enabled = true;
                btnTranFormat.Enabled = true;
                btnGEImport.Enabled = true;
                btnFrequencies.Enabled = true;
                btnGECheckError.Enabled = false;
                btnGECorrectError.Enabled = false;
                dgGEDetailInformation.ReadOnly = true;

                // Set can export 
                canExportCSV = true;
                isMustCheckRS = false;

            }
            else
            {
                // Check error
                btnGECheckError.Enabled = true;
                btnGEImport.Enabled = true;
                btnFrequencies.Enabled = false;
                btnTranFormat.Enabled = false;
                btnGEShow.Enabled = false;

            }
        }

        private void btnGEShow_Click(object sender, EventArgs e)
        {
            //OutFormatBO objFormat = new OutFormatBO();
            double dStep = Convert.ToDouble(cmbGEStep.SelectedItem.ToString());
            if (dgGEDetailInformation != null && dgGEDetailInformation.DataSource != null)
            {
                List<string> list = new List<string>();
                listGEWExport = new List<string>();
                DataTable tbGEWInfo = (DataTable)dgGEDetailInformation.DataSource;
                if (tbGEWInfo != null && tbGEWInfo.Rows.Count > 0)
                {
                    for (int i = 0; i < tbGEWInfo.Rows.Count; i++)
                    {
                        if (!isExportTran)
                        {

                            StringBuilder stbuilderRow = new StringBuilder();
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.TRANSMITTER_EXTERNAL_ID]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.FREQUENCY_EXTERNAL_ID]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.CENTRE_FREQUENCY]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.BANDWIDTH]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.CHANNEL_SPACE]);
                            stbuilderRow.Append(";");

                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.CHANNEL_NAME]);

                            list.Add(stbuilderRow.ToString());
                        }
                        else
                        {
                            StringBuilder stbuilderRow = new StringBuilder();
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.TRANSMITTER_EXTERNAL_ID]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.NAME]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.TYPE]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.LATITUDE]);
                            stbuilderRow.Append(";");
                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.LONGITUDE]);
                            stbuilderRow.Append(";");

                            stbuilderRow.Append(tbGEWInfo.Rows[i][Constants.TableExport.GEWTABLE.COMMENT]);

                            list.Add(stbuilderRow.ToString());
                        }
                    }
                }

                foreach (var openForm in Application.OpenForms)
                {
                    if (openForm.Equals(form1))
                    {

                    }
                    else
                    {
                        form1 = new Form2(list);
                        //form1.Show();
                    }
                }
                form1.Show();
                listGEWExport = list;
                // Enable button show
                buttonGExport.Enabled = true;
                btnGEShow.Enabled = true;
                btnGEImport.Enabled = true;

            }
        }

        private void buttonGExport_Click(object sender, EventArgs e)
        {
            if (dgGEDetailInformation != null && dgGEDetailInformation.DataSource != null)
            {
                //List<string> list = new List<string>();

                //Utilities utils = new Utilities();
                //DataTable dtRS = utils.GetTemplateTableRS();

                listGEWExport = new List<string>();
                DataTable tbGEWInfo = (DataTable)dgGEDetailInformation.DataSource;
                string frequenceRange = cmbGEChooseFreq.SelectedItem.ToString();

                foreach (var openForm in Application.OpenForms)
                {
                    if (openForm.Equals(GEWconfirmExport))
                    {

                    }
                    else
                    {
                        GEWconfirmExport = new GEWConfirmExport(tbGEWInfo, isExportTran, frequenceRange);
                        //form1.Show();
                    }
                }
                GEWconfirmExport.Show();

                // Enable button show
                buttonGExport.Enabled = true;
                btnGEShow.Enabled = true;
                btnGEImport.Enabled = true;
                btnTranFormat.Enabled = true;
                btnFrequencies.Enabled = true;

            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void imgTCI_Click(object sender, EventArgs e)
        {

        }

        private void imgRS_Click(object sender, EventArgs e)
        {

        }

        private void gEWToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnGEExport.SelectedIndex = 2;
        }

        private void gEWToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel file (*.xls)|*.xls";
            dialog.Title = "Open file Excel convert.";

            // Clean table dtSource
            dtTCISource = null;

            Utilities utilities = new Utilities();

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string sheetName = this.GetSheetName(dialog.FileName);
                //string sheetName = "sheet1";
                DataSet dsExcel = utilities.GetAllDataFromFileExcel(dialog.FileName, sheetName);

                if (dsExcel != null
                && dsExcel.Tables != null
                && dsExcel.Tables.Count > 0
                && dsExcel.Tables[0].Rows.Count > 0)
                {
                    //dgRSDetailInformation.DataSource = null;
                    if (dgDetailInformation.DataSource != null)
                    {
                        dgDetailInformation.DataSource = null;
                        dgDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    else
                    {
                        dgDetailInformation.DataSource = dsExcel.Tables[0];
                    }
                    if (allListRange != null && allListRange.Count > 0)
                    {
                        allListRange.Clear();
                    }

                    // dgDetailInformation.DataSource = dsExcel.Tables[0];

                    // if (allListRange != null && allListRange.Count > 0)
                    // {
                    //     allListRange.Clear();
                    // }
                }

                // Enable button
                imgTCI.Visible = false;
                btnCheckError.Enabled = true;
                btnFormat.Enabled = false;
                btnCorrectError.Enabled = false;
                btnShow.Enabled = false;
                button9.Enabled = false;
                btnExport.Enabled = false;
            }
        }
    }
}