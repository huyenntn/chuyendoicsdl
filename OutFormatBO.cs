using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;

namespace AVDApplication
{
    class OutFormatBO
    {
        Utilities objUti;

        /// <summary>
        /// Create table output
        /// Xu ly ID
        /// Xu ly tan so
        /// Xu ly kinh do, vi do
        /// </summary>
        /// <param name="inputTable"></param>
        /// <param name="listRange"></param>
        /// <returns></returns>
        public DataTable GetTCITableOutputByFreq(DataTable inputTable, Dictionary<int, ArrayList> listRange, string strFreq)
        {
            objUti = new Utilities();
            DataTable tciFormatTable = objUti.GetTemplateTable();

            if (inputTable != null && inputTable.Rows.Count > 0)
            {
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if (listRange.ContainsKey(i))
                    {
                        // Row i has range frequency
                        int indexBegin = 0;
                        if (tciFormatTable.Rows.Count > 0)
                        {
                            indexBegin = tciFormatTable.Rows.Count + i + 1;
                        }
                        else
                        {
                            indexBegin = i + 1;
                        }
                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = tciFormatTable.NewRow();

                            // ID and frequency
                            row[Constants.TableExport.ID] = indexBegin + j;

                            // Frequency
                            row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j]);


                            row[Constants.TableExport.SO_THAM_CHIEU] = Constants.ValueConstant.SPACE;
                            row[Constants.TableExport.DO_LECH_F] = Constants.ValueConstant.SPACE;

                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.GPNo].ToString()))
                            {
                                row[Constants.TableExport.GPNo] =
                                    inputTable.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.GPNo] = Constants.ValueConstant.SPACE;
                            }

                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                            {
                                row[Constants.TableExport.MAU_GIAY_PHEP] =
                                    inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.MAU_GIAY_PHEP] = Constants.ValueConstant.SPACE;
                            }

                            // Five columns normal
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                            {
                                row[Constants.TableExport.TEN_KHACH_HANG] =
                                    inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.TEN_KHACH_HANG] = Constants.ValueConstant.SPACE;
                            }


                            // Five columns normal
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString()))
                            {
                                row[Constants.TableExport.HO_HIEU] =
                                    inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.HO_HIEU] = Constants.ValueConstant.SPACE;
                            }


                            // Three columns insert space
                            row[Constants.TableExport.BRAND_UU_TIEN] = Constants.ValueConstant.SPACE;
                            row[Constants.TableExport.DO_RONG_KENH] = Constants.ValueConstant.SPACE;
                            row[Constants.TableExport.SO_KENH] = Constants.ValueConstant.SPACE;


                            // Xu ly kinh do vi do
                            // Longtitude and latitude
                            string[] kinhdoVidoArr =
                                objUti.GetKinhdoAndVido(inputTable.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
                            if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
                            {
                                // Get normal
                                row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);

                            }
                            else
                            {
                                // Call random value;
                                row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
                            }

                            if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
                            {
                                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
                            }
                            else
                            {
                                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
                            }
                            //row[Constants.TableExport.KINH_DO] = inputTable.Rows[i][Constants.TableExport.KINH_DO];
                            //row[Constants.TableExport.VI_DO] = inputTable.Rows[i][Constants.TableExport.VI_DO];

                            row[Constants.TableExport.TEN_MAY] = inputTable.Rows[i][Constants.TableExport.TEN_MAY];

                            tciFormatTable.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency
                        DataRow row = tciFormatTable.NewRow();

                        // Xu ly ID
                        // ID = Row count.
                        if (tciFormatTable.Rows.Count > 0)
                        {
                            row[Constants.TableExport.ID] = tciFormatTable.Rows.Count + 1;
                        }
                        else
                        {
                            row[Constants.TableExport.ID] = 1;
                        }
                        // Two columns insert space
                        row[Constants.TableExport.SO_THAM_CHIEU] = Constants.ValueConstant.SPACE;
                        row[Constants.TableExport.DO_LECH_F] = Constants.ValueConstant.SPACE;

                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.GPNo].ToString()))
                        {
                            row[Constants.TableExport.GPNo] =
                                inputTable.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Replace(
                                    ";", "");
                        }
                        else
                        {
                            row[Constants.TableExport.GPNo] = Constants.ValueConstant.SPACE;
                        }

                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                        {
                            row[Constants.TableExport.MAU_GIAY_PHEP] =
                                inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim().Replace(
                                    ";", "");
                        }
                        else
                        {
                            row[Constants.TableExport.MAU_GIAY_PHEP] = Constants.ValueConstant.SPACE;
                        }

                        //row[Constants.TableExport.TAN_SO] = inputTable.Rows[i][Constants.TableExport.TAN_SO];
                        row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString());

                        // Five columns normal
                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                        {
                            row[Constants.TableExport.TEN_KHACH_HANG] =
                                inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(
                                    ";", "");
                        }
                        else
                        {
                            row[Constants.TableExport.TEN_KHACH_HANG] = Constants.ValueConstant.SPACE;
                        }


                        // Five columns normal
                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString()))
                        {
                            row[Constants.TableExport.HO_HIEU] =
                                inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Replace(
                                    ";", "");
                        }
                        else
                        {
                            row[Constants.TableExport.HO_HIEU] = Constants.ValueConstant.SPACE;
                        }


                        // Three columns insert space
                        row[Constants.TableExport.BRAND_UU_TIEN] = Constants.ValueConstant.SPACE;
                        row[Constants.TableExport.DO_RONG_KENH] = Constants.ValueConstant.SPACE;
                        row[Constants.TableExport.SO_KENH] = Constants.ValueConstant.SPACE;


                        // Xu ly kinh do vi do
                        // Longtitude and latitude
                        string[] kinhdoVidoArr =
                            objUti.GetKinhdoAndVido(inputTable.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
                        if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
                        {
                            // Get normal
                            row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);

                        }
                        else
                        {
                            // Call random value;
                            row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
                        }

                        if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
                        {
                            row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
                        }
                        else
                        {
                            row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
                        }
                        //row[Constants.TableExport.KINH_DO] = inputTable.Rows[i][Constants.TableExport.KINH_DO];
                        //row[Constants.TableExport.VI_DO] = inputTable.Rows[i][Constants.TableExport.VI_DO];

                        row[Constants.TableExport.TEN_MAY] = inputTable.Rows[i][Constants.TableExport.TEN_MAY];

                        // Add row
                        tciFormatTable.Rows.Add(row);
                    }
                }
            }

            return tciFormatTable;
        }

        /// <summary>
        /// Create table output
        /// Xu ly ID
        /// Xu ly tan so
        /// Xu ly kinh do, vi do
        /// </summary>
        /// <param name="inputTable"></param>
        /// <param name="listRange"></param>
        /// <returns></returns>
        public DataTable GetTCITableOutput(DataTable inputTable, Dictionary<int, ArrayList> listRange)
        {
            objUti = new Utilities();
            DataTable tciFormatTable = objUti.GetTemplateTable();

            if(inputTable != null && inputTable.Rows.Count > 0)
            {
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if(listRange.ContainsKey(i))
                    {
                        // Row i has range frequency
                        int indexBegin = 0;
                        if (tciFormatTable.Rows.Count > 0)
                        {
                            indexBegin = tciFormatTable.Rows.Count + i + 1;
                        }
                        else
                        {
                            indexBegin = i + 1;
                        } 
                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = tciFormatTable.NewRow();

                            // ID and frequency
                            row[Constants.TableExport.ID] = indexBegin + j;

                            // Frequency
                            row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j]);


                            row[Constants.TableExport.SO_THAM_CHIEU] = Constants.ValueConstant.SPACE;
                            row[Constants.TableExport.DO_LECH_F] = Constants.ValueConstant.SPACE;

                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.GPNo].ToString()))
                            {
                                row[Constants.TableExport.GPNo] =
                                    inputTable.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.GPNo] = Constants.ValueConstant.SPACE;
                            }

                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                            {
                                row[Constants.TableExport.MAU_GIAY_PHEP] =
                                    inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.MAU_GIAY_PHEP] = Constants.ValueConstant.SPACE;
                            }

                            // Five columns normal
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                            {
                                row[Constants.TableExport.TEN_KHACH_HANG] =
                                    inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.TEN_KHACH_HANG] = Constants.ValueConstant.SPACE;
                            }


                            // Five columns normal
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString()))
                            {
                                row[Constants.TableExport.HO_HIEU] =
                                    inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.HO_HIEU] = Constants.ValueConstant.SPACE;
                            }


                            // Three columns insert space
                            row[Constants.TableExport.BRAND_UU_TIEN] = Constants.ValueConstant.SPACE;
                            row[Constants.TableExport.DO_RONG_KENH] = Constants.ValueConstant.SPACE;
                            row[Constants.TableExport.SO_KENH] = Constants.ValueConstant.SPACE;


                            // Xu ly kinh do vi do
                            // Longtitude and latitude
                            string[] kinhdoVidoArr =
                                objUti.GetKinhdoAndVido(inputTable.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
                            if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
                            {
                                // Get normal
                                row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);

                            }
                            else
                            {
                                // Call random value;
                                row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
                            }

                            if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
                            {
                                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
                            }
                            else
                            {
                                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
                            }
                            //row[Constants.TableExport.KINH_DO] = inputTable.Rows[i][Constants.TableExport.KINH_DO];
                            //row[Constants.TableExport.VI_DO] = inputTable.Rows[i][Constants.TableExport.VI_DO];
                            string tenmayTemp = inputTable.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Replace(";", "");
                            if (tenmayTemp.Length > 50)
                            {
                                row[Constants.TableExport.TEN_MAY] = tenmayTemp.Substring(0, 50);
                            }
                            else
                            {
                                row[Constants.TableExport.TEN_MAY] = tenmayTemp;
                            }

                            tciFormatTable.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency
                        DataRow row = tciFormatTable.NewRow();

                        // Xu ly ID
                        // ID = Row count.
                        if (tciFormatTable.Rows.Count > 0)
                        {
                            row[Constants.TableExport.ID] = tciFormatTable.Rows.Count + 1;
                        }
                        else
                        {
                            row[Constants.TableExport.ID] = 1;
                        }
                        // Two columns insert space
                        row[Constants.TableExport.SO_THAM_CHIEU] = Constants.ValueConstant.SPACE;
                        row[Constants.TableExport.DO_LECH_F] = Constants.ValueConstant.SPACE;

                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.GPNo].ToString()))
                        {
                            row[Constants.TableExport.GPNo] =
                                inputTable.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Replace(
                                    ";", "");
                        } 
                        else
                        {
                            row[Constants.TableExport.GPNo] = Constants.ValueConstant.SPACE;
                        }

                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                        {
                            row[Constants.TableExport.MAU_GIAY_PHEP] =
                                inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim().Replace(
                                    ";", "");
                        }
                        else
                        {
                            row[Constants.TableExport.MAU_GIAY_PHEP] = Constants.ValueConstant.SPACE;
                        }

                        //row[Constants.TableExport.TAN_SO] = inputTable.Rows[i][Constants.TableExport.TAN_SO];
                        row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString());

                        // Five columns normal
                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                        {
                            row[Constants.TableExport.TEN_KHACH_HANG] =
                                inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(
                                    ";", "");
                        }
                        else
                        {
                            row[Constants.TableExport.TEN_KHACH_HANG] = Constants.ValueConstant.SPACE;
                        }


                        // Five columns normal
                        if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString()))
                        {
                            row[Constants.TableExport.HO_HIEU] =
                                inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Replace(
                                    ";", "");
                        }
                        else
                        {
                            row[Constants.TableExport.HO_HIEU] = Constants.ValueConstant.SPACE;
                        }


                        // Three columns insert space
                        row[Constants.TableExport.BRAND_UU_TIEN] = Constants.ValueConstant.SPACE;
                        row[Constants.TableExport.DO_RONG_KENH] = Constants.ValueConstant.SPACE;
                        row[Constants.TableExport.SO_KENH] = Constants.ValueConstant.SPACE;


                        // Xu ly kinh do vi do
                        // Longtitude and latitude
                        string[] kinhdoVidoArr =
                            objUti.GetKinhdoAndVido(inputTable.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
                        if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
                        {
                            // Get normal
                            row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);

                        }
                        else
                        {
                            // Call random value;
                            row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
                        }

                        if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
                        {
                            row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
                        }
                        else
                        {
                            row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
                        }
                        //row[Constants.TableExport.KINH_DO] = inputTable.Rows[i][Constants.TableExport.KINH_DO];
                        //row[Constants.TableExport.VI_DO] = inputTable.Rows[i][Constants.TableExport.VI_DO];

                        string tenmayTemp = inputTable.Rows[i][Constants.TableExport.TEN_MAY].ToString().Trim().Replace(";", "");
                        if (tenmayTemp.Length > 50)
                        {
                            row[Constants.TableExport.TEN_MAY] = tenmayTemp.Substring(0, 50);
                        }
                        else
                        {
                            row[Constants.TableExport.TEN_MAY] = tenmayTemp;
                        }

                        // Add row
                        tciFormatTable.Rows.Add(row);
                    }
                }
            }

            return tciFormatTable;
        }
        public DataTable GetTCITableOutput_DFSCAN(DataTable inputTable, Dictionary<int, ArrayList> listRange)
        {
            objUti = new Utilities();
            DataTable tciFormatTable = objUti.GetTemplateTableDFScan();

            if (inputTable != null && inputTable.Rows.Count > 0)
            {
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if (listRange.ContainsKey(i))
                    {
                        // Row i has range frequency
                        int indexBegin = 0;
                        if (tciFormatTable.Rows.Count > 0)
                        {
                            indexBegin = tciFormatTable.Rows.Count + i + 1;
                        }
                        else
                        {
                            indexBegin = i + 1;
                        }
                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = tciFormatTable.NewRow();

                            // ID and frequency
                            row[Constants.TableExport.ID] = indexBegin + j;

                            // Frequency
                            row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j]);

                            tciFormatTable.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency
                        DataRow row = tciFormatTable.NewRow();

                        // Xu ly ID
                        // ID = Row count.
                        if (tciFormatTable.Rows.Count > 0)
                        {
                            row[Constants.TableExport.ID] = tciFormatTable.Rows.Count + 1;
                        }
                        else
                        {
                            row[Constants.TableExport.ID] = 1;
                        }
                       

                        // Add row
                        tciFormatTable.Rows.Add(row);
                    }
                }
            }

            return tciFormatTable;
        }

        public DataTable GetRSTableBeforeFormat(DataTable inputTable, Dictionary<int, ArrayList> listRange)
        {
            DataTable rsBeforeFormat = inputTable.Clone();
            //Utilities util = new Utilities();
            objUti = new Utilities();

            ArrayList listNameColumn = objUti.GetColumnName(inputTable);

            if (inputTable != null && inputTable.Rows.Count > 0)
            {
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if (listRange.ContainsKey(i))
                    {
                        // Row i has range frequency
                        int indexBegin = 0;
                        if (rsBeforeFormat.Rows.Count > 0)
                        {
                            indexBegin = rsBeforeFormat.Rows.Count + i + 1;
                        }
                        else
                        {
                            indexBegin = i + 1;
                        }
                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = rsBeforeFormat.NewRow();
                            //DataRow[] rowTemp = new DataRow[inputTable.Rows.Count];
                            //inputTable.Rows.CopyTo(rowTemp, i);
                            //row = rowTemp[0];

                            for (int col = 0; col < listNameColumn.Count; col++)
                            {
                                row[listNameColumn[col].ToString()] = inputTable.Rows[i][listNameColumn[col].ToString()];
                            }
                            
                            // Frequency
                            row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j])/1000000;
                            
                            rsBeforeFormat.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency
                        

                        string[] arrFrequence = inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString().Trim().Split(';');

                        if (arrFrequence.Length > 1)
                        {
                            
                            for (int ai = 0; ai < arrFrequence.Length; ai++)
                            {
                                if (!String.IsNullOrEmpty(arrFrequence[ai]))
                                {
                                    DataRow row = rsBeforeFormat.NewRow();
                                    for (int col = 0; col < listNameColumn.Count; col++)
                                    {
                                        row[listNameColumn[col].ToString()] = inputTable.Rows[i][listNameColumn[col].ToString()];
                                    }

                                    // Frequency
                                    row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(arrFrequence[ai]) / 1000000;

                                    // Add row
                                    rsBeforeFormat.Rows.Add(row);
                                }
                            }
                        }
                        else
                        {
                            //row = inputTable.Rows[i];
                            // Copy all data to new row
                            DataRow row = rsBeforeFormat.NewRow();
                            for (int col = 0; col < listNameColumn.Count; col++)
                            {
                                row[listNameColumn[col].ToString()] = inputTable.Rows[i][listNameColumn[col].ToString()];
                            }

                            // Frequency
                            row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString()) / 1000000;

                            // Add row
                            rsBeforeFormat.Rows.Add(row);
                        }
                    }
                }
            }

            // Reset stt
            for (int stt = 1; stt <= rsBeforeFormat.Rows.Count; stt++)
            {
                rsBeforeFormat.Rows[stt - 1][listNameColumn[0].ToString()] = stt;
            }


            return rsBeforeFormat;
        }

        public DataTable GetTCITableBeforeFormat(DataTable inputTable, Dictionary<int, ArrayList> listRange)
        {
            DataTable rsBeforeFormat = inputTable.Clone();
            //Utilities util = new Utilities();
            objUti = new Utilities();

            ArrayList listNameColumn = objUti.GetColumnName(inputTable);

            if (inputTable != null && inputTable.Rows.Count > 0)
            {
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if (listRange.ContainsKey(i))
                    {
                        // Row i has range frequency
                        int indexBegin = 0;
                        if (rsBeforeFormat.Rows.Count > 0)
                        {
                            indexBegin = rsBeforeFormat.Rows.Count + i + 1;
                        }
                        else
                        {
                            indexBegin = i + 1;
                        }
                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = rsBeforeFormat.NewRow();
                            //DataRow[] rowTemp = new DataRow[inputTable.Rows.Count];
                            //inputTable.Rows.CopyTo(rowTemp, i);
                            //row = rowTemp[0];

                            for (int col = 0; col < listNameColumn.Count; col++)
                            {
                                row[listNameColumn[col].ToString()] = inputTable.Rows[i][listNameColumn[col].ToString()];
                            }

                            // Frequency
                           // row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j]);

                            rsBeforeFormat.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency


                        string[] arrFrequence = inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString().Trim().Split(';');

                        if (arrFrequence.Length > 1)
                        {

                            for (int ai = 0; ai < arrFrequence.Length; ai++)
                            {
                                if (!String.IsNullOrEmpty(arrFrequence[ai]))
                                {
                                    DataRow row = rsBeforeFormat.NewRow();
                                    for (int col = 0; col < listNameColumn.Count; col++)
                                    {
                                        row[listNameColumn[col].ToString()] = inputTable.Rows[i][listNameColumn[col].ToString()];
                                    }

                                    // Frequency
                                 //   row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(arrFrequence[ai]);

                                    // Add row
                                    rsBeforeFormat.Rows.Add(row);
                                }
                            }
                        }
                        else
                        {
                            //row = inputTable.Rows[i];
                            // Copy all data to new row
                            DataRow row = rsBeforeFormat.NewRow();
                            for (int col = 0; col < listNameColumn.Count; col++)
                            {
                                row[listNameColumn[col].ToString()] = inputTable.Rows[i][listNameColumn[col].ToString()];
                            }

                            // Frequency
                           // row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString());

                            // Add row
                            rsBeforeFormat.Rows.Add(row);
                        }
                    }
                }
            }

            // Reset stt
            for (int stt = 1; stt <= rsBeforeFormat.Rows.Count; stt++)
            {
                rsBeforeFormat.Rows[stt - 1][listNameColumn[0].ToString()] = stt;
            }


            return rsBeforeFormat;
        }

        public DataTable GetRSTableOutput(DataTable inputTable, Dictionary<int, ArrayList> listRange)
        {
            objUti = new Utilities();
            DataTable rsFormatTable = objUti.GetTemplateTableRS();

            if (inputTable != null && inputTable.Rows.Count > 0)
            {
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if (listRange.ContainsKey(i))
                    {
                        // Row i has range frequency
                        int indexBegin = 0;
                        if (rsFormatTable.Rows.Count > 0)
                        {
                            indexBegin = rsFormatTable.Rows.Count + i + 1;
                        }
                        else
                        {
                            indexBegin = i + 1;
                        }
                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = rsFormatTable.NewRow();

                            #region 5 Fields of set 1
                            // Ten khach hang
                            row[Constants.TableExport.TEN_KHACH_HANG] =
                                inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(";", "");

                            // Frequency
                            row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j]);

                            // offset
                            if (inputTable.Columns.Contains(Constants.TableExport.OFFSET_FREQ))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.OFFSET_FREQ].ToString()))
                                {
                                    row[Constants.TableExport.OFFSET_FREQ] =
                                        inputTable.Rows[i][Constants.TableExport.OFFSET_FREQ].ToString().Trim().Replace(";", "");
                                }
                                else
                                {
                                    // Null or emty --> set to 0(Hz)
                                    row[Constants.TableExport.OFFSET_FREQ] = 0;
                                }
                            }
                            else
                            {
                                // Null or emty --> set to 0(Hz)
                                row[Constants.TableExport.OFFSET_FREQ] = 0;
                            }

                            // Dich vu
                            if (inputTable.Columns.Contains(Constants.TableExport.DICH_VU))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DICH_VU].ToString()))
                                {
                                    string dichvuValue = inputTable.Rows[i][Constants.TableExport.DICH_VU].ToString().Trim().Replace(";", "");
                                    if (dichvuValue.Length > 2)
                                    {
                                        dichvuValue = dichvuValue.Substring(0, 2);
                                    }

                                    row[Constants.TableExport.DICH_VU] = dichvuValue;
                                }
                                else
                                {
                                    // Null or emty --> set to BC
                                    row[Constants.TableExport.DICH_VU] = "BC";
                                }
                            }
                            else
                            {
                                // Null or emty --> set to BC
                                row[Constants.TableExport.DICH_VU] = "BC";
                            }

                            // Ky hieu
                            if (inputTable.Columns.Contains(Constants.TableExport.KY_HIEU))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.KY_HIEU].ToString()))
                                {
                                    string kyhieuValue = inputTable.Rows[i][Constants.TableExport.KY_HIEU].ToString().Trim().Replace(";", "");
                                    if (kyhieuValue.Length > 20)
                                    {
                                        kyhieuValue = kyhieuValue.Substring(0, 20);
                                    }

                                    row[Constants.TableExport.KY_HIEU] = kyhieuValue;
                                }
                                else
                                {
                                    // Null or emty -->space
                                    row[Constants.TableExport.KY_HIEU] = Constants.ValueConstant.SPACE;
                                }
                            }
                            else
                            {
                                row[Constants.TableExport.KY_HIEU] = Constants.ValueConstant.SPACE;
                            }
                            #endregion

                            #region 5 fields of set 2
                            // 5 fields of set 2
                            // Ho Hieu
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString()))
                            {
                                string hohieuValue = inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Replace(";", "");
                                if (hohieuValue.Length > 32)
                                {
                                    hohieuValue = hohieuValue.Substring(0, 32);
                                }

                                row[Constants.TableExport.HO_HIEU] = hohieuValue;
                            }
                            else
                            {
                                // Null or emty -->space
                                row[Constants.TableExport.HO_HIEU] = Constants.ValueConstant.SPACE;
                            }

                            // GP No
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.GPNo].ToString()))
                            {
                                string gpNoValue = inputTable.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Replace(";", "");
                                if (gpNoValue.Length > 32)
                                {
                                    gpNoValue = gpNoValue.Substring(0, 32);
                                }

                                row[Constants.TableExport.GPNo] = gpNoValue;
                            }
                            else
                            {
                                // Null or emty -->space
                                row[Constants.TableExport.GPNo] = Constants.ValueConstant.SPACE;
                            }

                            // Dien thoai
                            if (inputTable.Columns.Contains(Constants.TableExport.DIEN_THOAI))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DIEN_THOAI].ToString()))
                                {
                                    string dienthoaiValue = inputTable.Rows[i][Constants.TableExport.DIEN_THOAI].ToString().Trim().Replace(";", "");
                                    if (dienthoaiValue.Length > 20)
                                    {
                                        dienthoaiValue = dienthoaiValue.Substring(0, 20);
                                    }

                                    row[Constants.TableExport.DIEN_THOAI] = dienthoaiValue;
                                }
                                else
                                {
                                    // Null or emty -->space
                                    row[Constants.TableExport.DIEN_THOAI] = Constants.ValueConstant.SPACE;
                                }
                            }
                            else
                            {
                                row[Constants.TableExport.DIEN_THOAI] = Constants.ValueConstant.SPACE;
                            }

                            // Ten ma dat nuoc
                            if (inputTable.Columns.Contains(Constants.TableExport.TEN_MA_DAT_NUOC))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TEN_MA_DAT_NUOC].ToString()))
                                {
                                    string tenmacontryValue = inputTable.Rows[i][Constants.TableExport.TEN_MA_DAT_NUOC].ToString().Trim().Replace(";", "");
                                    if (tenmacontryValue.Length > 3)
                                    {
                                        tenmacontryValue = tenmacontryValue.Substring(0, 3);
                                    }

                                    row[Constants.TableExport.TEN_MA_DAT_NUOC] = tenmacontryValue;
                                }
                                else
                                {
                                    // Null or emty -->space
                                    row[Constants.TableExport.TEN_MA_DAT_NUOC] = Constants.ValueConstant.SPACE;
                                }
                            }
                            else
                            {
                                // Null or emty -->space
                                row[Constants.TableExport.TEN_MA_DAT_NUOC] = Constants.ValueConstant.SPACE;
                            }

                            // Zip code
                            if (inputTable.Columns.Contains(Constants.TableExport.ZIP_CODE))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.ZIP_CODE].ToString()))
                                {
                                    string zipcodeValue = inputTable.Rows[i][Constants.TableExport.ZIP_CODE].ToString().Trim().Replace(";", "");
                                    if (zipcodeValue.Length > 8)
                                    {
                                        zipcodeValue = zipcodeValue.Substring(0, 8);
                                    }

                                    row[Constants.TableExport.ZIP_CODE] = zipcodeValue;
                                }
                                else
                                {
                                    // Null or emty -->space
                                    row[Constants.TableExport.ZIP_CODE] = Constants.ValueConstant.SPACE;
                                }
                            }
                            else
                            {
                                // Null or emty -->space
                                row[Constants.TableExport.ZIP_CODE] = Constants.ValueConstant.SPACE;
                            }

                            #endregion

                            #region 5 fields of set 3
                            // 5 fields of set 3
                            // Tinh thanh pho
                            if (inputTable.Columns.Contains(Constants.TableExport.TINH_THANH))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TINH_THANH].ToString()))
                                {
                                    string tinhthanhValue = inputTable.Rows[i][Constants.TableExport.TINH_THANH].ToString().Trim().Replace(";", "");
                                    if (tinhthanhValue.Length > 40)
                                    {
                                        tinhthanhValue = tinhthanhValue.Substring(0, 40);
                                    }

                                    row[Constants.TableExport.TINH_THANH] = tinhthanhValue;
                                }
                                else
                                {
                                    // Null or emty -->space
                                    row[Constants.TableExport.TINH_THANH] = Constants.ValueConstant.SPACE;
                                }
                            }
                            else
                            {
                                // Null or emty -->space
                                row[Constants.TableExport.TINH_THANH] = Constants.ValueConstant.SPACE;
                            }

                            // Duong pho
                            if (inputTable.Columns.Contains(Constants.TableExport.DUONG_PHO))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DUONG_PHO].ToString()))
                                {
                                    string gpNoValue = inputTable.Rows[i][Constants.TableExport.DUONG_PHO].ToString().Trim().Replace(";", "");
                                    if (gpNoValue.Length > 32)
                                    {
                                        gpNoValue = gpNoValue.Substring(0, 32);
                                    }

                                    row[Constants.TableExport.DUONG_PHO] = gpNoValue;
                                }
                                else
                                {
                                    // Null or emty -->space
                                    row[Constants.TableExport.DUONG_PHO] = Constants.ValueConstant.SPACE;
                                }
                            }
                            else
                            {
                                // Null or emty -->space
                                row[Constants.TableExport.DUONG_PHO] = Constants.ValueConstant.SPACE;
                            }

                            // Kinh do, vi do
                            string[] kinhdoVidoArr =
                                objUti.GetKinhdoAndVido(inputTable.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
                            if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
                            {
                                // Get normal
                                row[Constants.TableExport.VI_DO] = objUti.FormatLatitudeRS(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);
                            }
                            else
                            {
                                // Call random value;
                                row[Constants.TableExport.VI_DO] = objUti.FormatLatitudeRS(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
                            }

                            if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
                            {
                                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitudeRS(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
                            }
                            else
                            {
                                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitudeRS(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
                            }

                            // Huong dai phat
                            if (inputTable.Columns.Contains(Constants.TableExport.HUONG_DAI_PHAT))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HUONG_DAI_PHAT].ToString()))
                                {

                                    row[Constants.TableExport.HUONG_DAI_PHAT] =
                                        Convert.ToDouble(
                                            inputTable.Rows[i][Constants.TableExport.HUONG_DAI_PHAT].ToString().Trim().
                                                Replace(";", ""));
                                }
                                else
                                {
                                    // Null or emty -->0
                                    row[Constants.TableExport.HUONG_DAI_PHAT] = 0;
                                }
                            }
                            else
                            {
                                // Null or emty -->0
                                row[Constants.TableExport.HUONG_DAI_PHAT] = 0;
                            }
                            #endregion

                            #region 5 fields of set 4
                            // 5 fields of set 4
                            // Khoang cach dai phat
                            if (inputTable.Columns.Contains(Constants.TableExport.KHOANG_CACH_DAI_PHAT))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.KHOANG_CACH_DAI_PHAT].ToString()))
                                {
                                    row[Constants.TableExport.KHOANG_CACH_DAI_PHAT] =
                                        Convert.ToDouble(
                                            inputTable.Rows[i][Constants.TableExport.KHOANG_CACH_DAI_PHAT].ToString().Trim()
                                                .Replace(";", ""));
                                }
                                else
                                {
                                    // Null or emty -->0
                                    row[Constants.TableExport.KHOANG_CACH_DAI_PHAT] = 0;
                                }
                            }
                            else
                            {
                                // Null or emty -->0
                                row[Constants.TableExport.KHOANG_CACH_DAI_PHAT] = 0;
                            }

                            // Min do lech tan so
                            if (inputTable.Columns.Contains(Constants.TableExport.KHOANG_CACH_DAI_PHAT))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MIN_DO_LECH_FREQ].ToString()))
                                {
                                    row[Constants.TableExport.MIN_DO_LECH_FREQ] =
                                        Convert.ToDouble(
                                            inputTable.Rows[i][Constants.TableExport.MIN_DO_LECH_FREQ].ToString().Trim().
                                                Replace(";", "")); ;
                                }
                                else
                                {
                                    // Null or emty -->0 Hz
                                    row[Constants.TableExport.MIN_DO_LECH_FREQ] = 0;
                                }
                            }
                            else
                            {
                                // Null or emty -->0 Hz
                                row[Constants.TableExport.MIN_DO_LECH_FREQ] = 0;
                            }

                            // Bang thong
                            if (inputTable.Columns.Contains(Constants.TableExport.BANG_THONG))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.BANG_THONG].ToString()))
                                {
                                    row[Constants.TableExport.BANG_THONG] =
                                        Convert.ToDouble(
                                            inputTable.Rows[i][Constants.TableExport.BANG_THONG].ToString().Trim().Replace(
                                                ";", "")); ;
                                }
                                else
                                {
                                    // Null or emty -->0 Hz
                                    row[Constants.TableExport.BANG_THONG] = 0;
                                }
                            }
                            else
                            {
                                // Null or emty -->0 Hz
                                row[Constants.TableExport.BANG_THONG] = 0;
                            }

                            // MIn dieu che
                            if (inputTable.Columns.Contains(Constants.TableExport.MIN_DIEU_CHE))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MIN_DIEU_CHE].ToString()))
                                {
                                    row[Constants.TableExport.MIN_DIEU_CHE] =
                                        Convert.ToDouble(
                                            inputTable.Rows[i][Constants.TableExport.MIN_DIEU_CHE].ToString().Trim().Replace
                                                (";", ""));
                                }
                                else
                                {
                                    // Null or emty -->0
                                    row[Constants.TableExport.MIN_DIEU_CHE] = 0;
                                }
                            }
                            else
                            {
                                // Null or emty -->0
                                row[Constants.TableExport.MIN_DIEU_CHE] = 0;
                            }

                            // Don vi dieu che
                            if (inputTable.Columns.Contains(Constants.TableExport.DON_VI_DIEU_CHE))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DON_VI_DIEU_CHE].ToString()))
                                {
                                    row[Constants.TableExport.DON_VI_DIEU_CHE] =
                                        inputTable.Rows[i][Constants.TableExport.DON_VI_DIEU_CHE].ToString().Trim().Replace(";", "");
                                }
                                else
                                {
                                    // Null or emty -->% or β
                                    row[Constants.TableExport.DON_VI_DIEU_CHE] = "%";
                                }
                            }
                            else
                            {
                                // Null or emty -->% or β
                                row[Constants.TableExport.DON_VI_DIEU_CHE] = "%";
                            }

                            #endregion
                            
                            rsFormatTable.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency
                        string[] arrFrequence = inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString().Trim().Split(';');

                        if (arrFrequence.Length > 1)
                        {

                            for (int ai = 0; ai < arrFrequence.Length; ai++)
                            {
                                if (!String.IsNullOrEmpty(arrFrequence[ai]))
                                {
                                    DataRow row = rsFormatTable.NewRow();
                                    row = this.RSDataRow(arrFrequence[ai], row, inputTable, i);

                                    // Frequency
                                    //row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(arrFrequence[ai]);

                                    // Add row
                                    rsFormatTable.Rows.Add(row);
                                }
                            }
                        }
                        else
                        {
                            //row = inputTable.Rows[i];
                            // Copy all data to new row
                            DataRow row = rsFormatTable.NewRow();
                            row = this.RSDataRow(inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString(), row, inputTable, i);

                            // Frequency
                            //row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(inputTable.Rows[i][Constants.TableExport.TAN_SO].ToString()) / 1000000;

                            // Add row
                            rsFormatTable.Rows.Add(row);
                        }
                    }
                }
            }

            return rsFormatTable;
        }

        public DataTable GetGEFrequencyTableOutput(DataTable inputTable, Dictionary<int, ArrayList> listRange)
        {
            objUti = new Utilities();
            DataTable freFormatTable = objUti.GetTemplateTableFrequency();

            if (inputTable != null && inputTable.Rows.Count > 0)
            {
                int k = 1;
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if (listRange.ContainsKey(i))
                    {

                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = freFormatTable.NewRow();

                            if (inputTable.Columns.Contains(Constants.TableExport.STT))
                            {
                                row[Constants.TableExport.GEWTABLE.FREQUENCY_EXTERNAL_ID] = inputTable.Rows[i][Constants.TableExport.STT].ToString();
                            }

                            //band width
                            row[Constants.TableExport.GEWTABLE.TRANSMITTER_EXTERNAL_ID] = k;
                            if (inputTable.Columns.Contains(Constants.TableExport.DAI_LL))
                            {
                                int width = 0;
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DAI_LL].ToString()))
                                {
                                    string bandwidth = inputTable.Rows[i][Constants.TableExport.DAI_LL].ToString().Trim().Replace(";", "");

                                    if (bandwidth.Contains(Constants.ValueDAILL._16K0F3E))
                                    {
                                        width = 25000;
                                    }
                                    else if (bandwidth.Contains(Constants.ValueDAILL._11K0F3E) || bandwidth.Contains(Constants.ValueDAILL._6K50) || bandwidth.Contains(Constants.ValueDAILL._6K5F3E))
                                    {
                                        width = 12500;
                                    }
                                    else
                                    {
                                        width = 25000;
                                    }
                                }
                                else
                                {
                                    width = 0;
                                }
                                row[Constants.TableExport.GEWTABLE.BANDWIDTH] = width;
                                row[Constants.TableExport.GEWTABLE.CHANNEL_SPACE] = width;
                            }

                            // Frequency
                            row[Constants.TableExport.GEWTABLE.CENTRE_FREQUENCY] = Convert.ToDouble(listfrequencyRange[j]);

                            // Dich vu
                            if (inputTable.Columns.Contains(Constants.TableExport.MUC_DICH_SU_DUNG))
                            {
                                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MUC_DICH_SU_DUNG].ToString()))
                                {
                                    string dichvuValue = inputTable.Rows[i][Constants.TableExport.MUC_DICH_SU_DUNG].ToString().Trim().Replace(";", "");
                                    row[Constants.TableExport.GEWTABLE.CHANNEL_NAME] = dichvuValue;
                                }
                                else
                                {
                                    // Null or emty --> set to BC
                                    row[Constants.TableExport.GEWTABLE.CHANNEL_NAME] = "";
                                }
                            }
                            else
                            {
                                // Null or emty --> set to BC
                                row[Constants.TableExport.GEWTABLE.CHANNEL_NAME] = "";
                            }
                            k++;
                            freFormatTable.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency
                        DataRow row = freFormatTable.NewRow();

                        row[Constants.TableExport.GEWTABLE.TRANSMITTER_EXTERNAL_ID] = k;
                        k++;
                        // Add row
                        freFormatTable.Rows.Add(row);
                    }
                }
            }

            return freFormatTable;

        }

        public DataTable GetGETranmisterTableOutput(DataTable inputTable, Dictionary<int, ArrayList> listRange)
        {
            objUti = new Utilities();
            DataTable freFormatTable = objUti.GetTemplateTableTranmister();

            if (inputTable != null && inputTable.Rows.Count > 0)
            {
                int k = 1;
                for (int i = 0; i < inputTable.Rows.Count; i++)
                {
                    if (listRange.ContainsKey(i))
                    {

                        ArrayList listfrequencyRange = listRange[i];

                        for (int j = 0; j < listfrequencyRange.Count; j++)
                        {
                            // This row does not had range frequency
                            DataRow row = freFormatTable.NewRow();

                            

                            //band width
                            row[Constants.TableExport.GEWTABLE.TRANSMITTER_EXTERNAL_ID] = k;
                            
                            // Five columns normal
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString()))
                            {
                                row[Constants.TableExport.GEWTABLE.NAME] =
                                    inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.GEWTABLE.NAME] = Constants.ValueConstant.SPACE;
                            }


                            // Five columns normal
                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString()))
                            {
                                row[Constants.TableExport.GEWTABLE.TYPE] =
                                    inputTable.Rows[i][Constants.TableExport.MAU_GIAY_PHEP].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.GEWTABLE.TYPE] = Constants.ValueConstant.SPACE;
                            }


                            // Xu ly kinh do vi do
                            // Longtitude and latitude
                            string[] kinhdoVidoArr =
                                objUti.GetKinhdoAndVido(inputTable.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
                            if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
                            {
                                // Get normal
                                row[Constants.TableExport.GEWTABLE.LATITUDE] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);

                            }
                            else
                            {
                                // Call random value;
                                row[Constants.TableExport.GEWTABLE.LATITUDE] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
                            }

                            if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
                            {
                                row[Constants.TableExport.GEWTABLE.LONGITUDE] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
                            }
                            else
                            {
                                row[Constants.TableExport.GEWTABLE.LONGITUDE] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
                            }

                            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MUC_DICH_SU_DUNG].ToString()))
                            {
                                row[Constants.TableExport.GEWTABLE.COMMENT] =
                                    inputTable.Rows[i][Constants.TableExport.MUC_DICH_SU_DUNG].ToString().Trim().Replace(
                                        ";", "");
                            }
                            else
                            {
                                row[Constants.TableExport.GEWTABLE.TYPE] = Constants.ValueConstant.SPACE;
                            }
                            k++;
                            freFormatTable.Rows.Add(row);
                        }

                    }
                    else
                    {
                        // This row does not had range frequency
                        DataRow row = freFormatTable.NewRow();

                        row[Constants.TableExport.GEWTABLE.TRANSMITTER_EXTERNAL_ID] = k;
                        k++;
                        // Add row
                        freFormatTable.Rows.Add(row);
                    }
                }
            }

            return freFormatTable;

        }

        

        private DataRow RSDataRow(string tanso, DataRow row, DataTable inputTable, int i)
        {
            #region 5 Fields of set 1
            // Ten khach hang
            row[Constants.TableExport.TEN_KHACH_HANG] =
                inputTable.Rows[i][Constants.TableExport.TEN_KHACH_HANG].ToString().Trim().Replace(";", "");

            // Frequency
            row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(tanso);
            //row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j]);

            // offset
            if (inputTable.Columns.Contains(Constants.TableExport.OFFSET_FREQ))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.OFFSET_FREQ].ToString()))
                {
                    row[Constants.TableExport.OFFSET_FREQ] =
                        inputTable.Rows[i][Constants.TableExport.OFFSET_FREQ].ToString().Trim().Replace(";", "");
                }
                else
                {
                    // Null or emty --> set to 200.000(Hz)
                    row[Constants.TableExport.OFFSET_FREQ] = 200000;
                }
            }
            else
            {
                // Null or emty --> set to 200.000(Hz)
                row[Constants.TableExport.OFFSET_FREQ] = 200000;
            }

            // Dich vu
            if (inputTable.Columns.Contains(Constants.TableExport.DICH_VU))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DICH_VU].ToString()))
                {
                    string dichvuValue = inputTable.Rows[i][Constants.TableExport.DICH_VU].ToString().Trim().Replace(";", "");
                    if (dichvuValue.Length > 2)
                    {
                        dichvuValue = dichvuValue.Substring(0, 2);
                    }

                    row[Constants.TableExport.DICH_VU] = dichvuValue;
                }
                else
                {
                    // Null or emty --> set to 200.000(Hz)
                    row[Constants.TableExport.DICH_VU] = "BC";
                }
            }
            else
            {
                // Null or emty --> set to 200.000(Hz)
                row[Constants.TableExport.DICH_VU] = "BC";
            }

            // Ky hieu
            if (inputTable.Columns.Contains(Constants.TableExport.KY_HIEU))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.KY_HIEU].ToString()))
                {
                    string kyhieuValue = inputTable.Rows[i][Constants.TableExport.KY_HIEU].ToString().Trim().Replace(";", "");
                    if (kyhieuValue.Length > 20)
                    {
                        kyhieuValue = kyhieuValue.Substring(0, 20);
                    }

                    row[Constants.TableExport.KY_HIEU] = kyhieuValue;
                }
                else
                {
                    // Null or emty -->space
                    row[Constants.TableExport.KY_HIEU] = Constants.ValueConstant.SPACE;
                }
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.KY_HIEU] = Constants.ValueConstant.SPACE;
            }

            #endregion

            #region 5 fields of set 2
            // 5 fields of set 2
            // Ho Hieu
            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString()))
            {
                string hohieuValue = inputTable.Rows[i][Constants.TableExport.HO_HIEU].ToString().Trim().Replace(";", "");
                if (hohieuValue.Length > 32)
                {
                    hohieuValue = hohieuValue.Substring(0, 32);
                }

                row[Constants.TableExport.HO_HIEU] = hohieuValue;
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.HO_HIEU] = Constants.ValueConstant.SPACE;
            }

            // GP No
            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.GPNo].ToString()))
            {
                string gpNoValue = inputTable.Rows[i][Constants.TableExport.GPNo].ToString().Trim().Replace(";", "");
                if (gpNoValue.Length > 32)
                {
                    gpNoValue = gpNoValue.Substring(0, 32);
                }

                row[Constants.TableExport.GPNo] = gpNoValue;
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.GPNo] = Constants.ValueConstant.SPACE;
            }

            // Dien thoai
            if (inputTable.Columns.Contains(Constants.TableExport.DIEN_THOAI))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DIEN_THOAI].ToString()))
                {
                    string dienthoaiValue = inputTable.Rows[i][Constants.TableExport.DIEN_THOAI].ToString().Trim().Replace(";", "");
                    if (dienthoaiValue.Length > 20)
                    {
                        dienthoaiValue = dienthoaiValue.Substring(0, 20);
                    }

                    row[Constants.TableExport.DIEN_THOAI] = dienthoaiValue;
                }
                else
                {
                    // Null or emty -->space
                    row[Constants.TableExport.DIEN_THOAI] = Constants.ValueConstant.SPACE;
                }
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.DIEN_THOAI] = Constants.ValueConstant.SPACE;
            }

            // Ten ma dat nuoc
            if (inputTable.Columns.Contains(Constants.TableExport.TEN_MA_DAT_NUOC))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TEN_MA_DAT_NUOC].ToString()))
                {
                    string tenmacontryValue = inputTable.Rows[i][Constants.TableExport.TEN_MA_DAT_NUOC].ToString().Trim().Replace(";", "");
                    if (tenmacontryValue.Length > 3)
                    {
                        tenmacontryValue = tenmacontryValue.Substring(0, 3);
                    }

                    row[Constants.TableExport.TEN_MA_DAT_NUOC] = tenmacontryValue;
                }
                else
                {
                    // Null or emty -->space
                    row[Constants.TableExport.TEN_MA_DAT_NUOC] = Constants.ValueConstant.SPACE;
                }
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.TEN_MA_DAT_NUOC] = Constants.ValueConstant.SPACE;
            }

            // Zip code
            if (inputTable.Columns.Contains(Constants.TableExport.ZIP_CODE))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.ZIP_CODE].ToString()))
                {
                    string zipcodeValue = inputTable.Rows[i][Constants.TableExport.ZIP_CODE].ToString().Trim().Replace(";", "");
                    if (zipcodeValue.Length > 8)
                    {
                        zipcodeValue = zipcodeValue.Substring(0, 8);
                    }

                    row[Constants.TableExport.ZIP_CODE] = zipcodeValue;
                }
                else
                {
                    // Null or emty -->space
                    row[Constants.TableExport.ZIP_CODE] = Constants.ValueConstant.SPACE;
                }
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.ZIP_CODE] = Constants.ValueConstant.SPACE;
            }

            #endregion

            #region 5 fields of set 3
            // 5 fields of set 3
            // Tinh thanh pho
            if (inputTable.Columns.Contains(Constants.TableExport.TINH_THANH))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.TINH_THANH].ToString()))
                {
                    string tinhthanhValue = inputTable.Rows[i][Constants.TableExport.TINH_THANH].ToString().Trim().Replace(";", "");
                    if (tinhthanhValue.Length > 40)
                    {
                        tinhthanhValue = tinhthanhValue.Substring(0, 40);
                    }

                    row[Constants.TableExport.TINH_THANH] = tinhthanhValue;
                }
                else
                {
                    // Null or emty -->space
                    row[Constants.TableExport.TINH_THANH] = Constants.ValueConstant.SPACE;
                }
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.TINH_THANH] = Constants.ValueConstant.SPACE;
            }

            // Duong pho
            if (inputTable.Columns.Contains(Constants.TableExport.DUONG_PHO))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DUONG_PHO].ToString()))
                {
                    string gpNoValue = inputTable.Rows[i][Constants.TableExport.DUONG_PHO].ToString().Trim().Replace(";", "");
                    if (gpNoValue.Length > 32)
                    {
                        gpNoValue = gpNoValue.Substring(0, 32);
                    }

                    row[Constants.TableExport.DUONG_PHO] = gpNoValue;
                }
                else
                {
                    // Null or emty -->space
                    row[Constants.TableExport.DUONG_PHO] = Constants.ValueConstant.SPACE;
                }
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.DUONG_PHO] = Constants.ValueConstant.SPACE;
            }

            // Kinh do, vi do
            string[] kinhdoVidoArr =
                objUti.GetKinhdoAndVido(inputTable.Rows[i][Constants.TableExport.KINHDO_VIDO].ToString().Trim());
            if (kinhdoVidoArr[1] != Constants.ValueConstant.RANDOM)
            {
                // Get normal
                row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.NORMAL);

            }
            else
            {
                // Call random value;
                row[Constants.TableExport.VI_DO] = objUti.FormatLatitude(kinhdoVidoArr[1].Trim(), Constants.ValueConstant.RANDOM);
            }

            if (kinhdoVidoArr[0] != Constants.ValueConstant.RANDOM)
            {
                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.NORMAL);
            }
            else
            {
                row[Constants.TableExport.KINH_DO] = objUti.FormatLongtitude(kinhdoVidoArr[0].Trim(), Constants.ValueConstant.RANDOM);
            }

            // Huong dai phat
            if (inputTable.Columns.Contains(Constants.TableExport.HUONG_DAI_PHAT))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.HUONG_DAI_PHAT].ToString()))
                {

                    row[Constants.TableExport.HUONG_DAI_PHAT] =
                        Convert.ToDouble(
                            inputTable.Rows[i][Constants.TableExport.HUONG_DAI_PHAT].ToString().Trim().
                                Replace(";", ""));
                }
                else
                {
                    // Null or emty -->0
                    row[Constants.TableExport.HUONG_DAI_PHAT] = 0;
                }
            }
            else
            {
                // Null or emty -->0
                row[Constants.TableExport.HUONG_DAI_PHAT] = 0;
            }

            #endregion

            #region 5 fields of set 4
            // 5 fields of set 4
            // Khoang cach dai phat
            if (inputTable.Columns.Contains(Constants.TableExport.KHOANG_CACH_DAI_PHAT))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.KHOANG_CACH_DAI_PHAT].ToString()))
                {
                    row[Constants.TableExport.KHOANG_CACH_DAI_PHAT] =
                        Convert.ToDouble(
                            inputTable.Rows[i][Constants.TableExport.KHOANG_CACH_DAI_PHAT].ToString().Trim()
                                .Replace(";", ""));
                }
                else
                {
                    // Null or emty -->0
                    row[Constants.TableExport.KHOANG_CACH_DAI_PHAT] = 0;
                }
            }
            else
            {
                // Null or emty -->0
                row[Constants.TableExport.KHOANG_CACH_DAI_PHAT] = 0;
            }
            // Min do lech tan so
            if (inputTable.Columns.Contains(Constants.TableExport.MIN_DO_LECH_FREQ))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MIN_DO_LECH_FREQ].ToString()))
                {
                    row[Constants.TableExport.MIN_DO_LECH_FREQ] =
                        Convert.ToDouble(
                            inputTable.Rows[i][Constants.TableExport.MIN_DO_LECH_FREQ].ToString().Trim().
                                Replace(";", "")); ;
                }
                else
                {
                    // Null or emty -->200 Hz
                    row[Constants.TableExport.MIN_DO_LECH_FREQ] = 200;
                }
            }
            else
            {
                // Null or emty -->200 Hz
                row[Constants.TableExport.MIN_DO_LECH_FREQ] = 200;
            }

            // Bang thong
            if (inputTable.Columns.Contains(Constants.TableExport.BANG_THONG))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.BANG_THONG].ToString()))
                {
                    row[Constants.TableExport.BANG_THONG] =
                        Convert.ToDouble(
                            inputTable.Rows[i][Constants.TableExport.BANG_THONG].ToString().Trim().Replace(
                                ";", "")); ;
                }
                else
                {
                    // Null or emty -->200.000 Hz
                    row[Constants.TableExport.BANG_THONG] = 200000;
                }
            }
            else
            {
                // Null or emty -->200.000 Hz
                row[Constants.TableExport.BANG_THONG] = 200000;
            }

            // MIn dieu che
            if (inputTable.Columns.Contains(Constants.TableExport.MIN_DIEU_CHE))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MIN_DIEU_CHE].ToString()))
                {
                    row[Constants.TableExport.MIN_DIEU_CHE] =
                        Convert.ToDouble(
                            inputTable.Rows[i][Constants.TableExport.MIN_DIEU_CHE].ToString().Trim().Replace
                                (";", ""));
                }
                else
                {
                    // Null or emty -->0
                    row[Constants.TableExport.MIN_DIEU_CHE] = 0;
                }
            }
            else
            {
                // Null or emty -->0
                row[Constants.TableExport.MIN_DIEU_CHE] = 0;
            }

            // Don vi dieu che
            if (inputTable.Columns.Contains(Constants.TableExport.DON_VI_DIEU_CHE))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DON_VI_DIEU_CHE].ToString()))
                {
                    row[Constants.TableExport.DON_VI_DIEU_CHE] =
                        inputTable.Rows[i][Constants.TableExport.DON_VI_DIEU_CHE].ToString().Trim().Replace(";", "");
                }
                else
                {
                    // Null or emty -->% or β
                    row[Constants.TableExport.DON_VI_DIEU_CHE] = "%";
                }
            }
            else
            {
                // Null or emty -->% or β
                row[Constants.TableExport.DON_VI_DIEU_CHE] = "%";
            }

            #endregion

            return row;
        }

        private DataRow FreDataRow(string tanso, DataRow row, DataTable inputTable, int i)
        {

            // Frequency
            row[Constants.TableExport.TAN_SO] = objUti.FormatFrequency(tanso);
            //row[Constants.TableExport.TAN_SO] = Convert.ToDouble(listfrequencyRange[j]);

            // Dich vu
            if (inputTable.Columns.Contains(Constants.TableExport.MUC_DICH_SU_DUNG))
            {
                if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.MUC_DICH_SU_DUNG].ToString()))
                {
                    string dichvuValue = inputTable.Rows[i][Constants.TableExport.MUC_DICH_SU_DUNG].ToString().Trim().Replace(";", "");

                    row[Constants.TableExport.MUC_DICH_SU_DUNG] = dichvuValue;
                }
                else
                {
                    // Null or emty --> set to 200.000(Hz)
                    row[Constants.TableExport.MUC_DICH_SU_DUNG] = "BC";
                }
            }
            else
            {
                // Null or emty --> set to 200.000(Hz)
                row[Constants.TableExport.MUC_DICH_SU_DUNG] = "BC";
            }

            

           
            // 5 fields of set 2
            // Ho Hieu
            if (!String.IsNullOrEmpty(inputTable.Rows[i][Constants.TableExport.DAI_LL].ToString()))
            {
                string hohieuValue = inputTable.Rows[i][Constants.TableExport.DAI_LL].ToString().Trim().Replace(";", "");
                

                row[Constants.TableExport.DAI_LL] = hohieuValue;
            }
            else
            {
                // Null or emty -->space
                row[Constants.TableExport.DAI_LL] = Constants.ValueConstant.SPACE;
            }

            

            return row;
        }

        public List<OutFormatBE> GetFormatOutput(DataTable inputTable,double step, string freq)
        {
            List<OutFormatBE> outList = new List<OutFormatBE>();
            //check input data and reformat data.
            Dictionary<Int64, OutFormatBE> dicOutFormat = new Dictionary<long, OutFormatBE>();
            OutFormatBE outBe;
            int index = 1;
            objUti = new Utilities();
            foreach (DataRow row in inputTable.Rows)
            {
                string gpNum = row[Constants.TableExport.GPNo].ToString();
                string gpType = row[Constants.TableExport.MAU_GIAY_PHEP].ToString();
                string refNum = row[Constants.TableExport.SO_THAM_CHIEU].ToString();
                string biasF = row[Constants.TableExport.DO_LECH_F].ToString();
                string prioband = String.Empty;
                string bandwidth = String.Empty;
                string noOfChanel = String.Empty;
                string customerName = row[Constants.TableExport.TEN_KHACH_HANG].ToString();
                string callName = row[Constants.TableExport.HO_HIEU].ToString();
                string temp = row[Constants.TableExport.KINHDO_VIDO].ToString();
                string[] kinhvido = GetLongitudeAndLatitude(temp);
                double longitude;
                double latitude;
                if (!Constants.ValueConstant.RANDOM.Equals(kinhvido[0]))
                {
                    longitude = FormatLongtitude(kinhvido[0], Constants.ValueConstant.NORMAL);
                }
                else
                {
                    longitude = FormatLongtitude(kinhvido[0], Constants.ValueConstant.RANDOM);
                }


                if (!Constants.ValueConstant.RANDOM.Equals(kinhvido[1]))
                {
                    latitude = FormatLatitude(kinhvido[1], Constants.ValueConstant.NORMAL);
                }
                else
                {
                    latitude = FormatLatitude(kinhvido[1], Constants.ValueConstant.RANDOM);
                }


                string machName = row[Constants.TableExport.TEN_MAY].ToString();

                string frequently = row[Constants.TableExport.TAN_SO].ToString();                

                List<double> listFreq = GetFrequencyByRange(frequently, step, freq);
                foreach (double outFre in listFreq)
                {
                    outBe = new OutFormatBE();
                    outBe.Index = index + 1;
                    outBe.GPNo = gpNum;
                    outBe.GPType = gpType;
                    outBe.RefNo = refNum;
                    outBe.FBias = biasF;
                    outBe.Frequency = outFre;
                    outBe.PrioBand = prioband;
                    outBe.BandWidth = bandwidth;
                    outBe.NoChanel = noOfChanel;
                    outBe.CustomerName = customerName;
                    outBe.Call = callName;
                    outBe.Longitude = longitude;
                    outBe.Latitude = latitude;
                    outBe.MachineName = machName;
                    outList.Add(outBe);
                }
            }
            return outList;
        }

        /// <summary>
        /// Get kinh do and vi do
        /// </summary>
        /// <param name="kinhdovido"></param>
        /// <returns></returns>
        private string[] GetLongitudeAndLatitude(string longAndLat)
        {
            string[] LongLatArr = new string[2];

            // Split kinhdovido
            // (105E50'57.35" /20N56'52.77" );
            // Replace character no need
            longAndLat = longAndLat.Replace("(", "");
            longAndLat = longAndLat.Replace(")", "");
            longAndLat = longAndLat.Replace(";", "");

            string[] arrTemp = longAndLat.Trim().Split('/');

            if (arrTemp.Length <= 2
                && arrTemp.Length > 0
                && arrTemp[0].Length > 3
                && arrTemp[1].Length > 2)
            {
                LongLatArr[0] = arrTemp[0];
                LongLatArr[1] = arrTemp[1];
            }
            else
            {
                // Set random longtitude and latitude
                LongLatArr[0] = Constants.ValueConstant.RANDOM;
                LongLatArr[1] = Constants.ValueConstant.RANDOM;
            }

            return LongLatArr;
        }

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


        // Save random value
        private Dictionary<double, double> randomLongtitudeDict = null;

        private Dictionary<double, double> randomLatitudeDict = null;


        private List<double> GetFrequencyByRange(string range, double step,string freq)
        {
            List<double> arr = new List<double>();
            return arr;
        }
    }
}
