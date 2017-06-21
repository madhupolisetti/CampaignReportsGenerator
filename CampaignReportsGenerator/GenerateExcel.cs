using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using NPOI.XSSF.UserModel;
using System.IO;

namespace CampaignReportsGenerator
{
    public class GenerateExcel 
    {
        SqlConnection sqlCon = (SqlConnection)null;
        SqlCommand sqlCmd = (SqlCommand)null;
        DataSet ds = (DataSet)null;
        SqlDataAdapter da = (SqlDataAdapter)null;
        SqlConnection updateSqlCon = (SqlConnection)null;
        SqlCommand updateSqlCmd = (SqlCommand)null;
        
        public void GenerateReports(CampaignReports campaignReportsObj)
        {
            try
            {
                sqlCon = new SqlConnection(SharedClass.GetConnectionString(campaignReportsObj.SourceDataBase));
                sqlCmd = new SqlCommand("Get_CampaignReports", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.Add("@CampaignScheduleId", SqlDbType.BigInt).Value = campaignReportsObj.CampaignScheduleId;
                sqlCmd.Parameters.Add("@FileName", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output;
                sqlCmd.Parameters.Add("@Success", SqlDbType.Bit).Direction = ParameterDirection.Output;
                sqlCmd.Parameters.Add("@Message", SqlDbType.VarChar, -1).Direction = ParameterDirection.Output;
                da = new SqlDataAdapter(sqlCmd);
                ds = new DataSet();
                da.Fill(ds);
                if(Convert.ToBoolean(sqlCmd.Parameters["@Success"].Value))
                {
                    if(ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        //string fileName = "Reports_" + campaignReportsObj.CampaignScheduleId.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xlsx";
                        string fileName = sqlCmd.Parameters["@FileName"].Value.ToString();
                        if (fileName.Length == 0)
                            fileName = "Reports_" + campaignReportsObj.CampaignScheduleId.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xlsx";
                        else
                            fileName = fileName + ".xlsx";     
                        bool isGenerated = GenerateExcelReports(ds.Tables[0], fileName , campaignReportsObj.SourceDataBase);
                        if(isGenerated)
                        {
                            this.UpdateRequest(requestId: campaignReportsObj.RequestId, sourceDataBase: campaignReportsObj.SourceDataBase, status: 2, message: "Generated", fileName: fileName);
                        }
                        else
                        {
                            this.UpdateRequest(requestId: campaignReportsObj.RequestId, sourceDataBase: campaignReportsObj.SourceDataBase, status: 3, message: "Error in Generating ExcelFile");
                        }
                    }
                    else
                    {
                        this.UpdateRequest(requestId: campaignReportsObj.RequestId, sourceDataBase: campaignReportsObj.SourceDataBase, status: 3, message: "No Records Found");
                    }
                }
                else
                {
                    this.UpdateRequest(requestId: campaignReportsObj.RequestId, sourceDataBase: campaignReportsObj.SourceDataBase, status: 1, message: sqlCmd.Parameters["@Message"].Value.ToString());
                }

            }
            catch(Exception ex)
            {
                SharedClass.Logger.Error("Exception in GenerateReports Reason: " + ex.ToString());
                this.UpdateRequest(requestId: campaignReportsObj.RequestId, sourceDataBase: campaignReportsObj.SourceDataBase, status: 3, message: ex.ToString());
            }
            finally
            {
                try { da.Dispose(); }
                catch (Exception ex) { SharedClass.Logger.Error("Exception in disposing DataAdapter is :" + ex.ToString()); }
                finally { da = null; }

                try { ds.Dispose(); }
                catch (Exception ex) { SharedClass.Logger.Error("Exception in disposing DataSet is :" + ex.ToString()); }
                finally { ds = null; }
                sqlCmd = null;
                sqlCon = null;

            }
        }
        
        public bool GenerateExcelReports(DataTable reportsTable, string fileName, SourceDatabase sourceDataBase)
        {
            Int16 tableRowIndex = 0, sheetRowIndex = 0, columnIndex = 0;
            Int64 maxRowsPerSheet = 0;
            XSSFSheet currentSheet = null;
            XSSFSheet previousSheet = null;
            Dictionary<string, XSSFSheet> sheetsDic = new Dictionary<string, XSSFSheet>();
            NPOI.SS.UserModel.IRow sheetRow = null;
            string cellValue = "", filePathFull = "";
            FileStream fileStream = null;
            XSSFWorkbook workBook = new XSSFWorkbook();

            try
            {
                maxRowsPerSheet = SharedClass.MaxRowsPerSheet;
                filePathFull = SharedClass.GetFileSavePath(sourceDataBase) + fileName;
                Dictionary<string, string> strDic = new Dictionary<string, string>();
                sheetsDic["Sheet1"] = workBook.CreateSheet("Sheet1") as XSSFSheet;
                currentSheet = sheetsDic["Sheet1"];
                SharedClass.Logger.Info("Started loading data into " + currentSheet.SheetName);
                Console.WriteLine("Started loading data into " + currentSheet.SheetName);
                previousSheet = currentSheet;
                sheetRow = currentSheet.CreateRow(sheetRowIndex);
                columnIndex = 0;
                foreach (DataColumn dataColumn in reportsTable.Columns)
                {
                    cellValue = dataColumn.ColumnName;
                    sheetRow.CreateCell(columnIndex).SetCellValue(cellValue);
                    columnIndex += 1;
                }
                sheetRowIndex += 1;
                foreach (DataRow dataRow in reportsTable.Rows)
                {

                    if (tableRowIndex <= maxRowsPerSheet)
                    {
                        currentSheet = sheetsDic["Sheet1"];
                    }
                    else if (tableRowIndex % maxRowsPerSheet == 0)
                    {
                        try
                        {
                            currentSheet = sheetsDic["Sheet" + (Convert.ToInt32(tableRowIndex / maxRowsPerSheet)).ToString()];
                        }
                        catch (Exception ex)
                        {
                            sheetsDic["Sheet" + (Convert.ToInt16(tableRowIndex / maxRowsPerSheet)).ToString()] = workBook.CreateSheet("Sheet" + Convert.ToInt16(tableRowIndex / maxRowsPerSheet).ToString()) as XSSFSheet;
                            currentSheet = sheetsDic["Sheet" + Convert.ToInt16(tableRowIndex / maxRowsPerSheet).ToString()];
                            SharedClass.Logger.Info("Started loading data into " + currentSheet.SheetName);
                            //Console.WriteLine("Started loading data into " + currentSheet.SheetName);
                        }
                    }
                    else
                    {
                        try
                        {
                            currentSheet = sheetsDic["Sheet" + (Convert.ToInt16(tableRowIndex / maxRowsPerSheet) + 1).ToString()];
                        }
                        catch (Exception ex)
                        {

                            sheetsDic["Sheet" + (Convert.ToInt32(tableRowIndex / maxRowsPerSheet) + 1).ToString()] = workBook.CreateSheet("Sheet" + (Convert.ToInt16(tableRowIndex / maxRowsPerSheet) + 1).ToString()) as XSSFSheet;
                            currentSheet = sheetsDic["Sheet" + (Convert.ToInt32(tableRowIndex / maxRowsPerSheet) + 1).ToString()];
                            SharedClass.Logger.Info("Started loading data into " + currentSheet.SheetName);
                            //Console.WriteLine("Started loading data into " + currentSheet.SheetName);
                        }
                    }

                    if (previousSheet.SheetName != currentSheet.SheetName)
                    {
                        sheetRowIndex = 0;
                        sheetRow = currentSheet.CreateRow(sheetRowIndex);
                        columnIndex = 0;
                        foreach (DataColumn dataColumn in reportsTable.Columns)
                        {
                            cellValue = dataColumn.ColumnName;
                            sheetRow.CreateCell(columnIndex).SetCellValue(cellValue);
                            columnIndex += 1;
                        }
                        sheetRowIndex += 1;
                    }
                    sheetRow = currentSheet.CreateRow(sheetRowIndex);
                    columnIndex = 0;
                    foreach (DataColumn dataColumn in dataRow.Table.Columns)
                    {
                        if (dataRow[dataColumn.ColumnName] != null)
                        {
                            cellValue = dataRow[dataColumn.ColumnName].ToString();
                        }
                        else
                        {
                            cellValue = "";
                        }
                        sheetRow.CreateCell(columnIndex).SetCellValue(cellValue);
                        columnIndex += 1;
                    }
                    sheetRowIndex += 1;
                    tableRowIndex += 1;
                    previousSheet = currentSheet;
                }
                SharedClass.Logger.Info("All Sheets were loaded into workbook");
                //Console.WriteLine("All Sheets Were Loaded Into WorkBook");
                SharedClass.Logger.Info("Creating file " + filePathFull);
                //Console.WriteLine("Creating File " + filePathFull);
                fileStream = new FileStream(filePathFull, FileMode.CreateNew);
                SharedClass.Logger.Info("File created");
                //Console.WriteLine("File Created");
                SharedClass.Logger.Info("Dumping data into " + filePathFull);
                //Console.WriteLine("Dumping Data Into " + filePathFull);
                workBook.Write(fileStream);
                fileStream.Close();
                //Console.WriteLine("Data Dumped");
                return true;
            }
            catch (Exception ex)
            {
                SharedClass.Logger.Error("Error in generating excel excel");
                //Console.WriteLine("Error In Generating Excel : " + ex.ToString());
                return false;
            }
        }

        public void UpdateRequest(long requestId , SourceDatabase sourceDataBase, Int16 status, string message, string fileName = "")
        {    
            try
            {
                SharedClass.Logger.Info("Updating the Request Of Id : " + requestId.ToString() + " Status :" + status.ToString() + " Reason: " + message.ToString() + " fileName: " + fileName.ToString() + " SourceDatBase:" + sourceDataBase.ToString());
                updateSqlCon = new SqlConnection(SharedClass.GetConnectionString(sourceDataBase));
                updateSqlCmd = new SqlCommand("Update_CampaignScheduleReports", updateSqlCon);
                updateSqlCmd.CommandType = CommandType.StoredProcedure;
                updateSqlCmd.Parameters.Add("@ReportsId", SqlDbType.BigInt).Value = requestId;
                updateSqlCmd.Parameters.Add("@Status", SqlDbType.TinyInt).Value = status;
                updateSqlCmd.Parameters.Add("@Reason", SqlDbType.VarChar, 1000).Value = message;
                if(fileName != "")
                    updateSqlCmd.Parameters.Add("@FileName", SqlDbType.VarChar, 500).Value = fileName;
                updateSqlCmd.Parameters.Add("@Success", SqlDbType.Bit).Direction = ParameterDirection.Output;
                updateSqlCmd.Parameters.Add("@Message", SqlDbType.VarChar, -1).Direction = ParameterDirection.Output;
                updateSqlCon.Open();
                updateSqlCmd.ExecuteNonQuery();
                if (!Convert.ToBoolean(updateSqlCmd.Parameters["@Success"].Value))
                    SharedClass.Logger.Error("Not Updated CampaignScheduleReports of Id " + requestId.ToString()  + " Reason : " + updateSqlCmd.Parameters["@Message"].Value.ToString());
            }
            catch(Exception ex)
            {
                SharedClass.Logger.Error("Exception in updating the CampaignScheduleReports Of Id " + requestId.ToString() + " Reason : " + ex.ToString());
            }
            finally
            {
                try { updateSqlCmd.Dispose(); }
                catch (Exception ex) { SharedClass.Logger.Error("Exception in disposing updateSqlCmd is :" + ex.ToString()); }
                finally { updateSqlCmd = null; }

                try { updateSqlCon.Dispose(); }
                catch (Exception ex) { SharedClass.Logger.Error("Exception in disposing updateSqlCon is :" + ex.ToString()); }
                finally { updateSqlCon = null; }
            }
        }

        

    }
}
