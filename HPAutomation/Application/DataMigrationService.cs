namespace HPAutomation.Application
{
    using HPAutomation.Infrastructure;
    using HPAutomation.Logging;
    using System;
    using System.Data;
    using System.Data.OleDb;
    using System.Data.SqlClient;
    using System.Text;

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class DataMigrationService
    {
        public void MigrateData()
        {
            var filePath = AppSettings.FileLocation;

            string extension = Path.GetExtension(filePath).ToLower();
            string[] validFileTypes = { ".xls", ".xlsx" };

            if (!validFileTypes.Contains(extension))
            {
                LogFile.WriteLog($"File extension {extension} is not valid type");
            }

            var connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";

            ImportExceltoDatabase(connString, filePath);
        }

        private void ImportExceltoDatabase(string connString, string filePath)
        {
            LogFile.WriteLog($"Log: ImportExceltoDatabase execution started");

            var excelConnection = new OleDbConnection(connString);
            string filename = Path.GetFileNameWithoutExtension(filePath);

            try
            {
                excelConnection.Open();
                DataTable dataTables = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                foreach (DataRow drSheet in dataTables.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    {
                        string sheetname = drSheet["TABLE_NAME"].ToString();

                        //Load the DataTable with Sheet Data
                        var oconn = new OleDbCommand("select * from [" + sheetname + "]", excelConnection);
                        var adp = new OleDbDataAdapter(oconn);
                        DataTable dt = new DataTable();
                        adp.Fill(dt);


                        string createTableScript = string.Empty;
                        createTableScript += "IF Not EXISTS (SELECT * FROM sys.objects WHERE object_id = ";
                        createTableScript += "OBJECT_ID(N'[dbo].[" + filename + "]') AND type in (N'U'))";
                        createTableScript += "Create table [" + filename + "]";
                        createTableScript += "(";

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            if (i != dt.Columns.Count - 1)
                                createTableScript += "[" + dt.Columns[i].ColumnName + "] " + "NVarchar(max)" + ",";
                            else
                                createTableScript += "[" + dt.Columns[i].ColumnName + "] " + "NVarchar(max)";
                        }

                        createTableScript += ")";

                        using (var sqlConnection = new SqlConnection(AppSettings.DbConnectionString))
                        {
                            sqlConnection.Open();

                            var tableCrationCommand = new SqlCommand(createTableScript, sqlConnection);
                            tableCrationCommand.ExecuteNonQuery();
                            LogFile.WriteLog($"Log: Table {filename} created successfully.");

                            var blk = new SqlBulkCopy(sqlConnection);
                            blk.DestinationTableName = "[" + filename + "]";
                            blk.WriteToServer(dt);
                            LogFile.WriteLog($"Log: Data inserted into table '{filename}' successfully.");

                            LogFile.WriteLog($"Log: automation_proc execution started");

                            string proc_ddl = "exec automation_proc " + filename;
                            var spExecutionCommand = new SqlCommand(proc_ddl, sqlConnection);
                            spExecutionCommand.ExecuteNonQuery();

                            LogFile.WriteLog($"Log: automation_proc execution completed");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var stringBuilder = new StringBuilder();

                while (ex != null)
                {
                    stringBuilder.Append(ex.Message);
                    ex = ex.InnerException;
                }

                LogFile.WriteLog($"Error: An error occurred while importing migrating the data. {stringBuilder.ToString()}");
            }
            finally
            {
                excelConnection.Close();
            }

            LogFile.WriteLog($"Log: ImportExceltoDatabase execution completed");
        }
    }
}
