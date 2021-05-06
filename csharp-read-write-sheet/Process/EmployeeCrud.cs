using csharp_read_write_sheet.Configuration;
using csharp_read_write_sheet.Employee;
using csharp_read_write_sheet.Helper;
using csharp_read_write_sheet.Helpers;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csharp_read_write_sheet
{
    public class EmployeeCrud:SheetConfiguration
    {
        static string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
        private const string Process = "Employee Crud";
        private const string Section = "employeecrud";

        private int RowsLinked;
        private int AccountBatchSize=3;
        static Dictionary<string, long> columnMap = new Dictionary<string, long>();
      
        public EmployeeCrud()
        {
            try
            {

                var settings = (NameValueCollection)ConfigurationManager.GetSection(Section);
                ConfigSheetId = Convert.ToInt64(ConfigSheetId);
                ConfigSheet = Client.GetSheet(ConfigSheetId);
                Logger.ErrorSheet = Client.GetSheet(ConfigSheetId);
                Logger.RunLogSheet = Client.GetSheet(ConfigSheetId);
                ConfigManager = new ConfigManager(ConfigSheet);

                this.InitLogs(this);
            }
            catch (Exception e)
            {
                throw new ApplicationException($"Unable to get configuration sheet due to exception: {e.Message}");
            }
        }

        public async Task Run()
        {
            Logger.LogToConsole($"Starting {Process}");
            try
            {
                BulkInsertDbDataToSmartSheet();
                FetchEmployeeDatas();
                CreateNewEmployeeDatas();
                UpdateEmployeeDatas();
                DeleteEmployeeDatas();
            }
            catch (Exception e)
            {
                if (e is ConfigurationErrorsException)
                {
                    Logger.LogToConsole($"Config initialization failed with exception: {e.Message}");
                    Logger.LogException(e, "Config initialization failed.");
                }
                else
                {
                    Logger.LogToConsole($"{Process} failed with Exception: {e.Message}");
                    Logger.LogException(e, $"{Process} failed");
                }

                var startTime = StartTime.ToString(CultureInfo.InvariantCulture);
                var endTime = DateTime.Now.ToString(CultureInfo.InvariantCulture);
                Logger.LogJobRun(startTime, endTime, $"{Process} failed.", true);
                throw e;
            }
            LogJobRun();
            Logger.LogToConsole($"Done...");
            Console.ReadLine();

        }

        public static DataTable FetchEmployeeDatas()
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                try
                {
                    //Logger.LogToConsole($"Started Fetching data from database...");
                    con.Open();
                    SqlCommand cmd = new SqlCommand("Select * from Employee", con);
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd.CommandText, con);
                    dataAdapter.Fill(dt);
                   // Logger.LogToConsole($"Fetching data process compeleted...");
                }
                catch (Exception ex)
                {
                    var message = $"Failed to get employee details: {ex.Message}";
                    Logger.LogException(ex, message);
                    throw new ApplicationException(message, ex);
                }

            }
            return dt;
        }
        IEnumerable<EmployeeModel> employeeList = FetchEmployeeDatas().AsEnumerable().Select(row => new EmployeeModel
        {
            EmployeeId = row.Field<int>("EmployeeId"),
            FirstName = row.Field<string>("FirstName"),
            LastName = row.Field<string>("LastName"),
            Email = row.Field<string>("Email"),
            Address = row.Field<string>("Address")
        }).ToList();
        public void CreateNewEmployeeDatas()
        {
            var sheet = Client.GetSheet(ConfigSheetId);
            foreach (Column column in sheet.Columns)
                columnMap.Add(column.Title, (long)column.Id);
            DataTable dt = FetchEmployeeDatas();
            int targetEmployeeval;
            List<int> sheetEmpIds = new List<int>();
            var accountsToCreate = new List<EmployeeModel>();
            for (int i = 0; i < sheet.Rows.Count; i++)
            {
                targetEmployeeval = Convert.ToInt32(sheet.Rows[i].GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN));
                sheetEmpIds.Add(targetEmployeeval);
            }
            if (dt.Rows.Count != sheet.Rows.Count)
            {
                foreach (var dbrow in employeeList)
                {

                    if (!sheetEmpIds.Contains(dbrow.EmployeeId))
                    {
                        //Cell[] newcell = new Cell[]
                        //{
                        //      new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE1_COLUMN],dbrow.EmployeeId).Build(),
                        //      new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE2_COLUMN],dbrow.FirstName).Build(),
                        //      new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE3_COLUMN],dbrow.LastName).Build(),
                        //      new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE4_COLUMN],dbrow.Email).Build(),
                        //      new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE5_COLUMN],dbrow.Address).Build(),
                        //};
                        //Row rowA = new Row.AddRowBuilder(null, true, null, null, null).SetCells(newcell).Build();
                        //Client.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });

                        accountsToCreate.Add(dbrow);
                    }
                }

                if (accountsToCreate.Any())
                {
                    var intakeRows = BuildNewIntakeRows(accountsToCreate, sheet);

                    while (intakeRows.Any())
                    {
                        var takeCount = intakeRows.Count < AccountBatchSize ? intakeRows.Count : AccountBatchSize;
                        var rows = intakeRows.Take(takeCount).ToList();

                        var importedRows = Client.SheetResources.RowResources.AddRows(sheet.Id.Value, rows);

                        Logger.LogToConsole($"Imported {rows.Count} rows to {sheet.Name}");

                        intakeRows = intakeRows.Except(rows).ToList();
                    }
                }
            }
        }

        private List<Row> BuildNewIntakeRows(List<EmployeeModel> accounts, Sheet sheet)
        {
            Logger.LogToConsole("Building new intake rows");

            return (from account in accounts
                    select new Row
                    {
                        Cells = new List<Cell>
                        {
                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE1_COLUMN],account.EmployeeId).Build(),
                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE2_COLUMN],account.FirstName).Build(),
                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE3_COLUMN],account.LastName).Build(),
                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE4_COLUMN],account.Email).Build(),
                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE5_COLUMN],account.Address).Build(),
                        },
                        ToBottom = true
                    }).ToList();
        }
        public void UpdateEmployeeDatas()
        {
            var sheet = Client.GetSheet(ConfigSheetId);
            var rows = new List<Row>();
            try
            {
                Logger.LogToConsole($"Started to Update employee details...");
                foreach (var emp in employeeList)
                {
                    var accountRow = FindAccountRow(sheet, emp);
                    //var FirstNameColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE2_COLUMN, false)?.Id;
                    //var LastNameColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE3_COLUMN, false)?.Id;
                    //var EmailColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE4_COLUMN, false)?.Id;
                    //var AddressColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE5_COLUMN, false)?.Id;
                    var FirstNameColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE2_COLUMN, false)?.Id;
                    var LastNameColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE3_COLUMN, false)?.Id;
                    var EmailColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE4_COLUMN, false)?.Id;
                    var AddressColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE5_COLUMN,false)?.Id;

                    rows.Add(new Row
                    {
                        Id = accountRow.Id,
                        Cells = new List<Cell>
                            {
                                new Cell
                                {
                                    ColumnId = FirstNameColumnId,
                                    Value = emp.FirstName

                                },
                                new Cell
                                {
                                    ColumnId = LastNameColumnId,
                                    Value = emp.LastName
                                },
                                new Cell
                                {
                                    ColumnId = EmailColumnId,
                                    Value = emp.Email
                                },
                                new Cell
                                {
                                    ColumnId = AddressColumnId,
                                    Value = emp.Address
                                }
                            }
                    });
                }
                var rowsUpdated = Client.SheetResources.RowResources.UpdateRows(sheet.Id.Value, rows).Count;
                Logger.LogToConsole($"Successfully {rowsUpdated:N0} rows as completed in sheet {sheet.Id.Value}");
                Logger.LogToConsole($"UpdateData's Compeleted...");
            }
            catch (Exception ex)
            {
                var message = $"Failed to update employee details: {ex.Message}";
                Logger.LogException(ex, message);
                throw new ApplicationException(message, ex);
            }
          
        }
        public void DeleteEmployeeDatas()
        {
            //***************To Remove All the Rows****************
            //var allRows = sheet.Rows.ToList();
            //while (allRows.Any())
            //{
            //    var takeCount = allRows.Count < 10 ? allRows.Count : 10;
            //    var rowsToRemove = allRows.Take(takeCount).ToList();
            //    var identifiersRemoved = smartsheet.SheetResources.RowResources.DeleteRows(sheet.Id.Value,
            //        rowsToRemove.Select(x => (long) x.Id).ToList(), true);
            //        allRows.RemoveAll(row => identifiersRemoved.Any(id => id == row.Id));
            //}
            //***************To Remove All the Rows****************
            try
            {
               
                Logger.LogToConsole($"Started Delete rows...");
                var sheet = Client.GetSheet(ConfigSheetId);
                var sourceEmployeeIdList = employeeList.Select(x => x.EmployeeId).ToList();
                foreach (var row in sheet.Rows)
                {
                    int? targetEmployeeId = Convert.ToInt32(row.GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN));
                    if (targetEmployeeId != null)
                    {
                        if (!sourceEmployeeIdList.Contains(targetEmployeeId.Value))
                        {
                            var rowsUpdated = Client.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { (long)row.Id }.ToList(), true).Count;
                            Client.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { (long)row.Id }.ToList(), true);
                        }
                    }
                    else
                        break;
                }
                Logger.LogToConsole($"delete employee data's completed...");
            }
            catch (Exception ex)
            {
                var message = $"Failed to Delet Row: {ex.Message}";
                Logger.LogException(ex, message);
                throw new ApplicationException(message, ex);
            }
        }
        public void BulkInsertDbDataToSmartSheet()
        {
            var sheet = Client.GetSheet(ConfigSheetId);
            DataTable dt = FetchEmployeeDatas();
            try
            {
                if (sheet.Rows.Count==0 && (dt.Rows.Count != sheet.Rows.Count))
                {
                    foreach (Column column in sheet.Columns)
                        columnMap.Add(column.Title, (long)column.Id);
                    Logger.LogToConsole($"Bulk Insert Started...");
                    Cell[] cellsA = null;
                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        for (int j = 0; j < sheet.Columns.Count; j++)
                        {
                            cellsA = new Cell[]
                            {
                                    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE1_COLUMN],dt.Rows[i][0]).Build(),
                                    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE2_COLUMN],dt.Rows[i][1]).Build(),
                                    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE3_COLUMN],dt.Rows[i][2]).Build(),
                                    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE4_COLUMN],dt.Rows[i][3]).Build(),
                                    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE5_COLUMN],dt.Rows[i][4]).Build(),
                                    //new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE6_COLUMN],"").Build(),
                                    //new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE7_COLUMN],"").Build(),
                                    //new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE8_COLUMN],"").Build(),
                                    //new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE9_COLUMN],"").Build(),
                            };
                        }
                        Row rowA = new Row.AddRowBuilder(true, null, null, null, null).SetCells(cellsA).Build();
                        Client.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
                    }
                    Logger.LogToConsole($"Bulk Insert completed...");
                }
            }
            catch(Exception ex)
            {
                var message = $"Failed to Insert BulkInsert Rows: {ex.Message}";
                Logger.LogException(ex, message);
                throw new ApplicationException(message, ex);
            }
        }
        static Row FindAccountRow(Sheet sheet, EmployeeModel employee)
        {
            foreach (var row in sheet.Rows)
            {
                var sourceEmployeeId = row.GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN);
                if (string.IsNullOrWhiteSpace(sourceEmployeeId))
                    continue;
                var RowExist = (sourceEmployeeId == employee.EmployeeId.ToString());
                if (RowExist)
                {
                    return row;
                }
            }
            throw new ApplicationException($"Cannot find account row for {employee} in sheet {sheet.Id.Value}");
        }
        private void LogJobRun()
        {
            Logger.LogToConsole($"{Process} complete");

            var startdate = StartTime.ToString(CultureInfo.InvariantCulture);
            var enddate = DateTime.Now.ToString(CultureInfo.InvariantCulture);
            var notes = $"{Process} complete. rows imported: {RowsLinked}";
           // Logger.LogJobRun(startdate, enddate, notes, false);
          
        }
    }
}
