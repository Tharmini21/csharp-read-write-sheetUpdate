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
    public class EmployeeCrud : SheetConfiguration
    {
        static string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
        private const string Process = "Employee Crud";
        private const string Section = "employeecrud";

        private int RowsLinked;
        private int AccountBatchSize = 5;
        public static int Currentpagesize = 10;
        public static int currentPageNumber = 1;
        public static int pageNumber = 0;
        // public static DataTable dt;
        public static bool Processflag = false;
        public static int totalPages = 0;
        public static int CurrentBatch = 1;
        List<int> existingRowIds = new List<int>();
        List<int> newlistRowIds = new List<int>();
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
            Logger.LogToConsole($"Starting {Process} ,Start Process Time: {DateTime.UtcNow}");
            try
            {
                //FetchEmployeeDatas();
                //sBulkInsertDbDataToSmartSheet();
                CreateNewEmployeeDatas();
                UpdateEmployeeDatas();
                DeleteEmployeeDatas();

                if ((totalPages - 1) == currentPageNumber)
                {
                    Processflag = false;
                }
                else
                {
                    Processflag = true;
                }
                dt = null;
                employeeList = null;
                if (Processflag == true)
                {
                    dt = FetchEmployeeDatas();
                    employeeList = dt.AsEnumerable().Select(row => new EmployeeModel
                    {
                        EmployeeId = row.Field<int>("EmployeeId"),
                        FirstName = row.Field<string>("FirstName"),
                        LastName = row.Field<string>("LastName"),
                        Email = row.Field<string>("Email"),
                        Address = row.Field<string>("Address")
                    }).ToList();
                    await Run();
                }
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
            DataTable data = new DataTable();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM Employee", con);
                    int totalcount = (Int32)cmd.ExecuteScalar();
                    totalPages = (int)Math.Ceiling((decimal)totalcount / (decimal)Currentpagesize);
                    currentPageNumber = currentPageNumber + pageNumber;
                    if (Processflag == true)
                    {
                        if (currentPageNumber <= 1)
                        {
                            currentPageNumber = 1;
                            pageNumber = currentPageNumber;
                        }
                        else if (currentPageNumber > totalPages)
                        {
                            currentPageNumber = totalPages;
                            pageNumber = currentPageNumber;
                        }
                        else
                        {
                            pageNumber = currentPageNumber;
                        }
                    }
                    string querySelect = "Select * from Employee ORDER BY EmployeeId Asc OFFSET " + Currentpagesize + " * " + pageNumber + " ROWS FETCH NEXT " + Currentpagesize + " ROWS ONLY";
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(querySelect, con);
                    dataAdapter.Fill(data);
                }
                catch (Exception ex)
                {
                    var message = $"Failed to get employee details: {ex.Message}";
                    Logger.LogException(ex, message);
                    throw new ApplicationException(message, ex);
                }
            }
            return data;
        }


        static IEnumerable<EmployeeModel> employeeList = FetchEmployeeDatas().AsEnumerable().Select(row => new EmployeeModel
        {
            EmployeeId = row.Field<int>("EmployeeId"),
            FirstName = row.Field<string>("FirstName"),
            LastName = row.Field<string>("LastName"),
            Email = row.Field<string>("Email"),
            Address = row.Field<string>("Address")
        }).ToList();

        DataTable dt = ConvertToDataTable(employeeList);
        public static DataTable ConvertToDataTable(IEnumerable<EmployeeModel> source)
        {
            var props = typeof(EmployeeModel).GetProperties();

            var dt = new DataTable();
            dt.Columns.AddRange(
              props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray()
            );

            source.ToList().ForEach(
              i => dt.Rows.Add(props.Select(p => p.GetValue(i, null)).ToArray())
            );

            return dt;
        }

        public void CreateNewEmployeeDatas()
        {
            int targetEmployeeval;
            List<int> sheetEmpIds = new List<int>();
            var accountsToCreate = new List<EmployeeModel>();
            try
            {
                var sheet = Client.GetSheet(ConfigSheetId);
                if (columnMap.Count == 0)
                {
                    foreach (Column column in sheet.Columns)
                        columnMap.Add(column.Title, (long)column.Id);
                }
                for (int i = 0; i < sheet.Rows.Count; i++)
                {
                    targetEmployeeval = Convert.ToInt32(sheet.Rows[i].GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN));
                    sheetEmpIds.Add(targetEmployeeval);
                }
                foreach (var dbrow in employeeList)
                {
                    if (!sheetEmpIds.Contains(dbrow.EmployeeId))
                    {
                        accountsToCreate.Add(dbrow);
                    }
                }
                if (accountsToCreate.Any())
                {
                    var intakeRows = BuildNewIntakeRows(accountsToCreate, sheet);

                    while (intakeRows.Any())
                    {
                        Logger.LogToConsole($"Batch Number " + CurrentBatch + " Running StartTime: {DateTime.UtcNow}");
                        Logger.LogToConsole($"Currently Running Batch Number :" + CurrentBatch);

                        var takeCount = intakeRows.Count < AccountBatchSize ? intakeRows.Count : AccountBatchSize;
                        var rows = intakeRows.Take(takeCount).ToList();

                        //var importedRows = Client.SheetResources.RowResources.AddRows(sheet.Id.Value, rows);

                        RowWrapper rowWrapper = new RowWrapper.InsertRowsBuilder().SetRows(rows).SetToBottom(true).Build();
                        //var rowwrap = Client.SheetResources.RowResources.AddRows(sheet.Id.Value, rows);
                        var rowwrap = Client.SheetResources.RowResources.AddRows(sheet.Id.Value, rows);
                        //smartsheet.Sheets().Rows().InsertRows(sheet.Id.Value, rowWrapper);
                        //sheet.Sheets().Rows().InsertRows(sheet.Id.Value, rowWrapper);

                        Logger.LogToConsole($"Imported {rows.Count} rows to {sheet.Name}");

                        intakeRows = intakeRows.Except(rows).ToList();
                        CurrentBatch++;

                    }
                    Logger.LogToConsole($"Batch Number " + CurrentBatch + " Running EndTime: {DateTime.UtcNow}");
                }
            }
            catch(Exception ex)
            {
                var message = $"Failed to Upload Batch Number "+CurrentBatch+" Data: {ex.Message}";
                Logger.LogException(ex, message);
                throw new ApplicationException(message, ex);
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
                    var FirstNameColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE2_COLUMN, false)?.Id;
                    var LastNameColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE3_COLUMN, false)?.Id;
                    var EmailColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE4_COLUMN, false)?.Id;
                    var AddressColumnId = sheet.GetColumnByTitle(ConfigManager.CONFIGURATION_VALUE5_COLUMN, false)?.Id;

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

                //Row rowA = new Row.UpdateRowBuilder().setCells(cellsB).setRowId(rowId).build();
                //List<Row> updatedRows = smartsheet.sheetResources().rowResources().updateRows(sheet.Id.Value, Arrays.asList(rows));
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
                Dictionary<int, long> sheetEmpIdDic = new Dictionary<int, long>();
                List<int> sheetEmpIds = new List<int>();
                for (int i = 0; i < sheet.Rows.Count; i++)
                {
                    int id = Convert.ToInt32(sheet.Rows[i].GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN));
                    long columId = Convert.ToInt64(sheet.Rows[i].Id);
                    sheetEmpIds.Add(id);
                    sheetEmpIdDic.Add(id, columId);
                }
                int skip = (pageNumber * Currentpagesize);
                if (pageNumber == 0)
                {
                    sheetEmpIds = sheetEmpIds.Skip(skip).Take(Currentpagesize).ToList();
                    sheetEmpIdDic = sheetEmpIdDic.Skip(skip).Take(Currentpagesize).ToDictionary(pair => pair.Key, pair => pair.Value);
                }
                else
                {
                    sheetEmpIds = sheetEmpIds.Skip(skip).Take(Currentpagesize).ToList();
                    sheetEmpIdDic = sheetEmpIdDic.Skip(skip).Take(Currentpagesize).ToDictionary(pair => pair.Key, pair => pair.Value);
                }
                foreach (var row in sheetEmpIds)
                {
                    if (!sourceEmployeeIdList.Contains(row))
                    {
                        if (sheetEmpIdDic.ContainsKey(row))
                        {
                            long smartcellId = Convert.ToInt64(sheetEmpIdDic[row]);
                            var rowsUpdated = Client.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { smartcellId }.ToList(), true).Count;
                        }
                    }
                }
                // Existing Code//
                //foreach (var row in sheet.Rows)
                //{
                //    int? targetEmployeeId = Convert.ToInt32(row.GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN));
                //    if (targetEmployeeId != null)
                //    {
                //        if (!sourceEmployeeIdList.Contains(targetEmployeeId.Value))
                //        {
                //            var rowsUpdated = Client.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { (long)row.Id }.ToList(), true).Count;
                //            // Client.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { (long)row.Id }.ToList(), true);
                //        }
                //    }
                //    else
                //        break;
                //}
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
            //DataTable dt = FetchEmployeeDatas();
            try
            {
                if (sheet.Rows.Count == 0 && (dt.Rows.Count != sheet.Rows.Count))
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
                            };
                        }
                        Row rowA = new Row.AddRowBuilder(true, null, null, null, null).SetCells(cellsA).Build();
                        Client.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
                    }
                    Logger.LogToConsole($"Bulk Insert completed...");
                }
            }
            catch (Exception ex)
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
            Logger.LogToConsole($"{Process} complete,End Process Time: {DateTime.UtcNow}");

            var startdate = StartTime.ToString(CultureInfo.InvariantCulture);
            var enddate = DateTime.Now.ToString(CultureInfo.InvariantCulture);
            var notes = $"{Process} complete. rows imported: {RowsLinked}";
            // Logger.LogJobRun(startdate, enddate, notes, false);

        }
    }
}
