using csharp_read_write_sheet.Configuration;
using csharp_read_write_sheet.Employee;
using csharp_read_write_sheet.Helper;
using csharp_read_write_sheet.Helpers;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
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
        private int RowsLinked;

        static Dictionary<string, long> columnMap = new Dictionary<string, long>();
        public EmployeeCrud()
        {
            try
            {
                ConfigSheetId = Convert.ToInt64(ConfigSheetId);
                ConfigSheet = Client.GetSheet(ConfigSheetId);
                ConfigManager = new ConfigManager(ConfigSheet);
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
            LogJubRun();
            Logger.LogToConsole($"Done...");

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
            var rowsToCreate = new List<Row>();
            DataTable dt = FetchEmployeeDatas();
            var sourceEmployeeList = employeeList.Select(x => x).ToList();
            if (dt.Rows.Count != sheet.Rows.Count)
            {
                foreach (var row in sheet.Rows)
                {
                    int? targetEmployeeval = Convert.ToInt32(row.GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN));
                    if (targetEmployeeval != null)
                    {
                        if (!sourceEmployeeList.Any(a => a.EmployeeId == targetEmployeeval))
                        {
                            Cell[] newcell = new Cell[]
                            {
                                  new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE1_COLUMN],dt.Rows.Count+1).Build(),
                                  new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE2_COLUMN],"CheckInsert").Build(),
                                  new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE3_COLUMN],"Lst").Build(),
                                  new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE4_COLUMN],"CheckInsert@gmail.com").Build(),
                                  new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE5_COLUMN],"AbcStreet").Build(),
                            };
                            Row rowA = new Row.AddRowBuilder(null, true, null, null, null).SetCells(newcell).Build();
                            Client.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
                        }
                    }
                }
            }
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

        //public void CreateNewEmployeeDatas()
        //{
        //    var sheet = Client.GetSheet(ConfigSheetId);
        //    var rowsToCreate = new List<Row>();
        //    DataTable dt = FetchEmployeeDatas();

        //    foreach (var account in sheet.Rows)
        //    {
        //        int? exists = Convert.ToInt32(sheet.Rows.Any(c => c.GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN) == account.Cells[0].ColumnId));
        //        if (exists)
        //        {
        //            rowsToCreate.Add(account);
        //        }
        //    }

        //    if (dt.Rows.Count != sheet.Rows.Count)
        //    {
        //        foreach (var emp in employeeList)
        //        {
        //            var accountRow = FindAccountRow(sheet, emp);
        //        }
        //        //foreach (var row in sheet.Rows)
        //        //{
        //        //    if (row.GetValueForColumnAsString(sheet, ConfigManager.CONFIGURATION_VALUE1_COLUMN))
        //        //    {
        //        //        continue;
        //        //    }

        //        //    //rowsToCreate.Add(this.BuildRow(row, currentYear, previousYear));
        //        //}
        //        //foreach(var row in dt.Rows)
        //        //{
        //        //    foreach (var s in sheet.Rows)
        //        //    {
        //        //        if(row!=s)
        //        //        {
        //        //            Console.WriteLine("true");
        //        //        }
        //        //    }
        //        //}
        //    }
        //    // List<int> sourceEmployeeList = employeeList.Select(x => x.EmployeeId).ToList();
        //    //List<int> = targetlist.Select(x=>x.)
        //    //    //foreach (var row in sheet.Rows)
        //    //    //{
        //    //    //    //var targetEmployeeval = Convert.ToInt32(row.GetValueForColumnAsString(sheet, CONFIGURATION_VALUE1_COLUMN));
        //    //    //    var targetEmployeeval = (row.GetValueForColumnAsString(sheet, CONFIGURATION_VALUE1_COLUMN));
        //    //    //    if (targetEmployeeval != null)
        //    //    //    {
        //    //    //        if (!targetEmployeeval.Contains(sourceEmployeeList.AsEnumerable().Select(x=>x.ToString()))
        //    //    //        {
        //    //    //        }
        //    //    //    }
        //    //    //}
        //    //List<int> newStudentsList = newStudents.Select(n => n.Id).ToList();
        //    //List<int> oldStudentsList = oldStudents.Select(o => o.Id).ToList();

        //    //var missingStudents = oldStudentsList.Except(newStudentsList).ToList();
        //    //Cell[] newcell = new Cell[]
        //    //{
        //    //    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE1_COLUMN],dt.Rows.Count+1).Build(),
        //    //    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE2_COLUMN],"CheckInsert").Build(),
        //    //    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE3_COLUMN],"Lst").Build(),
        //    //    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE4_COLUMN],"CheckInsert@gmail.com").Build(),
        //    //    new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE5_COLUMN],"AbcStreet").Build(),
        //    //};
        //    //Row rowA = new Row.AddRowBuilder(null, true, null, null, null).SetCells(newcell).Build();
        //    //Client.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
        //}
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
        private void LogJubRun()
        {
            Logger.LogToConsole($"{Process} complete");

            var startdate = StartTime.ToString(CultureInfo.InvariantCulture);
            var enddate = DateTime.Now.ToString(CultureInfo.InvariantCulture);
            var notes = $"{Process} complete. rows imported: {RowsLinked}";
           Logger.LogJobRun(startdate, enddate, notes, false);
          
            // Logger.LogJobRun(startTime, endTime, notes, false);
        }
    }
}
