using csharp_read_write_sheet.Employee;
using csharp_read_write_sheet.Helper;
using csharp_read_write_sheet.Helpers;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using csharp_read_write_sheet.Configuration;

namespace csharp_read_write_sheet
{
    public class DalClass: SheetConfiguration
    {
        public DalClass()
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
        static Dictionary<string, long> columnMap = new Dictionary<string, long>();
        static Dictionary<string, long> columnMapPMO = new Dictionary<string, long>();
        static string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
        //string cs = @"Data Source=ILD-CHN-LAP-024\SQLEXPRESS;Initial Catalog=EmployeeDB;Persist Security Info=True;User ID=sa;Password=Test@123";
        public static DataTable FetchEmployeeDatas()
        {
            DataTable dt = new DataTable();
           // string strConString = @"Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("Select * from Employee", con);
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd.CommandText, con);
                    dataAdapter.Fill(dt);
                }
                catch(Exception ex)
                {
                    Console.WriteLine("Error: {0}", ex.ToString());
                }
               
            }
            return dt;
        }
        IEnumerable<EmployeeModel> employeeList= FetchEmployeeDatas().AsEnumerable().Select(row => new EmployeeModel
        {
            EmployeeId = row.Field<int>("EmployeeId"),
            FirstName = row.Field<string>("FirstName"),
            LastName = row.Field<string>("LastName"),
            Email = row.Field<string>("Email"),
            Address = row.Field<string>("Address")
        }).ToList();
        public void UpdateEmployeeDatas()
        {
            var sheet = Client.GetSheet(ConfigSheetId);
            var rows = new List<Row>();
            foreach (var emp in employeeList)
            {
                var accountRow = FindAccountRow(sheet, emp);
                var FirstNameColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE2_COLUMN, false)?.Id;
                var LastNameColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE3_COLUMN, false)?.Id;
                var EmailColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE4_COLUMN, false)?.Id;
                var AddressColumnId = SheetExtension.GetColumnByTitle(sheet, ConfigManager.CONFIGURATION_VALUE5_COLUMN, false)?.Id;
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
            Console.WriteLine("Done...");
            Console.ReadLine();
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
                var sheet = Client.GetSheet(ConfigSheetId);
                var sourceEmployeeIdList = employeeList.Select(x => x.EmployeeId).ToList();
                foreach (var row in sheet.Rows)
                {
                    int? targetEmployeeId = Convert.ToInt32(row.GetValueForColumnAsString(sheet,ConfigManager.CONFIGURATION_VALUE1_COLUMN));
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
                Console.WriteLine("Done...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Failed to Delet Row...");
                throw ex;
            }
        }
        //public void BulkInsertDbDataToSmartSheet()
        //{
        //    var sheet = Client.GetSheet(ConfigSheetId);
        //   // DataTable dt = new DataTable();
        //   // FetchEmployeeDatas().
        //    if (dt.Rows.Count != sheet.Rows.Count)
        //    {
        //        Cell[] cellsA = null;
        //        for (int i = dt.Rows.Count - 1; i >= 0; i--)
        //        {
        //            for (int j = 0; j < sheet.Columns.Count; j++)
        //            {
        //                cellsA = new Cell[]
        //                {
        //                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE1_COLUMN],dt.Rows[i][0]).Build(),
        //                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE2_COLUMN],dt.Rows[i][1]).Build(),
        //                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE3_COLUMN],dt.Rows[i][2]).Build(),
        //                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE4_COLUMN],dt.Rows[i][3]).Build(),
        //                            new Cell.AddCellBuilder(columnMap[ConfigManager.CONFIGURATION_VALUE5_COLUMN],dt.Rows[i][4]).Build(),
        //                };
        //            }
        //            Row rowA = new Row.AddRowBuilder(true, null, null, null, null).SetCells(cellsA).Build();
        //            Client.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
        //        }
        //        Console.WriteLine("Done...");
        //        Console.ReadLine();
        //    }
        //}
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
            throw new ApplicationException($"Cannot find account row for {employee} in sheet {sheet.Id}");
        }
    }
}
