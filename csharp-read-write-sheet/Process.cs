using csharp_read_write_sheet.Employee;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csharp_read_write_sheet
{
    public class Process
    {

        static Dictionary<string, long> columnMap = new Dictionary<string, long>();
        static Dictionary<string, long> columnMapPMO = new Dictionary<string, long>();
        int Insertupdateddelete = 0;
        List<EmployeeModel> listofemployees = null;

                foreach (Column column in sheet.Columns)
                    columnMap.Add(column.Title, (long) column.Id);
                    IList<EmployeeModel> items = dt.AsEnumerable().Select(row => new EmployeeModel
                    {
                        EmployeeId = row.Field<int>("EmployeeId"),
                        FirstName = row.Field<string>("FirstName"),
                        LastName = row.Field<string>("LastName"),
                        Email = row.Field<string>("Email"),
                        Address = row.Field<string>("Address")
                    }).ToList();
        IEnumerable<EmployeeModel> employeeList = items;
        listofemployees = new List<EmployeeModel>();
                listofemployees = (from DataRow dr in dt.Rows
                               select new EmployeeModel()
        {
            EmployeeId = Convert.ToInt32(dr[CONFIGURATION_VALUE1_COLUMN]),
                                   FirstName = dr[CONFIGURATION_VALUE2_COLUMN].ToString(),
                                   LastName = dr[CONFIGURATION_VALUE3_COLUMN].ToString(),
                                   Email = dr[CONFIGURATION_VALUE4_COLUMN].ToString(),
                                   Address = dr[CONFIGURATION_VALUE5_COLUMN].ToString()
                               }).ToList();
                if (Insertupdateddelete == 1)
                {
                    Cell[] newcell = new Cell[]
                    {
                        new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE1_COLUMN],dt.Rows.Count+1).Build(),
                        new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE2_COLUMN],"CheckInsert").Build(),
                        new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE3_COLUMN],"Lst").Build(),
                        new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE4_COLUMN],"CheckInsert@gmail.com").Build(),
                        new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE5_COLUMN],"AbcStreet").Build(),
                    };
        Row rowA = new Row.AddRowBuilder(null, true, null, null, null).SetCells(newcell).Build();
        smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA
    });
                }
                else if (Insertupdateddelete == 2)
{
    var rows = new List<Row>();
    foreach (var emp in employeeList)
    {
        var accountRow = FindAccountRow(sheet, emp);
        var FirstNameColumnId = SheetExtension.GetColumnByTitle(sheet, CONFIGURATION_VALUE2_COLUMN, false)?.Id;
        var LastNameColumnId = SheetExtension.GetColumnByTitle(sheet, CONFIGURATION_VALUE3_COLUMN, false)?.Id;
        var EmailColumnId = SheetExtension.GetColumnByTitle(sheet, CONFIGURATION_VALUE4_COLUMN, false)?.Id;
        var AddressColumnId = SheetExtension.GetColumnByTitle(sheet, CONFIGURATION_VALUE5_COLUMN, false)?.Id;
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

    var rowsUpdated = smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, rows).Count;
    Console.WriteLine("Done...");
    Console.ReadLine();
}
else if (Insertupdateddelete == 3)
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

    var sourceEmployeeIdList = employeeList.Select(x => x.EmployeeId).ToList();
    foreach (var row in sheet.Rows)
    {
        int? targetEmployeeId = Convert.ToInt32(row.GetValueForColumnAsString(sheet, CONFIGURATION_VALUE1_COLUMN));
        if (targetEmployeeId != null)
        {
            if (!sourceEmployeeIdList.Contains(targetEmployeeId.Value))
            {
                var rowsUpdated = smartsheet.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { (long)row.Id }.ToList(), true).Count;
                smartsheet.SheetResources.RowResources.DeleteRows(sheet.Id.Value, new long[] { (long)row.Id }.ToList(), true);
            }
        }
        else
            break;
    }
    Console.WriteLine("Done...");
    Console.ReadLine();

}
else
{
    if (dt.Rows.Count != sheet.Rows.Count)
    {
        Cell[] cellsA = null;
        for (int i = dt.Rows.Count - 1; i >= 0; i--)
        {
            for (int j = 0; j < sheet.Columns.Count; j++)
            {
                cellsA = new Cell[]
                {
                                    new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE1_COLUMN],dt.Rows[i][0]).Build(),
                                    new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE2_COLUMN],dt.Rows[i][1]).Build(),
                                    new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE3_COLUMN],dt.Rows[i][2]).Build(),
                                    new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE4_COLUMN],dt.Rows[i][3]).Build(),
                                    new Cell.AddCellBuilder(columnMap[CONFIGURATION_VALUE5_COLUMN],dt.Rows[i][4]).Build(),
                };
            }
            Row rowA = new Row.AddRowBuilder(true, null, null, null, null).SetCells(cellsA).Build();
            smartsheet.SheetResources.RowResources.AddRows(sheet.Id.Value, new Row[] { rowA });
        }
        Console.WriteLine("Done...");
        Console.ReadLine();
    }
}
    }
}
