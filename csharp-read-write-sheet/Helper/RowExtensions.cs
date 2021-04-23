using System.Collections.Generic;
using System.Linq;
using Smartsheet.Api.Models;

namespace CostcoAutomation.Helpers
{
    public static class RowExtensions
    {
       
        public static Cell GetCellForColumn(this Row row, Sheet sheet, string columnName)
        {
            var cell = row.Cells.FirstOrDefault(c => c.ColumnId == sheet.Columns.FirstOrDefault(col => col.Title == columnName)?.Id);
            return cell;
        }
		
        public static Cell GetCellForColumn(this Row row, Column column)
        {
            return row.Cells[column.Index.Value];
        }
		
        public static object GetValueForColumn(this Row row, Sheet sheet, string columnName)
        {
            var cell = row.GetCellForColumn(sheet, columnName);
            return cell?.Value;
        }
		
        public static object GetValueForColumn(this Row row, Column column)
        {
            return row.GetCellForColumn(column).Value;
        }
        
        public static string GetValueForColumnAsString(this Row row, Sheet sheet, string columnName)
        {
            var cell = row.GetCellForColumn(sheet, columnName);
            return cell?.Value == null ? "" : cell?.Value.ToString();
        }

        public static List<Row> GetChildren(this Row row, Sheet sheet)
        {
            var result = new List<Row>();

            foreach (var sheetRow in sheet.Rows)
            {
                if (sheetRow.ParentId == row.Id)
                {
                    result.Add(sheetRow);
                }
            }

            return result;
        }
        
        
    }
}