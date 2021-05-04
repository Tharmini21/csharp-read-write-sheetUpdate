using System.Collections.Generic;
using System.Linq;
using csharp_read_write_sheet.Configuration;
using Smartsheet.Api.Models;

namespace csharp_read_write_sheet.Helpers
{
    public static class RowExtensions
    {
        public static Cell GetCellForColumn(this Row row, Sheet sheet, string columnName)
        {
           var cell = row?.Cells.FirstOrDefault(c => c.ColumnId == sheet.Columns.FirstOrDefault(col => col.Title == columnName)?.Id);
           return cell;
        }
        public static Cell GetCellForColumn(this Row row, Column column)
        {
            return row.Cells[column.Index.Value];
        }
        public static string GetValueForColumnAsString(this Row row, Sheet sheet, string columnName)
        {
            var cell = row.GetCellForColumn(sheet, columnName);
            return cell?.Value == null ? "" : cell?.Value.ToString();
        }
    }
}