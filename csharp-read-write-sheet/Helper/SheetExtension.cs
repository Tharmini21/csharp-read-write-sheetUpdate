using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csharp_read_write_sheet.Helper
{
    public class SheetExtension
    {
        public static Column GetColumnByTitle(Sheet sheet, string columnTitle, bool caseSensitive = false)
        {
            var column = sheet.Columns.FirstOrDefault(c => String.Equals(c.Title, columnTitle, caseSensitive ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase));
            if (column == null)
                throw new ArgumentException($"The sheet '{sheet.Name}' does not contain a column with the title '{columnTitle}'");

            return column;
        }
    }
}
