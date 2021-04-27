using System.Collections.Generic;
using System.Linq;
using Smartsheet.Api;
using Smartsheet.Api.Models;

namespace csharp_read_write_sheet.Helpers
{
    public static class SmartsheetClientExtensions
    {
        public static Sheet GetSheet(this SmartsheetClient client, long? sheetId)
        {
            return client.SheetResources.GetSheet(sheetId.Value, null, null, null, null, null, null, null);
        }
    }
}
