using Smartsheet.Api;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CostcoAutomation.Helpers;

namespace csharp_read_write_sheet
{
    public class Logger
    {
        private static Sheet ErrorSheet;
        private static Sheet RunLogSheet;
        private static SmartsheetClient Client;

        private static string Process;

        private static readonly object Locker = new object();

        //public static void LogException(Exception e, string message = "")
        //{
        //    var rows = new List<Row>()
        //    {
        //        new Row
        //        {
        //            Cells = new List<Cell>()
        //            {
        //                new Cell()
        //                {
        //                    Value = e?.Message ?? string.Empty,
        //                    ColumnId = ErrorSheet.GetColumnByTitle("Exception")?.Id
        //                },
        //                new Cell()
        //                {
        //                    Value = e?.InnerException?.Message ?? string.Empty,
        //                    ColumnId = ErrorSheet.GetColumnByTitle("Inner Exception")?.Id
        //                },
        //                new Cell()
        //                {
        //                    Value = message,
        //                    ColumnId = ErrorSheet.GetColumnByTitle("Message")?.Id
        //                },
        //                new Cell()
        //                {
        //                    Value = e?.Source ?? string.Empty,
        //                    ColumnId = ErrorSheet.GetColumnByTitle("Source")?.Id
        //                },
        //                new Cell()
        //                {
        //                    Value = e?.StackTrace ?? string.Empty,
        //                    ColumnId = ErrorSheet.GetColumnByTitle("Stack Trace")?.Id
        //                }
        //            }
        //        }
        //    };

        //    if (ErrorSheet?.TotalRowCount >= 3500)
        //    {
        //        var rowsToDelete = ErrorSheet.Rows.Take(100).ToList().Select(x => (long)x.Id).ToList();
        //        Client.SheetResources.RowResources.DeleteRows((long)ErrorSheet.Id, rowsToDelete, true);
        //    }

        //    Client.SheetResources.RowResources.AddRows((long)ErrorSheet.Id, rows);
        //}

        //public static void LogJobRun(string startTime, string finishTime, string notes, bool failed)
        //{
        //    var rowToAdd = new List<Row>
        //    {
        //        new Row
        //        {
        //            Cells = new List<Cell>()
        //            {
        //                new Cell()
        //                {
        //                    ColumnId = RunLogSheet.GetColumnByTitle("Job Start Time").Id,
        //                    Value = startTime
        //                },
        //                new Cell()
        //                {
        //                    ColumnId = RunLogSheet.GetColumnByTitle("Job Finish Time").Id,
        //                    Value = finishTime
        //                },
        //                new Cell()
        //                {
        //                    ColumnId = RunLogSheet.GetColumnByTitle("Notes").Id,
        //                    Value = notes
        //                },
        //                new Cell()
        //                {
        //                    ColumnId = RunLogSheet.GetColumnByTitle("Failed").Id,
        //                    Value = failed
        //                }
        //            }
        //        }
        //    };

        //    using (var stream = new FileStream(GetRunLogFile(), FileMode.Open, FileAccess.Read))
        //    {
        //        if (RunLogSheet.TotalRowCount >= 3500)
        //        {
        //            var rowsToDelete = RunLogSheet.Rows.Take(100).ToList().Select(x => (long)x.Id).ToList();
        //            Client.SheetResources.RowResources.DeleteRows((long)RunLogSheet.Id, rowsToDelete, true);
        //        }

        //        var runLogResult = Client.SheetResources.RowResources.AddRows((long)RunLogSheet.Id, rowToAdd);

        //        var rowId = runLogResult.LastOrDefault()?.Id;

        //        if (rowId == null)
        //        {
        //            throw new ApplicationException("Run log row ID null, unable to upload log attachment");
        //        }

        //        Client.SheetResources.RowResources.AttachmentResources.AttachFile((long)RunLogSheet.Id, (long)rowId,
        //            GetRunLogFile(), "text");
        //    }
        //}
    }
}
