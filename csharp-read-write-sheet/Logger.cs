using Smartsheet.Api;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using csharp_read_write_sheet.Helper;
using csharp_read_write_sheet.Helpers;
using System.Globalization;
using System.IO;

namespace csharp_read_write_sheet
{
    
    public static class Logger
    {
        private static readonly string AssemblyPath =
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        public static Sheet ErrorSheet;
        public static Sheet RunLogSheet;
        private static SmartsheetClient Client;
        private static string Process;
        private static readonly object Locker = new object();
        public static void LogException(Exception e, string message = "")
        {
            var rows = new List<Row>()
            {
                new Row
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            Value = e?.Message ?? string.Empty,
                            ColumnId = ErrorSheet.GetColumnByTitle("Exception")?.Id
                        },
                        new Cell()
                        {
                            Value = e?.InnerException?.Message ?? string.Empty,
                            ColumnId = ErrorSheet.GetColumnByTitle("Inner Exception")?.Id
                        },
                        new Cell()
                        {
                            Value = message,
                            ColumnId = ErrorSheet.GetColumnByTitle("Message")?.Id
                        },
                        new Cell()
                        {
                            Value = e?.Source ?? string.Empty,
                            ColumnId = ErrorSheet.GetColumnByTitle("Source")?.Id
                        },
                        new Cell()
                        {
                            Value = e?.StackTrace ?? string.Empty,
                            ColumnId = ErrorSheet.GetColumnByTitle("Stack Trace")?.Id
                        }
                    }
                }
            };

            if (ErrorSheet?.TotalRowCount >= 10)
            {
                var rowsToDelete = ErrorSheet.Rows.Take(10).ToList().Select(x => (long)x.Id).ToList();
                Client.SheetResources.RowResources.DeleteRows((long)ErrorSheet.Id, rowsToDelete, true);
            }

            Client.SheetResources.RowResources.AddRows((long)ErrorSheet.Id, rows);
        }

        public static void LogJobRun(string startTime, string finishTime, string notes, bool failed)
        {
            var rowToAdd = new List<Row>
            {
                new Row
                {
                    Cells = new List<Cell>()
                    {
                        new Cell()
                        {
                            ColumnId = RunLogSheet.GetColumnByTitle("Job Start Time").Id,
                            Value = startTime
                        },
                        new Cell()
                        {
                            ColumnId = RunLogSheet.GetColumnByTitle("Job Finish Time").Id,
                            Value = finishTime
                        },
                        new Cell()
                        {
                            ColumnId = RunLogSheet.GetColumnByTitle("Notes").Id,
                            Value = notes
                        },
                        new Cell()
                        {
                            ColumnId = RunLogSheet.GetColumnByTitle("Failed").Id,
                            Value = failed
                        }
                    }
                }
            };
            using (var stream = new FileStream(GetRunLogFile(), FileMode.Open, FileAccess.Read))
            {
                if (RunLogSheet.TotalRowCount >= 3500)
                {
                    var rowsToDelete = RunLogSheet.Rows.Take(100).ToList().Select(x => (long)x.Id).ToList();
                    Client.SheetResources.RowResources.DeleteRows((long)RunLogSheet.Id, rowsToDelete, true);
                }

                var runLogResult = Client.SheetResources.RowResources.AddRows((long)RunLogSheet.Id, rowToAdd);

                var rowId = runLogResult.LastOrDefault()?.Id;

                if (rowId == null)
                {
                    throw new ApplicationException("Run log row ID null, unable to upload log attachment");
                }

                Client.SheetResources.RowResources.AttachmentResources.AttachFile((long)RunLogSheet.Id, (long)rowId,
                    GetRunLogFile(), "text");
            }
        }

        public static void InitErrorLog(long templateId, long errorFolderId)
        {
            var name = DateTime.Now.ToString("MMMM yyyy") + " - Error Log";

            if(errorFolderId!=0)
            {
                var folder = Client.FolderResources.GetFolder(errorFolderId, null);
                if (folder.Sheets != null && folder.Sheets.Any(x => x.Name.Contains(name)))
                {
                    var errorSheetId = folder.Sheets.FirstOrDefault(x => x.Name.Contains(name))?.Id;
                    ErrorSheet =
                        Client.SheetResources.GetSheet((long)errorSheetId, null, null, null, null, null, null, null);
                }
            }
            else
            {
                var containerDestination = new ContainerDestination
                {
                    NewName = name,
                    DestinationId = errorFolderId,
                    DestinationType = DestinationType.FOLDER,
                };

                var errorSheetId =
                    Client.SheetResources.CopySheet(templateId, containerDestination, new[] { SheetCopyInclusion.DATA })?.Id;
                ErrorSheet = Client.SheetResources.GetSheet((long)errorSheetId, null, null, null, null, null, null, null);
            }
        }

        public static void InitRunLog(long templateId, long runFolderId)
        {
            var name = DateTime.Now.ToString("MMMM yyyy") + " - Run Log";

            if(runFolderId!=0)
            {
                var folder = Client.FolderResources.GetFolder(runFolderId, new List<FolderInclusion>());
                if (folder.Sheets != null && folder.Sheets.Any(x => x.Name.Contains(name)))
                {
                    var runSheetId = folder.Sheets.FirstOrDefault(x => x.Name.Contains(name))?.Id;
                    RunLogSheet =
                        Client.SheetResources.GetSheet((long)runSheetId, null, null, null, null, null, null, null);
                }
            }
            else
            {
                var containerDestination = new ContainerDestination
                {
                    NewName = name,
                    DestinationId = runFolderId,
                    DestinationType = DestinationType.FOLDER,
                };

                var runSheetId =
                    Client.SheetResources.CopySheet(templateId, containerDestination, new[] { SheetCopyInclusion.DATA })?.Id;
                RunLogSheet = Client.SheetResources.GetSheet((long)runSheetId, null, null, null, null, null, null, null);
            }
        }
        public static void InitLogging(SmartsheetClient client, SheetConfiguration automation)
        {
            Client = client;
            Process = automation.GetType().Name;
        }

        public static void LogToConsole(string message)
        {
            Console.WriteLine($"{DateTime.Now.ToString(CultureInfo.CurrentCulture)}: {message}");
            LogWrite(message);
        }

        public static void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write($"\r\n {DateTime.Now.ToString(CultureInfo.CurrentCulture)} : ");
                txtWriter.WriteLine("  :{0}", " " + logMessage);
            }
            catch (Exception ex)
            {
                LogWrite(ex.StackTrace);
            }
        }
        private static void LogWrite(string logMessage)
        {
            lock (Locker)
            {
                try
                {
                    using (var writer = File.AppendText(GetRunLogFile()))
                    {
                        Log(logMessage, writer);
                    }
                }
                catch (Exception ex)
                {
                    Console.Write(ex.Message);
                }
            }
        }
        private static string GetRunLogFile()
        {
            return Path.Combine(AssemblyPath, $"{Process} - log.txt");
        }

        public static void ClearLogFileContents()
        {
            File.WriteAllText(GetRunLogFile(), string.Empty);
        }

    }
}
