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
    public class Logger
    {
        private static readonly string AssemblyPath =
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

        private static Sheet ErrorSheet;
        private static Sheet RunLogSheet;
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
