using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
// Add nuget reference to smartsheet-csharp-sdk (https://www.nuget.org/packages/smartsheet-csharp-sdk/)
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using MySql.Data.MySqlClient;
using System.Data;
using System.Data.SqlClient;
using csharp_read_write_sheet.Configuration;
using csharp_read_write_sheet.Employee;
using csharp_read_write_sheet.Helper;
using csharp_read_write_sheet.Helpers;


namespace csharp_read_write_sheet
{
    class Program
    {
        public static async Task Main(string[] args)
        {
           
          var employeeprocess = new EmployeeCrud();
          Logger.ClearLogFileContents();
          employeeprocess.CreateNewEmployeeDatas();
          await employeeprocess.Run();
          return;
                  
        }
    }
}
