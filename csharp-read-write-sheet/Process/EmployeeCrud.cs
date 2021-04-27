using csharp_read_write_sheet.Configuration;
using csharp_read_write_sheet.Employee;
using csharp_read_write_sheet.Helpers;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csharp_read_write_sheet
{
    public class EmployeeCrud:SheetConfiguration
    {
        static string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
        public EmployeeCrud()
        {
            try
            {
               // var settings = (NameValueCollection)ConfigurationManager.GetSection(Section);
                ConfigSheetId = Convert.ToInt64(ConfigSheetId);
                ConfigSheet = Client.GetSheet(ConfigSheetId);
                ConfigManager = new ConfigManager(ConfigSheet);
                //this.InitLogs(this);
            }
            catch (Exception e)
            {
                throw new ApplicationException($"Unable to get configuration sheet due to exception: {e.Message}");
            }
        }

        static Dictionary<string, long> columnMap = new Dictionary<string, long>();
        static Dictionary<string, long> columnMapPMO = new Dictionary<string, long>();
        int Insertupdateddelete = 0;

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
                catch (Exception ex)
                {
                    Console.WriteLine("Error: {0}", ex.ToString());
                }

            }
            return dt;
        }
        IEnumerable<EmployeeModel> employeeList = FetchEmployeeDatas().AsEnumerable().Select(row => new EmployeeModel
        {
            EmployeeId = row.Field<int>("EmployeeId"),
            FirstName = row.Field<string>("FirstName"),
            LastName = row.Field<string>("LastName"),
            Email = row.Field<string>("Email"),
            Address = row.Field<string>("Address")
        }).ToList();
    }
}
