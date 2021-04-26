using csharp_read_write_sheet.Employee;
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
    public class DalClass
    {
        static Dictionary<string, long> columnMap = new Dictionary<string, long>();
        static Dictionary<string, long> columnMapPMO = new Dictionary<string, long>();
        static string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
        //string cs = @"Data Source=ILD-CHN-LAP-024\SQLEXPRESS;Initial Catalog=EmployeeDB;Persist Security Info=True;User ID=sa;Password=Test@123";
        int Insertupdateddelete = 0;
        List<EmployeeModel> listofemployees = null;
        public DataTable FetchEmployeeDatas()
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
                catch(Exception ex)
                {
                    Console.WriteLine("Error: {0}", ex.ToString());
                }
               
            }
            return dt;
        }  
        //foreach (Column column in sheet.Columns)
        //            columnMap.Add(column.Title, (long) column.Id);
        //        IList<EmployeeModel> items = dt.AsEnumerable().Select(row => new EmployeeModel
        //        {
        //            EmployeeId = row.Field<int>("EmployeeId"),
        //            FirstName = row.Field<string>("FirstName"),
        //            LastName = row.Field<string>("LastName"),
        //            Email = row.Field<string>("Email"),
        //            Address = row.Field<string>("Address")
        //        }).ToList();
        //IEnumerable<EmployeeModel> employeeList = items;
        //listofemployees = new List<EmployeeModel>();
        //        listofemployees = (from DataRow dr in dt.Rows
        //                           select new EmployeeModel()
        //{
        //    EmployeeId = Convert.ToInt32(dr[CONFIGURATION_VALUE1_COLUMN]),
        //                               FirstName = dr[CONFIGURATION_VALUE2_COLUMN].ToString(),
        //                               LastName = dr[CONFIGURATION_VALUE3_COLUMN].ToString(),
        //                               Email = dr[CONFIGURATION_VALUE4_COLUMN].ToString(),
        //                               Address = dr[CONFIGURATION_VALUE5_COLUMN].ToString()
        //                           }).ToList();
    }
}
