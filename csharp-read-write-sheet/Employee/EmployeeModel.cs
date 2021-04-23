using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csharp_read_write_sheet.Employee
{
    public class EmployeeModel
    {
        public int EmployeeId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        //public EmployeeModel(int employeeId,string firstName, string lastName, string email, string address)
        //{
        //    EmployeeId = employeeId;
        //    FirstName = firstName;
        //    LastName = lastName;
        //    Email = email;
        //    Address = address;
        //}
    }
}
