using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csharp_read_write_sheet.Configuration
{
    public class ConfigManager
    {
        private const string CONFIGURATION_KEY_COLUMN = "Key";
        private const string CONFIGURATION_VALUE1_COLUMN = "EmployeeId";
        private const string CONFIGURATION_VALUE2_COLUMN = "FirstName";
        private const string CONFIGURATION_VALUE3_COLUMN = "LastName";
        private const string CONFIGURATION_VALUE4_COLUMN = "Email";
        private const string CONFIGURATION_VALUE5_COLUMN = "Address";

        private Sheet ConfigSheet;

        public ConfigManager(Sheet configSheet)
        {
            this.ConfigSheet = configSheet;
        }
    }
}
