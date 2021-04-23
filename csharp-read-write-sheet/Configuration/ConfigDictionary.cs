using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csharp_read_write_sheet.Configuration
{
    public class ConfigDictionary
    {
        public string Key { get; set; }
        public Dictionary<dynamic, dynamic> Values { get; set; }
    }
}
