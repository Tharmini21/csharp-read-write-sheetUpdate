using csharp_read_write_sheet.Helpers;
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
        public const string CONFIGURATION_KEY_COLUMN = "Key";
        public const string CONFIGURATION_VALUE1_COLUMN = "EmployeeId";
        public const string CONFIGURATION_VALUE2_COLUMN = "FirstName";
        public const string CONFIGURATION_VALUE3_COLUMN = "LastName";
        public const string CONFIGURATION_VALUE4_COLUMN = "Email";
        public const string CONFIGURATION_VALUE5_COLUMN = "Address";

        private Sheet ConfigSheet;

        public ConfigManager(Sheet configSheet)
        {
            this.ConfigSheet = configSheet;
        }

        public ConfigList GetConfigList(string key)
        {
            var rows = ConfigSheet.Rows;
            var row = rows.Where(x => x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value != null)
                .FirstOrDefault(x =>
                    x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value.Equals(key) ?? throw new Exception($"No configuration found for {key}. Please check configuration."));
            //throw new (
            //    $"No configuration found for {key}. Please check configuration."));

            return new ConfigList()
            {
                Key = key,
                Values = rows.Where(r => r.ParentId == row?.Id)
                    .Select(r => r.GetValueForColumnAsString(ConfigSheet, CONFIGURATION_VALUE1_COLUMN)).ToList()
            };
        }

        private ConfigDictionary GetConfigDictionary(string key)
        {
            var rows = ConfigSheet.Rows;
            var row = rows.Where(x => x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value != null)
                .FirstOrDefault(x =>
                    x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value.Equals(key) ?? throw new Exception($"No configuration found for {key}. Please check configuration."));
            //throw new ConfigurationErrorsException(
            //    $"No configuration found for {key}. Please check configuration."));

            return new ConfigDictionary()
            {
                Key = key,
                Values = rows.Where(r => r.ParentId == row?.Id)
                    .Select(r => new
                    {
                        key = r.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE1_COLUMN)?.Value,
                        value = r.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE2_COLUMN)?.Value
                    }).ToDictionary(kvp => kvp.key, kvp => kvp.value)
            };
        }

        public ConfigItem GetConfigItem(string key)
        {
            var rows = ConfigSheet.Rows;
            var row = rows.Where(x => x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value != null)
                .FirstOrDefault(x =>
                    x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value?.Equals(key) ?? throw new Exception($"No configuration found for {key}. Please check configuration."));
                    //throw new ConfigurationErrorsException(
                    //    $"No configuration found for {key}. Please check configuration."));

            return new ConfigItem()
            {
                Key = key,
                Value1 = row.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE1_COLUMN)?.Value,
                Value2 = row.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE2_COLUMN)?.Value,
                Value3 = row.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE3_COLUMN)?.Value,
                Value4 = row.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE4_COLUMN)?.Value,
                Value5 = row.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE5_COLUMN)?.Value,
            };
        }
    }
}
