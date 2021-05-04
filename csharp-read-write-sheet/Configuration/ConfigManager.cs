using csharp_read_write_sheet.Helpers;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
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
        public const string CONFIGURATION_VALUE6_COLUMN = "Job Start Time";
        public const string CONFIGURATION_VALUE7_COLUMN = "Job Finish Time";
        public const string CONFIGURATION_VALUE8_COLUMN = "Notes";
        public const string CONFIGURATION_VALUE9_COLUMN = "Failed";
        private const string CONFIG_VALUE1_COLUMN = "Value1";
        private const string CONFIG_VALUE2_COLUMN = "Value2";
        private const string CONFIG_VALUE3_COLUMN = "Value3";
        private const string CONFIG_VALUE4_COLUMN = "Value4";
        private const string CONFIG_VALUE5_COLUMN = "Value5";

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
                    x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value.Equals(key) ??
                    throw new ConfigurationErrorsException(
                        $"No configuration found for {key}. Please check configuration."));

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
                    x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value.Equals(key) ??
                    throw new ConfigurationErrorsException(
                        $"No configuration found for {key}. Please check configuration."));

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
                    x.GetCellForColumn(ConfigSheet, CONFIGURATION_KEY_COLUMN)?.Value?.Equals(key) ??
                    throw new ConfigurationErrorsException(
                        $"No configuration found for {key}. Please check configuration."));

            //var row = rows.Where(x => x.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE1_COLUMN)?.Value != null)
            //    .FirstOrDefault(x =>
            //        x.GetCellForColumn(ConfigSheet, CONFIGURATION_VALUE1_COLUMN)?.Value?.Equals(key) ??
            //        throw new ConfigurationErrorsException(
            //            $"No configuration found for {key}. Please check configuration."));

            return new ConfigItem()
            {
                Key = CONFIGURATION_VALUE1_COLUMN,
                Value1 = row.GetCellForColumn(ConfigSheet, CONFIG_VALUE1_COLUMN)?.Value,
                Value2 = row.GetCellForColumn(ConfigSheet, CONFIG_VALUE2_COLUMN)?.Value,
                Value3 = row.GetCellForColumn(ConfigSheet, CONFIG_VALUE3_COLUMN)?.Value,
                Value4 = row.GetCellForColumn(ConfigSheet, CONFIG_VALUE4_COLUMN)?.Value,
                Value5 = row.GetCellForColumn(ConfigSheet, CONFIG_VALUE5_COLUMN)?.Value
            };
        }
    }
}
