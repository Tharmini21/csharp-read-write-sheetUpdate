using csharp_read_write_sheet.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;

public class Config
{
    public Config()
    {
        this.ConfigItems = new List<ConfigItem>();
        this.ConfigLists = new List<ConfigList>();
        this.ConfigDictionaries = new List<ConfigDictionary>();
    }
    protected internal List<ConfigItem> ConfigItems { get; private set; }
    protected internal List<ConfigList> ConfigLists { get; private set; }
    protected internal List<ConfigDictionary> ConfigDictionaries { get; private set; }

    public ConfigItem GetConfigItem(string key)
    {
        return this.ConfigItems.FirstOrDefault(x => x.Key == key);
    }

    public ConfigList GetConfigList(string key)
    {
        return this.ConfigLists.FirstOrDefault(x => x.Key == key);
    }

    public ConfigDictionary GetConfigDictionary(string key)
    {
        return this.ConfigDictionaries.FirstOrDefault(x => x.Key == key);
    }
}
