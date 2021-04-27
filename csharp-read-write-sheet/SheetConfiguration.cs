using csharp_read_write_sheet.Configuration;
using Smartsheet.Api;
// Add nuget reference to smartsheet-csharp-sdk (https://www.nuget.org/packages/smartsheet-csharp-sdk/)
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace csharp_read_write_sheet
{
    public class SheetConfiguration
    {

        private const string AUTH_TOKEN = "AccessToken";
        // public static readonly string smartsheetAPIToken = ConfigurationManager.AppSettings["AccessToken"];


        string smartsheetAPIToken = ConfigurationManager.AppSettings["AccessToken"];
        Token token = new Token();
        private string AuthToken { get; set; }
        private Dictionary<SmartsheetClient, bool> Clients;
        protected long ConfigSheetId { get; set; }
        protected SmartsheetClient Client;
        protected Sheet ConfigSheet;
        protected ConfigManager ConfigManager;
        protected Config Config;
        protected DateTime StartTime;

        private static readonly object Locker = new object();

        //protected SheetConfiguration()
        public SheetConfiguration()
        {
            token.AccessToken = smartsheetAPIToken;
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            this.ConfigSheetId = Convert.ToInt64(ConfigurationManager.AppSettings["SheetId"]);
            this.AuthToken = ConfigurationManager.AppSettings[AUTH_TOKEN];
            Client = new SmartsheetBuilder().SetAccessToken(this.AuthToken).Build();
            Client.SheetResources.GetSheet(ConfigSheetId, null, null, null, null, null, null, null);
            StartTime = DateTime.Now;
            Clients = new Dictionary<SmartsheetClient, bool>();
        }
    }
}
