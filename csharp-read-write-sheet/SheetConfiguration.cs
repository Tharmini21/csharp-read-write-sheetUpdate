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
        private const string ERROR_LOG_FOLDER = "Error Log Folder ID";
        private const string ERROR_LOG_TEMPLATE = "Error Log Template ID";

        private const string RUN_LOG_FOLDER = "Run Log Folder ID";
        private const string RUN_LOG_TEMPLATE = "Run Log Template ID";
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
            //SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            this.ConfigSheetId = Convert.ToInt64(ConfigurationManager.AppSettings["SheetId"]);
            this.AuthToken = ConfigurationManager.AppSettings[AUTH_TOKEN];
            Client = new SmartsheetBuilder().SetAccessToken(this.AuthToken).Build();
            //Sheet sheet = Client.SheetResources.ImportXlsSheet("../../../TBL_Employee_Updated.xlsx", null, 0, null);
            //sheet = Client.SheetResources.GetSheet(sheet.Id.Value, null, null, null, null, null, null, null);
            Client.SheetResources.GetSheet(ConfigSheetId, null, null, null, null, null, null, null);
            StartTime = DateTime.Now;
            Clients = new Dictionary<SmartsheetClient, bool>();
        }
        protected void InitLogs(SheetConfiguration automation)
        {
            //set client and log
            Logger.InitLogging(Client, automation);

            //set error logs
            //var errorFolderId = Convert.ToInt64(this.ConfigManager.GetConfigItem(ERROR_LOG_FOLDER)?.Value1);
            //var errorTemplateId = Convert.ToInt64(this.ConfigManager.GetConfigItem(ERROR_LOG_TEMPLATE)?.Value1);
            //Logger.InitErrorLog(errorTemplateId, errorFolderId);

            //set run logs
            //var runFolderId = Convert.ToInt64(this.ConfigManager.GetConfigItem(RUN_LOG_FOLDER)?.Value1);
            //var runTemplateId = Convert.ToInt64(this.ConfigManager.GetConfigItem(RUN_LOG_TEMPLATE)?.Value1);
            //Logger.InitRunLog(runTemplateId, runFolderId);


        }
    }
}
