using Smartsheet.Api;
// Add nuget reference to smartsheet-csharp-sdk (https://www.nuget.org/packages/smartsheet-csharp-sdk/)
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System;
using System.Configuration;

namespace csharp_read_write_sheet
{
    public class SheetConfiguration
    {
        
        
       public static readonly string smartsheetAPIToken = ConfigurationManager.AppSettings["AccessToken"];


        string key= ConfigurationManager.AppSettings["AccessToken"];
        Token token = new Token();
        token 
                SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
        //Sheet sheet = smartsheet.SheetResources.ImportXlsSheet("../../../TBL_Employee.xlsx", null, 0, null);
        //sheet = smartsheet.SheetResources.GetSheet(sheet.Id.Value, null, null, null, null, null, null, null);
        long SheetId = Convert.ToInt64(ConfigurationManager.AppSettings["SheetId"]);
        Sheet sheet = smartsheet.SheetResources.GetSheet(SheetId, null, null, null, null, null, null, null);
    }
}
