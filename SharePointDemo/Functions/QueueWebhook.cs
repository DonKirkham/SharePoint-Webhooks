using Microsoft.Azure;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using SharePointDemo.Common;
using SharePointDemo.Models;
using System;
using System.Collections.Generic;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace SharePointDemo.Functions
{
    public static class QueueWebhook
    {
        [FunctionName("QueueWebhook")]
        public static void Run([QueueTrigger("dklabswebhookdemo-queue", 
            Connection = "AzureWebJobsStorage")]string queueItem, ILogger log)
        {
            ProcessNotification(log, queueItem);
        }

        public static void ProcessNotification(ILogger log, string queueItem)
        {

            NotificationModel notification = JsonConvert.DeserializeObject<NotificationModel>(queueItem);
            log.LogInformation($"Processing notification: {notification.Resource}");
            #region Get Context
            string url = string.Format($"https://{CloudConfigurationManager.GetSetting("SP_Tenant_Name")}.sharepoint.com{notification.SiteUrl}");
            OfficeDevPnP.Core.AuthenticationManager am = new AuthenticationManager();
            ClientContext cc = am.GetAppOnlyAuthenticatedContext(
                url,
                CloudConfigurationManager.GetSetting("SPApp_ClientId"),
                CloudConfigurationManager.GetSetting("SPApp_ClientSecret"));
            #endregion

            #region Grab the list for which the web hook was triggered
            List changeList = cc.Web.GetListById(new Guid(notification.Resource));
            cc.ExecuteQueryRetry();
            if (changeList == null)
            {
                // list has been deleted in between the event being fired and the event being processed
                log.LogInformation($"List \"{notification.Resource}\" no longer exists.");
                return;
            }
            #endregion

            #region Get the Last Change Token
            CloudStorageAccount storageAccount = 
                CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("AzureWebJobsStorage"));
            CloudTableClient client = storageAccount.CreateCloudTableClient();
            CloudTable table = 
                client.GetTableReference(CloudConfigurationManager.GetSetting("LastChangeTokensTableName"));
            table.CreateIfNotExists();
            TableOperation retrieveOperation = TableOperation.Retrieve<LastChangeEntity>("LastChangeToken", notification.Resource);
            TableResult query = table.Execute(retrieveOperation);

            ChangeToken lastChangeToken = null;
            if (query.Result != null)
            {
                lastChangeToken = new ChangeToken() { StringValue = ((LastChangeEntity)query.Result).LastChangeToken };
            }
            if (lastChangeToken == null)
            {
                lastChangeToken = new ChangeToken() { StringValue = $"1;3;{notification.Resource};{DateTime.Now.AddMinutes(-60).ToUniversalTime().Ticks.ToString()};-1" };
            }
            #endregion

            #region Grab Changes since Last Change Token and Process
            ChangeQuery changeQuery = new ChangeQuery(false, true)
            {
                Item = true,
                FetchLimit = 1000 // Max value is 2000, default = 1000
            };

            // Start pulling down the changes
            bool allChangesRead = false;
            do
            {

                //Assign the change token to the query...this determines from what point in time we'll receive changes
                changeQuery.ChangeTokenStart = lastChangeToken;
                ChangeCollection changes = changeList.GetChanges(changeQuery);
                cc.Load(changes);
                cc.ExecuteQueryRetry();

                log.LogInformation($"Changes found: {changes.Count}");
                if (changes.Count > 0)
                {
                    try
                    {
                        List<int> handledListItems = new List<int>();
                        foreach (Change change in changes)
                        {
                            lastChangeToken = change.ChangeToken;
                            if (change is ChangeItem)
                            {
                                var listItemId = (change as ChangeItem).ItemId;
                                log.LogInformation($"-Item that changed: ItemId: {listItemId}");
                                if (handledListItems.Contains(listItemId))
                                {
                                    log.LogInformation("-ListItem already handled in this batch.");
                                    continue;
                                }
                                ListItem listItem = changeList.GetItemById((change as ChangeItem).ItemId);
                                try
                                {
                                    cc.Load(listItem);
                                    cc.ExecuteQueryRetry();
                                }
                                catch (Exception ex)
                                {
                                    log.LogInformation($"ERROR: {ex.Message}");
                                    continue;
                                }

                                //DO SOMETHING WITH LIST ITEM
                                //DoWork(log, cc, listItem);

                                UpdateLastChangeToken(notification, lastChangeToken, table);
                                RecordChangeInWebhookHistory(cc, changeList, change);
                                handledListItems.Add(listItem.Id);
                            }
                        }
                        if (changes.Count < changeQuery.FetchLimit)
                        {
                            allChangesRead = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"ERROR: {ex.Message}");
                    }
                }
                else
                {
                    allChangesRead = true;
                }
                // Are we done?
            } while (allChangesRead == false);
            #endregion

            #region "Update" the web hook expiration date when needed
            // Optionally add logic to "update" the expirationdatetime of the web hook
            // If the web hook is about to expire within the coming 5 days then prolong it
            try
            {
                if (notification.ExpirationDateTime.AddDays(-5) < DateTime.Now)
                {
                    DateTime newDate = DateTime.Now.AddMonths(3);
                    log.LogInformation($"Updating the Webhook expiration date to {newDate}");
                    WebHookManager webHookManager = new WebHookManager();
                    Task<bool> updateResult = Task.WhenAny(
                        webHookManager.UpdateListWebHookAsync(
                            url,
                            changeList.Id.ToString(),
                            notification.SubscriptionId,
                            CloudConfigurationManager.GetSetting("AzureWebJobsStorage"),
                            newDate,
                            cc.GetAccessToken())
                        ).Result;

                    if (updateResult.Result == false)
                    {
                        throw new Exception(string.Format("The expiration date of web hook {0} with endpoint {1} could not be updated", notification.SubscriptionId, CloudConfigurationManager.GetSetting("WebHookEndPoint")));
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogInformation($"ERROR: {ex.Message}");
                //throw new Exception($"ERROR: {ex.Message}");
            }
            #endregion

            cc.Dispose();
            log.LogInformation("Processing complete.");
        }

        private static void DoWork(ILogger log, ClientContext cc, ListItem listItem)
        {
            string[] libraries = { "Petitions", "Judgments" };
            foreach (string library in libraries)
            {
                //try
                //{
                log.LogInformation($"Updating library: {library}");
                Web web = cc.Web;
                cc.Load(web);
                cc.ExecuteQueryRetry();
                List doclib1 = web.GetListByTitle(library);
                cc.Load(doclib1);
                cc.ExecuteQueryRetry();
                string template = doclib1.DocumentTemplateUrl;
                cc.ExecuteQueryRetry();
                const string pattern = "^(\\W|_vti|_)|[^\\w]|(files|file|Dateien|fichiers|bestanden|archivos|filer|tiedostot|pliki|soubory|elemei|ficheiros|arquivos|dosyalar|datoteke|fitxers|failid|fails|bylos|fajlovi|fitxategiak)$"; const string str = "Doc with & opr.docx";
                var validatedFileName = Regex.Replace(listItem.FieldValues["Title"].ToString(), pattern, string.Empty);
                try
                {
                    File file = cc.Web.GetFileByServerRelativeUrl($"{doclib1.RootFolder.ServerRelativeUrl}/{validatedFileName}.docx");
                    cc.Load(file, f => f.Exists, f => f.ListItemAllFields);
                    cc.ExecuteQueryRetry();
                    if (file.Exists)
                    {
                        log.LogInformation($"-Document being replaced.");
                        file.DeleteObject();
                        cc.ExecuteQueryRetry();
                    }
                }
                catch
                {

                }
                log.LogInformation($"-Creating new document...");
                ListItem file2 = doclib1.CreateDocumentFromTemplate($"{validatedFileName}.docx", doclib1.RootFolder, template);
                cc.Load(file2);
                cc.ExecuteQueryRetry();


                file2["Title"] = listItem.FieldValues["Title"];
                file2["Attorney_x0028_Full_x0020_Name_x0029_"] = listItem.FieldValues["Attorney"];
                file2["AttorneyTyping"] = listItem.FieldValues["AttorneyTyping"];
                file2["AttorneyFees"] = String.Format("{0:0.00}", (double)(listItem.FieldValues["AttorneyFees"] ?? 0.00));
                file2["County"] = listItem.FieldValues["County"];
                file2["Court"] = listItem.FieldValues["Court"];
                file2["Court_x0020_Date"] = listItem.FieldValues["Court_x0020_Date"];
                file2["Court_x0020_Time"] = listItem.FieldValues["Court_x0020_Time"];
                file2["Judge_x0020__x0028_Full_x0020_Name_x0029_"] = listItem.FieldValues["Judge"];
                file2["Late_x0020_Fees"] = String.Format("{0:0.00}", (double)(listItem.FieldValues["Late_x0020_Fees"] ?? 0.00));
                file2["MonthlyRent"] = String.Format("{0:0.00}", (double)(listItem.FieldValues["MonthlyRent"] ?? 0.00));
                file2["PetitionSentDate"] = listItem.FieldValues["PetitionSentDate"];
                file2["Property_x0020_Owner"] = listItem.FieldValues["Property_x0020_Owner"];
                file2["Rent_x002b_Fees"] = String.Format("{0:0.00}", (double)(listItem.FieldValues["Rent_x002b_Fees"] ?? 0.00));
                file2["Tenant_x0020_Name_x0020_2"] = listItem.FieldValues["Tenant_x0020_Name_x0020_2"] ?? " ";
                file2["Tenant2csv"] = (listItem.FieldValues["Tenant_x0020_Name_x0020_2"] ?? "").ToString() == "" ? " " : $", {(listItem.FieldValues["Tenant_x0020_Name_x0020_2"] ?? "").ToString()}";
                file2["Tenant_x0020_Name_x0020_3"] = listItem.FieldValues["Tenant_x0020_Name_x0020_3"];
                file2["Total_x0020_Rent"] = String.Format("{0:0.00}", (double)(listItem.FieldValues["Total_x0020_Rent"] ?? 0.00));
                file2["TotalJudgment"] = String.Format("{0:0.00}", (double)(listItem.FieldValues["TotalJudgment"] ?? 0.00));
                file2["Oldest_x0020_Month_x0020_Past_x002d_Due_x0020_Rent"] = listItem.FieldValues["Oldest_x0020_Month_x0020_Past_x0"];
                file2["Property_x0020_Owner_x0020_Address_x0020__x0028_City_x002c__x0020_State_x0020_and_x0020_Zip_x0029_"] = listItem.FieldValues["Tenant_x0020_Address_x0020__x002"];
                file2["Property_x0020_Owner_x0020_Address_x0020__x0028_Street_x0029_"] = listItem.FieldValues["Property_x0020_Owner_x0020_Addre"];
                file2["Tenant_x0020_Address_x0020__x0028_City_x002c__x0020_State_x0020__x0026__x0020_Zip_x0029_"] = listItem.FieldValues["Tenant_x0020_Address_x0020__x002"];
                file2["Tenant_x0020_Address_x0020__x0028_Street_x0020_and_x0020_Unit_x0029_"] = listItem.FieldValues["Tenant_x0020_Address"];
                file2["Town"] = listItem.FieldValues["Town"];
                log.LogInformation("DEBUG: Updating Document Metadata...");
                file2.Update();
                cc.ExecuteQueryRetry();
                //}
                //catch (Exception ex)
                //{
                //    log.LogError($"ERROR: {ex.Message}");
                //    throw ex;
                //}
            }
        }

        private static SecureString SecureString(string v)
        {
            SecureString secure = new SecureString();
            foreach (char c in v)
            {
                secure.AppendChar(c);
            }
            return secure;
        }

        public static void UpdateLastChangeToken(NotificationModel notification, ChangeToken lastChangeToken, CloudTable table)
        {
            // Persist the last used changetoken as we'll start from that one when the next event hits our service
            LastChangeEntity lce = new LastChangeEntity("LastChangeToken", notification.Resource) { LastChangeToken = lastChangeToken.StringValue };
            TableOperation insertOperation = TableOperation.InsertOrReplace(lce);
            table.Execute(insertOperation);
        }

        public static void RecordChangeInWebhookHistory(ClientContext cc, List changeList, Change change)
        {
            #region Grab the list used to write the webhook history
            // Ensure reference to the history list, create when not available
            List historyList = null;
            string historyListName = CloudConfigurationManager.GetSetting("HistoryListName");
            if (!string.IsNullOrEmpty(historyListName))
            {
                historyList = cc.Web.GetListByTitle(historyListName);
                if (historyList == null)
                {
                    historyList = cc.Web.CreateList(ListTemplateType.GenericList, historyListName, false);
                    cc.ExecuteQueryRetry();
                }
            }
            #endregion

            if (historyList == null) return;
            try
            {
                ListItemCreationInformation newItem = new ListItemCreationInformation();
                ListItem item = historyList.AddItem(newItem);
                item["Title"] = string.Format($"List {changeList.Title} had a Change of type \"{change.ChangeType.ToString()}\" on the item with Id {(change as ChangeItem).ItemId}. Change Token: {(change as ChangeItem).ChangeToken.StringValue}");
                item.Update();
                cc.ExecuteQueryRetry();
            }
            catch (Exception ex)
            {
                //log.LogError($"ERROR: {ex.Message}");
            }
        }
    }
}
