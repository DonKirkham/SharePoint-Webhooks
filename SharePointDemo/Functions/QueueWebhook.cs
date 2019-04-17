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
            //log.LogInformation($"Queue item: {queueItem}");
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
                                RecordChangeInWebhookHistory(cc, changeList, change, log);
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
        }

        public static void UpdateLastChangeToken(NotificationModel notification, ChangeToken lastChangeToken, CloudTable table)
        {
            // Persist the last used changetoken as we'll start from that one when the next event hits our service
            LastChangeEntity lce = new LastChangeEntity("LastChangeToken", notification.Resource) { LastChangeToken = lastChangeToken.StringValue };
            TableOperation insertOperation = TableOperation.InsertOrReplace(lce);
            table.Execute(insertOperation);
        }

        public static void RecordChangeInWebhookHistory(ClientContext cc, List changeList, Change change, ILogger log)
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
                log.LogError($"ERROR: {ex.Message}");
            }
        }
    }
}
