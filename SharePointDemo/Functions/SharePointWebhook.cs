using Microsoft.Azure;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using SharePointDemo.Models;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace SharePointDemo.Functions
{
    public static class SharePointWebhook
    {
        [FunctionName("SharePointWebhook")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous,
                                                            "post",
                                                            Route = null)]
                                                            HttpRequestMessage req, TraceWriter log)
        {
            try
            {
                log.Info($"Triggered started...");
                string validationToken = req.GetQueryNameValuePairs()
                    .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                    .Value;
                if (validationToken != null)
                {
                    log.Info($"Validation token {validationToken} received");
                    HttpResponseMessage response = req.CreateResponse(HttpStatusCode.OK);
                    response.Content = new StringContent(validationToken);
                    return response;
                }

                string content = await req.Content.ReadAsStringAsync();
                log.Info($"Payload: {content}");

                System.Collections.Generic.List<NotificationModel> notifications =
                    JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;

                log.Info($"Notification count: {notifications.Count}");

                if (notifications.Count > 0)
                {
                    CloudStorageAccount storageAccount = CloudStorageAccount
                                    .Parse(CloudConfigurationManager.GetSetting("AzureWebJobsStorage"));
                    CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                    CloudQueue queue = queueClient.GetQueueReference(
                                        CloudConfigurationManager.GetSetting("WebhooksQueueName"));
                    queue.CreateIfNotExists();

                    foreach (NotificationModel notification in notifications)
                    {
                        log.Info($"Processing notification: {notification.Resource}");
                        queue.AddMessage(new CloudQueueMessage(JsonConvert.SerializeObject(notification)));
                    }
                }
                return new HttpResponseMessage(HttpStatusCode.OK);
            }
            catch (System.Exception ex)
            {
                throw ex;

            }
        }
    }
}
