using SharePointDemo.Functions;
using System;

namespace SharePointDemo.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            ProcessQueueWebhook();
        }

        private static void ProcessQueueWebhook()
        {
            TraceLogger log = new TraceLogger();
            var notification = "{'subscriptionId':'01505ae4-c8bc-4bf4-9c70-ce7bd331969b','clientState':'NACS','expirationDateTime':'2019-06-12T12:53:13.7590000Z','resource':'ca6b34a2-a78e-40d2-8639-8e5da4c093cf','tenantId':'873e586a-ce9d-4dc7-b8f8-e3389e8a30c8','siteUrl':'/sites/WebhookDemo','webId':'830a9155-f964-4932-81de-7da6768e4bc4'}";
            QueueWebhook.ProcessNotification(log, notification);
            System.Console.Write("Press any key to close program...");
            System.Console.ReadKey();
        }
    }
}
