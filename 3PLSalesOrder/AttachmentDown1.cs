using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace _3PLSalesOrder
{
    class AttachmentDown1
    {
        private const string clientId = "YOUR_CLIENT_ID";
        private const string tenantId = "YOUR_TENANT_ID";
        private const string clientSecret = "YOUR_CLIENT_SECRET";
        private const string emailSubject = "EMAIL_SUBJECT";
        private const string localFolderPath = @"C:\Downloads\";

        static async Task Main(string[] args)
        {
            // Create a new GraphServiceClient with client credentials authentication
            var confidentialClient = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();

            var authProvider = new ClientCredentialProvider(confidentialClient);

            var graphClient = new GraphServiceClient(authProvider);

            // Search for the email with the specified subject
            var messages = await graphClient.Me.Messages
                .Request()
                .Filter($"subject eq '{emailSubject}'")
                .GetAsync();

            var message = messages.FirstOrDefault();

            if (message == null)
            {
                Console.WriteLine($"No email found with subject: {emailSubject}");
                return;
            }

            // Download the attachments from the email
            var attachments = await graphClient.Me.Messages[message.Id]
                .Attachments
                .Request()
                .GetAsync();

            // Iterate through the attachments and save the xls file to a local folder
            foreach (var attachment in attachments)
            {
                if (attachment.ContentType == "application/vnd.ms-excel" &&
                    attachment.Name.EndsWith(".xls"))
                {
                    var attachmentStream = await graphClient.Me.Messages[message.Id]
                        .Attachments[attachment.Id]
                        .Content
                        .Request()
                        .GetAsync();

                    var filePath = Path.Combine(localFolderPath, attachment.Name);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await attachmentStream.CopyToAsync(fileStream);
                    }

                    Console.WriteLine($"Attachment saved to {filePath}");
                }
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
