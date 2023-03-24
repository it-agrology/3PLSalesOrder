using System;
using OpenPop.Pop3;
using OpenPop.Mime;
using System.IO;

namespace _3PLSalesOrder
{
    class ScriptMain
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("");

            string host = "smtp.office365.com";
            int port = 995;
            bool useSsl = true;
            string username = "info@agrology.eu";
            string password = "i2453i!!@";
            int emailIndex = 0; // the index of the email message containing the XLS file attachment
            string attachmentName = "CortevaSales.xls";
            string saveFilePath = @"C:\Downloads\";

            using (var client = new Pop3Client())
            {
                client.Connect(host, port, useSsl);
                client.Authenticate(username, password);

                var message = client.GetMessage(emailIndex);

                foreach (var attachment in message.FindAllAttachments())
                {
                    if (attachment.FileName.Equals(attachmentName))
                    {
                        using (var xlsFileStream = new FileStream(saveFilePath, FileMode.Create))
                        {
                            var xlsAttachment = attachment as MessagePart;
                            xlsAttachment.Save(xlsFileStream);
                        }

                        System.Console.WriteLine($"XLS file '{attachmentName}' is saved to {saveFilePath}");
                        break;
                    }
                }

                client.Disconnect();
            }
        }
    }
}
