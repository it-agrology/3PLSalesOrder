using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace _3PLSalesOrder
{
    class Program
    {
        static void Main(string[] args)
        {
            string specificSenderAddress = "a.protopsalti@agrology.eu";
            string attachmentFileExtension = ".xls";
            string targetDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            SaveAttachmentsFromSpecificSender(specificSenderAddress, attachmentFileExtension, targetDirectory);
        }

        private static void SaveAttachmentsFromSpecificSender(string senderAddress, string fileExtension, string targetDirectory)
        {
            Outlook.Application outlookApplication = new Outlook.Application();
            Outlook.NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");

            Outlook.MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items items = inboxFolder.Items;

            foreach (Outlook.MailItem mailItem in items)
            {
                if (mailItem.SenderEmailAddress.ToLower() == senderAddress.ToLower())
                {
                    foreach (Outlook.Attachment attachment in mailItem.Attachments)
                    {
                        if (Path.GetExtension(attachment.FileName).ToLower() == fileExtension.ToLower())
                        {
                            string targetFilePath = Path.Combine(targetDirectory, attachment.FileName);
                            attachment.SaveAsFile(targetFilePath);
                            Console.WriteLine($"Attachment '{attachment.FileName}' saved to '{targetFilePath}'.");
                        }
                    }
                }

                ReleaseIfWindows(mailItem);
            }

            ReleaseIfWindows(items);
            ReleaseIfWindows(inboxFolder);
            ReleaseIfWindows(outlookNamespace);
            ReleaseIfWindows(outlookApplication);
        }

        [SupportedOSPlatform("windows")]
        private static void ReleaseIfWindows(object obj)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Marshal.ReleaseComObject(obj);
            }
        }
    }
}
