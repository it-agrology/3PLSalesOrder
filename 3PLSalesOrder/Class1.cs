using System;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Excel;

namespace _3PLSalesOrder
{
    class ReceiveXLS
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

            // Set the folder to retrieve the email from
            Microsoft.Office.Interop.Outlook.MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            Microsoft.Office.Interop.Outlook.MailItem email = inboxFolder.Items[1] as Microsoft.Office.Interop.Outlook.MailItem; // Retrieve the first email in the folder

            // Retrieve the attachment
            Microsoft.Office.Interop.Outlook.Attachment attachment = email.Attachments["filename.xls"];

            // Save the attachment to the local Documents folder
            string documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string savePath = Path.Combine(documentsFolder, attachment.FileName);
            attachment.SaveAsFile(savePath);

            // Quit the Outlook application
            outlookApp.Quit();

            // Open the saved file in Excel
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(savePath);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
            worksheet.Activate();
        }
    }
}
