using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Diagnostics;

namespace OfficeConvert
{
    public class OutlookConverter : Converter
    {
        private Outlook.Application app;
        private Outlook.MailItem mailItem;

        public OutlookConverter()
        {                
        }

        public OutlookConverter(string inputFile, string outputFile)
        {
            this.Convert(inputFile, outputFile);
        }

        public override void Convert(String inputFile, String outputFile)
        {
            Object nothing = Type.Missing;
            try
            {
                if (!File.Exists(inputFile))
                {
                    throw new ConvertException("File not Exists");
                }

                if (IsPasswordProtected(inputFile))
                {
                    throw new ConvertException("Password Exist");
                }

                string tmpDocOutputFile = outputFile.Replace(".pdf", ".doc");

                if (inputFile.ToLower().EndsWith("eml"))
                {
                    ProcessOutlookMsg(inputFile);
                    // System.Diagnostics.Process.Start(inputFile);
                    // app = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");  // note that it returns an exception if Outlook is not running
                    // mailItem = (Outlook.MailItem)app.ActiveInspector().CurrentItem; // now pOfficeItem is the COM object that represents your .eml file
                    return;
                }
                else
                {
                    mailItem = app.Session.OpenSharedItem(inputFile) as Microsoft.Office.Interop.Outlook.MailItem;
                }
                mailItem.SaveAs(tmpDocOutputFile, Microsoft.Office.Interop.Outlook.OlSaveAsType.olDoc);

                new WordConverter().Convert(tmpDocOutputFile, outputFile);
            }
            catch (Exception e)
            {
                release();
                throw new ConvertException(e.Message);
            }

            release();
        }

        private void release()
        {
            if (mailItem != null)
            {
                try
                {
                    mailItem.Close(Outlook.OlInspectorClose.olDiscard);
                    releaseCOMObject(mailItem);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message + "\r\n" + e.ToString() + "\r\n" + e.StackTrace);
                }
            }

            if (app != null)
            {
                try
                {
                    app.Quit();
                    releaseCOMObject(app);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message + "\r\n" + e.ToString() + "\r\n" + e.StackTrace);
                }
            }
        }

        internal void ProcessOutlookMsg(string strPathToSavedEmailFile)
        {

            //Microsoft.Win32 namespace to get path from registry
            var strPathToOutlook = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe").GetValue("").ToString();
            //var strPathToOutlook=@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE";

            string strOutlookArgs;
            if (Path.GetExtension(strPathToSavedEmailFile) == ".eml")
            {
                strOutlookArgs = @"/eml " + strPathToSavedEmailFile;  // eml is an undocumented outlook switch to open .eml files
            }
            else
            {
                strOutlookArgs = @"/f " + strPathToSavedEmailFile;
            }

            Process p = new System.Diagnostics.Process();
            p.StartInfo = new ProcessStartInfo()
            {
                CreateNoWindow = false,
                FileName = strPathToOutlook,
                Arguments = strOutlookArgs
            };
            p.Start();

            //Wait for Outlook to open the file
            Task.Delay(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();

            Microsoft.Office.Interop.Outlook.Application appOutlook = new Microsoft.Office.Interop.Outlook.Application();
            //Microsoft.Office.Interop.Outlook.MailItem myMailItem  = (Microsoft.Office.Interop.Outlook.MailItem)appOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            var myMailItem = (Microsoft.Office.Interop.Outlook.MailItem)appOutlook.ActiveInspector().CurrentItem;

            //Get the Email Address of the Sender from the Mailitem 
            string strSenderEmail = string.Empty;
            if (myMailItem.SenderEmailType == "EX")
            {
                strSenderEmail = myMailItem.Sender.GetExchangeUser().PrimarySmtpAddress;
            }
            else
            {
                strSenderEmail = myMailItem.SenderEmailAddress;
            }

            //Get the Email Addresses of the To, CC, and BCC recipients from the Mailitem   
            var strToAddresses = string.Empty;
            var strCcAddresses = string.Empty;
            var strBccAddresses = string.Empty;
            foreach (Microsoft.Office.Interop.Outlook.Recipient recip in myMailItem.Recipients)
            {
                const string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                Microsoft.Office.Interop.Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                string eAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                if (recip.Type == 1)
                {
                    if (strToAddresses == string.Empty)
                    {
                        strToAddresses = eAddress;
                    }
                    else
                    {
                        strToAddresses = strToAddresses + "," + eAddress;
                    }
                };
                if (recip.Type == 2)
                {
                    if (strCcAddresses == string.Empty)
                    {
                        strCcAddresses = eAddress;
                    }
                    else
                    {
                        strCcAddresses = strCcAddresses + "," + eAddress;
                    }
                };
                if (recip.Type == 3)
                {
                    if (strBccAddresses == string.Empty)
                    {
                        strBccAddresses = eAddress;
                    }
                    else
                    {
                        strBccAddresses = strBccAddresses + "," + eAddress;
                    }
                };
            }
            Console.WriteLine(strToAddresses);
            Console.WriteLine(strCcAddresses);
            Console.WriteLine(strBccAddresses);
            foreach (Microsoft.Office.Interop.Outlook.Attachment mailAttachment in myMailItem.Attachments)
            {
                Console.WriteLine(mailAttachment.FileName);
            }
            Console.WriteLine(myMailItem.Subject);
            Console.WriteLine(myMailItem.Body);
        }

    }
}
