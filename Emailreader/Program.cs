using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using VirusTotalNET;
using VirusTotalNET.Results;
using VirusTotalNET.ResponseCodes;
using VirusTotalNET.Objects;
using System.Configuration;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using System.Linq;

namespace Emailreader
{

    class Program
    {
        public string ScanUrl = "";
        public object mapiNameSpace;
        public bool foundMalicious = false;
        public static string MaliciousFolder { get; private set; }
        public static Microsoft.Office.Interop.Outlook.MAPIFolder destFolder;
        public static Microsoft.Office.Interop.Outlook.MAPIFolder myInbox;
        public string file = "";

        static void Main(string[] args)
        {

            Microsoft.Office.Interop.Outlook.Application myApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");
            Microsoft.Office.Interop.Outlook.MAPIFolder myContacts = mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
            mapiNameSpace.Logon(null, null, false, false);//"suraj2.kumar@lntinfotech.com""14March96kr#"
            mapiNameSpace.Logon(ConfigurationSettings.AppSettings["MailId"], ConfigurationSettings.AppSettings["Password"], Missing.Value, true);
            Microsoft.Office.Interop.Outlook.Items myItems = myContacts.Items;
            myInbox = mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            destFolder = myInbox.Folders["MaliciousAttachments"];
            Microsoft.Office.Interop.Outlook.Items inboxItems = myInbox.Items;
            Microsoft.Office.Interop.Outlook.Explorer myexp = myInbox.GetExplorer(false);
            mapiNameSpace.Logon("Profile", Missing.Value, false, true);
            new Program().readingUnReadMailsAsync(inboxItems);
        }


        public void readingUnReadMailsAsync(Microsoft.Office.Interop.Outlook.Items inboxItems)
        {

            if (myInbox.Items.Count > 0)
            {
                Microsoft.Office.Interop.Outlook.Items items = myInbox.Items;
                inboxItems = inboxItems.Restrict("[UnRead] = true");
                Console.WriteLine(string.Format("Total Unread message {0}:", inboxItems.Count));
                int x = 0;

                foreach (Microsoft.Office.Interop.Outlook.MailItem item in inboxItems)
                {
                    var attachments = item.Attachments;
                    var urlhtml = item.HTMLBody;
                    
                    string Textbody = item.Body;
                    var matches = Regex.Matches(urlhtml, @"<a\shref=""(?<url>.*?)"">(?<text>.*?)</a>");
                    if (attachments.Count != 0)
                    {
                        new LogFile().Writelog("Unread Mail from your inbox :-");
                        ++x;
                        Console.WriteLine("{0} Unread Mail from your inbox", ++x);
                        new LogFile().Writelog("            x"); 
                        new LogFile().Writelog("from:-");
                        string SenderName = item.SenderName;
                        new LogFile().Writelog("SenderName");
                        string ReceiverName = item.ReceivedByName;
                        new LogFile().Writelog("ReceiverName");
                        string Subject = item.Subject;
                        new LogFile().Writelog("Subject");
                         Console.WriteLine(string.Format("from:-     {0}", item.SenderName));
                         Console.WriteLine(string.Format("To:-     {0}", item.ReceivedByName));
                         Console.WriteLine(string.Format("Subject:-     {0}", item.Subject));
                        // Console.WriteLine(string.Format("Message:-     {0}", item.Body));                       
                        for (int i = 1; i <= attachments.Count; i++)
                        {
                            new LogFile().Writelog("Attachments:-");
                            string Attachments = item.Attachments[i].FileName;
                            new LogFile().Writelog("Attachments");                           
                            file = Path.GetFullPath(item.Attachments[i].FileName);
                            new VirusChecking().processVirusTotal(file, item, ScanUrl, matches);
                        }
                    }
                    else if (matches.Count != 0)
                    {
                        foreach (Match url in matches)
                        {
                            ScanUrl = url.Groups["url"].Value;
                            new VirusChecking().processVirusTotal(file, item, ScanUrl, matches);
                        }
                    }
                    else
                    {
                        string path = @"C:\Users\10653836\Desktop\dll.txt";
                        string[] contents = File.ReadAllLines(path);
                        if (contents.Any(item.Subject.Contains))
                        {
                            new VirusChecking().MoveAttactment(item);
                        }
                    }

                }

            }
        }


    }
}


