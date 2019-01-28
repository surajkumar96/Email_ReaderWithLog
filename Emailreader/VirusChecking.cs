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
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Linq;

namespace Emailreader
{
    public class VirusChecking
    {
        public bool foundMalicious = false;
        public readonly Microsoft.Office.Interop.Outlook.MailItem item;
        private readonly MAPIFolder destFolder;

        public async Task processVirusTotal(string filepath, MailItem Items, string scanurl, MatchCollection matches)
        {

            VirusTotal virusTotal = new VirusTotal(ConfigurationSettings.AppSettings["Apikey"]);
            virusTotal.UseTLS = true;
            if (filepath != "") //For Attatchment
            {
                byte[] eicar = Encoding.ASCII.GetBytes(filepath);
                FileReport fileReport = await virusTotal.GetFileReportAsync(eicar);
                AttachmentScan(fileReport, filepath, Items, scanurl, matches);
                Console.WriteLine();
            }
            else if (filepath == "" && scanurl != "") //For Url
            {

                Task<UrlReport> urlReport = virusTotal.GetUrlReportAsync(scanurl);
                UrlScan(urlReport.Result, Items);

            }
            else
            {
                //To be Code for Body.
            }
            Console.WriteLine();
        }
        public void PrintScan(UrlScanResult scanResult)
        {
            new LogFile().Writelog("Scan ID: ");
            string ScanId = scanResult.ScanId;
            new LogFile().Writelog("ScanId");
            new LogFile().Writelog("Scan Message: ");
            string VerboseMsg = scanResult.VerboseMsg;
            new LogFile().Writelog("VerboseMsg");
            Console.WriteLine("Scan ID: " + scanResult.ScanId);
            Console.WriteLine("Message: " + scanResult.VerboseMsg);
            Console.WriteLine();
        }

        public void PrintScan(ScanResult scanResult)
        {
            new LogFile().Writelog("Scan ID: ");
            string ScanId = scanResult.ScanId;
            new LogFile().Writelog("ScanId");
            new LogFile().Writelog("Scan Message: ");
            string VerboseMsg = scanResult.VerboseMsg;
            new LogFile().Writelog("VerboseMsg");
            Console.WriteLine("Scan ID: " + scanResult.ScanId);
            Console.WriteLine("Message: " + scanResult.VerboseMsg);
            Console.WriteLine();
        }

        public void AttachmentScan(FileReport fileReport, string fileinfo, Microsoft.Office.Interop.Outlook.MailItem item, string scanurl, MatchCollection matches)
        {
            new LogFile().Writelog("Scan ID: ");
            string ScanId = fileReport.ScanId;
            new LogFile().Writelog("ScanId");
            new LogFile().Writelog("Scan Message: ");
            string VerboseMsg = fileReport.VerboseMsg;
            new LogFile().Writelog("VerboseMsg");
            Console.WriteLine("Scan ID: " + fileReport.ScanId);
            Console.WriteLine("Message: " + fileReport.VerboseMsg);


            if (fileReport.ResponseCode == FileReportResponseCode.Present)
            {

                foreach (KeyValuePair<string, ScanEngine> scan in fileReport.Scans)
                {
                    Console.WriteLine("{0,-25} Detected: {1}", scan.Key, scan.Value.Detected);
                    if (scan.Value.Detected == true)
                    {
                        MoveAttactment(item);
                        foundMalicious = true;
                    }
                }

                if (foundMalicious == false)
                {
                    if (matches.Count != 0)

                    {
                        foreach (Match url in matches)
                        {
                            scanurl = url.Groups["url"].Value;
                            MoveAttactment(item);
                            foundMalicious = true;
                        }
                    }
                }
                else if (foundMalicious == false)
                {
                    string path = @"C:\Users\10653836\Desktop\dll.txt";
                    string[] contents = File.ReadAllLines(path);
                    if (contents.Any(item.Subject.Contains))
                    {
                        MoveAttactment(item);
                    }

                }


            }
            Console.WriteLine();
        }




        public void MoveAttactment(Microsoft.Office.Interop.Outlook.MailItem item)
        {
            Microsoft.Office.Interop.Outlook.MailItem movemail = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder MaliciousAttachments = destFolder;
            movemail = item;
            if (movemail != null)
            {

                movemail.Move(MaliciousAttachments);
                item.Delete();

            }

        }

        public void UrlScan(UrlReport urlReport, Microsoft.Office.Interop.Outlook.MailItem item)
        {
            new LogFile().Writelog("Scan ID: ");
            string ScanId = urlReport.ScanId;
            new LogFile().Writelog("ScanId");
            new LogFile().Writelog("Scan Message: ");
            string VerboseMsg = urlReport.VerboseMsg;
            new LogFile().Writelog("VerboseMsg");
            Console.WriteLine("Scan ID: " + urlReport.ScanId);
            Console.WriteLine("Message: " + urlReport.VerboseMsg);

            if (urlReport.ResponseCode == UrlReportResponseCode.Present)
            {
                foreach (KeyValuePair<string, UrlScanEngine> scan in urlReport.Scans)
                {
                    Console.WriteLine("{0,-25} Detected: {1}", scan.Key, scan.Value.Detected);
                    if (scan.Value.Detected == true)
                    {
                        MoveAttactment(item);
                        foundMalicious = true;

                    }
                }
                if (foundMalicious == false)
                {

                    string path = AppDomain.CurrentDomain.BaseDirectory;
                    string[] contents = File.ReadAllLines(path);
                    if (contents.Any(item.Subject.Contains))
                    {
                        MoveAttactment(item);
                    }

                }
            }

            Console.WriteLine();
        }
    }
}

