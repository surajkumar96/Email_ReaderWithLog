using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace Emailreader
{
    class LogFile
    {
        static string m_BaseDir = null;
         public string FilePath = @"C:\Users\10653836\Desktop";
         public string FileName = "Log.txt";

        public void Writelog(string message)
        {
            
            try
            {
                if (!System.IO.File.Exists(FileName))
                {
                    string PathName = System.IO.Path.Combine(FilePath, FileName);
                    System.IO.StreamWriter sw = new System.IO.StreamWriter(FileName, true);
                    XElement xmlEntry = new XElement("LogEntry", new XElement("Date:", System.DateTime.Now.ToString()), new XElement("Message:",message));
                    sw.WriteLine(xmlEntry);
                    sw.Close();
                }
                else
                {
                    System.IO.StreamWriter sw = new System.IO.StreamWriter(FileName, true);
                    XElement xmlEntry = new XElement("LogEntry", new XElement("Date:", System.DateTime.Now.ToString()),new XElement("Message:", message));
                    sw.WriteLine(xmlEntry);
                    sw.Close();
                }
              

            }catch(Exception e)
            {
                LogException(e);
            }
        }
        public void LogException(Exception ex)
        {
            System.IO.StreamWriter sw = new System.IO.StreamWriter(FileName, true);
            XElement xmlEntry = new XElement("logEntry",
            new XElement("Date", System.DateTime.Now.ToString()),
            new XElement("Exception",
            new XElement("Source", ex.Source),
            new XElement("Message", ex.Message),
            new XElement("Stack", ex.StackTrace)
         )
         );
            xmlEntry.Element("Exception").Add(
                    new XElement("InnerException",
                        new XElement("Source", ex.InnerException.Source),
                        new XElement("Message", ex.InnerException.Message),
                        new XElement("Stack", ex.InnerException.StackTrace))
                    );

        }
    }
}
