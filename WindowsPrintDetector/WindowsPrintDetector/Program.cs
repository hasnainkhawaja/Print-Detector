using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace WindowsPrintDetector
{

    class Program
    {
        static List<string> listJobCollection = new List<string>();
        static void Main(string[] args)
        {

            StringCollection listePrinter = GetPrinterCollection();
            foreach (string nomPrinter in listePrinter)
            {
                Console.WriteLine(nomPrinter);

            }
            try
            {
                string query = "SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'";
                WqlEventQuery WQLquery = new WqlEventQuery(query);
                ManagementEventWatcher eventWatcher = new ManagementEventWatcher(WQLquery);
                Console.WriteLine("waiting for printjob...");
                eventWatcher.EventArrived += EventWatcher_EventArrived;
                eventWatcher.Start();
                //System.Threading.Thread.Sleep(10000);
                //eventWatcher.Stop();

            }
            catch (Exception e)
            {
                Console.WriteLine("{0}", e.Message);
            }
            Console.ReadLine();
        }

        private static void EventWatcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            Console.WriteLine("Printer Jobs detected");
            listJobCollection = GetPrintJobCollection();
            foreach (string jobName in listJobCollection.Distinct())
            {
                Console.WriteLine(jobName);
            }
        }

        public static StringCollection GetPrinterCollection()
        {
            StringCollection stringCollection = new StringCollection();
            string searchPrinter = "SELECT * FROM Win32_printer";
            ManagementObjectSearcher mos = new ManagementObjectSearcher(searchPrinter);
            ManagementObjectCollection printerCollection = mos.Get();
            foreach (ManagementObject mo in printerCollection)
            {
                stringCollection.Add(mo["Name"].ToString());
            }
            return stringCollection;
        }

        public static List<string> GetPrintJobCollection()
        {
            List<string> listPrintJob = new List<string>();
            string searchJob = "SELECT * FROM Win32_PrintJob";
            ManagementObjectSearcher mos = new ManagementObjectSearcher(searchJob);
            ManagementObjectCollection moc = mos.Get();
            foreach (ManagementBaseObject Mo in moc)
            {
                string jobName = Mo["Name"].ToString();
                string printerName = jobName.Split(',')[0];
                string docName = Mo["Document"].ToString();
                string jobid = Mo["JobId"].ToString();
                string StatusMask = Mo["StatusMask"].ToString();
                string Status = Mo["Status"].ToString();
                string TotalPages = Mo["TotalPages"].ToString();
                string TimeSubmitted = Mo["TimeSubmitted"].ToString();
                string PaperSize = Mo["PaperSize"].ToString();
                string PaperWidth = Mo["PaperWidth"].ToString();
                string Description = Mo["Description"].ToString();
                string Color = Mo["Color"].ToString();
                
                /***
                 * The DataType property indicates the format of the data for this print job. This instructs the printer driver to eithertranslate the data (generic text, PostScript, or PCL) before printing, or to print in a raw format (for graphics and pictures).

Example: TEXT*/
                string DataType = Mo["DataType"].ToString();
            }
            return listPrintJob;
        }
    }
}
