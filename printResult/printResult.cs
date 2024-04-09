using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace printResult
{

    [ComVisible(true)]
    [ProgId("printResult.printResult")]

    public class printResult : IWorkflowExtension, IEntityExtension
    {
        private const string Type = "2";
        private int _port = 9100;
        INautilusServiceProvider sp;
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        private IDataLayer dal;
        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {

                #region parms
                sp = Parameters["SERVICE_PROVIDER"];
                var workstationId = Parameters["WORKSTATION_ID"];
                var rs = Parameters["RECORDS"];
                var resultIdStr = rs.Fields["RESULT_ID"].Value;
                #endregion


                int resultId = Convert.ToInt32(resultIdStr);
                //Get Connection String
                var ntlCon = Utils.GetNtlsCon(sp);

                Utils.CreateConstring(ntlCon);
                dal = new DataLayer();
                dal.Connect();
                Result result = dal.GetResultById(resultId);
                Test test = result.Test;
                Workstation ws = dal.getWorkStaitionById(workstationId);

                ReportStation reportStation = dal.getReportStationByWorksAndType(ws.NAME, Type);
                string GoodIp = "";//removeBadChar(ip);
                //            string printerName = "";

                if (reportStation != null)
                {

                    if (reportStation.Destination != null)
                    {
                        //מקבל את ה IP של המדפסת
                        GoodIp = reportStation.Destination.ManualIP;
                    }
                    if (reportStation.Destination != null && reportStation.Destination.RawTcpipPort != null)
                    {
                        //מקבל את הפורט רק במקרה שהוא שונה מהדיפולט
                        _port = (int)reportStation.Destination.RawTcpipPort;
                    }
                    Result res = dal.GetResultById(resultId);
                    Aliquot aliq = res.Test.Aliquot;
                    var sampleName = test.Aliquot.Sample.Name;
                    var testcode = "";
                    testcode = getTestCode(aliq, testcode);
                    var mihol = res.DilutionFactor;
                    //הוספת תווים לזיהוי שיירדו בקליטה 
                    sampleName = sampleName + "_@";

                    Print(sampleName, "r" + resultId.ToString(), testcode, mihol.ToString(), GoodIp);

                }
                else
                {
                    MessageBox.Show("לא הוגדרה תחנה");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת מדבקה");
                Logger.WriteLogFile(ex);
            }
        }
        private string removeBadChar(string ip)
        {
            string ret = "";
            foreach (var c in ip)
            {
                int ascii = (int)c;
                if ((ascii >= 48 && ascii <= 57) || ascii == 44 || ascii == 46)
                    ret += c;
            }
            return ret;
        }
        public string GetIp(string printerName)
        {
            string query = string.Format("SELECT * from Win32_Printer WHERE Name LIKE '%{0}'", printerName);
            string ret = "";
            var searcher = new ManagementObjectSearcher(query);
            var coll = searcher.Get();
            foreach (ManagementObject printer in coll)
            {
                foreach (PropertyData property in printer.Properties)
                {
                    if (property.Name == "PortName")
                    {
                        ret = property.Value.ToString();
                    }
                }
            }
            return ret;
        }
        private static string ReverseString(string s)
        {
            var str = s;
            string[] strsubs = s.Split(Convert.ToChar(" "));
            var newstr = "";
            string substr = "";
            int i;
            int c = strsubs.Count();
            for (i = 0; i < c; ++i)
            {
                substr = strsubs[i];
                if (HasHebrewChar(strsubs[i]))
                {
                    substr = Reverse(substr);
                }

                newstr += substr + " ";
            }
            return newstr;
        }

        private static string Reverse(string s)
        {
            char[] arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }

        public static bool HasHebrewChar(string value)
        {
            return value.ToCharArray().Any(x => (x <= 'ת' && x >= 'א'));
        }


        public void Print(string name, string ID, string testcode, string mihol, string ip)
        {
            string ipAddress = ip;


            // ZPL Command(s)
            string ntxt = name;
            string tctxt = testcode;
            string mtxt = mihol;
            string itxt = ID;
            //string ZPLStringOld =
            //     "^XA" +
            //     "^LH10,10" +
            //     "^FO10,0" +
            //     "^A@N20,20" +
            //    string.Format("^FD{0}^FS", ntxt) +
            //     "^FO10,60" +
            //     "^A@N20,20" +
            //    string.Format("^FD{0}^FS", mtxt) +
            //    "^FO100,60" +
            //     "^A@N20,20" +
            //    string.Format("^FD{0}^FS", tctxt) +
            //    "^FO260,0" + "^BQN,4,3" +
            //     string.Format("^FDLA,{0}^FS", itxt) + //ברקוד
            //    "^XZ";
            string ZPLString =
                 "^XA" +
                 "^LH0,0" +
                 "^FO20,10" +
                 "^A@N30,30" +
                string.Format("^FD{0}^FS", ntxt) +
                 "^FO10,60" +
                 "^A@N30,30" +
                string.Format("^FD{0}^FS", mtxt) +
                "^FO100,60" +
                 "^A@N30,30" +
                string.Format("^FD{0}^FS", tctxt) +
                "^FO320,0" + "^BQN,4,3" +
                 string.Format("^FDLA,{0}^FS", itxt) + //ברקוד
                "^XZ";
            try
            {
                // Open connection
                var client = new System.Net.Sockets.TcpClient();
                client.Connect(ipAddress, _port);

                // Write ZPL String to connection
                var writer = new StreamWriter(client.GetStream());
                writer.Write(ZPLString);
                writer.Flush();

                // Close Connection
                writer.Close();
                client.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private string getTestCode(Aliquot aliq, string testcode)
        {

            if (aliq.Parent.Count == 0)// && aliq.U_CHARGE == "T")//ספי ביקש להוריד
            {
                testcode = aliq.ShortName;
            }
            if (aliq.Parent.Count != 0)
            {
                getTestCode(aliq.Parent.FirstOrDefault(), testcode);
            }
            return testcode;
        }

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            sp = Parameters["SERVICE_PROVIDER"];
            var rs = Parameters["RECORDS"];
            var resultIdStr = rs.Fields["RESULT_ID"].Value;
            int resultId = Convert.ToInt32(resultIdStr);
            //Get Connection String
            var ntlCon = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(ntlCon);
            dal = new DataLayer();
            dal.Connect();
            Result result = dal.GetResultById(resultId);
            PhraseHeader phraseHeader = dal.GetPhraseByName("CanExecuteResultTemplates");
            string templetName = result.ResultTemplate.Name;
            foreach (var item in phraseHeader.PhraseEntries)
            {
                if (item.PhraseName == templetName)
                {
                    return ExecuteExtension.exEnabled;
                }
            }

            dal.Close();
            return ExecuteExtension.exEnabled;
        }
    }
}
