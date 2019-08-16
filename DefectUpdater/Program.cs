using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DefectUpdater {
    class Program {
        static string UpdateRequestFrom { get; set; }
        static string ProjectName { get; set; }
        static string UpgradeName { get; set; }

        static void Main(string[] args) {
            Console.WriteLine("Getting data for processing...");
            string errorLogPath = @"\Error.log";
            string userName = "";
            try {
                UpdateRequestFrom = args[0];
                userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace("\\", "").ToUpper();
                ProjectName = GetProjectName(UpdateRequestFrom);
                UpgradeName = GetUpgradeName(UpdateRequestFrom);
                ExcelHandler excelHandler = new ExcelHandler();
                List<KnownDefect> knownDefects = new List<KnownDefect>();
                knownDefects.AddRange(excelHandler.GetDeviationsFromExcel(UpdateRequestFrom, ProjectName, UpgradeName));
                Console.WriteLine("Received  " + knownDefects.Count + " unique records with defects");
                int countUpdated = 0;
                int countInserted = 0;
                int countDeleted = 0;
                if (knownDefects.Count > 0) {
                    OraSession oraSession = new OraSession("", "", "", "", "");
                    oraSession.OpenConnection();
                    foreach (var defect in knownDefects) {
                        string defectNo = oraSession.GetDefectNoFromDB(defect);
                        if (defectNo == "" && defect.DefectNo != "0") {
                            oraSession.InsertIntoDefectsTable(defect, userName);
                            countInserted++;
                        } else if (defect.DefectNo == "0" && defectNo != "") {
                            oraSession.DeleteDefectsTable(defect);
                            countDeleted++;
                        } else if (defectNo != defect.DefectNo && defectNo != "") {
                            oraSession.UpdateDefectsTable(defect, userName);
                            countUpdated++;
                        }
                    }
                    oraSession.CloseConnection();
                } else {
                    Console.WriteLine("Nothing to do");
                }
                Console.WriteLine("-----------------------------------------------");
                Console.WriteLine("Updated " + countUpdated + " deviation(s)");
                Console.WriteLine("Inserted " + countInserted + " deviation(s)");
                Console.WriteLine("Deleted " + countDeleted + " deviation(s)");
                Console.WriteLine("-----------------------------------------------");
                Console.WriteLine("Task completed");
            } catch (Exception ex) {
                Console.WriteLine("-----------------------------------------------");
                Console.WriteLine("Unable to update database due to error =(");
                Console.WriteLine("The problem is registered and will be fixed soon.");
                Console.WriteLine("-----------------------------------------------");
                Console.WriteLine("Task failed");
                WriteLog(errorLogPath, ex.Message, userName);
            }
            Console.WriteLine();
            Console.WriteLine("Now you can close the application. Or it will close automatically after 10 seconds.");
            Timer t = new Timer(CloseApp, null, 10000, 10000);
            Console.ReadKey();
        }

        private static string GetProjectName(string filePath) {
            var r = filePath.Split('\\');
            return r[2].Trim();
        }

        private static string GetUpgradeName(string filePath) {
            var r = filePath.Split('\\');
            return r[3].Replace("'", "").Trim();
        }

        private static void WriteLog(string path, string message, string user) {
            try {
                List<string> content = new List<string>();
                content.Add("Time: " + DateTime.Now);
                content.Add("User: " + user);
                content.Add("ProjectName: " + ProjectName);
                content.Add("UpgradeName: " + UpgradeName);
                content.Add("ExcelPath: " + UpdateRequestFrom);
                content.Add("ErrorMessage: " + message);
                content.Add("--------------------------------------------------------------------------");
                File.AppendAllLines(path, content);
            } catch (Exception) { }
        }

        private static void CloseApp(object state) {
            Environment.Exit(0);
        }

    }
}
