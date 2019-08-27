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
        static List<double> UpgradeVersions { get; set; }

        static void Main(string[] args) {
            Console.WriteLine("Getting data for processing...");
            string errorLogPath = @"O:\DATA\COMMON\core\defects\Error.log";
            string userName = "";
            try {
                //UpdateRequestFrom = @"I:\VT Execution\BIA\Upgrade 6.3 to 19.04\Temp\Compared_LV_Transactions_22082019_FINAL.xlsm";     
                UpdateRequestFrom = args[0];
                userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace("SCDOM\\", "").ToUpper();
                ProjectName = GetProjectName(UpdateRequestFrom);
                UpgradeVersions = GetUpgradeName(UpdateRequestFrom);
                if(ProjectName == "" || UpgradeVersions == null) {
                    Console.WriteLine("-----------------------------------------------");
                    var errMessage = "Please put your result file into the project folder and then launch the update process";
                    Console.WriteLine(errMessage);
                    WriteLog(errorLogPath, errMessage, userName);
                    Timer tt = new Timer(CloseApp, null, 10000, 10000);
                    Console.ReadKey();
                    return;
                }
                ExcelHandler excelHandler = new ExcelHandler();
                var knownDefects = excelHandler.GetDeviationsFromExcel(UpdateRequestFrom, ProjectName, UpgradeVersions);
                if(knownDefects == null) {
                    Console.WriteLine("-----------------------------------------------");
                    var errMessage = "Cannot find identifiers, file structure is broken!";
                    Console.WriteLine(errMessage);
                    WriteLog(errorLogPath, errMessage, userName);
                    Timer tt = new Timer(CloseApp, null, 10000, 10000);
                    Console.ReadKey();
                    return;
                }
                Console.WriteLine("Received  " + knownDefects.Count + " unique records with defects");
                int countUpdated = 0;
                int countInserted = 0;
                int countDeleted = 0;
                if (knownDefects.Count > 0) {
                    OraSession oraSession = new OraSession("*", "*", "*", "*", "*");
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
            if (r.Length > 3) {
                return r[2].Trim();
            } else {
                return "";
            }
        }

        private static List<double> GetUpgradeName(string filePath) {
            var r = filePath.Split('\\');
            if (r.Length <= 3) {
                return null;
            }
            var upgrade = r[3];
            StringBuilder pattern = new StringBuilder();
            List<double> upgradeName = new List<double>();
            foreach (var item in upgrade) {
                if (char.IsDigit(item) || (pattern.Length > 0 && item == '.' || pattern.Length > 0 && item == ',')) {
                    pattern.Append(item);
                } else if (pattern.Length > 0 && item == ' ') {
                    double d;
                    var isDouble = double.TryParse(pattern.ToString().Replace('.', ','), out d);
                    if (isDouble) {
                        upgradeName.Add(d);
                        pattern.Clear();
                    } else {
                        return null;
                    }
                }
            }
            if (pattern.Length > 0) {
                double d;
                var isDouble = double.TryParse(pattern.ToString().Replace('.', ','), out d);
                if (isDouble) {
                    upgradeName.Add(d);
                    pattern.Clear();
                    return upgradeName;
                } else {
                    return null;
                }
            } else {
                return null;
            }
        }

        private static void WriteLog(string path, string message, string user) {
            try {
                List<string> content = new List<string>();
                content.Add("Time: " + DateTime.Now);
                content.Add("User: " + user);
                content.Add("ProjectName: " + ProjectName);
                if(UpgradeVersions != null && UpgradeVersions.Count == 2) {
                    content.Add("UpgradeName: " + UpgradeVersions[0] + "->" + UpgradeVersions[1]);
                } else {
                    content.Add("UpgradeName: null");
                }                
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
