using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DefectUpdater {
    public class ExcelHandler {
        List<string> KnownDefectsMerged { get; set; }

        private List<List<string>> ReadExcel(string path) {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path.Trim('\r', '\n'));
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;
            //xlWorksheet.Columns.AutoFilter(1, "<>", Excel.XlAutoFilterOperator.xlFilterValues);
            //Excel.Range visibleCells = xlWorksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            List<List<string>> res = new List<List<string>>();
            object[,] values = (object[,])xlRange.Value2;
            int NumRow = 1;
            while (NumRow <= values.GetLength(0)) {
                List<string> innerRes = new List<string>();
                for (int i = 1; i <= colCount; i++) {
                    innerRes.Add(Convert.ToString(values[NumRow, i]));
                }
                res.Add(innerRes);
                NumRow++;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(visibleCells);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            return res;
        }

        public List<KnownDefect> GetDeviationsFromExcel(string path, string proj, List<double> upgrade) {
            var excelData = ReadExcel(path);
            if (excelData.Count <= 1) {
                return null;
            }
            var headers = excelData.First();
            int mTransNoIndex = -1;
            int tTransNoIndex = -1;
            int secIdIndex = -1;
            for (int i = 0; i < headers.Count; i++) {
                if (headers[i].ToLower().Contains("m_trans")) {
                    mTransNoIndex = i;
                } else if (headers[i].ToLower().Contains("t_trans")) {
                    tTransNoIndex = i;
                } else if (headers[i].ToLower().Contains("sec_id") || headers[i].ToLower().Contains("secid") || headers[i].ToLower().Contains("secshort")) {
                    secIdIndex = i;
                }
            }
            int deviationColumnNameIndx = headers.IndexOf("Column Name");
            int masterValueIndx = headers.IndexOf("Master Value");
            int testValueIndx = headers.IndexOf("Test Value");
            if (deviationColumnNameIndx<0 || masterValueIndx <0|| testValueIndx<0) {
                return null;
            }
            List<KnownDefect> knownDefects = new List<KnownDefect>();
            foreach (var item in excelData.Skip(1)) {
                if (item[0] != "" && !IsKnownDefect(item[0])) {
                    KnownDefect knownDefect = new KnownDefect(
                            project: proj,
                            lowerVersion: upgrade[0],
                            upperVersion: upgrade[1],
                            defectNo: item[0].Trim(),
                            masterTransNo: mTransNoIndex == -1 ? "" : item[mTransNoIndex],
                            testTransNo: tTransNoIndex == -1 ? "" : item[tTransNoIndex],
                            secId: secIdIndex == -1 ? "" : item[secIdIndex],
                            deviationColumnName: item[deviationColumnNameIndx],
                            masterValue: item[masterValueIndx],
                            testValue: item[testValueIndx]
                   );
                    if (!knownDefects.Contains(knownDefect)) {
                        knownDefects.Add(knownDefect);
                    }
                }
            }
            return knownDefects;
        }

        private bool IsKnownDefect(string defect) {
            if(defect.Contains("TransMatch")
                || defect.Contains("SecMatch")
                || defect.Contains("ValMatch")
                || defect.Contains("UpgradeMatch")
                || defect.Contains("DeepMatch")) {
                return true;
            }else {
                return false;
            }
        }

        private string ReplaceTags(string defectNo) {
            var cl = defectNo.Split(':');
            string str = "";
            if (cl.Length > 1) {
                str = cl[1];
            }else {
                str = cl[0];
            }
            var defects = str.Split(',').Select(item => item.Trim()).Distinct();
            return string.Join(", ", defects);
        }

    }
}
