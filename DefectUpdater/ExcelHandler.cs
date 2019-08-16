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

        private List<List<string>> ReadExcel(string path, string proj, string upgrade) {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path.Trim('\r','\n'));
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

        public List<KnownDefect> GetDeviationsFromExcel(string path, string proj, string upgrade) {
            var excelData = ReadExcel(path, proj, upgrade);
            if (excelData.Count <= 1) {
                return null;
            }         
            var headers = excelData.First();
            int mTransNoIndex = -1;
            int tTransNoIndex = -1;
            int secIdIndex = -1;
            for (int i = 0; i < headers.Count; i++) {
                if (headers[i].ToLower().Contains("m_trans")){
                    mTransNoIndex = i;
                }else if (headers[i].ToLower().Contains("t_trans")) {
                    tTransNoIndex = i;
                } else if (headers[i].ToLower().Contains("sec_id") || headers[i].ToLower().Contains("secid") || headers[i].ToLower().Contains("secshort")) {
                    secIdIndex = i;
                }
            }
            List<KnownDefect> knownDefects = new List<KnownDefect>();
            foreach (var item in excelData.Skip(1)) {
                if (item[0] != "" && !item[0].Contains("?")) {
                    KnownDefect knownDefect = new KnownDefect(
                            project: proj,
                            upgrade: upgrade,
                            defectNo: ReplaceTags(item[0]),
                            masterTransNo: mTransNoIndex == -1 ? "" : item[mTransNoIndex],
                            testTransNo: tTransNoIndex == -1 ? "" : item[tTransNoIndex],
                            secId: secIdIndex == -1 ? "" : item[secIdIndex],
                            deviationColumnName: item[headers.Count - 3],
                            masterValue: item[headers.Count - 2],
                            testValue: item[headers.Count - 1]
                   );
                    if (!knownDefects.Contains(knownDefect)) {
                        knownDefects.Add(knownDefect);
                    }
                }
            }
            return knownDefects;
        }

        private string ReplaceTags(string defectNo) {
            return defectNo.Replace("TransMatch: ", "")
                .Replace("SecMatch: ", "")
                .Replace("ValMatch: ", "")
                .Replace("UpgradeMatch: ", "")
                .Replace("DeepMatch: ", "");
        }

        
    }

   
    
}
