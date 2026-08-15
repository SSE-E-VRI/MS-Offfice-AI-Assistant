using System;
using System.Text;
using MistralOfficeAddin.Core;
using Excel = NetOffice.ExcelApi;

namespace MistralOfficeAddin.Hosts
{
    public class ExcelController
    {
        private readonly object _rawAppObj;
        private Excel.Application _excelApp;

        public ExcelController(object appObj)
        {
            _rawAppObj = appObj;
        }

        private Excel.Application GetApp()
        {
            if (_excelApp != null) return _excelApp;
            if (_rawAppObj == null) return null;
            try
            {
                _excelApp = (_rawAppObj is Excel.Application)
                    ? (Excel.Application)_rawAppObj
                    : new Excel.Application(null, _rawAppObj);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.GetApp failed: {0}", ex.Message));
            }
            return _excelApp;
        }

        public string GetSelectedRangeText()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.Selection is Excel.Range)
                {
                    var range = (Excel.Range)app.Selection;
                    var sb = new StringBuilder();
                    object val = range.Value2;
                    if (val is object[,])
                    {
                        var array2D = (object[,])val;
                        int rows = array2D.GetLength(0);
                        int cols = array2D.GetLength(1);
                        for (int r = 1; r <= rows; r++)
                        {
                            var rowVals = new string[cols];
                            for (int c = 1; c <= cols; c++)
                            {
                                rowVals[c - 1] = Convert.ToString(array2D[r, c]);
                            }
                            sb.AppendLine(string.Join("\t", rowVals));
                        }
                        return sb.ToString().TrimEnd();
                    }
                    else if (val != null)
                    {
                        return Convert.ToString(val);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.GetSelectedRangeText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public string GetSelectedRangeValues()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.Selection is Excel.Range)
                {
                    var range = (Excel.Range)app.Selection;
                    var sb = new StringBuilder();
                    sb.AppendLine(string.Format("[Selected Range: {0}]", range.Address));
                    object val = range.Value2;
                    if (val is object[,])
                    {
                        var array2D = (object[,])val;
                        int rows = array2D.GetLength(0);
                        int cols = array2D.GetLength(1);
                        for (int r = 1; r <= rows; r++)
                        {
                            var rowVals = new string[cols];
                            for (int c = 1; c <= cols; c++)
                            {
                                string cell = Convert.ToString(array2D[r, c]) ?? "";
                                if (cell.Contains(",") || cell.Contains("\"") || cell.Contains("\n"))
                                {
                                    cell = "\"" + cell.Replace("\"", "\"\"") + "\"";
                                }
                                rowVals[c - 1] = cell;
                            }
                            sb.AppendLine(string.Join(",", rowVals));
                        }
                        return sb.ToString().TrimEnd();
                    }
                    else if (val != null)
                    {
                        sb.AppendLine(Convert.ToString(val));
                        return sb.ToString().TrimEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.GetSelectedRangeValues failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public void InsertText(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            try
            {
                var app = GetApp();
                Excel.Range targetRange = null;
                if (app != null && app.Selection is Excel.Range)
                {
                    targetRange = (Excel.Range)app.Selection;
                }
                else if (app != null && app.ActiveSheet is Excel.Worksheet)
                {
                    var ws = (Excel.Worksheet)app.ActiveSheet;
                    targetRange = ws.Range("A1");
                }

                if (targetRange != null)
                {
                    targetRange.Value2 = text;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.InsertText failed", ex);
                throw;
            }
        }

        public void WriteFormula(string formula, string cellAddress = null)
        {
            if (string.IsNullOrEmpty(formula)) return;

            try
            {
                var app = GetApp();
                Excel.Range targetRange = null;
                if (!string.IsNullOrEmpty(cellAddress) && app != null && app.ActiveSheet is Excel.Worksheet)
                {
                    var ws = (Excel.Worksheet)app.ActiveSheet;
                    targetRange = ws.Range(cellAddress);
                }
                else if (app != null && app.Selection is Excel.Range)
                {
                    targetRange = (Excel.Range)app.Selection;
                }

                if (targetRange != null)
                {
                    if (!formula.StartsWith("=")) formula = "=" + formula;
                    targetRange.Formula = formula;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.WriteFormula failed", ex);
                throw;
            }
        }

        public string GetCellFormula(string cellAddress)
        {
            try
            {
                var app = GetApp();
                if (!string.IsNullOrEmpty(cellAddress) && app != null && app.ActiveSheet is Excel.Worksheet)
                {
                    var ws = (Excel.Worksheet)app.ActiveSheet;
                    var range = ws.Range(cellAddress);
                    return Convert.ToString(range.Formula);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.GetCellFormula failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public string GetActiveWorkbookName()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveWorkbook != null)
                    return app.ActiveWorkbook.Name;
            }
            catch { }
            return "ExcelWorkbook";
        }
    }
}
