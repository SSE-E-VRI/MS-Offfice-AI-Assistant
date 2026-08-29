using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Hosts
{
    public class ExcelController : IOfficeHostController
    {
        private readonly object _rawAppObj;
        private static readonly Regex ExcelNameRegex = new Regex(@"^[A-Za-z_][A-Za-z0-9_]{0,249}$", RegexOptions.Compiled);

        public string HostType
        {
            get { return "Excel"; }
        }

        public object GetRawAppObj()
        {
            return _rawAppObj;
        }

        public ExcelController(object appObj)
        {
            _rawAppObj = appObj;
        }

        public string GetActiveDocumentName()
        {
            return GetActiveWorkbookName();
        }

        public string GetActiveWorkbookPath()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app != null && app.ActiveWorkbook != null)
                {
                    string path = null;
                    try { path = Convert.ToString(app.ActiveWorkbook.FullName); } catch { }
                    if (!string.IsNullOrWhiteSpace(path)) return path;
                    try { path = Convert.ToString(app.ActiveWorkbook.Path); } catch { }
                    return path ?? string.Empty;
                }
            }
            catch { }
            return string.Empty;
        }

        public string GetSelectedText()
        {
            return GetSelectedRangeValues();
        }

        public string GetDocumentContext(string prompt, int maxCharacters)
        {
            string snapshot = GetWorksheetSnapshot(70, 26);
            if (maxCharacters > 0 && snapshot.Length > maxCharacters)
            {
                return snapshot.Substring(0, maxCharacters);
            }
            return snapshot;
        }

        public bool Undo()
        {
            return UndoLastAction();
        }

        public static string IndexToColumnLetter(int colIndex)
        {
            return SpreadsheetActionParser.IndexToColumnLetter(colIndex);
        }

        public string GetSelectedRangeValues()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app != null && app.Selection != null)
                {
                    dynamic selection = app.Selection;
                    int cellCount = 1;
                    try { cellCount = Convert.ToInt32(selection.Count); } catch { }

                    string address = Convert.ToString(selection.Address);
                    object val = selection.Value2;
                    if (val is object[,])
                    {
                        var array2D = (object[,])val;
                        int rows = array2D.GetLength(0);
                        int cols = array2D.GetLength(1);

                        bool hasData = false;
                        var sb = new StringBuilder();
                        sb.AppendLine(string.Format("[Selected Range: {0}]", address));

                        int startCol = 1;
                        try { startCol = Convert.ToInt32(selection.Column); } catch { }

                        for (int r = 1; r <= rows; r++)
                        {
                            var rowVals = new string[cols];
                            for (int c = 1; c <= cols; c++)
                            {
                                string colLet = IndexToColumnLetter(startCol + c - 1);
                                string cell = Convert.ToString(array2D[r, c]) ?? "";
                                if (!string.IsNullOrWhiteSpace(cell)) hasData = true;
                                if (cell.Contains(",") || cell.Contains("\"") || cell.Contains("\n"))
                                {
                                    cell = "\"" + cell.Replace("\"", "\"\"") + "\"";
                                }
                                rowVals[c - 1] = string.Format("{0}={1}", colLet, cell);
                            }
                            sb.AppendLine(string.Join(", ", rowVals));
                        }

                        if (hasData)
                        {
                            return sb.ToString().TrimEnd();
                        }
                    }
                    else if (val != null)
                    {
                        string s = Convert.ToString(val).Trim();
                        if (!string.IsNullOrWhiteSpace(s))
                        {
                            return string.Format("[Selected Cell {0}]: {1}", address, s);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.GetSelectedRangeValues failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public string GetWorksheetSnapshot(int maxRows = 70, int maxCols = 26)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return string.Empty;

                dynamic ws = null;
                try { ws = app.ActiveSheet; } catch { }
                if (ws == null) return string.Empty;

                string sheetName = Convert.ToString(ws.Name);
                string activeCellAddr = "A1";
                string selAddr = "A1";

                try
                {
                    if (app.ActiveCell != null)
                        activeCellAddr = Convert.ToString(app.ActiveCell.Address);
                    if (app.Selection != null)
                        selAddr = Convert.ToString(app.Selection.Address);
                }
                catch { }

                dynamic usedRange = ws.UsedRange;
                if (usedRange == null)
                {
                    return string.Format("[Worksheet: '{0}', ActiveCell: {1}, Selection: {2} - Empty Sheet]", sheetName, activeCellAddr, selAddr);
                }

                object val = usedRange.Value2;
                object formulaVal = null;
                try { formulaVal = usedRange.Formula; } catch { }

                if (val is object[,])
                {
                    var array2D = (object[,])val;
                    var formula2D = formulaVal as object[,];

                    int totalRows = array2D.GetLength(0);
                    int totalCols = array2D.GetLength(1);

                    int startRow = 1;
                    int startCol = 1;
                    try { startRow = Convert.ToInt32(usedRange.Row); } catch { }
                    try { startCol = Convert.ToInt32(usedRange.Column); } catch { }

                    int rows = Math.Min(totalRows, maxRows);
                    int cols = Math.Min(totalCols, maxCols);

                    var sb = new StringBuilder();
                    sb.AppendLine(string.Format("[Excel Context | Worksheet: '{0}' | ActiveCell: {1} | Selection: {2} | UsedRange: {3}]",
                        sheetName, activeCellAddr, selAddr, Convert.ToString(usedRange.Address)));

                    // 1. Header mapping: Explicit Column Letters with Header names
                    var headers = new List<string>();
                    for (int c = 1; c <= cols; c++)
                    {
                        string colLet = IndexToColumnLetter(startCol + c - 1);
                        string headerText = Convert.ToString(array2D[1, c]) ?? "";
                        if (string.IsNullOrWhiteSpace(headerText))
                            headerText = "(empty)";
                        headers.Add(string.Format("Col {0} [{1}]", colLet, headerText.Trim()));
                    }
                    sb.AppendLine("Columns: " + string.Join(" | ", headers));
                    sb.AppendLine("---");

                    // 2. Data rows with cell coordinates and values
                    bool hasData = false;
                    for (int r = 1; r <= rows; r++)
                    {
                        int actualRow = startRow + r - 1;
                        var rowItems = new List<string>();
                        bool rowHasContent = false;

                        for (int c = 1; c <= cols; c++)
                        {
                            string colLet = IndexToColumnLetter(startCol + c - 1);
                            string cellAddr = string.Format("{0}{1}", colLet, actualRow);
                            string cellVal = Convert.ToString(array2D[r, c]) ?? "";
                            string cellForm = formula2D != null ? Convert.ToString(formula2D[r, c]) : null;

                            if (!string.IsNullOrWhiteSpace(cellVal))
                            {
                                rowHasContent = true;
                                hasData = true;
                                if (!string.IsNullOrEmpty(cellForm) && cellForm.StartsWith("=") && cellForm != cellVal)
                                {
                                    rowItems.Add(string.Format("{0}={1} (formula: {2})", cellAddr, cellVal.Trim(), cellForm.Trim()));
                                }
                                else
                                {
                                    rowItems.Add(string.Format("{0}={1}", cellAddr, cellVal.Trim()));
                                }
                            }
                        }

                        if (rowHasContent || r == 1)
                        {
                            sb.AppendLine(string.Format("Row {0}: {1}", actualRow, string.Join(", ", rowItems)));
                        }
                    }

                    if (hasData)
                    {
                        // Workbook-wide overview (multi-sheet context + table enumeration) — bounded, best-effort
                        try
                        {
                            dynamic wb = app.ActiveWorkbook;
                            if (wb != null && wb.Worksheets != null)
                            {
                                int wsCount = 0;
                                try { wsCount = Convert.ToInt32(wb.Worksheets.Count); } catch { wsCount = 0; }
                                if (wsCount > 1)
                                {
                                    sb.AppendLine("--- Workbook Overview ---");
                                    int capSheets = Math.Min(wsCount, 10);
                                    for (int wi = 1; wi <= capSheets; wi++)
                                    {
                                        try
                                        {
                                            dynamic sh = wb.Worksheets[wi];
                                            string shName = Convert.ToString(sh.Name);
                                            string shAddr = "";
                                            try { shAddr = Convert.ToString(sh.UsedRange.Address); } catch { shAddr = "A1"; }
                                            int shRows = 0, shCols = 0;
                                            try { shRows = Convert.ToInt32(sh.UsedRange.Rows.Count); } catch { }
                                            try { shCols = Convert.ToInt32(sh.UsedRange.Columns.Count); } catch { }
                                            sb.AppendLine(string.Format("Sheet '{0}': {1} ({2} rows x {3} cols)", shName, shAddr, shRows, shCols));
                                            // Table enumeration per sheet
                                            try
                                            {
                                                dynamic los = sh.ListObjects;
                                                int loCount = 0;
                                                try { loCount = Convert.ToInt32(los.Count); } catch { loCount = 0; }
                                                for (int li = 1; li <= Math.Min(loCount, 5); li++)
                                                {
                                                    try
                                                    {
                                                        dynamic lo = los[li];
                                                        string loName = Convert.ToString(lo.Name);
                                                        string loRange = Convert.ToString(lo.Range.Address);
                                                        // Headers
                                                        var hdrVals = new List<string>();
                                                        try
                                                        {
                                                            dynamic hdrRow = lo.HeaderRowRange;
                                                            if (hdrRow != null)
                                                            {
                                                                object hv = hdrRow.Value2;
                                                                if (hv is object[,])
                                                                {
                                                                    var harr = (object[,])hv;
                                                                    for (int hc = 1; hc <= harr.GetLength(1); hc++)
                                                                    {
                                                                        string hvStr = Convert.ToString(harr[1, hc]) ?? "";
                                                                        if (!string.IsNullOrWhiteSpace(hvStr)) hdrVals.Add(hvStr.Trim());
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        catch { }
                                                        if (hdrVals.Count > 0)
                                                            sb.AppendLine(string.Format("  Table '{0}' {1} headers: {2}", loName, loRange, string.Join(", ", hdrVals.ToArray())));
                                                        else
                                                            sb.AppendLine(string.Format("  Table '{0}' {1}", loName, loRange));
                                                    }
                                                    catch { }
                                                }
                                            }
                                            catch { }
                                        }
                                        catch { }
                                    }
                                    if (wsCount > capSheets) sb.AppendLine(string.Format("  ... and {0} more sheets", wsCount - capSheets));
                                }
                                else
                                {
                                    // Single sheet: still enumerate tables for context
                                    try
                                    {
                                        dynamic sh = ws;
                                        dynamic los = sh.ListObjects;
                                        int loCount = 0;
                                        try { loCount = Convert.ToInt32(los.Count); } catch { loCount = 0; }
                                        for (int li = 1; li <= Math.Min(loCount, 5); li++)
                                        {
                                            try
                                            {
                                                dynamic lo = los[li];
                                                string loName = Convert.ToString(lo.Name);
                                                string loRange = Convert.ToString(lo.Range.Address);
                                                sb.AppendLine(string.Format("Table '{0}' {1}", loName, loRange));
                                            }
                                            catch { }
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { }
                        return sb.ToString().TrimEnd();
                    }
                }
                else if (val != null)
                {
                    return string.Format("[Worksheet: '{0}', ActiveCell: {1}, Selection: {2}]: {3}",
                        sheetName, activeCellAddr, selAddr, Convert.ToString(val));
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.GetWorksheetSnapshot failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public string GetContextReadout()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return string.Empty;

                dynamic ws = null;
                try { ws = app.ActiveSheet; } catch { }
                if (ws == null) return string.Empty;

                string sheetName = Convert.ToString(ws.Name);
                string selAddr = "A1";
                try
                {
                    if (app.Selection != null)
                        selAddr = Convert.ToString(app.Selection.Address);
                }
                catch { }

                return string.Format("{0}!{1}", sheetName, selAddr);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.GetContextReadout failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        private HostOperationResult ResolveTargetRange(string targetAddress, out dynamic app, out dynamic ws, out dynamic targetRange)
        {
            app = null;
            ws = null;
            targetRange = null;

            if (string.IsNullOrEmpty(targetAddress) || !SpreadsheetActionParser.IsSafeTarget(targetAddress))
                return HostOperationResult.Failed(string.Format("Invalid or unsafe cell/range address: '{0}'.", targetAddress), 0, targetAddress);

            string sheetName = null;
            string rangePart = null;
            SpreadsheetActionParser.TryParseSheetQualifiedTarget(targetAddress, out sheetName, out rangePart);

            try
            {
                app = _rawAppObj;
                if (app == null)
                    return HostOperationResult.Failed("No active Excel workbook found.", 0, targetAddress);

                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    try { ws = app.Worksheets[sheetName]; } catch { }
                    if (ws == null)
                        return HostOperationResult.Failed(string.Format("Worksheet '{0}' not found.", sheetName), 0, targetAddress);
                }
                else
                {
                    try { ws = app.ActiveSheet; } catch { }
                }
                if (ws == null)
                    return HostOperationResult.Failed("No active Excel worksheet found.", 0, targetAddress);

                EnsureWorksheetEditable(ws);

                string effectiveRange = !string.IsNullOrWhiteSpace(rangePart) ? rangePart : targetAddress;
                try { targetRange = ws.Range(effectiveRange); } catch { }
                if (targetRange == null)
                    return HostOperationResult.Failed(string.Format("Could not resolve target range: '{0}'.", targetAddress), 0, targetAddress);

                return HostOperationResult.Ok(null, targetAddress);
            }
            catch (Exception ex)
            {
                return HostOperationResult.FromException(ex, "ExcelController.ResolveTargetRange", targetAddress);
            }
        }

        public HostOperationResult ExecuteSpreadsheetAction(SpreadsheetAction action)
        {
            if (action == null)
                return HostOperationResult.Failed("Spreadsheet action cannot be null.");

            var oa = MSOfficeAIAssistant.Core.Actions.OfficeAction.FromSpreadsheetAction(action);
            var res = ToolRegistry.Execute(this, oa);
            if (res.Success)
            {
                action.Status = SpreadsheetActionStatus.Applied;
                action.ResultText = res.Value != null ? Convert.ToString(res.Value) : "Applied successfully";
                action.ErrorMessage = null;
                Logger.Info(string.Format("Applied {0} to {1}: {2}", action.Type, action.Target, action.ResultText));
            }
            else
            {
                action.Status = SpreadsheetActionStatus.Error;
                action.ErrorMessage = res.ErrorMessage;
                Logger.Error(string.Format("Failed to apply {0} to {1}: {2}", action.Type, action.Target, res.ErrorMessage));
            }
            return res;
        }

        public HostOperationResult ExecuteWriteFormula(string formula, string cellAddress)
        {
            if (string.IsNullOrEmpty(formula))
                return HostOperationResult.Failed("Formula cannot be empty.", 0, cellAddress);

            dynamic app, ws, targetRange;
            var resolveRes = ResolveTargetRange(cellAddress, out app, out ws, out targetRange);
            if (!resolveRes.Success) return resolveRes;

            try
            {
                if (!formula.StartsWith("=")) formula = "=" + formula;
                targetRange.Formula = formula;
                try { ExcelChangeHighlighter.ApplyHighlight(ws, targetRange); } catch { }
                string res = ReadRangeResult(targetRange);
                return HostOperationResult.Ok(res, cellAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteWriteFormula failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteWriteFormula", cellAddress);
            }
        }

        public HostOperationResult ExecuteWriteValue(string value, string cellAddress)
        {
            dynamic app, ws, targetRange;
            var resolveRes = ResolveTargetRange(cellAddress, out app, out ws, out targetRange);
            if (!resolveRes.Success) return resolveRes;

            try
            {
                targetRange.Value2 = value ?? string.Empty;
                try { ExcelChangeHighlighter.ApplyHighlight(ws, targetRange); } catch { }
                return HostOperationResult.Ok("Value written", cellAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteWriteValue failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteWriteValue", cellAddress);
            }
        }

        public HostOperationResult ExecuteFillDown(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                ApplyFillDown(targetRange, content);
                return HostOperationResult.Ok(string.Format("Filled {0}", targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteFillDown failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteFillDown", targetAddress);
            }
        }

        public HostOperationResult ExecuteTable(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                WriteTable(ws, targetRange, content);
                return HostOperationResult.Ok("Table written", targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteTable failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteTable", targetAddress);
            }
        }

        public HostOperationResult ExecuteCreateTable(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = CreateExcelTable(ws, targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteCreateTable failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteCreateTable", targetAddress);
            }
        }

        public HostOperationResult ExecuteConditionalFormat(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = ApplyConditionalFormatting(targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteConditionalFormat failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteConditionalFormat", targetAddress);
            }
        }

        public HostOperationResult ExecuteSort(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = ApplySort(ws, targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteSort failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteSort", targetAddress);
            }
        }

        public HostOperationResult ExecuteFilter(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = ApplyFilter(targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteFilter failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteFilter", targetAddress);
            }
        }

        public HostOperationResult ExecuteDataValidation(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = ApplyDataValidation(targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteDataValidation failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteDataValidation", targetAddress);
            }
        }

        public HostOperationResult ExecuteCreateChart(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = CreateChart(ws, targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteCreateChart failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteCreateChart", targetAddress);
            }
        }

        public HostOperationResult ExecuteCreatePivotTable(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = CreatePivotTable(app, ws, targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteCreatePivotTable failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteCreatePivotTable", targetAddress);
            }
        }

        public HostOperationResult ExecuteNamedRange(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = CreateNamedRange(app, ws, targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteNamedRange failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteNamedRange", targetAddress);
            }
        }

        public HostOperationResult ExecuteRemoveDuplicates(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                string r = RemoveDuplicates(targetRange, content);
                return HostOperationResult.Ok(r, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteRemoveDuplicates failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteRemoveDuplicates", targetAddress);
            }
        }

        // D-13 Tier 2: mutation methods that lacked a structured HostOperationResult wrapper.
        public HostOperationResult ExecuteUndoLastAction()
        {
            try
            {
                bool ok = UndoLastAction();
                return ok ? HostOperationResult.Ok("Undid the last Excel action.") : HostOperationResult.Failed("UndoLastAction returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteUndoLastAction failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteUndoLastAction");
            }
        }

        public HostOperationResult ExecuteInsertText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return HostOperationResult.Failed("Text cannot be empty.");

            try
            {
                bool ok = InsertText(text);
                return ok ? HostOperationResult.Ok("Inserted text into the active worksheet.") : HostOperationResult.Failed("InsertText returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteInsertText failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteInsertText");
            }
        }

        public const int MaxBeforeStateSnapshotCells = 5000;

        public HostOperationResult CaptureRangeFormulas(string targetAddress, out object capturedFormulas)
        {
            capturedFormulas = null;
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                int cellCount = 1;
                try { cellCount = Convert.ToInt32(targetRange.Count); } catch { }

                if (cellCount > MaxBeforeStateSnapshotCells)
                {
                    return HostOperationResult.Failed(
                        string.Format("Target range ({0} cells) exceeds snapshot capacity limit of {1} cells.", cellCount, MaxBeforeStateSnapshotCells),
                        0, targetAddress);
                }

                object formulaVal = targetRange.Formula;
                if (formulaVal is object[,])
                {
                    object[,] arr = (object[,])formulaVal;
                    int rows = arr.GetLength(0);
                    int cols = arr.GetLength(1);
                    var list2D = new List<List<string>>();
                    for (int r = 1; r <= rows; r++)
                    {
                        var rowList = new List<string>();
                        for (int c = 1; c <= cols; c++)
                        {
                            rowList.Add(Convert.ToString(arr[r, c]));
                        }
                        list2D.Add(rowList);
                    }
                    capturedFormulas = list2D;
                }
                else
                {
                    capturedFormulas = Convert.ToString(formulaVal);
                }

                return HostOperationResult.Ok(capturedFormulas, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.CaptureRangeFormulas failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.CaptureRangeFormulas", targetAddress);
            }
        }

        public HostOperationResult RestoreRangeFormulas(string targetAddress, object formulas)
        {
            if (formulas == null)
                return HostOperationResult.Failed("Formulas to restore cannot be null.", 0, targetAddress);

            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;

            try
            {
                if (formulas is List<List<string>>)
                {
                    var list2D = (List<List<string>>)formulas;
                    int rows = list2D.Count;
                    int cols = rows > 0 ? list2D[0].Count : 0;
                    if (rows > 0 && cols > 0)
                    {
                        object[,] arr = new object[rows, cols];
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                arr[r, c] = list2D[r][c];
                            }
                        }
                        targetRange.Formula = arr;
                    }
                }
                else if (formulas is Newtonsoft.Json.Linq.JArray)
                {
                    var jarr = (Newtonsoft.Json.Linq.JArray)formulas;
                    int rows = jarr.Count;
                    if (rows > 0 && jarr[0] is Newtonsoft.Json.Linq.JArray)
                    {
                        int cols = ((Newtonsoft.Json.Linq.JArray)jarr[0]).Count;
                        object[,] arr = new object[rows, cols];
                        for (int r = 0; r < rows; r++)
                        {
                            var rowJ = (Newtonsoft.Json.Linq.JArray)jarr[r];
                            for (int c = 0; c < cols; c++)
                            {
                                arr[r, c] = Convert.ToString(rowJ[c]);
                            }
                        }
                        targetRange.Formula = arr;
                    }
                    else
                    {
                        targetRange.Formula = Convert.ToString(formulas);
                    }
                }
                else
                {
                    targetRange.Formula = Convert.ToString(formulas);
                }

                return HostOperationResult.Ok("Restored formulas for " + targetAddress, targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.RestoreRangeFormulas failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.RestoreRangeFormulas", targetAddress);
            }
        }

        public bool UndoLastAction()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return false;
                app.Undo();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.UndoLastAction failed: {0}", ex.Message));
                return false;
            }
        }

        private static string EnsureFormula(string content)
        {
            string formula = (content ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(formula)) throw new InvalidOperationException("A formula is required.");
            return formula.StartsWith("=") ? formula : "=" + formula;
        }

        private static void EnsureWorksheetEditable(dynamic worksheet)
        {
            try
            {
                if (worksheet != null && Convert.ToBoolean(worksheet.ProtectContents))
                    throw new InvalidOperationException("The active worksheet is protected. Unprotect it before applying AI changes.");
            }
            catch (InvalidOperationException) { throw; }
            catch { }
        }

        private static int GetRangeRows(dynamic range)
        {
            try { return Convert.ToInt32(range.Rows.Count); } catch { return 0; }
        }

        private static int GetRangeColumns(dynamic range)
        {
            try { return Convert.ToInt32(range.Columns.Count); } catch { return 0; }
        }

        private static string ReadRangeResult(dynamic range)
        {
            try
            {
                string result = Convert.ToString(range.Text);
                if (!string.IsNullOrWhiteSpace(result)) return result;
                result = Convert.ToString(range.Value2);
                return string.IsNullOrWhiteSpace(result) ? "Formula applied" : result;
            }
            catch { return "Formula applied"; }
        }

        private static void ApplyFillDown(dynamic targetRange, string content)
        {
            if (GetRangeRows(targetRange) < 1) throw new InvalidOperationException("Fill-down requires a valid range.");
            dynamic firstRow = targetRange.Rows[1];
            firstRow.Formula = EnsureFormula(content);
            if (GetRangeRows(targetRange) > 1) targetRange.FillDown();
        }

        private void WriteTable(dynamic worksheet, dynamic targetRange, string content)
        {
            var tableRows = ParseMarkdownOrCsvTable(content);
            if (tableRows.Count == 0) throw new InvalidOperationException("The table action did not contain a Markdown table.");

            int startRow = Convert.ToInt32(targetRange.Row);
            int startCol = Convert.ToInt32(targetRange.Column);
            int targetRows = GetRangeRows(targetRange);
            int targetCols = GetRangeColumns(targetRange);

            // If target is a single cell (1x1), reject multi-cell table writes without explicit multi-cell range
            if (targetRows == 1 && targetCols == 1 && (tableRows.Count > 1 || (tableRows.Count > 0 && tableRows[0].Count > 1)))
            {
                int endRowAdvised = startRow + tableRows.Count - 1;
                int endColAdvised = startCol + tableRows[0].Count - 1;
                throw new InvalidOperationException(string.Format(
                    "Target is a single cell, but the table contains {0} rows × {1} columns. Declare a bounded multi-cell range (for example {2}) to apply this table.",
                    tableRows.Count,
                    tableRows[0].Count,
                    SpreadsheetActionParser.BuildRangeAddress(startCol, startRow, endColAdvised, endRowAdvised)));
            }

            // If a multi-cell range was explicitly declared (e.g. A1:B2), clip write to that extent.
            bool isExplicitBoundedRange = targetRows > 1 || targetCols > 1;
            int maxRowsToWrite = isExplicitBoundedRange ? Math.Min(tableRows.Count, targetRows) : tableRows.Count;
            int maxColsToWrite = isExplicitBoundedRange ? targetCols : 0;
            if (!isExplicitBoundedRange)
            {
                for (int r = 0; r < maxRowsToWrite; r++)
                {
                    if (tableRows[r].Count > maxColsToWrite) maxColsToWrite = tableRows[r].Count;
                }
            }

            int endRow = startRow + maxRowsToWrite - 1;
            int endCol = startCol + Math.Max(1, maxColsToWrite) - 1;

            if (!SpreadsheetActionParser.IsSafeRangeBounds(startCol, startRow, endCol, endRow))
            {
                throw new InvalidOperationException(string.Format(
                    "The generated table extent ({0}) exceeds Excel safe bounds or cell count limits.",
                    SpreadsheetActionParser.BuildRangeAddress(startCol, startRow, endCol, endRow)));
            }

            int totalCells = 0;
            for (int r = 0; r < maxRowsToWrite; r++)
            {
                var row = tableRows[r];
                int colsInRow = isExplicitBoundedRange ? Math.Min(row.Count, targetCols) : row.Count;
                for (int c = 0; c < colsInRow; c++)
                {
                    totalCells++;
                    if (totalCells > 100000) throw new InvalidOperationException("The generated table is too large to apply safely.");
                    worksheet.Cells[startRow + r, startCol + c].Value2 = row[c];
                }
            }
        }

        private static string CreateExcelTable(dynamic worksheet, dynamic targetRange, string content)
        {
            int rows = GetRangeRows(targetRange);
            int columns = GetRangeColumns(targetRange);
            if (rows < 2 || columns < 1) throw new InvalidOperationException("An Excel Table requires a header row and at least one data row.");
            EnsureUniqueNonBlankHeaders(targetRange, columns);

            dynamic tables = worksheet.ListObjects;
            dynamic table = tables.Add(1, targetRange, Type.Missing, 1, Type.Missing, Type.Missing);
            string requestedName = GetNamedOption(content, "name");
            if (!string.IsNullOrWhiteSpace(requestedName))
            {
                if (!IsSafeExcelName(requestedName)) throw new InvalidOperationException("The requested Excel Table name is invalid.");
                table.Name = requestedName;
            }
            return string.IsNullOrWhiteSpace(requestedName) ? "Excel Table created" : "Excel Table created: " + requestedName;
        }

        private static void EnsureUniqueNonBlankHeaders(dynamic targetRange, int columns)
        {
            var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int column = 1; column <= columns; column++)
            {
                string header = Convert.ToString(targetRange.Cells[1, column].Value2);
                if (string.IsNullOrWhiteSpace(header)) throw new InvalidOperationException("Excel Table headers cannot be blank.");
                if (!headers.Add(header.Trim())) throw new InvalidOperationException("Excel Table headers must be unique.");
            }
        }

        private static string ApplyConditionalFormatting(dynamic targetRange, string content)
        {
            string rule = (content ?? string.Empty).Trim().ToLowerInvariant();
            dynamic formats = targetRange.FormatConditions;
            if (rule == "" || rule == "color_scale" || rule == "colorscale")
            {
                formats.AddColorScale(3);
                return "Three-color scale added";
            }
            if (rule == "data_bar" || rule == "databar")
            {
                formats.AddDatabar();
                return "Data bar added";
            }

            // Optional custom highlight color (palette like apply_theme or hex)
            int customColor = 13551615; // default pink
            string colorOpt = GetNamedOption(content, "color");
            if (string.IsNullOrWhiteSpace(colorOpt)) colorOpt = GetNamedOption(content, "fill");
            int parsedCol;
            if (!string.IsNullOrWhiteSpace(colorOpt) && TryParseColor(colorOpt, out parsedCol)) customColor = parsedCol;

            // Extended rules: top_n, duplicates, contains / text_contains, between
            string topN = GetValueAfterPrefix(content, "top_n");
            if (!string.IsNullOrWhiteSpace(topN))
            {
                int n;
                if (!int.TryParse(topN.Trim(), out n) || n < 1 || n > 1000)
                    throw new InvalidOperationException("top_n requires a number between 1 and 1000.");
                dynamic cond = formats.AddTop10();
                try { cond.Top10.Rank = n; cond.Top10.Percent = false; } catch { }
                try { cond.Interior.Color = customColor; } catch { }
                return string.Format("Top {0} conditional formatting added", n);
            }

            string dup = GetValueAfterPrefix(content, "duplicates");
            if (!string.IsNullOrWhiteSpace(dup) || string.Equals(rule, "duplicates", StringComparison.OrdinalIgnoreCase))
            {
                dynamic cond = null;
                try { cond = formats.AddUniqueValues(); } catch { }
                if (cond != null)
                {
                    try { cond.DupeUnique = 0; cond.Interior.Color = customColor; } catch { }
                    return "Duplicate values highlighted";
                }
                throw new InvalidOperationException("Duplicate highlighting not supported on this Excel version.");
            }

            string contains = GetValueAfterPrefix(content, "contains");
            if (string.IsNullOrWhiteSpace(contains)) contains = GetValueAfterPrefix(content, "text_contains");
            if (!string.IsNullOrWhiteSpace(contains))
            {
                // 9 = xlTextString; TextOperator 3 = xlContains
                dynamic cond = formats.Add(9, Type.Missing, contains);
                try { cond.TextOperator = 3; } catch { } // 3 = xlContains
                try { cond.Interior.Color = customColor; } catch { }
                return string.Format("Text contains '{0}' highlighted", contains);
            }

            string between = GetValueAfterPrefix(content, "between");
            if (!string.IsNullOrWhiteSpace(between))
            {
                string[] parts = between.Split(',');
                if (parts.Length != 2) throw new InvalidOperationException("between requires min,max (e.g. between:10,100).");
                dynamic cond = formats.Add(1, 1, parts[0].Trim(), parts[1].Trim()); // 1=xlCellValue, 1=xlBetween
                try { cond.Interior.Color = customColor; } catch { }
                return string.Format("Between {0} and {1} highlighted", parts[0].Trim(), parts[1].Trim());
            }

            string iconSet = GetValueAfterPrefix(content, "icon_set");
            if (!string.IsNullOrWhiteSpace(iconSet) || string.Equals(rule, "icon_set", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    string setName = !string.IsNullOrWhiteSpace(iconSet) ? iconSet.Trim().ToLowerInvariant() : "3arrows";
                    dynamic cond = formats.AddIconSetCondition();
                    // IconSet IDs: 4=3Arrows, 1=3TrafficLights, etc. Map common names
                    int setId = 4;
                    if (setName == "3trafficlights") setId = 1;
                    else if (setName == "3arrows") setId = 4;
                    else if (setName == "3flags") setId = 7;
                    try { cond.IconSet.Id = setId; } catch { }
                    return "Icon set conditional formatting added";
                }
                catch (Exception ex) { throw new InvalidOperationException("Icon set not supported: " + ex.Message); }
            }

            string criterion = GetValueAfterPrefix(content, "greater_than");
            int comparison = 5; // xlGreater
            if (string.IsNullOrWhiteSpace(criterion))
            {
                criterion = GetValueAfterPrefix(content, "less_than");
                comparison = 6; // xlLess
            }
            if (string.IsNullOrWhiteSpace(criterion))
            {
                criterion = GetValueAfterPrefix(content, "equal_to");
                comparison = 3; // xlEqual
            }
            if (string.IsNullOrWhiteSpace(criterion))
                throw new InvalidOperationException("Use color_scale, data_bar, greater_than:value, less_than:value, equal_to:value, top_n:N, duplicates, contains:text, between:min,max, or icon_set for conditional formatting.");

            dynamic condition = formats.Add(1, comparison, criterion);
            try { condition.Interior.Color = customColor; }
            catch (Exception ex) { Logger.Warn(string.Format("Conditional format fill color not applied: {0}", ex.Message)); }
            return "Conditional formatting rule added";
        }

        private static string ApplySort(dynamic worksheet, dynamic targetRange, string content)
        {
            int rows = GetRangeRows(targetRange);
            int columns = GetRangeColumns(targetRange);
            if (rows < 2 || columns < 1) throw new InvalidOperationException("Sort requires a header row and at least one data row.");

            int field = GetIntegerOption(content, "field", 1);
            if (field < 1 || field > columns) throw new InvalidOperationException("The requested sort field is outside the target range.");
            bool descending = string.Equals(GetNamedOption(content, "direction"), "descending", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals((content ?? string.Empty).Trim(), "descending", StringComparison.OrdinalIgnoreCase);
            dynamic sort = worksheet.Sort;
            sort.SortFields.Clear();
            sort.SortFields.Add(targetRange.Columns[field], 0, descending ? 2 : 1, Type.Missing, 0);
            sort.SetRange(targetRange);
            sort.Header = 1; // xlYes
            sort.MatchCase = false;
            sort.Orientation = 1; // xlTopToBottom
            sort.Apply();
            return descending ? "Sorted descending" : "Sorted ascending";
        }

        private static string ApplyFilter(dynamic targetRange, string content)
        {
            int rows = GetRangeRows(targetRange);
            int columns = GetRangeColumns(targetRange);
            if (rows < 2 || columns < 1) throw new InvalidOperationException("Filter requires a header row and at least one data row.");

            int field = GetIntegerOption(content, "field", 1);
            string criteria = GetNamedOption(content, "criteria");
            if (field < 1 || field > columns || string.IsNullOrWhiteSpace(criteria))
                throw new InvalidOperationException("Filter requires a valid field number and criteria (for example field:2;criteria:Open).");
            targetRange.AutoFilter(field, criteria, Type.Missing, Type.Missing, true);
            return string.Format("Filter applied to field {0}", field);
        }

        private static string ApplyDataValidation(dynamic targetRange, string content)
        {
            string specification = (content ?? string.Empty).Trim();
            string listValues = GetValueAfterPrefix(specification, "list");
            dynamic validation = targetRange.Validation;
            validation.Delete();
            if (!string.IsNullOrWhiteSpace(listValues))
            {
                if (listValues.Length > 255) throw new InvalidOperationException("An inline validation list cannot exceed 255 characters.");
                validation.Add(3, 1, 1, listValues, Type.Missing); // list, stop, between
                return "List validation applied";
            }

            string wholeNumbers = GetValueAfterPrefix(specification, "whole_number");
            string decimals = GetValueAfterPrefix(specification, "decimal");
            string dates = GetValueAfterPrefix(specification, "date");
            string between = GetValueAfterPrefix(specification, "between");
            string custom = GetValueAfterPrefix(specification, "custom");
            if (!string.IsNullOrWhiteSpace(custom))
            {
                // Custom formula validation: type 7 = xlValidateCustom
                validation.Add(7, 1, 1, custom, Type.Missing);
                return "Custom data validation applied";
            }
            string rangeValues = !string.IsNullOrWhiteSpace(wholeNumbers) ? wholeNumbers : (!string.IsNullOrWhiteSpace(decimals) ? decimals : (!string.IsNullOrWhiteSpace(dates) ? dates : between));
            if (string.IsNullOrWhiteSpace(rangeValues))
                throw new InvalidOperationException("Use list:values, whole_number:min,max, decimal:min,max, date:start,end, between:min,max, or custom:formula for data validation.");
            if (!string.IsNullOrWhiteSpace(between))
            {
                string[] bounds = between.Split(',');
                if (bounds.Length != 2) throw new InvalidOperationException("Between validation requires a minimum and maximum value.");
                // Use whole-number validation only if both bounds are integers; otherwise decimal.
                int minInt, maxInt;
                bool bothWhole = int.TryParse(bounds[0].Trim(), out minInt) && int.TryParse(bounds[1].Trim(), out maxInt);
                int vType = bothWhole ? 1 : 2; // 1 = xlValidateWholeNumber, 2 = xlValidateDecimal
                validation.Add(vType, 1, 1, bounds[0].Trim(), bounds[1].Trim());
                return "Data validation (between) applied";
            }
            string[] bounds2 = rangeValues.Split(',');
            if (bounds2.Length != 2) throw new InvalidOperationException("Range validation requires a minimum and maximum value.");
            int validationType = !string.IsNullOrWhiteSpace(wholeNumbers) ? 1 : (!string.IsNullOrWhiteSpace(decimals) ? 2 : 4);
            validation.Add(validationType, 1, 1, bounds2[0].Trim(), bounds2[1].Trim());
            return "Data validation applied";
        }

        private static string CreateChart(dynamic worksheet, dynamic targetRange, string content)
        {
            if (GetRangeRows(targetRange) < 2 || GetRangeColumns(targetRange) < 1)
                throw new InvalidOperationException("A chart requires a header row and data.");
            string chartType = (content ?? string.Empty).Trim().ToLowerInvariant();
            // Extract type token before any semicolon/colon
            string typeToken = chartType.Split(new char[] { ';', ':', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "column";
            int excelChartType = 51; // xlColumnClustered
            if (typeToken == "line") excelChartType = 4;
            else if (typeToken == "pie") excelChartType = 5;
            else if (typeToken == "bar") excelChartType = 57;
            else if (typeToken == "scatter") excelChartType = -4169; // xlXYScatter
            else if (typeToken == "area") excelChartType = 1; // xlArea
            else if (typeToken == "doughnut" || typeToken == "donut") excelChartType = -4120;
            else if (typeToken == "stacked" || typeToken == "column_stacked") excelChartType = 52; // xlColumnStacked
            else if (typeToken == "combo") excelChartType = 51;
            double left = Convert.ToDouble(targetRange.Left) + Convert.ToDouble(targetRange.Width) + 18;
            double top = Convert.ToDouble(targetRange.Top);
            dynamic chartObject = worksheet.ChartObjects().Add(left, top, 420, 260);
            chartObject.Chart.SetSourceData(targetRange);
            chartObject.Chart.ChartType = excelChartType;
            // Optional title handling: content may contain "title:My Title"
            string title = GetNamedOption(content, "title");
            if (!string.IsNullOrWhiteSpace(title))
            {
                try { chartObject.Chart.HasTitle = true; chartObject.Chart.ChartTitle.Text = title.Trim(); } catch { }
            }
            return "Chart created (" + typeToken + ")";
        }

        private static string CreatePivotTable(dynamic app, dynamic worksheet, dynamic targetRange, string content)
        {
            if (GetRangeRows(targetRange) < 2 || GetRangeColumns(targetRange) < 2)
                throw new InvalidOperationException("A PivotTable requires a tabular source with headers and data.");
            EnsureUniqueNonBlankHeaders(targetRange, GetRangeColumns(targetRange));
            string destinationAddress = GetNamedOption(content, "destination");
            if (string.IsNullOrWhiteSpace(destinationAddress)) destinationAddress = GetNamedOption(content, "target");
            if (!SpreadsheetActionParser.IsSafeTarget(destinationAddress) || destinationAddress.IndexOf(':') >= 0)
                throw new InvalidOperationException("PivotTable requires a single, bounded destination cell (for example destination:H2).");
            dynamic destination = worksheet.Range(destinationAddress);
            if (!string.IsNullOrWhiteSpace(Convert.ToString(destination.Value2)))
                throw new InvalidOperationException("The PivotTable destination cell must be empty.");
            if (RangesOverlap(targetRange, destination))
                throw new InvalidOperationException("The PivotTable destination cannot overlap its source data.");

            string sourceAddress = string.Format("'{0}'!{1}", Convert.ToString(worksheet.Name).Replace("'", "''"), Convert.ToString(targetRange.Address));
            dynamic workbook = app.ActiveWorkbook;
            dynamic cache = workbook.PivotCaches().Create(1, sourceAddress);
            string pivotName = GetNamedOption(content, "name");
            if (!IsSafeExcelName(pivotName)) pivotName = "AIPivot" + DateTime.UtcNow.Ticks.ToString();
            dynamic pivotTable = null;
            try { pivotTable = cache.CreatePivotTable(destination, pivotName); } catch (Exception ex) { throw new InvalidOperationException("CreatePivotTable failed: " + ex.Message, ex); }
            // Resolve actual PivotTable object (CreatePivotTable returns void on some interop; fetch by name)
            if (pivotTable == null)
            {
                try { pivotTable = worksheet.PivotTables(pivotName); } catch { }
                if (pivotTable == null) { try { pivotTable = worksheet.PivotTables(1); } catch { } }
            }
            // Apply field config if requested: rows:/values:/columns:/filters: (comma-separated header names or 1-based indices)
            if (pivotTable != null)
            {
                try
                {
                    // Build header→index map for name resolution
                    var headerToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    int cols = GetRangeColumns(targetRange);
                    for (int ci = 1; ci <= cols; ci++)
                    {
                        string h = Convert.ToString(targetRange.Cells[1, ci].Value2);
                        if (!string.IsNullOrWhiteSpace(h)) headerToIndex[h.Trim()] = ci;
                    }
                    string rowsSpec = GetNamedOption(content, "rows");
                    if (string.IsNullOrWhiteSpace(rowsSpec)) rowsSpec = GetNamedOption(content, "row");
                    string colsSpec = GetNamedOption(content, "columns");
                    if (string.IsNullOrWhiteSpace(colsSpec)) colsSpec = GetNamedOption(content, "column");
                    string valsSpec = GetNamedOption(content, "values");
                    if (string.IsNullOrWhiteSpace(valsSpec)) valsSpec = GetNamedOption(content, "vals");
                    if (string.IsNullOrWhiteSpace(valsSpec)) valsSpec = GetNamedOption(content, "value");
                    string filtersSpec = GetNamedOption(content, "filters");
                    if (string.IsNullOrWhiteSpace(filtersSpec)) filtersSpec = GetNamedOption(content, "filter");
                    // Also parse generic colon syntax like "rows:Region,Sales" already handled; split by comma/semicolon
                    ApplyPivotFields(pivotTable, worksheet, rowsSpec, 1, headerToIndex); // 1 = xlRowField
                    ApplyPivotFields(pivotTable, worksheet, colsSpec, 2, headerToIndex); // 2 = xlColumnField
                    ApplyPivotFields(pivotTable, worksheet, filtersSpec, 3, headerToIndex); // 3 = xlPageField
                    ApplyPivotDataFields(pivotTable, valsSpec, headerToIndex);
                }
                catch (Exception pfEx) { Logger.Warn(string.Format("Pivot field config partially applied: {0}", pfEx.Message)); }
            }
            return "PivotTable created: " + pivotName;
        }

        private static void ApplyPivotFields(dynamic pivotTable, dynamic worksheet, string spec, int orientation, Dictionary<string, int> headerToIndex)
        {
            if (string.IsNullOrWhiteSpace(spec) || pivotTable == null) return;
            string[] tokens = spec.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string raw in tokens)
            {
                string token = raw.Trim();
                if (string.IsNullOrWhiteSpace(token)) continue;
                // strip optional "rows:" prefix if caller passed full content string
                int colon = token.IndexOf(':');
                if (colon >= 0 && token.Length > colon + 1 && (token.ToLowerInvariant().StartsWith("rows") || token.ToLowerInvariant().StartsWith("columns") || token.ToLowerInvariant().StartsWith("filter")))
                    token = token.Substring(colon + 1).Trim();
                // token may be "Sales" or "2" or "Sales:sum"
                string fieldName = token;
                int fieldIdx;
                if (int.TryParse(token, out fieldIdx))
                {
                    // 1-based index into source columns
                }
                else if (!headerToIndex.TryGetValue(token, out fieldIdx))
                {
                    // Try case-insensitive contains fallback
                    bool found = false;
                    foreach (var kv in headerToIndex) if (kv.Key.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0) { fieldIdx = kv.Value; found = true; break; }
                    if (!found) { Logger.Warn(string.Format("Pivot field '{0}' not found in headers", token)); continue; }
                }
                try
                {
                    dynamic pf = pivotTable.PivotFields(fieldIdx);
                    if (pf == null) pf = pivotTable.PivotFields(token);
                    if (pf != null) pf.Orientation = orientation;
                }
                catch (Exception ex) { Logger.Warn(string.Format("Pivot field orientation failed for '{0}': {1}", token, ex.Message)); }
            }
        }

        private static void ApplyPivotDataFields(dynamic pivotTable, string spec, Dictionary<string, int> headerToIndex)
        {
            if (string.IsNullOrWhiteSpace(spec) || pivotTable == null) return;
            string[] tokens = spec.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string raw in tokens)
            {
                string token = raw.Trim();
                if (string.IsNullOrWhiteSpace(token)) continue;
                int colon = token.IndexOf(':');
                string func = "sum";
                if (colon >= 0)
                {
                    string maybeFunc = token.Substring(colon + 1).Trim().ToLowerInvariant();
                    // allow "Sales:sum" or "values:Sales:sum"
                    if (maybeFunc == "sum" || maybeFunc == "count" || maybeFunc == "average" || maybeFunc == "avg" || maybeFunc == "max" || maybeFunc == "min")
                    {
                        func = maybeFunc;
                        token = token.Substring(0, colon).Trim();
                        // strip leading "values:" if present after first split
                        int c2 = token.IndexOf(':');
                        if (c2 >= 0) token = token.Substring(c2 + 1).Trim();
                    }
                }
                int fieldIdx;
                if (!int.TryParse(token, out fieldIdx))
                {
                    if (!headerToIndex.TryGetValue(token, out fieldIdx))
                    {
                        bool found = false;
                        foreach (var kv in headerToIndex) if (kv.Key.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0) { fieldIdx = kv.Value; found = true; break; }
                        if (!found) { Logger.Warn(string.Format("Pivot data field '{0}' not found", token)); continue; }
                    }
                }
                try
                {
                    dynamic pf = pivotTable.PivotFields(fieldIdx);
                    int funcConst = -4157; // xlSum
                    if (func == "count") funcConst = -4112;
                    else if (func == "average" || func == "avg") funcConst = -4106;
                    else if (func == "max") funcConst = -4136;
                    else if (func == "min") funcConst = -4139;
                    pivotTable.AddDataField(pf, token + " " + func, funcConst);
                }
                catch (Exception ex) { Logger.Warn(string.Format("Pivot AddDataField failed for '{0}': {1}", token, ex.Message)); }
            }
        }

        private static string CreateNamedRange(dynamic app, dynamic worksheet, dynamic targetRange, string content)
        {
            string name = GetNamedOption(content, "name");
            if (!IsSafeExcelName(name)) throw new InvalidOperationException("Named ranges must start with a letter or underscore and cannot look like a cell address.");
            dynamic workbook = app.ActiveWorkbook;
            try
            {
                dynamic existing = workbook.Names.Item(name);
                if (existing != null) throw new InvalidOperationException("A workbook name with that name already exists.");
            }
            catch (InvalidOperationException) { throw; }
            catch { }
            workbook.Names.Add(name, targetRange);
            return "Named range created: " + name;
        }

        private static string RemoveDuplicates(dynamic targetRange, string content)
        {
            int columns = GetRangeColumns(targetRange);
            int rows = GetRangeRows(targetRange);
            if (rows < 2 || columns < 1) throw new InvalidOperationException("Remove duplicates requires a header row and at least one data row.");
            List<int> selectedColumns = ParseColumnList(GetNamedOption(content, "columns"), columns);
            if (selectedColumns.Count == 0) selectedColumns.Add(1);
            object[] fields = new object[selectedColumns.Count];
            for (int i = 0; i < selectedColumns.Count; i++) fields[i] = selectedColumns[i];
            bool noHeader = string.Equals(GetNamedOption(content, "header"), "no", StringComparison.OrdinalIgnoreCase);
            targetRange.RemoveDuplicates(fields, noHeader ? 2 : 1);
            return "Duplicate rows removed";
        }

        private static string ReplaceIgnoreCase(string input, string oldValue, string newValue)
        {
            if (string.IsNullOrEmpty(oldValue)) return input;
            var sb = new System.Text.StringBuilder();
            int start = 0;
            int idx;
            while ((idx = input.IndexOf(oldValue, start, StringComparison.OrdinalIgnoreCase)) >= 0)
            {
                sb.Append(input, start, idx - start);
                sb.Append(newValue);
                start = idx + oldValue.Length;
            }
            sb.Append(input, start, input.Length - start);
            return sb.ToString();
        }

        public HostOperationResult ExecuteFindReplace(string targetAddress, string find, string replace)
        {
            if (string.IsNullOrWhiteSpace(find))
                return HostOperationResult.Failed("Find text cannot be empty.", 0, targetAddress);
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                string f = find ?? string.Empty;
                string r = replace ?? string.Empty;
                bool found = false;
                try
                {
                    // Use Excel's Replace on the range. Signature: Replace(What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat)
                    // LookAt=2 => xlPart, MatchCase false by default
                    object ret = targetRange.Replace(f, r, 2, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing);
                    if (ret is bool) found = (bool)ret;
                    else found = true;
                }
                catch
                {
                    // Fallback: manual cell iteration via Value2 array for cases where Replace throws (e.g., protected shapes)
                    int rows = GetRangeRows(targetRange);
                    int cols = GetRangeColumns(targetRange);
                    object val = targetRange.Value2;
                    object fval = null;
                    try { fval = targetRange.Formula; } catch { }
                    bool changed = false;
                    if (val is object[,])
                    {
                        var arr = (object[,])val;
                        var farr = fval as object[,];
                        int rr = arr.GetLength(0);
                        int cc = arr.GetLength(1);
                        var outArr = new object[rr, cc];
                        for (int r2 = 1; r2 <= rr; r2++)
                        {
                            for (int c2 = 1; c2 <= cc; c2++)
                            {
                                string formula = farr != null ? Convert.ToString(farr[r2, c2]) : null;
                                if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
                                {
                                    outArr[r2 - 1, c2 - 1] = formula;
                                    continue;
                                }
                                string cell = Convert.ToString(arr[r2, c2]) ?? string.Empty;
                                if (cell.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    outArr[r2 - 1, c2 - 1] = ReplaceIgnoreCase(cell, f, r);
                                    changed = true;
                                    found = true;
                                }
                                else
                                {
                                    outArr[r2 - 1, c2 - 1] = arr[r2, c2];
                                }
                            }
                        }
                        if (changed) targetRange.Value2 = outArr;
                    }
                    else if (val != null)
                    {
                        string cell = Convert.ToString(val) ?? string.Empty;
                        string formula = Convert.ToString(fval);
                        if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
                        {
                            if (cell.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                targetRange.Value2 = ReplaceIgnoreCase(cell, f, r);
                                found = true;
                            }
                        }
                    }
                }
                if (found) return HostOperationResult.Ok(string.Format("Replaced '{0}' with '{1}' in {2}.", f, r, targetAddress), targetAddress);
                return HostOperationResult.Ok(string.Format("Text '{0}' not found in {1}; no changes made.", f, targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteFindReplace failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteFindReplace", targetAddress);
            }
        }

        public HostOperationResult ExecuteSetCase(string targetAddress, string caseType)
        {
            if (string.IsNullOrWhiteSpace(caseType))
                return HostOperationResult.Failed("Case type is required (title, sentence, upper, lower).", 0, targetAddress);
            string normalized = caseType.Trim().ToLowerInvariant();
            bool isTitle = normalized == "title" || normalized == "title_case" || normalized == "titlecase";
            bool isSentence = normalized == "sentence" || normalized == "sentence_case";
            bool isUpper = normalized == "upper" || normalized == "upper_case" || normalized == "uppercase";
            bool isLower = normalized == "lower" || normalized == "lower_case" || normalized == "lowercase";
            if (!isTitle && !isSentence && !isUpper && !isLower)
                return HostOperationResult.Failed("Case type must be one of: title, sentence, upper, lower.", 0, targetAddress);
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                int rows = GetRangeRows(targetRange);
                int cols = GetRangeColumns(targetRange);
                object val = targetRange.Value2;
                object fval = null;
                try { fval = targetRange.Formula; } catch { }
                if (val is object[,])
                {
                    var arr = (object[,])val;
                    var farr = fval as object[,];
                    int rr = arr.GetLength(0);
                    int cc = arr.GetLength(1);
                    var outArr = new object[rr, cc];
                    for (int r = 1; r <= rr; r++)
                    {
                        for (int c = 1; c <= cc; c++)
                        {
                            string formula = farr != null ? Convert.ToString(farr[r, c]) : null;
                            if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
                            {
                                outArr[r - 1, c - 1] = arr[r, c];
                                continue;
                            }
                            string cell = Convert.ToString(arr[r, c]) ?? string.Empty;
                            if (!string.IsNullOrEmpty(cell))
                                outArr[r - 1, c - 1] = ConvertExcelCase(cell, isTitle, isSentence, isUpper, isLower);
                            else
                                outArr[r - 1, c - 1] = arr[r, c];
                        }
                    }
                    targetRange.Value2 = outArr;
                }
                else if (val != null)
                {
                    string formula = Convert.ToString(fval);
                    if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
                    {
                        string cell = Convert.ToString(val) ?? string.Empty;
                        targetRange.Value2 = ConvertExcelCase(cell, isTitle, isSentence, isUpper, isLower);
                    }
                }
                return HostOperationResult.Ok(string.Format("Changed {0} to {1} case.", targetAddress, normalized), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteSetCase failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteSetCase", targetAddress);
            }
        }

        public HostOperationResult ExecuteTrimRange(string targetAddress)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                object val = targetRange.Value2;
                object fval = null;
                try { fval = targetRange.Formula; } catch { }
                if (val is object[,])
                {
                    var arr = (object[,])val;
                    var farr = fval as object[,];
                    int rr = arr.GetLength(0);
                    int cc = arr.GetLength(1);
                    var outArr = new object[rr, cc];
                    for (int r = 1; r <= rr; r++)
                    {
                        for (int c = 1; c <= cc; c++)
                        {
                            string formula = farr != null ? Convert.ToString(farr[r, c]) : null;
                            if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
                            {
                                outArr[r - 1, c - 1] = arr[r, c];
                                continue;
                            }
                            string cell = Convert.ToString(arr[r, c]) ?? string.Empty;
                            outArr[r - 1, c - 1] = string.IsNullOrEmpty(cell) ? arr[r, c] : cell.Trim();
                        }
                    }
                    targetRange.Value2 = outArr;
                }
                else if (val != null)
                {
                    string formula = Convert.ToString(fval);
                    if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
                    {
                        string cell = Convert.ToString(val) ?? string.Empty;
                        targetRange.Value2 = cell.Trim();
                    }
                }
                return HostOperationResult.Ok(string.Format("Trimmed whitespace in {0}.", targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteTrimRange failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteTrimRange", targetAddress);
            }
        }

        public HostOperationResult ExecuteNormalizeWhitespace(string targetAddress)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                object val = targetRange.Value2;
                object fval = null;
                try { fval = targetRange.Formula; } catch { }
                if (val is object[,])
                {
                    var arr = (object[,])val;
                    var farr = fval as object[,];
                    int rr = arr.GetLength(0);
                    int cc = arr.GetLength(1);
                    var outArr = new object[rr, cc];
                    for (int r = 1; r <= rr; r++)
                    {
                        for (int c = 1; c <= cc; c++)
                        {
                            string formula = farr != null ? Convert.ToString(farr[r, c]) : null;
                            if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
                            {
                                outArr[r - 1, c - 1] = arr[r, c];
                                continue;
                            }
                            string cell = Convert.ToString(arr[r, c]) ?? string.Empty;
                            if (!string.IsNullOrEmpty(cell))
                            {
                                string norm = Regex.Replace(cell, @"[ \t]{2,}", " ");
                                norm = norm.Trim();
                                outArr[r - 1, c - 1] = norm;
                            }
                            else
                            {
                                outArr[r - 1, c - 1] = arr[r, c];
                            }
                        }
                    }
                    targetRange.Value2 = outArr;
                }
                else if (val != null)
                {
                    string formula = Convert.ToString(fval);
                    if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
                    {
                        string cell = Convert.ToString(val) ?? string.Empty;
                        string norm = Regex.Replace(cell, @"[ \t]{2,}", " ");
                        targetRange.Value2 = norm.Trim();
                    }
                }
                return HostOperationResult.Ok(string.Format("Normalized whitespace in {0}.", targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteNormalizeWhitespace failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteNormalizeWhitespace", targetAddress);
            }
        }

        private static string ConvertExcelCase(string text, bool isTitle, bool isSentence, bool isUpper, bool isLower)
        {
            if (string.IsNullOrEmpty(text)) return text;
            if (isUpper) return text.ToUpperInvariant();
            if (isLower) return text.ToLowerInvariant();
            if (isTitle)
            {
                System.Globalization.TextInfo ti = System.Globalization.CultureInfo.CurrentCulture.TextInfo;
                return ti.ToTitleCase(text.ToLowerInvariant());
            }
            if (isSentence)
            {
                string lower = text.ToLowerInvariant();
                System.Text.StringBuilder sb = new System.Text.StringBuilder(lower.Length);
                bool capNext = true;
                for (int i = 0; i < lower.Length; i++)
                {
                    char c = lower[i];
                    if (capNext && char.IsLetter(c))
                    {
                        sb.Append(char.ToUpperInvariant(c));
                        capNext = false;
                    }
                    else
                    {
                        sb.Append(c);
                    }
                    if (c == '.' || c == '!' || c == '?') capNext = true;
                }
                return sb.ToString();
            }
            return text;
        }

        public HostOperationResult ExecuteTextToColumns(string targetAddress, string delimiter)
        {
            if (string.IsNullOrWhiteSpace(delimiter))
                return HostOperationResult.Failed("Delimiter is required (e.g. ',', ';', '|', 'space', 'tab').", 0, targetAddress);
            string delim = delimiter.Trim();
            if (string.Equals(delim, "space", StringComparison.OrdinalIgnoreCase)) delim = " ";
            else if (string.Equals(delim, "tab", StringComparison.OrdinalIgnoreCase)) delim = "\t";
            // Only single-char delimiters supported for predictable split; multi-char falls back to that exact string
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                int cols = GetRangeColumns(targetRange);
                if (cols != 1)
                    return HostOperationResult.Failed("Text to columns requires a single-column range (e.g. A2:A100).", 0, targetAddress);
                int startRow = Convert.ToInt32(targetRange.Row);
                int startCol = Convert.ToInt32(targetRange.Column);
                int rows = GetRangeRows(targetRange);
                object val = targetRange.Value2;
                List<string[]> splitRows = new List<string[]>();
                int maxPieces = 1;
                if (val is object[,])
                {
                    var arr = (object[,])val;
                    int rr = arr.GetLength(0);
                    for (int r = 1; r <= rr; r++)
                    {
                        string cell = Convert.ToString(arr[r, 1]) ?? string.Empty;
                        string[] pieces;
                        if (delim.Length == 1)
                            pieces = cell.Split(new char[] { delim[0] });
                        else
                            pieces = cell.Split(new string[] { delim }, StringSplitOptions.None);
                        for (int i = 0; i < pieces.Length; i++) pieces[i] = pieces[i].Trim();
                        splitRows.Add(pieces);
                        if (pieces.Length > maxPieces) maxPieces = pieces.Length;
                    }
                }
                else if (val != null)
                {
                    string cell = Convert.ToString(val) ?? string.Empty;
                    string[] pieces = delim.Length == 1 ? cell.Split(new char[] { delim[0] }) : cell.Split(new string[] { delim }, StringSplitOptions.None);
                    for (int i = 0; i < pieces.Length; i++) pieces[i] = pieces[i].Trim();
                    splitRows.Add(pieces);
                    maxPieces = pieces.Length;
                }
                else
                {
                    return HostOperationResult.Ok(string.Format("No data to split in {0}.", targetAddress), targetAddress);
                }

                if (maxPieces <= 1)
                    return HostOperationResult.Ok(string.Format("No delimiter '{0}' found in {1}; no changes made.", delimiter, targetAddress), targetAddress);

                if (startCol + maxPieces - 1 > 16384)
                    return HostOperationResult.Failed("Split would exceed Excel column limit.", 0, targetAddress);

                // Safety: check if destination beyond first column would overwrite non-empty cells
                try
                {
                    if (rows > 0 && maxPieces > 1)
                    {
                        dynamic checkRange = ws.Range(SpreadsheetActionParser.BuildRangeAddress(startCol + 1, startRow, startCol + maxPieces - 1, startRow + rows - 1));
                        object checkVal = checkRange.Value2;
                        bool hasData = false;
                        if (checkVal is object[,])
                        {
                            var ca = (object[,])checkVal;
                            for (int r = 1; r <= ca.GetLength(0); r++)
                                for (int c = 1; c <= ca.GetLength(1); c++)
                                    if (!string.IsNullOrWhiteSpace(Convert.ToString(ca[r, c]))) { hasData = true; break; }
                        }
                        else if (checkVal != null && !string.IsNullOrWhiteSpace(Convert.ToString(checkVal)))
                        {
                            hasData = true;
                        }
                        if (hasData)
                            return HostOperationResult.Failed("Split would overwrite existing data to the right of the target column. Clear the adjacent columns first.", 0, targetAddress);
                    }
                }
                catch { }

                for (int r = 0; r < splitRows.Count; r++)
                {
                    string[] pieces = splitRows[r];
                    for (int c = 0; c < pieces.Length; c++)
                    {
                        try { ws.Cells[startRow + r, startCol + c].Value2 = pieces[c]; } catch { }
                    }
                }
                return HostOperationResult.Ok(string.Format("Split {0} into {1} columns by '{2}'.", targetAddress, maxPieces, delimiter), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteTextToColumns failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteTextToColumns", targetAddress);
            }
        }

        // ==================== E2.2/E2.3 Worksheet & Row/Column Operations ====================

        public HostOperationResult ExecuteAddWorksheet(string content)
        {
            string name = GetNamedOption(content, "name");
            if (string.IsNullOrWhiteSpace(name)) name = GetNamedOption(content, "value");
            if (!string.IsNullOrWhiteSpace(name) && !SpreadsheetActionParser.IsSafeSheetName(name))
                return HostOperationResult.Failed(string.Format("Invalid worksheet name '{0}'. Names must be 1-31 chars and cannot contain : \\ / ? * [ ].", name));
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.");
                dynamic workbook = null;
                try { workbook = app.ActiveWorkbook; } catch { }
                if (workbook == null) return HostOperationResult.Failed("No active workbook.");
                dynamic newSheet = null;
                try
                {
                    dynamic activeSheet = null;
                    try { activeSheet = app.ActiveSheet; } catch { }
                    if (activeSheet != null)
                        newSheet = workbook.Worksheets.Add(Type.Missing, activeSheet);
                    else
                        newSheet = workbook.Worksheets.Add();
                }
                catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddWorksheet"); }
                if (!string.IsNullOrWhiteSpace(name))
                {
                    try { newSheet.Name = name.Trim(); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddWorksheet"); }
                }
                string finalName = Convert.ToString(newSheet.Name);
                return HostOperationResult.Ok(string.Format("Worksheet created: {0}", finalName), finalName);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteAddWorksheet failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddWorksheet");
            }
        }

        public HostOperationResult ExecuteRenameWorksheet(string targetSheet, string newName)
        {
            if (string.IsNullOrWhiteSpace(newName)) return HostOperationResult.Failed("New worksheet name cannot be empty.", 0, targetSheet);
            if (!SpreadsheetActionParser.IsSafeSheetName(newName)) return HostOperationResult.Failed(string.Format("Invalid worksheet name '{0}'.", newName), 0, targetSheet);
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetSheet);
                dynamic ws = null;
                if (!string.IsNullOrWhiteSpace(targetSheet))
                {
                    try { ws = app.Worksheets[targetSheet.Trim()]; } catch { }
                    if (ws == null) return HostOperationResult.Failed(string.Format("Worksheet '{0}' not found.", targetSheet), 0, targetSheet);
                }
                else
                {
                    try { ws = app.ActiveSheet; } catch { }
                    if (ws == null) return HostOperationResult.Failed("No active worksheet.", 0, targetSheet);
                }
                string oldName = Convert.ToString(ws.Name);
                ws.Name = newName.Trim();
                return HostOperationResult.Ok(string.Format("Renamed worksheet '{0}' to '{1}'.", oldName, newName.Trim()), newName.Trim());
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteRenameWorksheet failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteRenameWorksheet", targetSheet);
            }
        }

        public HostOperationResult ExecuteDeleteWorksheet(string targetSheet)
        {
            if (string.IsNullOrWhiteSpace(targetSheet)) return HostOperationResult.Failed("Worksheet name to delete is required.", 0, targetSheet);
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetSheet);
                dynamic workbook = null;
                try { workbook = app.ActiveWorkbook; } catch { }
                if (workbook == null) return HostOperationResult.Failed("No active workbook.", 0, targetSheet);
                int count = 0;
                try { count = Convert.ToInt32(workbook.Worksheets.Count); } catch { }
                if (count <= 1) return HostOperationResult.Failed("Cannot delete the last worksheet in the workbook.", 0, targetSheet);
                dynamic ws = null;
                try { ws = app.Worksheets[targetSheet.Trim()]; } catch { }
                if (ws == null) return HostOperationResult.Failed(string.Format("Worksheet '{0}' not found.", targetSheet), 0, targetSheet);
                try { app.DisplayAlerts = false; } catch { }
                try { ws.Delete(); } finally { try { app.DisplayAlerts = true; } catch { } }
                return HostOperationResult.Ok(string.Format("Deleted worksheet '{0}'.", targetSheet), targetSheet);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteDeleteWorksheet failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteDeleteWorksheet", targetSheet);
            }
        }

        public HostOperationResult ExecuteDuplicateWorksheet(string targetSheet, string newName)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetSheet);
                dynamic workbook = null;
                try { workbook = app.ActiveWorkbook; } catch { }
                if (workbook == null) return HostOperationResult.Failed("No active workbook.", 0, targetSheet);
                dynamic ws = null;
                if (!string.IsNullOrWhiteSpace(targetSheet))
                {
                    try { ws = app.Worksheets[targetSheet.Trim()]; } catch { }
                    if (ws == null) return HostOperationResult.Failed(string.Format("Worksheet '{0}' not found.", targetSheet), 0, targetSheet);
                }
                else
                {
                    try { ws = app.ActiveSheet; } catch { }
                    if (ws == null) return HostOperationResult.Failed("No active worksheet.", 0, targetSheet);
                }
                string safeNewName = !string.IsNullOrWhiteSpace(newName) ? newName.Trim() : null;
                if (!string.IsNullOrWhiteSpace(safeNewName) && !SpreadsheetActionParser.IsSafeSheetName(safeNewName))
                    return HostOperationResult.Failed(string.Format("Invalid worksheet name '{0}'.", safeNewName), 0, targetSheet);
                // Copy after the source sheet
                try { ws.Copy(Type.Missing, ws); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteDuplicateWorksheet", targetSheet); }
                // The copy becomes the active sheet
                dynamic newWs = null;
                try { newWs = app.ActiveSheet; } catch { }
                if (!string.IsNullOrWhiteSpace(safeNewName) && newWs != null)
                {
                    try { newWs.Name = safeNewName; } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteDuplicateWorksheet", targetSheet); }
                }
                string finalName = newWs != null ? Convert.ToString(newWs.Name) : (safeNewName ?? "Copy");
                return HostOperationResult.Ok(string.Format("Duplicated worksheet to '{0}'.", finalName), finalName);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteDuplicateWorksheet failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteDuplicateWorksheet", targetSheet);
            }
        }

        public HostOperationResult ExecuteSetTabColor(string targetSheet, string color)
        {
            if (string.IsNullOrWhiteSpace(color)) return HostOperationResult.Failed("Color is required (e.g. #FF0000, red, blue).", 0, targetSheet);
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetSheet);
                dynamic ws = null;
                if (!string.IsNullOrWhiteSpace(targetSheet))
                {
                    try { ws = app.Worksheets[targetSheet.Trim()]; } catch { }
                    if (ws == null) return HostOperationResult.Failed(string.Format("Worksheet '{0}' not found.", targetSheet), 0, targetSheet);
                }
                else
                {
                    try { ws = app.ActiveSheet; } catch { }
                    if (ws == null) return HostOperationResult.Failed("No active worksheet.", 0, targetSheet);
                }
                int oleColor;
                if (!TryParseColor(color, out oleColor))
                    return HostOperationResult.Failed(string.Format("Invalid color '{0}'. Use hex #RRGGBB or names red/green/blue/yellow/orange/purple.", color), 0, targetSheet);
                try { ws.Tab.Color = oleColor; } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteSetTabColor", targetSheet); }
                return HostOperationResult.Ok(string.Format("Set tab color for '{0}'.", Convert.ToString(ws.Name)), targetSheet);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteSetTabColor failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteSetTabColor", targetSheet);
            }
        }

        private static bool TryParseColor(string input, out int oleColor)
        {
            oleColor = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;
            string s = input.Trim().ToLowerInvariant();
            if (s == "red") { oleColor = 255; return true; }
            if (s == "green") { oleColor = 5287936; return true; }
            if (s == "blue") { oleColor = 16711680; return true; }
            if (s == "yellow") { oleColor = 65535; return true; }
            if (s == "orange") { oleColor = 39423; return true; }
            if (s == "purple") { oleColor = 10498160; return true; }
            if (s.StartsWith("#")) s = s.Substring(1);
            if (s.StartsWith("0x")) s = s.Substring(2);
            if (s.Length == 6)
            {
                try
                {
                    int r = Convert.ToInt32(s.Substring(0, 2), 16);
                    int g = Convert.ToInt32(s.Substring(2, 2), 16);
                    int b = Convert.ToInt32(s.Substring(4, 2), 16);
                    oleColor = r + (g << 8) + (b << 16);
                    return true;
                }
                catch { return false; }
            }
            return false;
        }

        public HostOperationResult ExecuteInsertRows(string targetAddress, string content)
        {
            int count = GetIntegerOption(content, "count", 1);
            if (count < 1) count = 1;
            if (count > 1000) return HostOperationResult.Failed("Insert count cannot exceed 1000.", 0, targetAddress);
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                for (int i = 0; i < count; i++)
                {
                    try { targetRange.EntireRow.Insert(1, 0); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteInsertRows", targetAddress); }
                }
                return HostOperationResult.Ok(string.Format("Inserted {0} row(s) at {1}.", count, targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteInsertRows failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteInsertRows", targetAddress);
            }
        }

        public HostOperationResult ExecuteDeleteRows(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                try { targetRange.EntireRow.Delete(Type.Missing); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteDeleteRows", targetAddress); }
                return HostOperationResult.Ok(string.Format("Deleted rows at {0}.", targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteDeleteRows failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteDeleteRows", targetAddress);
            }
        }

        public HostOperationResult ExecuteInsertColumns(string targetAddress, string content)
        {
            int count = GetIntegerOption(content, "count", 1);
            if (count < 1) count = 1;
            if (count > 100) return HostOperationResult.Failed("Insert count cannot exceed 100.", 0, targetAddress);
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                for (int i = 0; i < count; i++)
                {
                    try { targetRange.EntireColumn.Insert(1, 0); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteInsertColumns", targetAddress); }
                }
                return HostOperationResult.Ok(string.Format("Inserted {0} column(s) at {1}.", count, targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteInsertColumns failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteInsertColumns", targetAddress);
            }
        }

        public HostOperationResult ExecuteDeleteColumns(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                try { targetRange.EntireColumn.Delete(Type.Missing); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteDeleteColumns", targetAddress); }
                return HostOperationResult.Ok(string.Format("Deleted columns at {0}.", targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteDeleteColumns failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteDeleteColumns", targetAddress);
            }
        }

        public HostOperationResult ExecuteHideUnhide(string targetAddress, bool hide, bool isRow)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                if (isRow) targetRange.EntireRow.Hidden = hide;
                else targetRange.EntireColumn.Hidden = hide;
                string action = hide ? "Hid" : "Unhid";
                string what = isRow ? "rows" : "columns";
                return HostOperationResult.Ok(string.Format("{0} {1} at {2}.", action, what, targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteHideUnhide failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteHideUnhide", targetAddress);
            }
        }

        public HostOperationResult ExecuteMergeCells(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                bool unmerge = string.Equals(GetNamedOption(content, "action"), "unmerge", StringComparison.OrdinalIgnoreCase) ||
                               string.Equals((content ?? string.Empty).Trim(), "unmerge", StringComparison.OrdinalIgnoreCase);
                if (unmerge)
                {
                    try { targetRange.UnMerge(); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteMergeCells", targetAddress); }
                    return HostOperationResult.Ok(string.Format("Unmerged {0}.", targetAddress), targetAddress);
                }
                try { targetRange.Merge(Type.Missing); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteMergeCells", targetAddress); }
                return HostOperationResult.Ok(string.Format("Merged {0}.", targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteMergeCells failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteMergeCells", targetAddress);
            }
        }

        public HostOperationResult ExecuteFormatCells(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                Dictionary<string, string> opts = ParseOptions(content);
                string bold = GetNamedOption(content, "bold");
                string italic = GetNamedOption(content, "italic");
                string fontColor = GetNamedOption(content, "font_color");
                if (string.IsNullOrWhiteSpace(fontColor)) fontColor = GetNamedOption(content, "font_colour");
                if (string.IsNullOrWhiteSpace(fontColor)) fontColor = GetNamedOption(content, "font");
                string fillColor = GetNamedOption(content, "fill");
                if (string.IsNullOrWhiteSpace(fillColor)) fillColor = GetNamedOption(content, "fill_color");
                if (string.IsNullOrWhiteSpace(fillColor)) fillColor = GetNamedOption(content, "fill_colour");
                if (string.IsNullOrWhiteSpace(fillColor)) fillColor = GetNamedOption(content, "bg");
                string border = GetNamedOption(content, "border");
                string numFmt = GetNamedOption(content, "number_format");
                if (string.IsNullOrWhiteSpace(numFmt)) numFmt = GetNamedOption(content, "numberformat");
                if (string.IsNullOrWhiteSpace(numFmt)) numFmt = GetNamedOption(content, "format");
                string hAlign = GetNamedOption(content, "align");
                if (string.IsNullOrWhiteSpace(hAlign)) hAlign = GetNamedOption(content, "horizontal_alignment");
                string fontSize = GetNamedOption(content, "font_size");

                int applied = 0;
                if (!string.IsNullOrWhiteSpace(bold))
                {
                    bool b = string.Equals(bold.Trim(), "true", StringComparison.OrdinalIgnoreCase) || bold.Trim() == "1" || string.Equals(bold.Trim(), "yes", StringComparison.OrdinalIgnoreCase);
                    try { targetRange.Font.Bold = b; applied++; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(italic))
                {
                    bool it = string.Equals(italic.Trim(), "true", StringComparison.OrdinalIgnoreCase) || italic.Trim() == "1" || string.Equals(italic.Trim(), "yes", StringComparison.OrdinalIgnoreCase);
                    try { targetRange.Font.Italic = it; applied++; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(fontSize))
                {
                    int sz;
                    if (int.TryParse(fontSize.Trim(), out sz) && sz >= 6 && sz <= 72)
                    {
                        try { targetRange.Font.Size = sz; applied++; } catch { }
                    }
                }
                if (!string.IsNullOrWhiteSpace(fontColor))
                {
                    int c;
                    if (TryParseColor(fontColor, out c))
                    {
                        try { targetRange.Font.Color = c; applied++; } catch { }
                    }
                }
                if (!string.IsNullOrWhiteSpace(fillColor))
                {
                    int c;
                    if (TryParseColor(fillColor, out c))
                    {
                        try { targetRange.Interior.Color = c; applied++; } catch { }
                    }
                }
                if (!string.IsNullOrWhiteSpace(border))
                {
                    int lineStyle = 1; // xlContinuous
                    string bLower = border.Trim().ToLowerInvariant();
                    if (bLower == "none" || bLower == "0") lineStyle = -4142; // xlNone
                    else if (bLower == "thin") lineStyle = 1;
                    else if (bLower == "thick") lineStyle = 4;
                    else if (bLower == "double") lineStyle = -4119;
                    else if (bLower == "dashed") lineStyle = -4115;
                    try { targetRange.Borders.LineStyle = lineStyle; applied++; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(numFmt))
                {
                    try { targetRange.NumberFormat = numFmt.Trim(); applied++; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(hAlign))
                {
                    int align = 1; // xlGeneral
                    string a = hAlign.Trim().ToLowerInvariant();
                    if (a == "left") align = -4131;
                    else if (a == "center") align = -4108;
                    else if (a == "right") align = -4152;
                    else if (a == "justify") align = -4130;
                    try { targetRange.HorizontalAlignment = align; applied++; } catch { }
                }
                if (applied == 0)
                    return HostOperationResult.Failed("No valid formatting options found. Use bold:true/false, italic:true/false, font_color:#RRGGBB, fill:#RRGGBB, border:thin/thick/none, number_format:string, align:left/center/right, font_size:12.", 0, targetAddress);
                return HostOperationResult.Ok(string.Format("Formatted {0} ({1} properties).", targetAddress, applied), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteFormatCells failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteFormatCells", targetAddress);
            }
        }

        public HostOperationResult ExecuteAutofitColumns(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                try { targetRange.Columns.AutoFit(); } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteAutofitColumns", targetAddress); }
                try { targetRange.Rows.AutoFit(); } catch { }
                return HostOperationResult.Ok(string.Format("Autofitted {0}.", targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteAutofitColumns failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteAutofitColumns", targetAddress);
            }
        }

        public HostOperationResult ExecuteFreezePanes(string targetAddress, string content)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetAddress);
                bool freeze = true;
                string val = (content ?? string.Empty).Trim().ToLowerInvariant();
                if (val == "unfreeze" || val == "false" || val == "0" || val == "off" || val == "no")
                    freeze = false;
                if (!string.IsNullOrWhiteSpace(targetAddress) && SpreadsheetActionParser.IsSafeTarget(targetAddress))
                {
                    string sh; string rng;
                    SpreadsheetActionParser.TryParseSheetQualifiedTarget(targetAddress, out sh, out rng);
                    try
                    {
                        dynamic ws2 = null;
                        if (!string.IsNullOrWhiteSpace(sh)) ws2 = app.Worksheets[sh];
                        else ws2 = app.ActiveSheet;
                        if (ws2 != null)
                        {
                            dynamic r = ws2.Range(rng ?? targetAddress);
                            try { r.Select(); } catch { }
                            try { ws2.Activate(); } catch { }
                        }
                    }
                    catch { }
                }
                try { app.ActiveWindow.FreezePanes = freeze; } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteFreezePanes", targetAddress); }
                return HostOperationResult.Ok(freeze ? "Froze panes." : "Unfroze panes.", targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteFreezePanes failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteFreezePanes", targetAddress);
            }
        }

        public HostOperationResult ExecuteAddSummaryRow(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                string op = GetNamedOption(content, "operation");
                if (string.IsNullOrWhiteSpace(op)) op = GetNamedOption(content, "op");
                if (string.IsNullOrWhiteSpace(op)) op = GetNamedOption(content, "function");
                if (string.IsNullOrWhiteSpace(op)) op = (content ?? string.Empty).Trim().ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(op)) op = "sum";
                string columnOpt = GetNamedOption(content, "column");
                if (string.IsNullOrWhiteSpace(columnOpt)) columnOpt = GetNamedOption(content, "col");
                int targetCol = 1;
                if (!string.IsNullOrWhiteSpace(columnOpt))
                {
                    // column may be letter like "C" or number "3"
                    if (int.TryParse(columnOpt.Trim(), out targetCol))
                    {
                        // numeric
                    }
                    else
                    {
                        string innerColLetter = columnOpt.Trim().ToUpperInvariant();
                        targetCol = SpreadsheetActionParser.ColumnLetterToIndex(innerColLetter) != 0 ? GetRangeColumns(targetRange) : 1;
                        // Actually map letter to index within range: find column position
                        try
                        {
                            int absCol = SpreadsheetActionParser.ColumnLetterToIndex(innerColLetter);
                            int absStartCol = Convert.ToInt32(targetRange.Column);
                            targetCol = absCol - absStartCol + 1;
                            if (targetCol < 1 || targetCol > GetRangeColumns(targetRange)) targetCol = 1;
                        }
                        catch { targetCol = 1; }
                    }
                }
                else
                {
                    // default: first numeric column? just use 1
                    targetCol = 1;
                }

                int rows = GetRangeRows(targetRange);
                int cols = GetRangeColumns(targetRange);
                if (rows < 2) return HostOperationResult.Failed("Summary row requires a header row and at least one data row.", 0, targetAddress);
                int startRow = Convert.ToInt32(targetRange.Row);
                int startCol = Convert.ToInt32(targetRange.Column);
                int summaryRow = startRow + rows;
                // Check if summary row would exceed limits
                if (summaryRow > 1048576) return HostOperationResult.Failed("Summary row would exceed Excel row limit.", 0, targetAddress);
                // Build A1 for column within range
                string colLetter = SpreadsheetActionParser.IndexToColumnLetter(startCol + targetCol - 1);
                string startData = string.Format("{0}{1}", colLetter, startRow + 1);
                string endData = string.Format("{0}{1}", colLetter, startRow + rows - 1);
                string formula = string.Empty;
                string lo = op.Trim().ToLowerInvariant();
                if (lo == "sum") formula = string.Format("=SUM({0}:{1})", startData, endData);
                else if (lo == "average" || lo == "avg") formula = string.Format("=AVERAGE({0}:{1})", startData, endData);
                else if (lo == "count") formula = string.Format("=COUNT({0}:{1})", startData, endData);
                else if (lo == "max") formula = string.Format("=MAX({0}:{1})", startData, endData);
                else if (lo == "min") formula = string.Format("=MIN({0}:{1})", startData, endData);
                else return HostOperationResult.Failed("Operation must be one of: sum, average, count, max, min.", 0, targetAddress);

                dynamic sumCell = null;
                try { sumCell = ws.Cells[summaryRow, startCol + targetCol - 1]; } catch { }
                if (sumCell == null) return HostOperationResult.Failed("Could not resolve summary cell.", 0, targetAddress);
                sumCell.Formula = formula;
                // Optionally label first column as "Total"
                try
                {
                    dynamic labelCell = ws.Cells[summaryRow, startCol];
                    string labelVal = Convert.ToString(labelCell.Value2);
                    if (string.IsNullOrWhiteSpace(labelVal))
                        labelCell.Value2 = "Total";
                }
                catch { }
                string resVal = ReadRangeResult(sumCell);
                return HostOperationResult.Ok(string.Format("Summary {0} added at {1}{2}: {3}", lo, colLetter, summaryRow, resVal), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteAddSummaryRow failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddSummaryRow", targetAddress);
            }
        }

        public HostOperationResult ExecuteWritePython(string targetAddress, string code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return HostOperationResult.Failed("Python code cannot be empty.", 0, targetAddress);
            string python = code.Trim();
            // Multi-line support: normalize line breaks to semicolons (Excel =PY() is single-cell but Python statements can be separated by ";")
            if (python.IndexOf('\r') >= 0 || python.IndexOf('\n') >= 0)
            {
                python = Regex.Replace(python, @"\r\n|\n|\r", "; ");
                python = Regex.Replace(python, @";\s*;\s*", "; ");
                python = python.Trim().Trim(';').Trim();
                if (string.IsNullOrWhiteSpace(python))
                    return HostOperationResult.Failed("Python code cannot be empty after normalizing line breaks.", 0, targetAddress);
            }
            // Basic length guard
            if (python.Length > 8000)
                return HostOperationResult.Failed("Python code exceeds single-cell limit (8000 chars).", 0, targetAddress);

            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                // Probe: check if Formula2 is supported (Python in Excel uses Formula2 property)
                bool hasFormula2 = false;
                try
                {
                    var probe = targetRange.Formula2;
                    hasFormula2 = true;
                }
                catch
                {
                    hasFormula2 = false;
                }
                if (!hasFormula2)
                {
                    return HostOperationResult.Failed("Python in Excel is not available on this host/version. It requires Microsoft 365 with Python in Excel enabled (Formulas > Python). The workbook will not be modified.", 0, targetAddress);
                }

                // Escape double quotes for Excel formula: " -> ""
                string escaped = python.Replace("\"", "\"\"");
                string formula = string.Format("=PY(\"{0}\")", escaped);

                // Single-cell only: enforce 1x1 target
                int rows = GetRangeRows(targetRange);
                int cols = GetRangeColumns(targetRange);
                if (rows != 1 || cols != 1)
                    return HostOperationResult.Failed("Python write requires a single-cell target (e.g. H2), not a multi-cell range.", 0, targetAddress);

                try
                {
                    // Use Formula2 for dynamic array / PY support, fallback to Formula if Formula2 fails
                    try { targetRange.Formula2 = formula; }
                    catch { targetRange.Formula = formula; }
                }
                catch (Exception ex)
                {
                    // Detect common "Python not enabled" HRESULTs: 0x800A03EC, 0x80020005, 0x800AC472
                    string msg = ex.Message ?? string.Empty;
                    if (msg.IndexOf("PY", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Formula2", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return HostOperationResult.Failed("Python in Excel is not enabled or not supported on this machine. Enable Python in Excel or update Microsoft 365. Details: " + msg, 0, targetAddress);
                    }
                    return HostOperationResult.FromException(ex, "ExcelController.ExecuteWritePython", targetAddress);
                }

                string sandboxNote = " (Python runs in Excel's managed sandbox; data may leave local session per Microsoft's Python runtime.)";
                string readBack = string.Empty;
                try { readBack = Convert.ToString(targetRange.Value2) ?? string.Empty; } catch { }
                if (!string.IsNullOrWhiteSpace(readBack) && readBack.Length > 120) readBack = readBack.Substring(0, 120) + "...";
                return HostOperationResult.Ok(string.Format("Python formula written to {0}{1} {2}", targetAddress, string.IsNullOrWhiteSpace(readBack) ? string.Empty : ": " + readBack, sandboxNote), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteWritePython failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteWritePython", targetAddress);
            }
        }

        public HostOperationResult ExecuteApplyTheme(string targetAddress, string content)
        {
            // Theme is a fixed enum palette per review verdict: blue, green, grey, railway
            string palette = GetNamedOption(content, "palette");
            if (string.IsNullOrWhiteSpace(palette)) palette = GetNamedOption(content, "theme");
            if (string.IsNullOrWhiteSpace(palette)) palette = (content ?? string.Empty).Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(palette)) palette = "blue";
            palette = palette.Trim().ToLowerInvariant();
            HashSet<string> allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "blue", "green", "grey", "gray", "railway", "neutral" };
            if (!allowed.Contains(palette))
                return HostOperationResult.Failed("Palette must be one of: blue, green, grey, railway. (Fixed enum, no freeform hex.)", 0, targetAddress);

            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                EnsureWorksheetEditable(ws);
                int oleHeaderFill, oleHeaderFont, oleBandFill;
                if (palette == "blue") { TryParseColor("#4472C4", out oleHeaderFill); TryParseColor("#FFFFFF", out oleHeaderFont); TryParseColor("#D9E1F2", out oleBandFill); }
                else if (palette == "green") { TryParseColor("#548235", out oleHeaderFill); TryParseColor("#FFFFFF", out oleHeaderFont); TryParseColor("#E2EFDA", out oleBandFill); }
                else if (palette == "grey" || palette == "gray" || palette == "neutral") { TryParseColor("#404040", out oleHeaderFill); TryParseColor("#FFFFFF", out oleHeaderFont); TryParseColor("#F2F2F2", out oleBandFill); }
                else { TryParseColor("#0F2A44", out oleHeaderFill); TryParseColor("#FFFFFF", out oleHeaderFont); TryParseColor("#EAF0F6", out oleBandFill); }

                int rows = GetRangeRows(targetRange);
                int cols = GetRangeColumns(targetRange);
                if (rows < 2) return HostOperationResult.Failed("Theme requires a header row and at least one data row.", 0, targetAddress);
                int startRow = Convert.ToInt32(targetRange.Row);
                int startCol = Convert.ToInt32(targetRange.Column);

                // Header row formatting
                try
                {
                    dynamic header = ws.Range(SpreadsheetActionParser.BuildRangeAddress(startCol, startRow, startCol + cols - 1, startRow));
                    header.Interior.Color = oleHeaderFill;
                    header.Font.Color = oleHeaderFont;
                    header.Font.Bold = true;
                    header.HorizontalAlignment = -4108; // xlCenter
                    header.Borders.LineStyle = 1;
                }
                catch (Exception ex) { Logger.Warn(string.Format("ApplyTheme header failed: {0}", ex.Message)); }

                // Banded rows
                try
                {
                    for (int r = startRow + 1; r < startRow + rows; r++)
                    {
                        if ((r - startRow) % 2 == 0)
                        {
                            dynamic rowRange = ws.Range(SpreadsheetActionParser.BuildRangeAddress(startCol, r, startCol + cols - 1, r));
                            rowRange.Interior.Color = oleBandFill;
                        }
                    }
                }
                catch (Exception ex) { Logger.Warn(string.Format("ApplyTheme banding failed: {0}", ex.Message)); }

                // Borders for whole range
                try { targetRange.Borders.LineStyle = 1; } catch { }
                try { targetRange.Columns.AutoFit(); } catch { }

                return HostOperationResult.Ok(string.Format("Theme '{0}' applied to {1}.", palette, targetAddress), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteApplyTheme failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteApplyTheme", targetAddress);
            }
        }

        // ==================== Analysis Mode (Local, Read-Only) ====================
        public HostOperationResult ExecuteAnalyzeRange(string targetAddress, string content)
        {
            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            try
            {
                int rows = GetRangeRows(targetRange);
                int cols = GetRangeColumns(targetRange);
                if (rows < 2 || cols < 1) return HostOperationResult.Failed("Analyze requires a header row and at least one data row.", 0, targetAddress);
                // Strict range-size limit for analysis to prevent OOM / COM timeout on large selections
                const int MaxAnalysisCells = 10000;
                long cellCount = (long)rows * cols;
                if (cellCount > MaxAnalysisCells)
                    return HostOperationResult.Failed(string.Format("Analysis range {0} contains {1} cells, exceeding limit of {2} for analysis. Narrow the range.", targetAddress, cellCount, MaxAnalysisCells), 0, targetAddress);
                string detail = GetNamedOption(content, "detail");
                if (string.IsNullOrWhiteSpace(detail)) detail = (content ?? string.Empty).Trim().ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(detail)) detail = "summary";
                else detail = detail.Trim().ToLowerInvariant();
                bool isFull = detail == "full" || detail == "detailed" || detail == "verbose";
                object val = targetRange.Value2;
                if (!(val is object[,])) return HostOperationResult.Failed("Cannot analyze single-cell range.", 0, targetAddress);
                var arr = (object[,])val;
                int rr = arr.GetLength(0);
                int cc = arr.GetLength(1);
                int startRow = Convert.ToInt32(targetRange.Row);
                int startCol = Convert.ToInt32(targetRange.Column);
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(string.Format("Analysis for {0} ({1} rows x {2} cols) [{3}]:", targetAddress, rows, cols, isFull ? "full" : "summary"));
                // For each column, compute stats
                for (int c = 1; c <= cc; c++)
                {
                    string colLetter = IndexToColumnLetter(startCol + c - 1);
                    string header = Convert.ToString(arr[1, c]) ?? string.Format("Col{0}", colLetter);
                    header = header.Trim();
                    if (string.IsNullOrEmpty(header)) header = string.Format("Col{0}", colLetter);
                    List<double> nums = new List<double>();
                    List<int> numRows = new List<int>();
                    List<string> texts = new List<string>();
                    for (int r = 2; r <= rr; r++)
                    {
                        string s = Convert.ToString(arr[r, c]) ?? string.Empty;
                        s = s.Trim();
                        if (string.IsNullOrEmpty(s)) continue;
                        double d;
                        if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out d) ||
                            double.TryParse(s, out d))
                        {
                            nums.Add(d);
                            numRows.Add(startRow + r - 1);
                        }
                        else
                            texts.Add(s);
                    }
                    sb.AppendLine(string.Format("  {0} [{1}]:", colLetter, header));
                    if (nums.Count > 0)
                    {
                        double sum = 0, min = nums[0], max = nums[0];
                        for (int i = 0; i < nums.Count; i++) { sum += nums[i]; if (nums[i] < min) min = nums[i]; if (nums[i] > max) max = nums[i]; }
                        double avg = sum / nums.Count;
                        // stdev
                        double var = 0;
                        for (int i = 0; i < nums.Count; i++) var += (nums[i] - avg) * (nums[i] - avg);
                        double stdev = nums.Count > 1 ? Math.Sqrt(var / (nums.Count - 1)) : 0;
                        // trend: simple linear slope (use original ordering, skip blanks)
                        double slope = 0;
                        if (nums.Count >= 2)
                        {
                            double sx = 0, sy = 0, sxx = 0, sxy = 0;
                            for (int i = 0; i < nums.Count; i++) { double x = i; double y = nums[i]; sx += x; sy += y; sxx += x * x; sxy += x * y; }
                            double n = nums.Count;
                            double denom = n * sxx - sx * sx;
                            if (Math.Abs(denom) > 1e-9) slope = (n * sxy - sx * sy) / denom;
                        }
                        string trend = Math.Abs(slope) < 1e-6 ? "flat" : (slope > 0 ? "upward" : "downward");
                        sb.AppendLine(string.Format("    Numeric: n={0} sum={1:F2} avg={2:F2} min={3:F2} max={4:F2} stdev={5:F2} trend={6} (slope {7:F4})", nums.Count, sum, avg, min, max, stdev, trend, slope));
                        // outliers: > avg+2*stdev or < avg-2*stdev - use correct original row addresses even when blanks were skipped
                        List<string> outliers = new List<string>();
                        if (stdev > 1e-9)
                        {
                            for (int i = 0; i < nums.Count; i++)
                            {
                                if (Math.Abs(nums[i] - avg) > 2 * stdev)
                                {
                                    int actualRow = numRows[i];
                                    outliers.Add(string.Format("{0}{1}={2}", colLetter, actualRow, nums[i]));
                                    if (outliers.Count >= 5) break;
                                }
                            }
                        }
                        if (outliers.Count > 0) sb.AppendLine(string.Format("    Outliers (2σ): {0}", string.Join(", ", outliers.ToArray())));
                        // distribution: count above/below avg
                        int above = 0;
                        for (int i = 0; i < nums.Count; i++) if (nums[i] > avg) above++;
                        sb.AppendLine(string.Format("    Distribution: {0} above avg, {1} below avg", above, nums.Count - above));
                        if (isFull && nums.Count > 0)
                        {
                            // Full detail: include min/max rows addresses
                            int minIdx = 0, maxIdx = 0;
                            for (int i = 1; i < nums.Count; i++) { if (nums[i] < nums[minIdx]) minIdx = i; if (nums[i] > nums[maxIdx]) maxIdx = i; }
                            sb.AppendLine(string.Format("    Min at {0}{1}, Max at {0}{2}", colLetter, numRows[minIdx], numRows[maxIdx]));
                        }
                    }
                    if (texts.Count > 0)
                    {
                        // distinct top values
                        Dictionary<string, int> freq = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        for (int i = 0; i < texts.Count; i++) { string t = texts[i]; if (freq.ContainsKey(t)) freq[t]++; else freq[t] = 1; }
                        List<KeyValuePair<string, int>> sorted = new List<KeyValuePair<string, int>>(freq);
                        sorted.Sort(delegate(KeyValuePair<string, int> a, KeyValuePair<string, int> b) { return b.Value.CompareTo(a.Value); });
                        int topN = Math.Min(isFull ? 5 : 3, sorted.Count);
                        List<string> top = new List<string>();
                        for (int i = 0; i < topN; i++) top.Add(string.Format("{0} ({1})", sorted[i].Key, sorted[i].Value));
                        if (top.Count > 0) sb.AppendLine(string.Format("    Top values: {0}", string.Join(", ", top.ToArray())));
                        if (freq.Count > 10) sb.AppendLine(string.Format("    Distinct: {0} values", freq.Count));
                    }
                }
                // Suggestions for chart/pivot - detail controls verbosity
                sb.AppendLine("  Suggestions:");
                if (isFull)
                {
                    sb.AppendLine("    - Chart: column chart of first numeric column by header (use excel.create_chart)");
                    sb.AppendLine("    - Chart: line chart if trend is upward/downward to visualize trajectory");
                    sb.AppendLine("    - Pivot: summarize by first text column with sum of first numeric column");
                    sb.AppendLine("    - Conditional format: highlight outliers with top_n or color_scale");
                }
                else
                {
                    sb.AppendLine("    - Chart: column chart of first numeric column by header");
                    sb.AppendLine("    - Pivot: summarize by first text column with sum of first numeric column");
                }
                return HostOperationResult.Ok(sb.ToString().TrimEnd(), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteAnalyzeRange failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteAnalyzeRange", targetAddress);
            }
        }

        // ==================== Explain Formula (Read-Only) ====================
        public HostOperationResult ExecuteGetFormulaDetails(string targetAddress, string content)
        {
            try
            {
                // Enforce single-cell target and reuse safe resolver
                string addr = targetAddress;
                if (string.IsNullOrWhiteSpace(addr))
                {
                    dynamic appProbe = _rawAppObj;
                    if (appProbe == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetAddress);
                    try { addr = Convert.ToString(appProbe.ActiveCell.Address); } catch { addr = "A1"; }
                    if (!string.IsNullOrEmpty(addr)) addr = addr.Replace("$", "");
                }
                else
                {
                    addr = addr.Replace("$", "").Trim();
                }
                // Use safe resolver - this already validates IsSafeTarget, sheet existence, and protection
                dynamic app, ws, targetRange;
                var resolve = ResolveTargetRange(addr, out app, out ws, out targetRange);
                if (!resolve.Success) return resolve;
                // Enforce single cell
                int rRows = GetRangeRows(targetRange);
                int rCols = GetRangeColumns(targetRange);
                if (rRows != 1 || rCols != 1)
                    return HostOperationResult.Failed("Formula details requires a single cell address (e.g. B7 or Sheet1!B7), not a multi-cell range.", 0, addr);
                dynamic cell = targetRange;
                string sh;
                string rng;
                SpreadsheetActionParser.TryParseSheetQualifiedTarget(addr, out sh, out rng);
                string effective = !string.IsNullOrWhiteSpace(rng) ? rng : addr;
                string formula = string.Empty;
                string value = string.Empty;
                string addrResolved = effective;
                try { formula = Convert.ToString(cell.Formula); } catch { }
                try { value = Convert.ToString(cell.Value2); } catch { }
                try { addrResolved = Convert.ToString(cell.Address).Replace("$", ""); } catch { }
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(string.Format("Formula details for {0}:", string.IsNullOrWhiteSpace(sh) ? addrResolved : sh + "!" + addrResolved));
                sb.AppendLine(string.Format("  Formula: {0}", string.IsNullOrEmpty(formula) ? "(empty)" : formula));
                sb.AppendLine(string.Format("  Value: {0}", string.IsNullOrEmpty(value) ? "(empty)" : value));
                // dependencies: try Precedents / Dependents
                try
                {
                    dynamic prec = cell.Precedents;
                    if (prec != null)
                    {
                        string pAddr = Convert.ToString(prec.Address);
                        if (!string.IsNullOrWhiteSpace(pAddr)) sb.AppendLine(string.Format("  Precedents: {0}", pAddr.Replace("$", "")));
                    }
                }
                catch { }
                try
                {
                    dynamic dep = cell.Dependents;
                    if (dep != null)
                    {
                        string dAddr = Convert.ToString(dep.Address);
                        if (!string.IsNullOrWhiteSpace(dAddr)) sb.AppendLine(string.Format("  Dependents: {0}", dAddr.Replace("$", "")));
                    }
                }
                catch { }
                // simple function list via regex
                if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
                {
                    MatchCollection funcs = Regex.Matches(formula, @"\b([A-Z][A-Z0-9]*)\s*\(");
                    HashSet<string> uniq = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (Match m in funcs) uniq.Add(m.Groups[1].Value.ToUpperInvariant());
                    if (uniq.Count > 0) sb.AppendLine(string.Format("  Functions: {0}", string.Join(", ", new List<string>(uniq).ToArray())));
                    // cell refs
                    MatchCollection refs = Regex.Matches(formula, @"(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?");
                    if (refs.Count > 0)
                    {
                        List<string> rlist = new List<string>();
                        foreach (Match m in refs) { string r = m.Value.Replace("$", ""); if (!rlist.Contains(r)) rlist.Add(r); if (rlist.Count >= 10) break; }
                        sb.AppendLine(string.Format("  References: {0}", string.Join(", ", rlist.ToArray())));
                    }
                }
                return HostOperationResult.Ok(sb.ToString().TrimEnd(), addr);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteGetFormulaDetails failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteGetFormulaDetails", targetAddress);
            }
        }

        // ==================== Text Analysis Column ====================
        public HostOperationResult ExecuteAddAnalysisColumn(string targetAddress, string content)
        {
            // target = header cell for new column (e.g. G1), content = "source:A2:A100;type:sentiment|classify|topic|summarize;header:Sentiment"
            string source = GetNamedOption(content, "source");
            if (string.IsNullOrWhiteSpace(source)) source = GetNamedOption(content, "range");
            if (string.IsNullOrWhiteSpace(source)) return HostOperationResult.Failed("Source range is required (e.g. source:A2:A100).", 0, targetAddress);
            string type = GetNamedOption(content, "type");
            if (string.IsNullOrWhiteSpace(type)) type = "classify";
            type = type.Trim().ToLowerInvariant();
            HashSet<string> allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "sentiment", "classify", "topic", "summarize", "category", "tag" };
            if (!allowed.Contains(type)) return HostOperationResult.Failed("Type must be one of: sentiment, classify, topic, summarize.", 0, targetAddress);
            string header = GetNamedOption(content, "header");
            if (string.IsNullOrWhiteSpace(header)) header = GetNamedOption(content, "title");
            if (string.IsNullOrWhiteSpace(header))
            {
                if (type == "sentiment") header = "Sentiment";
                else if (type == "topic") header = "Topic";
                else if (type == "summarize") header = "Summary";
                else header = "Category";
            }
            dynamic app, ws, targetRange, sourceRange;
            var resT = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!resT.Success) return resT;
            string shT, rngT;
            SpreadsheetActionParser.TryParseSheetQualifiedTarget(targetAddress, out shT, out rngT);
            dynamic wsSource = ws;
            string shS, rngS;
            SpreadsheetActionParser.TryParseSheetQualifiedTarget(source, out shS, out rngS);
            if (!string.IsNullOrWhiteSpace(shS))
            {
                try { wsSource = app.Worksheets[shS]; } catch { return HostOperationResult.Failed(string.Format("Source worksheet '{0}' not found.", shS), 0, targetAddress); }
            }
            try
            {
                sourceRange = wsSource.Range(string.IsNullOrWhiteSpace(rngS) ? source : rngS);
            }
            catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddAnalysisColumn", targetAddress); }
            if (sourceRange == null) return HostOperationResult.Failed("Could not resolve source range.", 0, targetAddress);
            if (GetRangeRows(targetRange) != 1 || GetRangeColumns(targetRange) != 1)
                return HostOperationResult.Failed("Target must be a single header cell (e.g. G1) for the new analysis column.", 0, targetAddress);
            try
            {
                // Write header
                targetRange.Value2 = header;
                try { targetRange.Font.Bold = true; } catch { }
                try { ExcelChangeHighlighter.ApplyHighlight(ws, targetRange); } catch { }
                int srcRows = GetRangeRows(sourceRange);
                int srcCols = GetRangeColumns(sourceRange);
                if (srcCols != 1) return HostOperationResult.Failed("Source must be a single column for analysis.", 0, targetAddress);
                // For local demo, fill new column with placeholder analysis values derived deterministically
                // (real AI classification would be inserted via subsequent write actions; this creates the column structure)
                int startRow = Convert.ToInt32(sourceRange.Row) + 1; // skip header of source
                int endRow = Convert.ToInt32(sourceRange.Row) + srcRows - 1;
                int tgtCol = Convert.ToInt32(targetRange.Column);
                int tgtRowStart = Convert.ToInt32(targetRange.Row) + 1;
                object srcVal = sourceRange.Value2;
                int filled = 0;
                if (srcVal is object[,])
                {
                    var arr = (object[,])srcVal;
                    int rr = arr.GetLength(0);
                    for (int r = 2; r <= rr; r++)
                    {
                        string txt = Convert.ToString(arr[r, 1]) ?? string.Empty;
                        txt = txt.Trim();
                        string analysis = string.Empty;
                        if (string.IsNullOrEmpty(txt)) analysis = "";
                        else if (type == "sentiment")
                        {
                            string lower = txt.ToLowerInvariant();
                            // Expanded deterministic lexicon - still heuristic placeholder, AI can overwrite via follow-up write_value actions
                            int pos = 0, neg = 0;
                            string[] posWords = { "good", "excellent", "positive", "great", "love", "like", "awesome", "fantastic", "satisfied", "happy", "pleased", "wonderful", "best", "amazing" };
                            string[] negWords = { "bad", "poor", "negative", "terrible", "awful", "hate", "disappoint", "horrible", "worst", "sad", "angry", "frustrat", "complain", "issue", "problem" };
                            foreach (var w in posWords) if (lower.Contains(w)) pos++;
                            foreach (var w in negWords) if (lower.Contains(w)) neg++;
                            if (pos > neg) analysis = "Positive";
                            else if (neg > pos) analysis = "Negative";
                            else analysis = "Neutral";
                            // Add note that this is heuristic
                        }
                        else if (type == "summarize")
                        {
                            // Slightly smarter: take first sentence or 40 chars, not naive truncation
                            string firstSentence = txt.Split(new[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? txt;
                            firstSentence = firstSentence.Trim();
                            if (firstSentence.Length > 50) analysis = firstSentence.Substring(0, 47) + "...";
                            else analysis = firstSentence;
                            if (string.IsNullOrWhiteSpace(analysis)) analysis = txt.Length > 40 ? txt.Substring(0, 40) + "..." : txt;
                        }
                        else // classify/topic - deterministic keyword clustering placeholder
                        {
                            string lower = txt.ToLowerInvariant();
                            // Simple keyword -> category heuristic (local placeholder, AI overwrites via follow-up)
                            if (lower.Contains("train") || lower.Contains("rail") || lower.Contains("track") || lower.Contains("station") || lower.Contains("ohe") || lower.Contains("signal")) analysis = "Rail Operations";
                            else if (lower.Contains("invoice") || lower.Contains("payment") || lower.Contains("cost") || lower.Contains("price") || lower.Contains("budget") || lower.Contains("revenue")) analysis = "Finance";
                            else if (lower.Contains("delay") || lower.Contains("cancel") || lower.Contains("schedule") || lower.Contains("time") || lower.Contains("late")) analysis = "Schedule";
                            else if (lower.Contains("safety") || lower.Contains("incident") || lower.Contains("accident") || lower.Contains("risk")) analysis = "Safety";
                            else if (lower.Contains("customer") || lower.Contains("passenger") || lower.Contains("feedback") || lower.Contains("complaint") || lower.Contains("service")) analysis = "Customer";
                            else if (lower.Contains("maintenance") || lower.Contains("repair") || lower.Contains("breakdown") || lower.Contains("failure") || lower.Contains("pm") || lower.Contains("cm")) analysis = "Maintenance";
                            else
                            {
                                string snippet = txt.Length > 24 ? txt.Substring(0, 24) : txt;
                                analysis = "Topic: " + snippet;
                            }
                        }
                        try { ws.Cells[tgtRowStart + r - 2, tgtCol].Value2 = analysis; filled++; } catch { }
                        if (filled >= 10000) break;
                    }
                }
                return HostOperationResult.Ok(string.Format("Analysis column '{0}' created at {1} from {2} ({3}, {4} rows prepared — AI can overwrite via follow-up actions).", header, targetAddress, source, type, filled), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteAddAnalysisColumn failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddAnalysisColumn", targetAddress);
            }
        }

        // ==================== Local Workbook Import ====================
        public HostOperationResult ExecuteImportWorksheet(string targetAddress, string content)
        {
            string sourcePath = GetNamedOption(content, "source");
            if (string.IsNullOrWhiteSpace(sourcePath)) sourcePath = GetNamedOption(content, "sourcePath");
            if (string.IsNullOrWhiteSpace(sourcePath)) sourcePath = GetNamedOption(content, "path");
            if (string.IsNullOrWhiteSpace(sourcePath)) return HostOperationResult.Failed("Source file path is required (e.g. source:C:\\Data\\Report.xlsx).", 0, targetAddress);
            sourcePath = sourcePath.Trim().Trim('"').Trim('\'');
            string sourceSheet = GetNamedOption(content, "sheet");
            if (string.IsNullOrWhiteSpace(sourceSheet)) sourceSheet = GetNamedOption(content, "sourceSheet");
            // targetAddress is where to import — if it looks like a sheet name (no A1), create new sheet
            try
            {
                if (!File.Exists(sourcePath)) return HostOperationResult.Failed(string.Format("Source file not found: {0}", sourcePath), 0, targetAddress);
                string ext = Path.GetExtension(sourcePath).ToLowerInvariant();
                if (ext != ".xlsx") return HostOperationResult.Failed("Only .xlsx files are supported for local import.", 0, targetAddress);
                FileInfo fi = new FileInfo(sourcePath);
                if (fi.Length > 20 * 1024 * 1024) return HostOperationResult.Failed("Source file exceeds 20 MB limit.", 0, targetAddress);
                // Validate source not the same as active workbook (avoid self-copy deadlock)
                dynamic app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetAddress);
                string activePath = string.Empty;
                try { activePath = Convert.ToString(app.ActiveWorkbook.FullName); } catch { }
                if (!string.IsNullOrWhiteSpace(activePath) && string.Equals(Path.GetFullPath(activePath).TrimEnd('\\'), Path.GetFullPath(sourcePath).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                    return HostOperationResult.Failed("Source and target workbook are the same file. Use a different file.", 0, targetAddress);

                // Read source via Zip/OpenXml preserving original cell addresses (sparse-aware) and formulas
                List<Tuple<int, int, string>> sparseCells = new List<Tuple<int, int, string>>();
                int? minRow = null, minCol = null, maxRow = null, maxCol = null;
                string resolvedSheetName = sourceSheet;
                using (var zip = System.IO.Compression.ZipFile.OpenRead(sourcePath))
                {
                    // sheet names
                    List<string> sheetNames = new List<string>();
                    var wbEntry = zip.GetEntry("xl/workbook.xml");
                    if (wbEntry != null)
                    {
                        using (var s = wbEntry.Open())
                        {
                            var doc = System.Xml.Linq.XDocument.Load(s);
                            System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                            foreach (var sh in doc.Descendants(ns + "sheet"))
                            {
                                string n = (string)sh.Attribute("name");
                                if (!string.IsNullOrWhiteSpace(n)) sheetNames.Add(n);
                            }
                        }
                    }
                    if (string.IsNullOrWhiteSpace(sourceSheet) && sheetNames.Count > 0) resolvedSheetName = sheetNames[0];
                    else if (string.IsNullOrWhiteSpace(sourceSheet)) resolvedSheetName = "Sheet1";
                    // shared strings
                    List<string> shared = new List<string>();
                    var sst = zip.GetEntry("xl/sharedStrings.xml");
                    if (sst != null)
                    {
                        using (var s = sst.Open())
                        {
                            var doc = System.Xml.Linq.XDocument.Load(s);
                            System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                            foreach (var si in doc.Descendants(ns + "si")) shared.Add(string.Concat(si.Descendants(ns + "t").Select(t => t.Value)));
                        }
                    }
                    // find sheet file via workbook.xml.rels (accurate) with fallback to ordinal sheetN.xml
                    System.IO.Compression.ZipArchiveEntry sheetEntry = null;
                    string resolvedSheetFile = null;
                    try
                    {
                        var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                        if (relsEntry != null && wbEntry != null)
                        {
                            var rIdToTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            using (var rs = relsEntry.Open())
                            {
                                var rdoc = System.Xml.Linq.XDocument.Load(rs);
                                System.Xml.Linq.XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                                foreach (var rel in rdoc.Descendants(relNs + "Relationship"))
                                {
                                    string id = (string)rel.Attribute("Id");
                                    string target = (string)rel.Attribute("Target");
                                    if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                                    {
                                        string normalized;
                                        if (target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                                            normalized = target;
                                        else if (target.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
                                            normalized = target.TrimStart('/');
                                        else if (target.StartsWith("/"))
                                            normalized = target.TrimStart('/');
                                        else
                                            normalized = "xl/" + target.TrimStart('/');
                                        // Ensure xl/ prefix
                                        if (!normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                                            normalized = "xl/" + normalized.TrimStart('/');
                                        // Normalize slashes
                                        normalized = normalized.Replace("\\", "/");
                                        rIdToTarget[id] = normalized;
                                    }
                                }
                            }
                            using (var ws = wbEntry.Open())
                            {
                                var wdoc = System.Xml.Linq.XDocument.Load(ws);
                                System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                                System.Xml.Linq.XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                                foreach (var sh in wdoc.Descendants(ns + "sheet"))
                                {
                                    string name = (string)sh.Attribute("name");
                                    if (!string.Equals(name, resolvedSheetName, StringComparison.OrdinalIgnoreCase)) continue;
                                    string rId = (string)sh.Attribute(rNs + "id");
                                    string targetPath = null;
                                    if (!string.IsNullOrWhiteSpace(rId) && rIdToTarget.TryGetValue(rId, out targetPath))
                                    {
                                        resolvedSheetFile = targetPath;
                                        sheetEntry = zip.GetEntry(targetPath);
                                        // Also try with/without leading xl/
                                        if (sheetEntry == null && !targetPath.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                                            sheetEntry = zip.GetEntry("xl/" + targetPath.TrimStart('/'));
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                    if (sheetEntry == null)
                    {
                        int sheetIdx = -1;
                        for (int i = 0; i < sheetNames.Count; i++) if (string.Equals(sheetNames[i], resolvedSheetName, StringComparison.OrdinalIgnoreCase)) { sheetIdx = i + 1; break; }
                        if (sheetIdx < 0) sheetIdx = 1;
                        sheetEntry = zip.GetEntry(string.Format("xl/worksheets/sheet{0}.xml", sheetIdx));
                        // Also try resolving via alternative rels path fallback enumeration
                        if (sheetEntry == null)
                        {
                            var candidates = zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).OrderBy(e => e.FullName).ToList();
                            if (candidates.Count >= sheetIdx) sheetEntry = candidates[sheetIdx - 1];
                        }
                    }
                    if (sheetEntry == null) return HostOperationResult.Failed(string.Format("Sheet '{0}' not found in source file (resolved file: {1}).", resolvedSheetName, resolvedSheetFile ?? "unknown"), 0, targetAddress);
                    using (var s = sheetEntry.Open())
                    {
                        var doc = System.Xml.Linq.XDocument.Load(s);
                        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        foreach (var row in doc.Descendants(ns + "row"))
                        {
                            foreach (var c in row.Descendants(ns + "c"))
                            {
                                string addr = (string)c.Attribute("r");
                                if (string.IsNullOrWhiteSpace(addr)) continue;
                                // parse column letters and row number from address like C5
                                int colIdx = 0;
                                int rowIdx = 0;
                                string colLetters = string.Empty;
                                string rowDigits = string.Empty;
                                for (int ai = 0; ai < addr.Length; ai++) { char ch = addr[ai]; if (char.IsLetter(ch)) colLetters += ch; else if (char.IsDigit(ch)) rowDigits += ch; }
                                if (!string.IsNullOrWhiteSpace(colLetters)) colIdx = SpreadsheetActionParser.ColumnLetterToIndex(colLetters);
                                if (!string.IsNullOrWhiteSpace(rowDigits)) int.TryParse(rowDigits, out rowIdx);
                                if (colIdx <= 0 || rowIdx <= 0) continue;
                                if (colIdx > 26 || rowIdx > 70) continue; // cap import window
                                string t = (string)c.Attribute("t");
                                string f = c.Element(ns + "f") != null ? c.Element(ns + "f").Value : null;
                                string v = c.Element(ns + "v") != null ? c.Element(ns + "v").Value : string.Empty;
                                string cellValue;
                                if (!string.IsNullOrWhiteSpace(f))
                                {
                                    // Prefer formula import; Excel will evaluate; preserve as =FORMULA
                                    cellValue = "=" + f.Trim();
                                }
                                else if (t == "s")
                                {
                                    int idx;
                                    if (int.TryParse(v, out idx) && idx >= 0 && idx < shared.Count) cellValue = shared[idx];
                                    else cellValue = v ?? string.Empty;
                                }
                                else if (t == "inlineStr")
                                    cellValue = string.Concat(c.Descendants(ns + "t").Select(x => x.Value));
                                else
                                    cellValue = v ?? string.Empty;
                                if (cellValue == null) cellValue = string.Empty;
                                sparseCells.Add(Tuple.Create(rowIdx, colIdx, cellValue));
                                if (minRow == null || rowIdx < minRow) minRow = rowIdx;
                                if (maxRow == null || rowIdx > maxRow) maxRow = rowIdx;
                                if (minCol == null || colIdx < minCol) minCol = colIdx;
                                if (maxCol == null || colIdx > maxCol) maxCol = colIdx;
                                if (sparseCells.Count >= 70 * 26) break;
                            }
                            if (sparseCells.Count >= 70 * 26) break;
                        }
                    }
                }
                if (sparseCells.Count == 0) return HostOperationResult.Failed("Source sheet is empty or could not be read.", 0, targetAddress);

                // Determine target: if targetAddress is a sheet name without A1, create new sheet
                bool targetIsSheetName = !SpreadsheetActionParser.IsSafeTarget(targetAddress) && SpreadsheetActionParser.IsSafeSheetName(targetAddress);
                dynamic workbook = app.ActiveWorkbook;
                dynamic targetWs = null;
                string targetAddrForWrite = "A1";
                if (targetIsSheetName)
                {
                    string newName = targetAddress.Trim();
                    // ensure unique
                    try
                    {
                        // check exists
                        bool exists = false;
                        try { dynamic test = app.Worksheets[newName]; exists = test != null; } catch { exists = false; }
                        if (exists) return HostOperationResult.Failed(string.Format("Worksheet '{0}' already exists.", newName), 0, targetAddress);
                        dynamic newSh = workbook.Worksheets.Add(Type.Missing, app.ActiveSheet);
                        newSh.Name = newName;
                        targetWs = newSh;
                    }
                    catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteImportWorksheet", targetAddress); }
                }
                else
                {
                    dynamic tmpApp, tmpWs, tmpRange;
                    var r = ResolveTargetRange(targetAddress, out tmpApp, out tmpWs, out tmpRange);
                    if (!r.Success) return r;
                    targetWs = tmpWs;
                    string shTmp, rngTmp;
                    SpreadsheetActionParser.TryParseSheetQualifiedTarget(targetAddress, out shTmp, out rngTmp);
                    targetAddrForWrite = !string.IsNullOrWhiteSpace(rngTmp) ? rngTmp : targetAddress;
                }
                // Write sparse cells preserving original offsets and formulas
                int startRow = 1, startCol = 1;
                if (!targetIsSheetName)
                {
                    try { startRow = Convert.ToInt32(targetWs.Range(targetAddrForWrite).Row); } catch { startRow = 1; }
                    try { startCol = Convert.ToInt32(targetWs.Range(targetAddrForWrite).Column); } catch { startCol = 1; }
                }
                int baseRow = minRow ?? 1;
                int baseCol = minCol ?? 1;
                int maxRowDelta = (maxRow ?? baseRow) - baseRow;
                int maxColDelta = (maxCol ?? baseCol) - baseCol;
                if (maxRowDelta < 0) maxRowDelta = 0;
                if (maxColDelta < 0) maxColDelta = 0;
                if (!SpreadsheetActionParser.IsSafeRangeBounds(startCol, startRow, startCol + maxColDelta, startRow + maxRowDelta))
                    return HostOperationResult.Failed("Import would exceed safe bounds.", 0, targetAddress);
                int baseDeltaRow = startRow - baseRow;
                int baseDeltaCol = startCol - baseCol;
                int written = 0;
                foreach (var cell in sparseCells)
                {
                    int deltaRow = cell.Item1 - baseRow;
                    int deltaCol = cell.Item2 - baseCol;
                    int tgtRow = startRow + deltaRow;
                    int tgtCol = startCol + deltaCol;
                    string val = cell.Item3 ?? string.Empty;
                    try
                    {
                        if (!string.IsNullOrEmpty(val) && val.StartsWith("="))
                        {
                            string translated = TranslateFormulaForImport(val, baseDeltaRow, baseDeltaCol);
                            try { targetWs.Cells[tgtRow, tgtCol].Formula = translated; }
                            catch { targetWs.Cells[tgtRow, tgtCol].Value2 = translated; }
                        }
                        else
                        {
                            targetWs.Cells[tgtRow, tgtCol].Value2 = val;
                        }
                        written++;
                    }
                    catch { }
                }
                string tgtName = Convert.ToString(targetWs.Name);
                int srcRows = maxRowDelta + 1;
                int srcCols = maxColDelta + 1;
                return HostOperationResult.Ok(string.Format("Imported {0} cells ({1} rows x {2} cols sparse) from '{3}'!{4}:{5} into '{6}'!{7} ({8} cells written, gaps preserved, formulas retained).", sparseCells.Count, srcRows, srcCols, resolvedSheetName, SpreadsheetActionParser.IndexToColumnLetter(baseCol) + baseRow, SpreadsheetActionParser.IndexToColumnLetter(maxCol ?? baseCol) + (maxRow ?? baseRow), tgtName, targetAddrForWrite, written), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteImportWorksheet failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteImportWorksheet", targetAddress);
            }
        }

        // ==================== Shape Support (Local, No Cloud) ====================
        public HostOperationResult ExecuteCreateShape(string targetAddress, string content)
        {
            // content: type:rectangle|oval|textbox; text:Hello; left:100; top:50; width:200; height:100
            string type = GetNamedOption(content, "type");
            if (string.IsNullOrWhiteSpace(type)) type = GetNamedOption(content, "shape");
            if (string.IsNullOrWhiteSpace(type)) type = "rectangle";
            type = type.Trim().ToLowerInvariant();
            string text = GetNamedOption(content, "text");
            if (string.IsNullOrWhiteSpace(text)) text = GetNamedOption(content, "label");
            int left = GetIntegerOption(content, "left", 100);
            int top = GetIntegerOption(content, "top", 50);
            int width = GetIntegerOption(content, "width", 200);
            int height = GetIntegerOption(content, "height", 80);
            if (width < 10 || width > 2000 || height < 10 || height > 2000)
                return HostOperationResult.Failed("Width/Height must be 10..2000.", 0, targetAddress);
            dynamic app, ws, dummy;
            // targetAddress may be sheet name or range; resolve sheet
            string sh, rng;
            SpreadsheetActionParser.TryParseSheetQualifiedTarget(targetAddress, out sh, out rng);
            try
            {
                app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetAddress);
                if (!string.IsNullOrWhiteSpace(sh))
                {
                    try { ws = app.Worksheets[sh]; } catch { return HostOperationResult.Failed(string.Format("Worksheet '{0}' not found.", sh), 0, targetAddress); }
                }
                else
                {
                    try { ws = app.ActiveSheet; } catch { return HostOperationResult.Failed("No active worksheet.", 0, targetAddress); }
                }
                EnsureWorksheetEditable(ws);
                int msoType = 1; // msoShapeRectangle
                if (type == "oval" || type == "circle" || type == "ellipse") msoType = 9; // msoShapeOval
                else if (type == "textbox") msoType = -1; // special
                else if (type == "rounded_rectangle") msoType = 5;
                else if (type == "diamond") msoType = 4;
                dynamic shape = null;
                if (msoType == -1)
                {
                    shape = ws.Shapes.AddTextbox(1, left, top, width, height); // msoTextOrientationHorizontal =1
                }
                else
                {
                    shape = ws.Shapes.AddShape(msoType, left, top, width, height);
                }
                if (!string.IsNullOrWhiteSpace(text))
                {
                    try { shape.TextFrame.Characters().Text = text; } catch { try { shape.TextFrame2.TextRange.Text = text; } catch { } }
                }
                // accessible alt text
                try { shape.AlternativeText = text ?? type; } catch { }
                string shapeName = string.Empty;
                try { shapeName = Convert.ToString(shape.Name); } catch { shapeName = type; }
                return HostOperationResult.Ok(string.Format("Shape '{0}' ({1}) created at {2},{3}.", shapeName, type, left, top), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteCreateShape failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteCreateShape", targetAddress);
            }
        }

        public HostOperationResult ExecuteUpdateShape(string targetAddress, string content)
        {
            // targetAddress = shape name or index? We use content shape:Name; text:NewText; fill:#FF0000
            string shapeName = GetNamedOption(content, "name");
            if (string.IsNullOrWhiteSpace(shapeName)) shapeName = GetNamedOption(content, "shape");
            if (string.IsNullOrWhiteSpace(shapeName)) shapeName = (content ?? string.Empty).Trim().Split(';')[0].Trim();
            if (string.IsNullOrWhiteSpace(shapeName)) return HostOperationResult.Failed("Shape name is required (e.g. name:Rectangle 1).", 0, targetAddress);
            string newText = GetNamedOption(content, "text");
            string fill = GetNamedOption(content, "fill");
            if (string.IsNullOrWhiteSpace(fill)) fill = GetNamedOption(content, "fill_color");
            string line = GetNamedOption(content, "line");
            dynamic app, ws, dummy;
            string sh, rng;
            SpreadsheetActionParser.TryParseSheetQualifiedTarget(targetAddress, out sh, out rng);
            try
            {
                app = _rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.", 0, targetAddress);
                if (!string.IsNullOrWhiteSpace(sh))
                {
                    try { ws = app.Worksheets[sh]; } catch { return HostOperationResult.Failed(string.Format("Worksheet '{0}' not found.", sh), 0, targetAddress); }
                }
                else
                {
                    try { ws = app.ActiveSheet; } catch { return HostOperationResult.Failed("No active worksheet.", 0, targetAddress); }
                }
                dynamic shape = null;
                try { shape = ws.Shapes[shapeName]; } catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteUpdateShape", targetAddress); }
                if (shape == null) return HostOperationResult.Failed(string.Format("Shape '{0}' not found.", shapeName), 0, targetAddress);
                if (!string.IsNullOrWhiteSpace(newText))
                {
                    try { shape.TextFrame.Characters().Text = newText; } catch { try { shape.TextFrame2.TextRange.Text = newText; } catch { } }
                }
                if (!string.IsNullOrWhiteSpace(fill))
                {
                    int c;
                    if (TryParseColor(fill, out c))
                    {
                        try { shape.Fill.ForeColor.RGB = c; } catch { try { shape.Fill.BackColor.RGB = c; } catch { } }
                    }
                }
                if (!string.IsNullOrWhiteSpace(line))
                {
                    int c2;
                    if (TryParseColor(line, out c2))
                    {
                        try { shape.Line.ForeColor.RGB = c2; } catch { }
                    }
                }
                return HostOperationResult.Ok(string.Format("Shape '{0}' updated.", shapeName), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteUpdateShape failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteUpdateShape", targetAddress);
            }
        }

        public HostOperationResult ExecuteSetWorkbookRule(string targetAddress, string content)
        {
            // content: key:value or key=value; targetAddress optionally workbook name, else active workbook
            string key = GetNamedOption(content, "key");
            if (string.IsNullOrWhiteSpace(key)) key = GetNamedOption(content, "rule");
            if (string.IsNullOrWhiteSpace(key))
            {
                string[] parts = (content ?? string.Empty).Split(new char[] { ':', '=' }, 2);
                if (parts.Length == 2) { key = parts[0].Trim(); content = parts[1].Trim(); }
            }
            string value = GetNamedOption(content, "value");
            if (string.IsNullOrWhiteSpace(value)) value = GetNamedOption(content, "val");
            if (string.IsNullOrWhiteSpace(value) && !string.IsNullOrWhiteSpace(content) && !string.IsNullOrWhiteSpace(key) && content.IndexOf(':') < 0)
                value = content.Trim();
            if (string.IsNullOrWhiteSpace(key)) return HostOperationResult.Failed("Rule key is required (e.g. key:preferred_table_style value:blue).", 0, targetAddress);
            if (string.IsNullOrWhiteSpace(value)) return HostOperationResult.Failed("Rule value is required.", 0, targetAddress);
            try
            {
                dynamic app = _rawAppObj;
                string wbName = targetAddress;
                if (string.IsNullOrWhiteSpace(wbName))
                {
                    try { wbName = Convert.ToString(app.ActiveWorkbook.Name); } catch { wbName = "DefaultWorkbook"; }
                }
                if (string.IsNullOrWhiteSpace(wbName)) wbName = "DefaultWorkbook";
                var rules = WorkbookRulesStore.LoadRules(wbName);
                rules[key.Trim()] = value.Trim();
                WorkbookRulesStore.SaveRules(wbName, rules);
                return HostOperationResult.Ok(string.Format("Rule '{0}' set for workbook '{1}'.", key.Trim(), wbName), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteSetWorkbookRule failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteSetWorkbookRule", targetAddress);
            }
        }

        public HostOperationResult ExecuteGetWorkbookRules(string targetAddress, string content)
        {
            try
            {
                dynamic app = _rawAppObj;
                string wbName = targetAddress;
                if (string.IsNullOrWhiteSpace(wbName))
                {
                    try { wbName = Convert.ToString(app.ActiveWorkbook.Name); } catch { wbName = "DefaultWorkbook"; }
                }
                if (string.IsNullOrWhiteSpace(wbName)) wbName = "DefaultWorkbook";
                string fileRules = WorkbookRulesStore.FormatRulesForPrompt(wbName);
                string sheetRules = WorkbookRulesStore.TryReadRulesSheet(this);
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                if (!string.IsNullOrWhiteSpace(sheetRules)) sb.AppendLine(sheetRules);
                if (!string.IsNullOrWhiteSpace(fileRules)) sb.AppendLine(fileRules);
                if (sb.Length == 0) return HostOperationResult.Ok(string.Format("No rules found for workbook '{0}'.", wbName), targetAddress);
                return HostOperationResult.Ok(sb.ToString().TrimEnd(), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteGetWorkbookRules failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteGetWorkbookRules", targetAddress);
            }
        }

        public HostOperationResult ExecuteClearWorkbookRules(string targetAddress, string content)
        {
            try
            {
                dynamic app = _rawAppObj;
                string wbName = targetAddress;
                if (string.IsNullOrWhiteSpace(wbName))
                {
                    try { wbName = Convert.ToString(app.ActiveWorkbook.Name); } catch { wbName = "DefaultWorkbook"; }
                }
                if (string.IsNullOrWhiteSpace(wbName)) wbName = "DefaultWorkbook";
                string path = WorkbookRulesStore.GetRulesFilePath(wbName);
                if (System.IO.File.Exists(path)) System.IO.File.Delete(path);
                return HostOperationResult.Ok(string.Format("Cleared rules for workbook '{0}'.", wbName), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteClearWorkbookRules failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteClearWorkbookRules", targetAddress);
            }
        }

        public HostOperationResult ExecuteAddSparkline(string targetAddress, string content)
        {
            if (string.IsNullOrWhiteSpace(content)) return HostOperationResult.Failed("Sparkline requires source range (e.g. source:A2:E2;type:line).", 0, targetAddress);
            string source = GetNamedOption(content, "source");
            if (string.IsNullOrWhiteSpace(source)) source = GetNamedOption(content, "data");
            if (string.IsNullOrWhiteSpace(source)) source = GetNamedOption(content, "range");
            if (string.IsNullOrWhiteSpace(source)) return HostOperationResult.Failed("Source range is required for sparkline (e.g. source:A2:E2).", 0, targetAddress);
            string type = GetNamedOption(content, "type");
            if (string.IsNullOrWhiteSpace(type)) type = "line";
            type = type.Trim().ToLowerInvariant();
            int sparkType = 1; // xlSparkLine
            if (type == "column" || type == "col") sparkType = 2;
            else if (type == "winloss" || type == "win_loss") sparkType = 3;

            dynamic app, ws, targetRange;
            var res = ResolveTargetRange(targetAddress, out app, out ws, out targetRange);
            if (!res.Success) return res;
            string sh, rng;
            SpreadsheetActionParser.TryParseSheetQualifiedTarget(source, out sh, out rng);
            dynamic srcRange = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(sh))
                {
                    dynamic srcWs = null;
                    try { srcWs = app.Worksheets[sh]; } catch { return HostOperationResult.Failed(string.Format("Source worksheet '{0}' not found.", sh), 0, targetAddress); }
                    srcRange = srcWs.Range(string.IsNullOrWhiteSpace(rng) ? source : rng);
                }
                else
                {
                    srcRange = ws.Range(source);
                }
            }
            catch (Exception ex) { return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddSparkline", targetAddress); }
            if (srcRange == null) return HostOperationResult.Failed("Could not resolve source range for sparkline.", 0, targetAddress);
            try
            {
                EnsureWorksheetEditable(ws);
                string srcAddr = string.Format("'{0}'!{1}", Convert.ToString(srcRange.Worksheet.Name).Replace("'", "''"), Convert.ToString(srcRange.Address));
                // SparklineGroups.Add signature: Add(Type, SourceData)
                try { targetRange.SparklineGroups.Add(sparkType, srcAddr); }
                catch
                {
                    // Fallback via Worksheet.SparklineGroups
                    dynamic sGroups = targetRange.SparklineGroups;
                    sGroups.Add(sparkType, srcAddr);
                }
                return HostOperationResult.Ok(string.Format("Sparkline ({0}) added at {1} from {2}.", type, targetAddress, source), targetAddress);
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.ExecuteAddSparkline failed", ex);
                return HostOperationResult.FromException(ex, "ExcelController.ExecuteAddSparkline", targetAddress);
            }
        }

        private static bool RangesOverlap(dynamic first, dynamic second)
        {
            int firstTop = Convert.ToInt32(first.Row);
            int firstLeft = Convert.ToInt32(first.Column);
            int firstBottom = firstTop + GetRangeRows(first) - 1;
            int firstRight = firstLeft + GetRangeColumns(first) - 1;
            int secondTop = Convert.ToInt32(second.Row);
            int secondLeft = Convert.ToInt32(second.Column);
            int secondBottom = secondTop + GetRangeRows(second) - 1;
            int secondRight = secondLeft + GetRangeColumns(second) - 1;
            return firstTop <= secondBottom && secondTop <= firstBottom && firstLeft <= secondRight && secondLeft <= firstRight;
        }

        private static Dictionary<string, string> ParseOptions(string content)
        {
            var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string text = (content ?? string.Empty).Trim();
            if (text.Length == 0) return options;
            string[] segments = text.Split(new[] { ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string rawSegment in segments)
            {
                string segment = rawSegment.Trim();
                int separator = segment.IndexOf(':');
                if (separator < 0) separator = segment.IndexOf('=');
                if (separator > 0)
                    options[segment.Substring(0, separator).Trim()] = segment.Substring(separator + 1).Trim();
                else if (!options.ContainsKey("value"))
                    options["value"] = segment;
            }
            return options;
        }

        private static string GetNamedOption(string content, string key)
        {
            Dictionary<string, string> options = ParseOptions(content);
            string value;
            if (options.TryGetValue(key, out value)) return value;
            if (options.TryGetValue("value", out value)) return value;
            return string.Empty;
        }

        private static int GetIntegerOption(string content, string key, int fallback)
        {
            int result;
            return int.TryParse(GetNamedOption(content, key), out result) ? result : fallback;
        }

        private static string GetValueAfterPrefix(string content, string prefix)
        {
            Dictionary<string, string> options = ParseOptions(content);
            string value;
            return options.TryGetValue(prefix, out value) ? value : string.Empty;
        }

        private static bool IsSafeExcelName(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || !ExcelNameRegex.IsMatch(value.Trim())) return false;
            return !SpreadsheetActionParser.IsSafeTarget(value.Trim()) &&
                   !string.Equals(value.Trim(), "R", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(value.Trim(), "C", StringComparison.OrdinalIgnoreCase);
        }

        private static List<int> ParseColumnList(string value, int maximum)
        {
            var columns = new List<int>();
            if (string.IsNullOrWhiteSpace(value)) return columns;
            string[] pieces = value.Split(',');
            foreach (string piece in pieces)
            {
                int number;
                if (!int.TryParse(piece.Trim(), out number) || number < 1 || number > maximum)
                    throw new InvalidOperationException("Each duplicate-removal column must be inside the target range.");
                if (!columns.Contains(number)) columns.Add(number);
            }
            return columns;
        }

        private static string TranslateFormulaForImport(string formula, int deltaRow, int deltaCol)
        {
            if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("=") || (deltaRow == 0 && deltaCol == 0)) return formula;
            try
            {
                // Split by double-quoted string literals to avoid translating inside strings
                string[] parts = formula.Split('"');
                for (int i = 0; i < parts.Length; i += 2)
                {
                    string segment = parts[i];
                    // Regex for A1 cell references with optional sheet prefix and $ absolute markers
                    // Matches optional 'Sheet'! prefix, then $col, col letters, $row, row digits
                    // Uses negative lookbehind/ahead to avoid matching inside names, but simple approach works for formulas
                    segment = Regex.Replace(segment,
                        @"(?:(?<sheet>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)\!))?(?<colAbs>\$?)(?<col>[A-Z]{1,3})(?<rowAbs>\$?)(?<row>\d+)",
                        delegate(Match m)
                        {
                            string sheet = m.Groups["sheet"].Value ?? string.Empty;
                            string colAbs = m.Groups["colAbs"].Value;
                            string colLetters = m.Groups["col"].Value;
                            string rowAbs = m.Groups["rowAbs"].Value;
                            string rowDigits = m.Groups["row"].Value;
                            // Validate col/row are plausible cell references (not named range fragment)
                            // Require at least one of them to be relative? No, we still need to handle absolute.
                            // Filter out matches where colLetters is part of a larger word: check char before/after already handled by regex boundaries via sheet handling, but we add extra check:
                            int colIdx = SpreadsheetActionParser.ColumnLetterToIndex(colLetters);
                            int rowIdx;
                            if (colIdx <= 0 || !int.TryParse(rowDigits, out rowIdx) || rowIdx < 1 || rowIdx > 1048576 || colIdx > 16384)
                                return m.Value;
                            int newCol = string.Equals(colAbs, "$", StringComparison.Ordinal) ? colIdx : colIdx + deltaCol;
                            int newRow = string.Equals(rowAbs, "$", StringComparison.Ordinal) ? rowIdx : rowIdx + deltaRow;
                            if (newCol < 1 || newCol > 16384 || newRow < 1 || newRow > 1048576)
                                return m.Value; // out of bounds -> keep original to avoid invalid ref
                            string newColLetters = SpreadsheetActionParser.IndexToColumnLetter(newCol);
                            if (string.IsNullOrEmpty(newColLetters)) return m.Value;
                            return sheet + colAbs + newColLetters + rowAbs + newRow.ToString();
                        }, RegexOptions.IgnoreCase);
                    parts[i] = segment;
                }
                return string.Join("\"", parts);
            }
            catch
            {
                return formula;
            }
        }

        public bool InsertText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return false;
                EnsureWorksheetEditable(app.ActiveSheet);

                string targetCellAddress;
                string valueToInsert = ExtractCleanExcelContent(text, out targetCellAddress);

                dynamic targetRange = null;
                if (!string.IsNullOrEmpty(targetCellAddress))
                {
                    if (!SpreadsheetActionParser.IsSafeTarget(targetCellAddress))
                        throw new InvalidOperationException(string.Format("Target address '{0}' is unsafe or outside Excel boundaries.", targetCellAddress));

                    // An explicit, validated target address must resolve or the whole call fails —
                    // silently falling back to Selection/A1 here would write to a *different* cell
                    // than the one requested with no indication anything went differently than asked.
                    targetRange = app.ActiveSheet.Range(targetCellAddress);
                    if (targetRange == null)
                    {
                        Logger.Warn(string.Format("ExcelController.InsertText: target address '{0}' resolved to null.", targetCellAddress));
                        return false;
                    }
                }
                else
                {
                    // No explicit target was given — defaulting to the current selection (or A1)
                    // is a legitimate fallback only when the caller never specified where to write.
                    try { targetRange = app.Selection; } catch { }
                    if (targetRange == null)
                    {
                        try { targetRange = app.ActiveSheet.Range("A1"); } catch { }
                    }
                }

                if (targetRange == null) return false;

                // 1. If value is an Excel formula, write Formula property
                if (valueToInsert.StartsWith("=") && !valueToInsert.Contains("\n"))
                {
                    // Guard against filling an entire multi-cell selection with a single formula
                    if (GetRangeRows(targetRange) > 1 || GetRangeColumns(targetRange) > 1)
                    {
                        try
                        {
                            if (app.ActiveCell != null) targetRange = app.ActiveCell;
                            else targetRange = targetRange.Cells[1, 1];
                        }
                        catch
                        {
                            targetRange = targetRange.Cells[1, 1];
                        }
                    }
                    targetRange.Formula = valueToInsert;
                    return true;
                }

                // 2. If text contains a markdown table (| Col 1 | Col 2 |), route through WriteTable
                var tableRows = ParseMarkdownOrCsvTable(valueToInsert);
                if (tableRows.Count > 0)
                {
                    dynamic ws = targetRange.Worksheet;
                    WriteTable(ws, targetRange, valueToInsert);
                    return true;
                }

                // 3. Otherwise write clean text/value into target cell
                targetRange.Value2 = valueToInsert;
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.InsertText failed", ex);
                return false;
            }
        }

        public static string ExtractCleanExcelContent(string rawText, out string targetCellAddress)
        {
            targetCellAddress = null;
            if (string.IsNullOrWhiteSpace(rawText)) return string.Empty;

            string text = rawText.Trim();

            // 1. Detect target cell from explicit cell directives (e.g. "cell K19", "in cell K19", "to cell K19", "Range("K19")")
            var matchCell = Regex.Match(text, @"(?:Range\(\""|\b(?:cell|in\s+cell|to\s+cell)\s+)([A-Za-z]{1,3}\d{1,7})\b", RegexOptions.IgnoreCase);
            if (matchCell.Success)
            {
                targetCellAddress = matchCell.Groups[1].Value.ToUpperInvariant();
            }

            // 2. Check for VBA Range("K19").Value = "1225" pattern
            var matchVba = Regex.Match(text, @"Range\(\""([A-Za-z]{1,3}\d{1,7})\""\)\.Value\s*=\s*\""?([^""\r\n]+)\""?", RegexOptions.IgnoreCase);
            if (matchVba.Success)
            {
                targetCellAddress = matchVba.Groups[1].Value.ToUpperInvariant();
                return matchVba.Groups[2].Value.Trim();
            }

            // 3. Check if there is a fenced code block ```...```
            var matchCodeBlock = Regex.Match(text, @"```(?:excel|csv|tsv|text)?\s*([\s\S]*?)\s*```", RegexOptions.IgnoreCase);
            if (matchCodeBlock.Success)
            {
                return matchCodeBlock.Groups[1].Value.Trim();
            }

            // 4. Check if there is a Markdown Table (| Col 1 | Col 2 |)
            var tableLines = new List<string>();
            var allLines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var l in allLines)
            {
                string tl = l.Trim();
                if (tl.StartsWith("|") && tl.EndsWith("|"))
                {
                    tableLines.Add(tl);
                }
            }
            if (tableLines.Count > 0)
            {
                return string.Join("\n", tableLines);
            }

            // 5. Strip conversational preamble and postamble instructions
            var cleanLines = new List<string>();
            bool inInstruction = false;
            foreach (var l in allLines)
            {
                string tl = l.Trim();
                if (Regex.IsMatch(tl, @"^(?:Here(?:'s| is)|Sure|Okay|Certainly|Below is|Note:)", RegexOptions.IgnoreCase))
                    continue;
                if (Regex.IsMatch(tl, @"^(?:How to insert|Method \d|Instructions|Steps to|To insert|Verification|Let me know)", RegexOptions.IgnoreCase))
                {
                    inInstruction = true;
                    continue;
                }
                if (inInstruction) continue;

                // Strip markdown bold/ticks
                string cleaned = Regex.Replace(tl, @"\*\*([^*]+)\*\*", "$1").Replace("`", "").Trim();
                if (!string.IsNullOrWhiteSpace(cleaned))
                {
                    cleanLines.Add(cleaned);
                }
            }

            if (cleanLines.Count > 0)
            {
                return string.Join("\n", cleanLines);
            }

            // 6. Fallback: extract single number or formula if present
            var matchNumOrFormula = Regex.Match(text, @"(?m)^(?:=[\w\(\):,]+|\d+(?:\.\d+)?)$");
            if (matchNumOrFormula.Success)
            {
                return matchNumOrFormula.Value.Trim();
            }

            return text;
        }

        private List<List<string>> ParseMarkdownOrCsvTable(string rawText)
        {
            var result = new List<List<string>>();
            if (string.IsNullOrWhiteSpace(rawText)) return result;

            var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Skip markdown separator row |---|---|
                if (Regex.IsMatch(line, @"^\|?\s*[-:]+\s*\|[\s-:|]*$")) continue;

                if (line.StartsWith("|") && line.EndsWith("|"))
                {
                    var parts = line.Substring(1, line.Length - 2).Split('|');
                    var row = new List<string>();
                    foreach (var p in parts)
                    {
                        string cell = p.Trim();
                        // Clean bold / italic markdown
                        cell = Regex.Replace(cell, @"\*\*([^*]+)\*\*", "$1");
                        cell = Regex.Replace(cell, @"\*([^*]+)\*", "$1");
                        row.Add(cell);
                    }
                    if (row.Count > 0) result.Add(row);
                }
            }

            return result;
        }

        public string GetActiveWorkbookName()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app != null && app.ActiveWorkbook != null)
                    return Convert.ToString(app.ActiveWorkbook.Name);
            }
            catch { }
            return "ExcelWorkbook";
        }

        /// <summary>
        /// Navigate to a specific cell by sheet name and cell address (A1 notation).
        /// If sheetName is null or empty, uses the active sheet.
        /// Defensive: returns false on any error without throwing.
        /// </summary>
        public bool NavigateToCell(string sheetName, string cellAddress)
        {
            try
            {
                // Validate cell address first
                if (string.IsNullOrEmpty(cellAddress) || !SpreadsheetActionParser.IsSafeTarget(cellAddress))
                    return false;

                dynamic app = _rawAppObj;
                if (app == null)
                    return false;

                dynamic ws = null;
                if (string.IsNullOrEmpty(sheetName))
                {
                    // Use active sheet
                    try { ws = app.ActiveSheet; } catch { }
                }
                else
                {
                    // Activate the specified worksheet by name
                    try { ws = app.Worksheets[sheetName]; } catch { }
                    if (ws != null)
                    {
                        try { ws.Activate(); } catch { }
                    }
                }

                if (ws == null)
                    return false;

                // Select and activate the cell
                dynamic targetRange = null;
                try { targetRange = ws.Range(cellAddress); } catch { }
                if (targetRange == null)
                    return false;

                // Select is the actual navigation — must not be silently swallowed, or a real
                // failure here (e.g. protected sheet) would still report success. Let it
                // propagate to the outer catch, which correctly logs and returns false.
                // Activate() is secondary window-focus polish and stays optional.
                targetRange.Select();
                try { targetRange.Activate(); } catch { }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelController.NavigateToCell failed: {0}", ex.Message));
                return false;
            }
        }
    }
}
