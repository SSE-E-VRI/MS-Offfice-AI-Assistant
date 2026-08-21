using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
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

        public ExcelController(object appObj)
        {
            _rawAppObj = appObj;
        }

        public string GetActiveDocumentName()
        {
            return GetActiveWorkbookName();
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

        private HostOperationResult ResolveTargetRange(string targetAddress, out dynamic app, out dynamic ws, out dynamic targetRange)
        {
            app = null;
            ws = null;
            targetRange = null;

            if (string.IsNullOrEmpty(targetAddress) || !SpreadsheetActionParser.IsSafeTarget(targetAddress))
                return HostOperationResult.Failed(string.Format("Invalid or unsafe cell/range address: '{0}'.", targetAddress), 0, targetAddress);

            try
            {
                app = _rawAppObj;
                if (app == null)
                    return HostOperationResult.Failed("No active Excel workbook found.", 0, targetAddress);

                try { ws = app.ActiveSheet; } catch { }
                if (ws == null)
                    return HostOperationResult.Failed("No active Excel worksheet found.", 0, targetAddress);

                EnsureWorksheetEditable(ws);

                try { targetRange = ws.Range(targetAddress); } catch { }
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

        public bool ApplySpreadsheetAction(SpreadsheetAction action)
        {
            if (action == null) return false;
            action.Status = SpreadsheetActionStatus.Applying;
            action.ErrorMessage = string.Empty;
            action.ResultText = string.Empty;

            var res = ExecuteSpreadsheetAction(action);
            return res.Success;
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

        public string CreatePreviewDescription(SpreadsheetAction action)
        {
            if (action == null) return "No spreadsheet action is available.";
            string description = string.IsNullOrWhiteSpace(action.Description) ? "No additional description provided." : action.Description;
            return string.Format("Excel will apply {0} to {1}.\n\nDescription: {2}\n\nConfiguration:\n{3}",
                action.Type, action.Target, description, action.Content ?? string.Empty);
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
                throw new InvalidOperationException("Use color_scale, data_bar, greater_than:value, less_than:value, or equal_to:value for conditional formatting.");

            dynamic condition = formats.Add(1, comparison, criterion);
            try { condition.Interior.Color = 13551615; } catch { }
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
            string rangeValues = !string.IsNullOrWhiteSpace(wholeNumbers) ? wholeNumbers : (!string.IsNullOrWhiteSpace(decimals) ? decimals : dates);
            if (string.IsNullOrWhiteSpace(rangeValues))
                throw new InvalidOperationException("Use list:values, whole_number:min,max, decimal:min,max, or date:start,end for data validation.");
            string[] bounds = rangeValues.Split(',');
            if (bounds.Length != 2) throw new InvalidOperationException("Range validation requires a minimum and maximum value.");
            int validationType = !string.IsNullOrWhiteSpace(wholeNumbers) ? 1 : (!string.IsNullOrWhiteSpace(decimals) ? 2 : 4);
            validation.Add(validationType, 1, 1, bounds[0].Trim(), bounds[1].Trim());
            return "Data validation applied";
        }

        private static string CreateChart(dynamic worksheet, dynamic targetRange, string content)
        {
            if (GetRangeRows(targetRange) < 2 || GetRangeColumns(targetRange) < 1)
                throw new InvalidOperationException("A chart requires a header row and data.");
            string chartType = (content ?? string.Empty).Trim().ToLowerInvariant();
            int excelChartType = chartType == "line" ? 4 : (chartType == "pie" ? 5 : (chartType == "bar" ? 57 : 51));
            double left = Convert.ToDouble(targetRange.Left) + Convert.ToDouble(targetRange.Width) + 18;
            double top = Convert.ToDouble(targetRange.Top);
            dynamic chartObject = worksheet.ChartObjects().Add(left, top, 420, 260);
            chartObject.Chart.SetSourceData(targetRange);
            chartObject.Chart.ChartType = excelChartType;
            return "Chart created";
        }

        private static string CreatePivotTable(dynamic app, dynamic worksheet, dynamic targetRange, string content)
        {
            if (GetRangeRows(targetRange) < 2 || GetRangeColumns(targetRange) < 2)
                throw new InvalidOperationException("A PivotTable requires a tabular source with headers and data.");
            EnsureUniqueNonBlankHeaders(targetRange, GetRangeColumns(targetRange));
            string destinationAddress = GetNamedOption(content, "destination");
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
            cache.CreatePivotTable(destination, pivotName);
            return "PivotTable created: " + pivotName;
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
                    try { targetRange = app.ActiveSheet.Range(targetCellAddress); } catch { }
                }

                if (targetRange == null)
                {
                    try { targetRange = app.Selection; } catch { }
                }
                if (targetRange == null)
                {
                    try { targetRange = app.ActiveSheet.Range("A1"); } catch { }
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

        public void WriteFormula(string formula, string cellAddress = null)
        {
            if (string.IsNullOrEmpty(formula)) return;

            if (!string.IsNullOrEmpty(cellAddress) && !SpreadsheetActionParser.IsSafeTarget(cellAddress))
            {
                Logger.Warn(string.Format("ExcelController.WriteFormula: Rejected unsafe cell address: {0}", cellAddress));
                return;
            }

            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return;

                dynamic targetRange = null;
                if (!string.IsNullOrEmpty(cellAddress))
                {
                    targetRange = app.ActiveSheet.Range(cellAddress);
                }
                else
                {
                    targetRange = app.Selection;
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
    }
}
