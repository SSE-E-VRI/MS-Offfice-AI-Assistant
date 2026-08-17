using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.Hosts
{
    public class ExcelController
    {
        private readonly object _rawAppObj;

        public ExcelController(object appObj)
        {
            _rawAppObj = appObj;
        }

        public static string IndexToColumnLetter(int colIndex)
        {
            if (colIndex <= 0) return "A";
            int div = colIndex;
            string colLetter = string.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
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

        public bool ApplySpreadsheetAction(SpreadsheetAction action)
        {
            if (action == null || string.IsNullOrWhiteSpace(action.Target))
                return false;

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveSheet == null)
                {
                    action.Status = SpreadsheetActionStatus.Error;
                    action.ErrorMessage = "No active Excel worksheet found.";
                    return false;
                }

                dynamic ws = app.ActiveSheet;
                dynamic targetRange = ws.Range(action.Target);
                if (targetRange == null)
                {
                    action.Status = SpreadsheetActionStatus.Error;
                    action.ErrorMessage = string.Format("Invalid range target: {0}", action.Target);
                    return false;
                }

                action.Status = SpreadsheetActionStatus.Applying;

                switch (action.Type)
                {
                    case SpreadsheetActionType.Formula:
                        string formula = action.Content.Trim();
                        if (!formula.StartsWith("=")) formula = "=" + formula;
                        targetRange.Formula = formula;

                        // Read evaluated calculation result
                        string evalResult = "";
                        try
                        {
                            evalResult = Convert.ToString(targetRange.Text) ?? Convert.ToString(targetRange.Value2);
                        }
                        catch { }

                        action.ResultText = !string.IsNullOrEmpty(evalResult) ? evalResult : "OK";
                        action.Status = SpreadsheetActionStatus.Applied;
                        Logger.Info(string.Format("Applied formula '{0}' to {1}. Evaluated value: {2}", formula, action.Target, evalResult));
                        return true;

                    case SpreadsheetActionType.FillDown:
                        string fillFormula = action.Content.Trim();
                        if (!fillFormula.StartsWith("=")) fillFormula = "=" + fillFormula;
                        targetRange.Formula = fillFormula;

                        action.ResultText = string.Format("Filled {0}", action.Target);
                        action.Status = SpreadsheetActionStatus.Applied;
                        Logger.Info(string.Format("Applied filldown formula '{0}' to range {1}", fillFormula, action.Target));
                        return true;

                    case SpreadsheetActionType.Table:
                        var tableRows = ParseMarkdownOrCsvTable(action.Content);
                        if (tableRows.Count > 0)
                        {
                            int startRow = Convert.ToInt32(targetRange.Row);
                            int startCol = Convert.ToInt32(targetRange.Column);

                            for (int r = 0; r < tableRows.Count; r++)
                            {
                                var row = tableRows[r];
                                for (int c = 0; c < row.Count; c++)
                                {
                                    try
                                    {
                                        ws.Cells[startRow + r, startCol + c].Value2 = row[c];
                                    }
                                    catch { }
                                }
                            }
                        }
                        action.ResultText = "Table Written";
                        action.Status = SpreadsheetActionStatus.Applied;
                        return true;

                    default: // Value
                        targetRange.Value2 = action.Content;
                        action.ResultText = action.Content;
                        action.Status = SpreadsheetActionStatus.Applied;
                        return true;
                }
            }
            catch (Exception ex)
            {
                action.Status = SpreadsheetActionStatus.Error;
                action.ErrorMessage = ex.Message;
                Logger.Error(string.Format("ApplySpreadsheetAction failed on target {0}", action.Target), ex);
                return false;
            }
        }

        public void InsertText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return;

                string targetCellAddress;
                string valueToInsert = ExtractCleanExcelContent(text, out targetCellAddress);

                dynamic targetRange = null;
                if (!string.IsNullOrEmpty(targetCellAddress))
                {
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

                if (targetRange == null) return;

                // 1. If value is an Excel formula, write Formula property
                if (valueToInsert.StartsWith("=") && !valueToInsert.Contains("\n"))
                {
                    targetRange.Formula = valueToInsert;
                    return;
                }

                // 2. If text contains a markdown table (| Col 1 | Col 2 |), parse into cells
                var tableRows = ParseMarkdownOrCsvTable(valueToInsert);
                if (tableRows.Count > 0)
                {
                    int startRow = Convert.ToInt32(targetRange.Row);
                    int startCol = Convert.ToInt32(targetRange.Column);
                    dynamic ws = targetRange.Worksheet;

                    for (int r = 0; r < tableRows.Count; r++)
                    {
                        var row = tableRows[r];
                        for (int c = 0; c < row.Count; c++)
                        {
                            try
                            {
                                ws.Cells[startRow + r, startCol + c].Value2 = row[c];
                            }
                            catch { }
                        }
                    }
                    return;
                }

                // 3. Otherwise write clean text/value into target cell
                targetRange.Value2 = valueToInsert;
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelController.InsertText failed", ex);
                throw;
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
