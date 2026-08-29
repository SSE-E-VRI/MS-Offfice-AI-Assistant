using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Hosts
{
    /// <summary>
    /// E4.3 Change highlighting — mirrors Copilot's "green tab + grid" visual cue.
    /// IMPORTANT: This touches only Excel COM (Tab.Color + Interior.Color). It deliberately
    /// does NOT touch WPF focus plumbing at ChatSidebar.xaml.cs:2009-2069 (cached HWND + SetFocus),
    /// which must remain untouched per review verdict.
    /// Highlights are best-effort, never block mutation success, and are clearable via
    /// excel.clear_highlights. RollbackExecutor can also clear highlights on rollback batch.
    /// </summary>
    public static class ExcelChangeHighlighter
    {
        private static int _highlightColor;
        private static int _tabColor;
        private static bool _colorsInitialized;
        // Keyed by "WorkbookName!SheetName" (not bare sheet name) so highlights from one
        // workbook can never be cleared against a different, unrelated workbook that
        // happens to share a default sheet name (e.g. "Sheet1").
        private static readonly HashSet<string> TrackedSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static string TrackKey(string workbookName, string sheetName)
        {
            return (workbookName ?? string.Empty) + "!" + (sheetName ?? string.Empty);
        }

        private static void EnsureColors()
        {
            if (_colorsInitialized) return;
            int hl, tab;
            // Light green fill like Copilot's change highlight
            if (!TryParseHex("#C6EFCE", out hl)) hl = 0xCEF6C6; // fallback OLE for #C6EFCE
            if (!TryParseHex("#00B050", out tab)) tab = 0x50B000; // green tab
            _highlightColor = hl;
            _tabColor = tab;
            _colorsInitialized = true;
        }

        private static bool TryParseHex(string hex, out int ole)
        {
            ole = 0;
            if (string.IsNullOrWhiteSpace(hex)) return false;
            string s = hex.Trim().TrimStart('#');
            if (s.StartsWith("0x", StringComparison.OrdinalIgnoreCase)) s = s.Substring(2);
            if (s.Length != 6) return false;
            try
            {
                int r = Convert.ToInt32(s.Substring(0, 2), 16);
                int g = Convert.ToInt32(s.Substring(2, 2), 16);
                int b = Convert.ToInt32(s.Substring(4, 2), 16);
                ole = r + (g << 8) + (b << 16);
                return true;
            }
            catch { return false; }
        }

        public static void ApplyHighlight(dynamic ws, dynamic targetRange)
        {
            if (ws == null || targetRange == null) return;
            EnsureColors();
            try
            {
                // Mark sheet tab green
                try
                {
                    ws.Tab.Color = _tabColor;
                    string sheetName = Convert.ToString(ws.Name);
                    string workbookName = null;
                    try { workbookName = Convert.ToString(ws.Parent.Name); } catch { }
                    if (!string.IsNullOrWhiteSpace(sheetName) && !string.IsNullOrWhiteSpace(workbookName))
                        lock (TrackedSheets) { TrackedSheets.Add(TrackKey(workbookName, sheetName)); }
                }
                catch (Exception ex) { Logger.Warn(string.Format("ExcelChangeHighlighter tab color failed: {0}", ex.Message)); }
                // Mark grid range light green (best-effort; format mutations may overwrite)
                try { targetRange.Interior.Color = _highlightColor; } catch (Exception ex) { Logger.Warn(string.Format("ExcelChangeHighlighter interior failed: {0}", ex.Message)); }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExcelChangeHighlighter.ApplyHighlight failed: {0}", ex.Message));
            }
        }

        public static HostOperationResult ClearHighlights(object rawAppObj)
        {
            try
            {
                dynamic app = rawAppObj;
                if (app == null) return HostOperationResult.Failed("No active Excel workbook found.");
                dynamic workbook = null;
                try { workbook = app.ActiveWorkbook; } catch { }
                if (workbook == null) return HostOperationResult.Failed("No active workbook.");
                string activeWorkbookName = null;
                try { activeWorkbookName = Convert.ToString(workbook.Name); } catch { }
                // Clear tab colors only for sheets tracked under the active workbook; leave
                // entries for other (still-open or already-closed) workbooks untouched.
                List<string> toClear = new List<string>();
                List<string> sheetNamesForActiveWorkbook = new List<string>();
                lock (TrackedSheets)
                {
                    foreach (string key in TrackedSheets)
                    {
                        if (!string.IsNullOrWhiteSpace(activeWorkbookName) &&
                            key.StartsWith(activeWorkbookName + "!", StringComparison.OrdinalIgnoreCase))
                        {
                            toClear.Add(key);
                            sheetNamesForActiveWorkbook.Add(key.Substring(activeWorkbookName.Length + 1));
                        }
                    }
                    foreach (string key in toClear) TrackedSheets.Remove(key);
                }
                toClear = sheetNamesForActiveWorkbook;
                if (toClear.Count == 0)
                {
                    try
                    {
                        dynamic ws = app.ActiveSheet;
                        if (ws != null)
                        {
                            try { ws.Tab.ColorIndex = -4142; } catch { try { ws.Tab.Color = 0; } catch { } } // xlColorIndexNone
                            try { ws.Cells.Interior.ColorIndex = -4142; } catch { }
                        }
                    }
                    catch { }
                }
                else
                {
                    foreach (string name in toClear)
                    {
                        try
                        {
                            dynamic ws = app.Worksheets[name];
                            if (ws != null)
                            {
                                try { ws.Tab.ColorIndex = -4142; } catch { try { ws.Tab.Color = 0; } catch { } }
                            }
                        }
                        catch { }
                    }
                    // Also clear interior highlights on active sheet's used range (best-effort)
                    try
                    {
                        dynamic ws = app.ActiveSheet;
                        dynamic used = ws.UsedRange;
                        if (used != null) { try { used.Interior.ColorIndex = -4142; } catch { } }
                    }
                    catch { }
                }
                return HostOperationResult.Ok("Cleared change highlights.");
            }
            catch (Exception ex)
            {
                Logger.Error("ExcelChangeHighlighter.ClearHighlights failed", ex);
                return HostOperationResult.FromException(ex, "ExcelChangeHighlighter.ClearHighlights");
            }
        }
    }
}
