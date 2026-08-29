using System;
using System.Collections.Generic;
using System.Linq;
using MSOfficeAIAssistant.Core.Actions;
using MSOfficeAIAssistant.Hosts;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Central registry of all host capabilities and tools per SSOT §5.2 and §5.3.
    /// Provides dynamic allow-list generation (D-5), schema validation, execution dispatching, and host guarding.
    /// </summary>
    public static class ToolRegistry
    {
        private static readonly Dictionary<string, ToolDefinition> ToolsByName =
            new Dictionary<string, ToolDefinition>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, ToolDefinition> ToolsByAlias =
            new Dictionary<string, ToolDefinition>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, ToolDefinition> ToolsByHostAndAlias =
            new Dictionary<string, ToolDefinition>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, List<ToolDefinition>> ToolsByHost =
            new Dictionary<string, List<ToolDefinition>>(StringComparer.OrdinalIgnoreCase);

        static ToolRegistry()
        {
            RegisterAllBuiltInTools();
        }

        public static void Register(ToolDefinition tool)
        {
            if (tool == null || string.IsNullOrWhiteSpace(tool.Name)) return;

            ToolsByName[tool.Name] = tool;
            ToolsByAlias[tool.Name] = tool;

            string host = tool.Host ?? "General";
            ToolsByHostAndAlias[host + ":" + tool.Name] = tool;

            if (tool.Aliases != null)
            {
                foreach (var alias in tool.Aliases)
                {
                    if (!string.IsNullOrWhiteSpace(alias))
                    {
                        ToolsByAlias[alias] = tool;
                        ToolsByHostAndAlias[host + ":" + alias] = tool;
                    }
                }
            }

            if (!ToolsByHost.ContainsKey(host))
            {
                ToolsByHost[host] = new List<ToolDefinition>();
            }
            if (!ToolsByHost[host].Any(t => string.Equals(t.Name, tool.Name, StringComparison.OrdinalIgnoreCase)))
            {
                ToolsByHost[host].Add(tool);
            }
        }

        public static ToolDefinition GetTool(string nameOrAlias, string host = null)
        {
            ToolDefinition tool;
            if (TryGetTool(nameOrAlias, out tool, host))
            {
                return tool;
            }
            return null;
        }

        public static bool TryGetTool(string nameOrAlias, out ToolDefinition tool, string host = null)
        {
            tool = null;
            if (string.IsNullOrWhiteSpace(nameOrAlias)) return false;

            if (!string.IsNullOrEmpty(host))
            {
                string hostKey = host + ":" + nameOrAlias;
                if (ToolsByHostAndAlias.TryGetValue(hostKey, out tool)) return true;
                if (ToolsByName.TryGetValue(nameOrAlias, out tool) && string.Equals(tool.Host, host, StringComparison.OrdinalIgnoreCase)) return true;
            }

            if (ToolsByName.TryGetValue(nameOrAlias, out tool)) return true;
            if (ToolsByAlias.TryGetValue(nameOrAlias, out tool)) return true;

            return false;
        }

        public static IReadOnlyList<ToolDefinition> GetToolsForHost(string host)
        {
            if (string.IsNullOrWhiteSpace(host)) return new List<ToolDefinition>();
            List<ToolDefinition> list;
            if (ToolsByHost.TryGetValue(host, out list))
            {
                return list.AsReadOnly();
            }
            return new List<ToolDefinition>();
        }

        public static IReadOnlyList<ToolDefinition> GetAllTools()
        {
            return ToolsByName.Values.ToList().AsReadOnly();
        }

        /// <summary>
        /// Dispatches an action's execution to its registered tool handler, enforcing strict host and controller type guards.
        /// </summary>
        public static HostOperationResult Execute(object controller, OfficeAction action)
        {
            if (action == null)
            {
                return HostOperationResult.Failed("Action cannot be null.");
            }

            if (controller == null)
            {
                return HostOperationResult.Failed("Controller cannot be null.", 0, action.TargetDisplay);
            }

            string host = !string.IsNullOrEmpty(action.Host) ? action.Host : null;
            ToolDefinition tool = GetTool(action.Operation, host);
            if (tool == null)
            {
                return HostOperationResult.Failed(
                    string.Format("Unknown tool or operation '{0}' for host '{1}'.", action.Operation, host ?? "unspecified"),
                    0, action.TargetDisplay);
            }

            // Host Guard: verify action.Host matches tool.Host if specified
            if (!string.IsNullOrEmpty(action.Host) && !string.Equals(action.Host, tool.Host, StringComparison.OrdinalIgnoreCase))
            {
                return HostOperationResult.Failed(
                    string.Format("Host mismatch: action specifies host '{0}' but tool '{1}' belongs to host '{2}'.", action.Host, tool.Name, tool.Host),
                    0, action.TargetDisplay);
            }

            // Controller Type Guard: verify controller matches expected host controller type
            string ctrlTypeName = controller.GetType().Name;
            if (string.Equals(tool.Host, "Excel", StringComparison.OrdinalIgnoreCase))
            {
                if (!(controller is ExcelController) && !ctrlTypeName.Contains("Excel"))
                {
                    return HostOperationResult.Failed(
                        string.Format("Controller type mismatch: tool '{0}' requires Excel controller but received '{1}'.", tool.Name, ctrlTypeName),
                        0, action.TargetDisplay);
                }
            }
            else if (string.Equals(tool.Host, "PowerPoint", StringComparison.OrdinalIgnoreCase))
            {
                if (!(controller is PowerPointController) && !ctrlTypeName.Contains("PowerPoint"))
                {
                    return HostOperationResult.Failed(
                        string.Format("Controller type mismatch: tool '{0}' requires PowerPoint controller but received '{1}'.", tool.Name, ctrlTypeName),
                        0, action.TargetDisplay);
                }
            }
            else if (string.Equals(tool.Host, "Word", StringComparison.OrdinalIgnoreCase))
            {
                if (!(controller is WordController) && !ctrlTypeName.Contains("Word"))
                {
                    return HostOperationResult.Failed(
                        string.Format("Controller type mismatch: tool '{0}' requires Word controller but received '{1}'.", tool.Name, ctrlTypeName),
                        0, action.TargetDisplay);
                }
            }

            if (tool.Handler == null)
            {
                return HostOperationResult.Failed(
                    string.Format("No execution handler registered for tool '{0}'.", tool.Name),
                    0, action.TargetDisplay);
            }

            try
            {
                return tool.Handler(controller, action);
            }
            catch (Exception ex)
            {
                return HostOperationResult.FromException(ex, tool.Name, action.TargetDisplay);
            }
        }

        /// <summary>
        /// Returns the standard list of action type keywords for prompt generation (resolving D-5).
        /// </summary>
        public static string FormatActionTypesList(string host)
        {
            if (string.Equals(host, "Excel", StringComparison.OrdinalIgnoreCase))
            {
                return "formula, value, filldown, table, create_table, conditional_format, sort, filter, data_validation, chart, pivot_table, named_range, remove_duplicates, find_replace, set_case, trim_range, normalize_whitespace, text_to_columns, add_worksheet, rename_worksheet, delete_worksheet, duplicate_worksheet, set_tab_color, insert_rows, delete_rows, insert_columns, delete_columns, hide_rows, unhide_rows, hide_columns, unhide_columns, merge_cells, format_cells, autofit_columns, freeze_panes, add_summary_row, write_python, apply_theme, clear_highlights, add_sparkline, analyze_range, get_formula_details, add_analysis_column, import_worksheet, create_shape, update_shape, set_workbook_rule, get_workbook_rules, and clear_workbook_rules";
            }
            if (string.Equals(host, "PowerPoint", StringComparison.OrdinalIgnoreCase))
            {
                return "move_slide, create_section, rename_section, set_notes, create_slide, insert_image, delete_slide, duplicate_slide, hide_slide, unhide_slide, apply_layout, set_shape_text, replace_text, add_table, add_chart, add_shape, set_font, fit_content, translate_deck, audit_deck, audit_alt_text, and set_alt_text";
            }
            if (string.Equals(host, "Word", StringComparison.OrdinalIgnoreCase))
            {
                return "add_comment, list_comments, delete_comment, edit_comment, list_revisions, accept_revision, reject_revision, compare_documents, translate, insert_table, format_table, find_replace, apply_style, set_font, set_paragraph_format, set_case, reorganize_paragraphs, normalize_whitespace, insert_break, set_page_setup, set_header_footer, insert_page_number, insert_hyperlink, insert_bookmark, insert_image, insert_toc, update_toc, export_pdf, save_as, toggle_track_changes, list_styles, set_proofing_language, merge_document, set_watermark, insert_caption, delete, apply_list, and readability_stats";
            }
            return string.Empty;
        }

        /// <summary>
        /// Generates OpenAI-compatible function schema definitions for all tools registered for the given host.
        /// </summary>
        public static List<Dictionary<string, object>> GetOpenAiToolsForHost(string host)
        {
            var result = new List<Dictionary<string, object>>();
            var tools = GetToolsForHost(host);
            foreach (var tool in tools)
            {
                result.Add(tool.ToOpenAiFunctionSchema());
            }
            return result;
        }

        private static void RegisterAllBuiltInTools()
        {
            // === EXCEL TOOLS ===
            Register(new ToolDefinition("excel.write_formula", "Excel", "Writes or updates a calculation formula in an Excel cell or range.", 2, true, true)
                .WithParameter("target", "string", "Target cell or range address (e.g. 'K20' or 'B2:B100')", true)
                .WithParameter("formula", "string", "The Excel formula starting with '=' (e.g. '=SUM(A1:A10)')", true)
                .WithParameter("description", "string", "Explanation of what the formula computes")
                .WithAlias("formula")
                .WithAlias("write_formula")
                .WithAlias("excel_write_formula")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteWriteFormula(act.GetParameterString("formula") ?? act.ContentDisplay, act.TargetDisplay)));

            Register(new ToolDefinition("excel.write_value", "Excel", "Writes literal text, numbers, or dates into an Excel cell or range.", 1, true, true)
                .WithParameter("target", "string", "Target cell or range address (e.g. 'A1' or 'B2:D10')", true)
                .WithParameter("value", "string", "Literal content to write", true)
                .WithParameter("description", "string", "Explanation of the value")
                .WithAlias("value")
                .WithAlias("write_value")
                .WithAlias("excel_write_value")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteWriteValue(act.GetParameterString("value") ?? act.ContentDisplay, act.TargetDisplay)));

            Register(new ToolDefinition("excel.fill_down", "Excel", "Propagates a formula downward through a designated column range.", 2, true, true)
                .WithParameter("target", "string", "Target column range (e.g. 'G2:G27')", true)
                .WithParameter("formula", "string", "Formula template for the top cell", true)
                .WithParameter("description", "string", "Explanation of the fill operation")
                .WithAlias("filldown")
                .WithAlias("fill_down")
                .WithAlias("excel_fill_down")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteFillDown(act.TargetDisplay, act.GetParameterString("formula") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.table", "Excel", "Formats an existing range as an Excel table.", 2, true, true)
                .WithParameter("target", "string", "Target data range (e.g. 'A1:E50')", true)
                .WithParameter("description", "string", "Explanation of the table")
                .WithAlias("table")
                .WithAlias("excel_table")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteTable(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.create_table", "Excel", "Creates and formats a new Excel ListObject table.", 2, true, false)
                .WithParameter("target", "string", "Target data range (e.g. 'A1:E50')", true)
                .WithParameter("name", "string", "Optional table name")
                .WithParameter("description", "string", "Explanation of the table")
                .WithAlias("create_table")
                .WithAlias("createtable")
                .WithAlias("excel_create_table")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteCreateTable(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.conditional_format", "Excel", "Applies conditional formatting rules to a range (supports custom color: hex #RRGGBB).", 2, true, true)
                .WithParameter("target", "string", "Target cell or range address", true)
                .WithParameter("rule", "string", "Formatting rule (e.g. 'highlight_gt:50000' or 'top_n:10;color:#FF0000')", true)
                .WithParameter("color", "string", "Optional highlight color hex #RRGGBB (default pink)")
                .WithParameter("description", "string", "Explanation of the formatting")
                .WithAlias("conditional_format")
                .WithAlias("conditionalformat")
                .WithAlias("format")
                .WithAlias("excel_conditional_format")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteConditionalFormat(act.TargetDisplay, act.GetParameterString("rule") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.sort", "Excel", "Sorts a range by a specified column and direction.", 2, true, true)
                .WithParameter("target", "string", "Target range to sort", true)
                .WithParameter("order", "string", "Sort direction ('ascending' or 'descending')", false, "ascending")
                .WithParameter("column", "integer", "1-based column index to sort by", false, 1)
                .WithAlias("sort")
                .WithAlias("excel_sort")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteSort(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.filter", "Excel", "Applies AutoFilter criteria to a worksheet table or range.", 2, true, true)
                .WithParameter("target", "string", "Target range with headers", true)
                .WithParameter("criteria", "string", "Filter criteria (e.g. 'ColB:>100')", true)
                .WithAlias("filter")
                .WithAlias("excel_filter")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteFilter(act.TargetDisplay, act.GetParameterString("criteria") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.data_validation", "Excel", "Configures input validation rules (e.g. dropdown lists).", 2, true, true)
                .WithParameter("target", "string", "Target cell or range address", true)
                .WithParameter("rule", "string", "Validation rule (e.g. 'list:Open,Closed,Pending')", true)
                .WithAlias("data_validation")
                .WithAlias("datavalidation")
                .WithAlias("validate")
                .WithAlias("excel_data_validation")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteDataValidation(act.TargetDisplay, act.GetParameterString("rule") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.create_chart", "Excel", "Inserts an Excel chart based on a source range.", 2, true, false)
                .WithParameter("target", "string", "Source data range for the chart", true)
                .WithParameter("chart_type", "string", "Chart type: column/bar/line/pie/scatter/area/doughnut/stacked", false, "column")
                .WithParameter("title", "string", "Chart title")
                .WithAlias("chart")
                .WithAlias("create_chart")
                .WithAlias("excel_create_chart")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteCreateChart(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.create_pivot_table", "Excel", "Creates a Pivot Table summary from a source table with optional field config.", 2, true, false)
                .WithParameter("target", "string", "Source data range", true)
                .WithParameter("destination", "string", "Target cell for pivot top-left (e.g. H2)", true)
                .WithParameter("rows", "string", "Row fields: comma-separated header names or 1-based indices (e.g. Region,Sales)")
                .WithParameter("columns", "string", "Column fields")
                .WithParameter("values", "string", "Data fields with optional :sum/:count/:average/:max/:min (e.g. Sales:sum,Qty:count)")
                .WithParameter("filters", "string", "Filter/page fields")
                .WithParameter("name", "string", "PivotTable name")
                .WithAlias("pivot_table")
                .WithAlias("pivottable")
                .WithAlias("pivot")
                .WithAlias("create_pivot_table")
                .WithAlias("excel_create_pivot_table")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteCreatePivotTable(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.named_range", "Excel", "Defines a global or worksheet-scoped named range.", 2, true, false)
                .WithParameter("target", "string", "Cell or range address", true)
                .WithParameter("name", "string", "Identifier name for the range", true)
                .WithAlias("named_range")
                .WithAlias("namedrange")
                .WithAlias("name")
                .WithAlias("excel_named_range")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteNamedRange(act.TargetDisplay, act.GetParameterString("name") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.remove_duplicates", "Excel", "Deletes duplicate rows from a designated range.", 3, true, false)
                .WithParameter("target", "string", "Target range to deduplicate", true)
                .WithParameter("columns", "string", "Column indices to check (e.g. 'columns:1,2')")
                .WithAlias("remove_duplicates")
                .WithAlias("removeduplicates")
                .WithAlias("dedupe")
                .WithAlias("excel_remove_duplicates")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteRemoveDuplicates(act.TargetDisplay, act.GetParameterString("columns") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.find_replace", "Excel", "Finds and replaces text within a range.", 2, true, true)
                .WithParameter("target", "string", "Target range to search (e.g. 'A1:Z100')", true)
                .WithParameter("find", "string", "Text to find", true)
                .WithParameter("replace", "string", "Replacement text", true)
                .WithAlias("find_replace")
                .WithAlias("excel_find_replace")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteFindReplace(act.TargetDisplay, act.GetParameterString("find") ?? act.GetParameterString("target_text") ?? string.Empty, act.GetParameterString("replace") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.set_case", "Excel", "Changes the case of text in a range (upper, lower, title, sentence).", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A1:A100')", true)
                .WithParameter("case_type", "string", "Case type: title, sentence, upper, lower", true)
                .WithAlias("set_case")
                .WithAlias("excel_set_case")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteSetCase(act.TargetDisplay, act.GetParameterString("case_type") ?? act.GetParameterString("case") ?? "sentence")));

            Register(new ToolDefinition("excel.trim_range", "Excel", "Trims leading and trailing whitespace from each cell in a range.", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A1:Z100')", true)
                .WithAlias("trim_range")
                .WithAlias("trim")
                .WithAlias("excel_trim_range")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteTrimRange(act.TargetDisplay)));

            Register(new ToolDefinition("excel.normalize_whitespace", "Excel", "Collapses multiple spaces and trims whitespace in a range.", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A1:Z100')", true)
                .WithAlias("normalize_whitespace")
                .WithAlias("normalize")
                .WithAlias("excel_normalize_whitespace")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteNormalizeWhitespace(act.TargetDisplay)));

            Register(new ToolDefinition("excel.text_to_columns", "Excel", "Splits a single column into multiple columns by a delimiter.", 2, true, true)
                .WithParameter("target", "string", "Single-column source range (e.g. 'A2:A100')", true)
                .WithParameter("delimiter", "string", "Delimiter character or word (e.g. ',' ';' '|' 'space' 'tab')", true)
                .WithAlias("text_to_columns")
                .WithAlias("split_column")
                .WithAlias("split")
                .WithAlias("excel_text_to_columns")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteTextToColumns(act.TargetDisplay, act.GetParameterString("delimiter") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.add_worksheet", "Excel", "Creates a new worksheet in the workbook.", 2, true, false)
                .WithParameter("name", "string", "Optional name for the new worksheet")
                .WithAlias("add_worksheet")
                .WithAlias("add_sheet")
                .WithAlias("excel_add_worksheet")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteAddWorksheet(act.ContentDisplay)));

            Register(new ToolDefinition("excel.rename_worksheet", "Excel", "Renames a worksheet.", 2, true, true)
                .WithParameter("target", "string", "Current worksheet name", true)
                .WithParameter("name", "string", "New worksheet name", true)
                .WithAlias("rename_worksheet")
                .WithAlias("rename_sheet")
                .WithAlias("excel_rename_worksheet")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteRenameWorksheet(act.TargetDisplay, act.GetParameterString("name") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.delete_worksheet", "Excel", "Deletes a worksheet. Cannot delete the last sheet.", 3, true, false)
                .WithParameter("target", "string", "Worksheet name to delete", true)
                .WithAlias("delete_worksheet")
                .WithAlias("delete_sheet")
                .WithAlias("excel_delete_worksheet")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteDeleteWorksheet(act.TargetDisplay)));

            Register(new ToolDefinition("excel.duplicate_worksheet", "Excel", "Duplicates a worksheet.", 2, true, false)
                .WithParameter("target", "string", "Source worksheet name", true)
                .WithParameter("name", "string", "Optional name for the duplicate")
                .WithAlias("duplicate_worksheet")
                .WithAlias("duplicate_sheet")
                .WithAlias("excel_duplicate_worksheet")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteDuplicateWorksheet(act.TargetDisplay, act.GetParameterString("name") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.set_tab_color", "Excel", "Sets the tab color of a worksheet.", 1, true, true)
                .WithParameter("target", "string", "Worksheet name (or empty for active sheet)", false)
                .WithParameter("color", "string", "Color hex #RRGGBB or name red/green/blue", true)
                .WithAlias("set_tab_color")
                .WithAlias("tab_color")
                .WithAlias("excel_set_tab_color")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteSetTabColor(act.TargetDisplay, act.GetParameterString("color") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.insert_rows", "Excel", "Inserts rows above a target range.", 2, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A5' or 'A5:A10')", true)
                .WithParameter("count", "integer", "Number of rows to insert", false, 1)
                .WithAlias("insert_rows")
                .WithAlias("excel_insert_rows")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteInsertRows(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.delete_rows", "Excel", "Deletes rows at a target range.", 3, true, false)
                .WithParameter("target", "string", "Target range (e.g. 'A5:A10')", true)
                .WithAlias("delete_rows")
                .WithAlias("excel_delete_rows")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteDeleteRows(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.insert_columns", "Excel", "Inserts columns to the left of a target range.", 2, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'C1' or 'C:C')", true)
                .WithParameter("count", "integer", "Number of columns to insert", false, 1)
                .WithAlias("insert_columns")
                .WithAlias("excel_insert_columns")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteInsertColumns(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.delete_columns", "Excel", "Deletes columns at a target range.", 3, true, false)
                .WithParameter("target", "string", "Target range (e.g. 'C1')", true)
                .WithAlias("delete_columns")
                .WithAlias("excel_delete_columns")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteDeleteColumns(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.hide_rows", "Excel", "Hides rows in a target range.", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A5:A10')", true)
                .WithAlias("hide_rows")
                .WithAlias("excel_hide_rows")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteHideUnhide(act.TargetDisplay, true, true)));

            Register(new ToolDefinition("excel.unhide_rows", "Excel", "Unhides rows in a target range.", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A5:A10')", true)
                .WithAlias("unhide_rows")
                .WithAlias("excel_unhide_rows")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteHideUnhide(act.TargetDisplay, false, true)));

            Register(new ToolDefinition("excel.hide_columns", "Excel", "Hides columns in a target range.", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'C1')", true)
                .WithAlias("hide_columns")
                .WithAlias("excel_hide_columns")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteHideUnhide(act.TargetDisplay, true, false)));

            Register(new ToolDefinition("excel.unhide_columns", "Excel", "Unhides columns in a target range.", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'C1')", true)
                .WithAlias("unhide_columns")
                .WithAlias("excel_unhide_columns")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteHideUnhide(act.TargetDisplay, false, false)));

            Register(new ToolDefinition("excel.merge_cells", "Excel", "Merges or unmerges a target range.", 2, true, true)
                .WithParameter("target", "string", "Target range to merge (e.g. 'A1:C1')", true)
                .WithParameter("action", "string", "'merge' or 'unmerge'", false, "merge")
                .WithAlias("merge_cells")
                .WithAlias("merge")
                .WithAlias("excel_merge_cells")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteMergeCells(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.format_cells", "Excel", "Applies cell formatting (bold, italic, colors, borders, number format, alignment).", 2, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A1:C10')", true)
                .WithParameter("bold", "string", "true/false")
                .WithParameter("italic", "string", "true/false")
                .WithParameter("font_color", "string", "Hex #RRGGBB or name")
                .WithParameter("fill", "string", "Background hex #RRGGBB")
                .WithParameter("border", "string", "Border style thin/thick/none/dashed/double")
                .WithParameter("number_format", "string", "Excel number format string")
                .WithParameter("align", "string", "left/center/right/justify")
                .WithAlias("format_cells")
                .WithAlias("excel_format_cells")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteFormatCells(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.autofit_columns", "Excel", "Autofits columns/rows to content.", 1, true, true)
                .WithParameter("target", "string", "Target range (e.g. 'A1:Z100')", true)
                .WithAlias("autofit_columns")
                .WithAlias("autofit")
                .WithAlias("excel_autofit_columns")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteAutofitColumns(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.freeze_panes", "Excel", "Freezes or unfreezes panes at a target cell.", 1, true, true)
                .WithParameter("target", "string", "Target cell for freeze anchor (e.g. 'B2')", true)
                .WithParameter("action", "string", "'freeze' or 'unfreeze'", false, "freeze")
                .WithAlias("freeze_panes")
                .WithAlias("freeze")
                .WithAlias("excel_freeze_panes")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteFreezePanes(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.add_summary_row", "Excel", "Adds a summary formula row (sum/average/count/max/min) below a table.", 2, true, true)
                .WithParameter("target", "string", "Table range including header (e.g. 'A1:D50')", true)
                .WithParameter("operation", "string", "sum, average, count, max, min", false, "sum")
                .WithParameter("column", "string", "Column letter or 1-based index within target", false, "1")
                .WithAlias("add_summary_row")
                .WithAlias("summary_row")
                .WithAlias("excel_add_summary_row")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteAddSummaryRow(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.write_python", "Excel", "Writes a Python formula =PY() to a single cell. Multi-line code is normalized to semicolons; runs in Microsoft-managed sandbox.", 2, true, true)
                .WithParameter("target", "string", "Single cell target (e.g. 'H2')", true)
                .WithParameter("code", "string", "Python code (multi-line supported; newlines become ';')", true)
                .WithAlias("write_python")
                .WithAlias("python")
                .WithAlias("excel_write_python")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteWritePython(act.TargetDisplay, act.GetParameterString("code") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.apply_theme", "Excel", "Applies a coordinated theme palette (fixed enum: blue, green, grey, railway) to a table range — header, banded rows, borders, autofit.", 2, true, true)
                .WithParameter("target", "string", "Table range including header (e.g. 'A1:D20')", true)
                .WithParameter("palette", "string", "Palette name: blue, green, grey, railway", false, "blue")
                .WithAlias("apply_theme")
                .WithAlias("theme")
                .WithAlias("excel_apply_theme")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteApplyTheme(act.TargetDisplay, act.GetParameterString("palette") ?? act.ContentDisplay)));

            Register(new ToolDefinition("excel.clear_highlights", "Excel", "Clears change-highlight tab colors and grid highlights applied by previous AI edits.", 1, true, true)
                .WithParameter("target", "string", "Optional sheet name to clear (or empty for active sheet)", false)
                .WithAlias("clear_highlights")
                .WithAlias("clear_highlight")
                .WithAlias("excel_clear_highlights")
                .WithHandler((ctrl, act) => ExcelChangeHighlighter.ClearHighlights(((ExcelController)ctrl).GetRawAppObj())));

            Register(new ToolDefinition("excel.add_sparkline", "Excel", "Adds a sparkline (line/column/winloss) at target cell from source data range.", 2, true, true)
                .WithParameter("target", "string", "Target cell for sparkline (e.g. 'F2')", true)
                .WithParameter("source", "string", "Source data range (e.g. 'B2:E2')", true)
                .WithParameter("type", "string", "Sparkline type: line, column, winloss", false, "line")
                .WithAlias("add_sparkline")
                .WithAlias("sparkline")
                .WithAlias("excel_add_sparkline")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteAddSparkline(act.TargetDisplay, act.ContentDisplay)));

            // --- Analysis / Explain / AI Columns / Import / Shapes (Local, No Cloud) ---
            Register(new ToolDefinition("excel.analyze_range", "Excel", "Analyzes a table range for trends, outliers, min/max, distributions and suggests charts/PivotTables. Read-only.", 0, false, false)
                .WithParameter("target", "string", "Table range including header (e.g. 'A1:D50')", true)
                .WithParameter("detail", "string", "Detail level: summary|full", false, "summary")
                .WithAlias("analyze_range")
                .WithAlias("analyze")
                .WithAlias("excel_analyze_range")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteAnalyzeRange(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.get_formula_details", "Excel", "Reads formula, value, precedents and dependents for a cell. Read-only.", 0, false, false)
                .WithParameter("target", "string", "Cell address (e.g. 'B7' or 'Sheet1!B7'; empty for ActiveCell)", false)
                .WithAlias("get_formula_details")
                .WithAlias("explain_formula")
                .WithAlias("excel_get_formula_details")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteGetFormulaDetails(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.add_analysis_column", "Excel", "Creates a new analysis column (sentiment, classify, topic, summarize) next to a source column. Deterministic local placeholder — AI can overwrite via follow-up actions.", 2, true, true)
                .WithParameter("target", "string", "Header cell for new column (e.g. 'G1')", true)
                .WithParameter("source", "string", "Source single-column range to analyze (e.g. 'A2:A100')", true)
                .WithParameter("type", "string", "Analysis type: sentiment|classify|topic|summarize", false, "classify")
                .WithParameter("header", "string", "Header title for new column", false)
                .WithAlias("add_analysis_column")
                .WithAlias("analysis_column")
                .WithAlias("excel_add_analysis_column")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteAddAnalysisColumn(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.import_worksheet", "Excel", "Imports data from another local .xlsx file into a new or existing sheet/range. No cloud.", 2, true, true)
                .WithParameter("target", "string", "Target sheet name (new) or range (e.g. 'Imported' or 'A1')", true)
                .WithParameter("source", "string", "Absolute local path to source .xlsx", true)
                .WithParameter("sheet", "string", "Source sheet name (default first sheet)", false)
                .WithAlias("import_worksheet")
                .WithAlias("import")
                .WithAlias("excel_import_worksheet")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteImportWorksheet(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.create_shape", "Excel", "Creates a shape/text box on a worksheet.", 2, true, true)
                .WithParameter("target", "string", "Worksheet name or sheet-qualified target (e.g. 'Sheet1')", true)
                .WithParameter("type", "string", "Shape type: rectangle, oval, textbox, rounded_rectangle, diamond", false, "rectangle")
                .WithParameter("text", "string", "Text inside shape", false)
                .WithParameter("left", "integer", "Left position in points", false, 100)
                .WithParameter("top", "integer", "Top position in points", false, 50)
                .WithParameter("width", "integer", "Width in points", false, 200)
                .WithParameter("height", "integer", "Height in points", false, 80)
                .WithAlias("create_shape")
                .WithAlias("excel_create_shape")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteCreateShape(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.update_shape", "Excel", "Updates text/fill/line of an existing shape.", 2, true, true)
                .WithParameter("target", "string", "Worksheet name", true)
                .WithParameter("name", "string", "Shape name (e.g. 'Rectangle 1')", true)
                .WithParameter("text", "string", "New text", false)
                .WithParameter("fill", "string", "Fill color hex #RRGGBB", false)
                .WithAlias("update_shape")
                .WithAlias("excel_update_shape")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteUpdateShape(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.set_workbook_rule", "Excel", "Sets a per-workbook local rule/personalization (e.g. preferred_table_style). Stored locally, no cloud.", 1, true, true)
                .WithParameter("target", "string", "Workbook name (or empty for active)", false)
                .WithParameter("key", "string", "Rule key", true)
                .WithParameter("value", "string", "Rule value", true)
                .WithAlias("set_workbook_rule")
                .WithAlias("set_rule")
                .WithAlias("excel_set_workbook_rule")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteSetWorkbookRule(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.get_workbook_rules", "Excel", "Reads local workbook rules and .Rules sheet. Read-only.", 0, false, false)
                .WithParameter("target", "string", "Workbook name (or empty for active)", false)
                .WithAlias("get_workbook_rules")
                .WithAlias("get_rules")
                .WithAlias("excel_get_workbook_rules")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteGetWorkbookRules(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.clear_workbook_rules", "Excel", "Clears per-workbook local rules.", 1, true, false)
                .WithParameter("target", "string", "Workbook name (or empty for active)", false)
                .WithAlias("clear_workbook_rules")
                .WithAlias("clear_rules")
                .WithAlias("excel_clear_workbook_rules")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteClearWorkbookRules(act.TargetDisplay, act.ContentDisplay)));

            // === POWERPOINT TOOLS ===
            Register(new ToolDefinition("powerpoint.move_slide", "PowerPoint", "Moves a slide to a new ordinal position in the presentation.", 2, true, true)
                .WithParameter("source", "integer", "1-based source slide index", true)
                .WithParameter("target", "integer", "1-based destination slide index", true)
                .WithAlias("move_slide")
                .WithAlias("move")
                .WithAlias("powerpoint_move_slide")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int src = act.GetParameterInt("source");
                    int tgt = act.GetParameterInt("target");
                    return p.ExecuteMoveSlide(src, tgt);
                }));

            Register(new ToolDefinition("powerpoint.create_section", "PowerPoint", "Creates a named section header before a designated slide.", 2, true, true)
                .WithParameter("name", "string", "Name of the new section", true)
                .WithParameter("slide", "integer", "1-based slide index before which to create the section", true)
                .WithAlias("create_section")
                .WithAlias("section+")
                .WithAlias("powerpoint_create_section")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.Target != null && act.Target.Slide.HasValue ? act.Target.Slide.Value : act.GetParameterInt("slide");
                    string n = act.GetParameterString("name") ?? "Section";
                    return p.ExecuteCreateSectionBeforeSlide(n, s);
                }));

            Register(new ToolDefinition("powerpoint.rename_section", "PowerPoint", "Renames an existing section header.", 2, true, true)
                .WithParameter("section", "integer", "1-based section index", true)
                .WithParameter("name", "string", "New name for the section", true)
                .WithAlias("rename_section")
                .WithAlias("section")
                .WithAlias("powerpoint_rename_section")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("section");
                    string n = act.GetParameterString("name") ?? "Section";
                    return p.ExecuteRenameSectionInPlace(s, n);
                }));

            Register(new ToolDefinition("powerpoint.set_notes", "PowerPoint", "Updates or appends speaker notes for a designated slide.", 1, true, true)
                .WithParameter("slide", "integer", "1-based slide index", true)
                .WithParameter("notes", "string", "Speaker notes text", true)
                .WithAlias("set_notes")
                .WithAlias("notes")
                .WithAlias("powerpoint_set_notes")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.Target != null && act.Target.Slide.HasValue ? act.Target.Slide.Value : act.GetParameterInt("slide");
                    string n = act.GetParameterString("notes") ?? act.ContentDisplay;
                    return p.ExecuteSetSpeakerNotesInPlace(s, n);
                }));

            Register(new ToolDefinition("powerpoint.create_slide", "PowerPoint", "Inserts a new slide with structured title, bullets, and speaker notes.", 2, true, true)
                .WithParameter("title", "string", "Slide title (optional if bullets or notes supplied)", false)
                .WithParameter("bullets", "object", "List of bullet points")
                .WithParameter("layout", "string", "Slide layout name (TitleOnly, TitleAndContent, SectionHeader, TwoContent, Comparison, TitleSlide, Blank, Custom)")
                .WithParameter("index", "integer", "1-based insertion index (optional, appends if omitted)")
                .WithParameter("notes", "string", "Speaker notes for the new slide")
                .WithAlias("create_slide")
                .WithAlias("slide")
                .WithAlias("powerpoint_create_slide")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    string title = act.GetParameterString("title");
                    string layout = act.GetParameterString("layout");
                    string notes = act.GetParameterString("notes") ?? act.GetParameterString("speaker_notes");
                    int index = act.GetParameterInt("index");
                    if (index <= 0) index = act.GetParameterInt("slide");
                    if (index <= 0 && act.Target != null && act.Target.Slide.HasValue) index = act.Target.Slide.Value;
                    List<string> bullets = null;
                    try
                    {
                        object bulletsObj = null;
                        if (act.Parameters != null && (act.Parameters.TryGetValue("bullets", out bulletsObj) || act.Parameters.TryGetValue("content", out bulletsObj) || act.Parameters.TryGetValue("points", out bulletsObj)) && bulletsObj != null)
                        {
                            bullets = ParseStringArray(bulletsObj);
                            if (bullets == null && bulletsObj is string)
                            {
                                string s = Convert.ToString(bulletsObj);
                                if (!string.IsNullOrWhiteSpace(s))
                                    bullets = new List<string>(s.Split(new[] { '\r', '\n', ';' }, StringSplitOptions.RemoveEmptyEntries));
                            }
                        }
                    }
                    catch { }
                    // Fallback: ContentDisplay is outline markdown if no structured title present
                    if (string.IsNullOrWhiteSpace(title) && (bullets == null || bullets.Count == 0) && !string.IsNullOrWhiteSpace(act.ContentDisplay) && act.Parameters != null && !act.Parameters.ContainsKey("title") && !act.Parameters.ContainsKey("bullets"))
                    {
                        return p.ExecuteCreateDeckFromOutline(act.ContentDisplay);
                    }
                    return p.ExecuteCreateSlide(title, bullets, layout, index, notes);
                }));

            Register(new ToolDefinition("powerpoint.insert_image", "PowerPoint", "Inserts an image file onto a slide.", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index (optional, active slide if omitted)", false)
                .WithParameter("image_path", "string", "Absolute path to the image file", true)
                .WithParameter("alt_text", "string", "Alternative text for accessibility")
                .WithAlias("insert_image")
                .WithAlias("image")
                .WithAlias("powerpoint_insert_image")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    string path = act.GetParameterString("image_path") ?? act.GetParameterString("path") ?? act.GetParameterString("file") ?? act.ContentDisplay;
                    // Guard: if ContentDisplay is composite k=v string, extract image_path value
                    if (!string.IsNullOrWhiteSpace(path) && path.Contains("=") && path.Contains("image_path"))
                    {
                        string extracted = act.GetParameterString("image_path");
                        if (!string.IsNullOrWhiteSpace(extracted)) path = extracted;
                    }
                    string alt = act.GetParameterString("alt_text") ?? act.GetParameterString("alt") ?? act.GetParameterString("description");
                    int slide = act.GetParameterInt("slide");
                    if (slide <= 0 && act.Target != null && act.Target.Slide.HasValue) slide = act.Target.Slide.Value;
                    if (slide <= 0) slide = act.GetParameterInt("index");
                    return p.ExecuteInsertImage(path, alt, slide);
                }));

            Register(new ToolDefinition("powerpoint.delete_slide", "PowerPoint", "Deletes a slide at the specified 1-based index.", 3, true, false)
                .WithParameter("slide", "integer", "1-based slide index to delete", true)
                .WithParameter("index", "integer", "Alias for slide index", false)
                .WithAlias("delete_slide")
                .WithAlias("powerpoint_delete_slide")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0) s = act.GetParameterInt("index");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    return p.ExecuteDeleteSlide(s);
                }));

            Register(new ToolDefinition("powerpoint.duplicate_slide", "PowerPoint", "Duplicates a slide in place (copy appears directly after original).", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index to duplicate", true)
                .WithParameter("index", "integer", "Alias for slide index", false)
                .WithAlias("duplicate_slide")
                .WithAlias("powerpoint_duplicate_slide")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0) s = act.GetParameterInt("index");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    return p.ExecuteDuplicateSlide(s);
                }));

            Register(new ToolDefinition("powerpoint.hide_slide", "PowerPoint", "Hides a slide from the slide show.", 1, true, true)
                .WithParameter("slide", "integer", "1-based slide index to hide", true)
                .WithAlias("hide_slide")
                .WithAlias("powerpoint_hide_slide")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    return p.ExecuteHideSlide(s);
                }));

            Register(new ToolDefinition("powerpoint.unhide_slide", "PowerPoint", "Unhides a previously hidden slide.", 1, true, true)
                .WithParameter("slide", "integer", "1-based slide index to unhide", true)
                .WithAlias("unhide_slide")
                .WithAlias("powerpoint_unhide_slide")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    return p.ExecuteUnhideSlide(s);
                }));

            Register(new ToolDefinition("powerpoint.apply_layout", "PowerPoint", "Applies a slide layout or master layout to an existing slide.", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index", true)
                .WithParameter("layout", "string", "Layout name (TitleOnly, TitleAndContent, SectionHeader, TwoContent, Comparison, TitleSlide, Blank)", true)
                .WithAlias("apply_layout")
                .WithAlias("set_layout")
                .WithAlias("powerpoint_apply_layout")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    string layout = act.GetParameterString("layout") ?? act.ContentDisplay;
                    return p.ExecuteApplyLayout(s, layout);
                }));

            Register(new ToolDefinition("powerpoint.set_shape_text", "PowerPoint", "Edits text of a specific shape/text box on a slide.", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index", true)
                .WithParameter("shape", "string", "Shape name or 1-based shape index (optional)", false)
                .WithParameter("text", "string", "Replacement text for the shape", true)
                .WithAlias("set_shape_text")
                .WithAlias("edit_text")
                .WithAlias("powerpoint_set_shape_text")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    string shape = act.GetParameterString("shape") ?? act.GetParameterString("shape_name") ?? act.GetParameterString("target");
                    string txt = act.GetParameterString("text") ?? act.ContentDisplay;
                    return p.ExecuteSetShapeText(s, shape, txt);
                }));

            Register(new ToolDefinition("powerpoint.replace_text", "PowerPoint", "Replaces the currently selected text or active shape text with new content (true replacement, not append).", 2, true, true)
                .WithParameter("text", "string", "Replacement text", true)
                .WithAlias("replace_text")
                .WithAlias("powerpoint_replace_text")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    string txt = act.GetParameterString("text") ?? act.ContentDisplay;
                    return p.ExecuteReplaceSelectedText(txt);
                }));

            Register(new ToolDefinition("powerpoint.add_table", "PowerPoint", "Adds a native PowerPoint table to a slide.", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index (optional, active slide if omitted)", false)
                .WithParameter("rows", "integer", "Number of rows", true)
                .WithParameter("cols", "integer", "Number of columns", true)
                .WithParameter("data", "object", "2D array of cell strings (optional)")
                .WithAlias("add_table")
                .WithAlias("powerpoint_add_table")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    int rows = act.GetParameterInt("rows");
                    int cols = act.GetParameterInt("cols");
                    if (rows <= 0) rows = 2;
                    if (cols <= 0) cols = 2;
                    List<List<string>> tableData = null;
                    try
                    {
                        object dataObj = null;
                        if (act.Parameters != null && (act.Parameters.TryGetValue("data", out dataObj) || act.Parameters.TryGetValue("values", out dataObj) || act.Parameters.TryGetValue("table", out dataObj)) && dataObj != null)
                        {
                            tableData = Parse2DArray(dataObj);
                            if (tableData != null && tableData.Count > 0)
                            {
                                if (rows < tableData.Count) rows = tableData.Count;
                                foreach (var r in tableData) if (cols < r.Count) cols = r.Count;
                            }
                        }
                        object headersObj = null;
                        if (act.Parameters != null && act.Parameters.TryGetValue("headers", out headersObj) && headersObj != null)
                        {
                            var headers = ParseStringArray(headersObj);
                            if (headers != null && headers.Count > 0)
                            {
                                if (tableData == null) tableData = new List<List<string>>();
                                tableData.Insert(0, headers);
                                if (rows < tableData.Count) rows = tableData.Count;
                                if (cols < headers.Count) cols = headers.Count;
                            }
                        }
                    }
                    catch { }
                    return p.ExecuteAddTable(s, rows, cols, tableData);
                }));

            Register(new ToolDefinition("powerpoint.add_chart", "PowerPoint", "Inserts a native PowerPoint chart onto a slide.", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index (optional)", false)
                .WithParameter("chart_type", "string", "Chart type: column, line, bar, pie, area, scatter", false, "column")
                .WithParameter("title", "string", "Chart title (optional)")
                .WithAlias("add_chart")
                .WithAlias("powerpoint_add_chart")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    string ctype = act.GetParameterString("chart_type") ?? act.GetParameterString("type") ?? "column";
                    string t = act.GetParameterString("title");
                    return p.ExecuteAddChart(s, ctype, t);
                }));

            Register(new ToolDefinition("powerpoint.add_shape", "PowerPoint", "Creates a shape or text box on a slide (SmartArt via shapes).", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index (optional)", false)
                .WithParameter("type", "string", "Shape type: rectangle, rounded_rectangle, oval, diamond, triangle, textbox", false, "rectangle")
                .WithParameter("text", "string", "Text inside shape (optional)")
                .WithAlias("add_shape")
                .WithAlias("powerpoint_add_shape")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    string stype = act.GetParameterString("type") ?? act.GetParameterString("shape_type") ?? "rectangle";
                    string txt = act.GetParameterString("text") ?? act.ContentDisplay;
                    return p.ExecuteAddShape(s, stype, txt);
                }));

            Register(new ToolDefinition("powerpoint.set_font", "PowerPoint", "Sets font properties for the currently selected text or shape.", 1, true, true)
                .WithParameter("font_name", "string", "Font family (e.g., Calibri, Arial)")
                .WithParameter("font_size", "string", "Font size 6-72")
                .WithParameter("bold", "string", "true/false")
                .WithParameter("italic", "string", "true/false")
                .WithParameter("color", "string", "Hex #RRGGBB or named color")
                .WithAlias("set_font")
                .WithAlias("powerpoint_set_font")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    return p.ExecuteSetFont(act.GetParameterString("font_name"), act.GetParameterString("font_size"), act.GetParameterString("bold"), act.GetParameterString("italic"), act.GetParameterString("color"));
                }));

            Register(new ToolDefinition("powerpoint.fit_content", "PowerPoint", "Fits and auto-sizes text content to slide shapes to prevent overflow.", 1, true, true)
                .WithParameter("slide", "integer", "1-based slide index (optional, active slide if omitted)", false)
                .WithAlias("fit_content")
                .WithAlias("powerpoint_fit_content")
                .WithHandler((ctrl, act) => {
                    var p = (PowerPointController)ctrl;
                    int s = act.GetParameterInt("slide");
                    if (s <= 0 && act.Target != null && act.Target.Slide.HasValue) s = act.Target.Slide.Value;
                    return p.ExecuteFitContent(s);
                }));

            // === WORD TOOLS ===
            Register(new ToolDefinition("word.add_comment", "Word", "Adds a review comment anchored to the active selection, search match, or document.", 1, true, true)
                .WithParameter("target_text", "string", "Optional text snippet to anchor the comment to")
                .WithParameter("comment", "string", "Comment body text", true)
                .WithAlias("add_comment")
                .WithAlias("comment")
                .WithAlias("word_add_comment")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteAddComment(act.GetParameterString("comment") ?? act.GetParameterString("comment_text") ?? act.ContentDisplay, act.GetParameterString("target_text"))));

            Register(new ToolDefinition("word.insert_table", "Word", "Inserts a structured grid table with specified rows, columns, and optional headers.", 2, true, true)
                .WithParameter("rows", "integer", "Number of rows", true)
                .WithParameter("cols", "integer", "Number of columns", true)
                .WithParameter("headers", "object", "Optional header column names (array)")
                .WithParameter("data", "object", "Optional 2D array of cell values")
                .WithAlias("insert_table")
                .WithAlias("table")
                .WithAlias("word_insert_table")
                .WithHandler((ctrl, act) => {
                    int rows = act.GetParameterInt("rows");
                    int cols = act.GetParameterInt("cols");
                    if (rows <= 0) rows = 2;
                    if (cols <= 0) cols = 2;
                    // Prefer explicit data/headers if supplied
                    List<List<string>> tableData = null;
                    try
                    {
                        object headersObj = null;
                        if (act.Parameters != null && act.Parameters.TryGetValue("headers", out headersObj) && headersObj != null)
                        {
                            var headers = ParseStringArray(headersObj);
                            if (headers != null && headers.Count > 0)
                            {
                                if (tableData == null) tableData = new List<List<string>>();
                                tableData.Add(headers);
                                // Adjust rows/cols if not explicitly set
                                if (cols < headers.Count) cols = headers.Count;
                            }
                        }
                        object dataObj = null;
                        if (act.Parameters != null && (act.Parameters.TryGetValue("data", out dataObj) || act.Parameters.TryGetValue("rows_data", out dataObj) || act.Parameters.TryGetValue("values", out dataObj)) && dataObj != null)
                        {
                            var parsed = Parse2DArray(dataObj);
                            if (parsed != null && parsed.Count > 0)
                            {
                                if (tableData == null) tableData = new List<List<string>>();
                                tableData.AddRange(parsed);
                                if (rows < tableData.Count) rows = tableData.Count;
                                foreach (var r in tableData) if (cols < r.Count) cols = r.Count;
                            }
                        }
                    }
                    catch { }
                    return ((WordController)ctrl).ExecuteInsertTable(rows, cols, tableData);
                }));

            Register(new ToolDefinition("word.find_replace", "Word", "Finds and replaces text throughout the Word document.", 2, true, true)
                .WithParameter("find", "string", "Text to find", true)
                .WithParameter("replace", "string", "Replacement text", true)
                .WithParameter("match_case", "boolean", "Match case")
                .WithAlias("find_replace")
                .WithAlias("replace")
                .WithAlias("word_find_replace")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteFindReplace(act.GetParameterString("find") ?? act.GetParameterString("target_text") ?? string.Empty, act.GetParameterString("replace") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.apply_style", "Word", "Applies a Word style (e.g., Heading 1, Normal, Title) to a paragraph.", 1, true, true)
                .WithParameter("target", "string", "Paragraph index (1-based) or target text to locate", false)
                .WithParameter("style", "string", "Style name to apply (e.g., Heading 1)", true)
                .WithParameter("paragraph", "integer", "Paragraph index (1-based)")
                .WithAlias("apply_style")
                .WithAlias("style")
                .WithAlias("word_apply_style")
                .WithHandler((ctrl, act) => {
                    int para = act.GetParameterInt("paragraph");
                    if (para <= 0) para = act.GetParameterInt("target");
                    string style = act.GetParameterString("style") ?? act.ContentDisplay;
                    string targetText = act.GetParameterString("target_text") ?? act.GetParameterString("target");
                    if (para > 0) return ((WordController)ctrl).ExecuteApplyStyle(para, style);
                    if (!string.IsNullOrWhiteSpace(targetText)) return ((WordController)ctrl).ExecuteApplyStyleByText(targetText, style);
                    return ((WordController)ctrl).ExecuteApplyStyle(1, style);
                }));

            Register(new ToolDefinition("word.set_case", "Word", "Changes the case of text (Title Case, Sentence case, UPPER, lower).", 1, true, true)
                .WithParameter("target", "string", "Target text or range")
                // Not marked required: the handler below defaults a missing case_type to "sentence",
                // so PreVerify must not reject an action that omits it.
                .WithParameter("case_type", "string", "Case type: title, sentence, upper, lower", false)
                .WithAlias("set_case")
                .WithAlias("word_set_case")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSetCase(act.GetParameterString("target") ?? string.Empty, act.GetParameterString("case_type") ?? act.GetParameterString("case") ?? "sentence")));

            Register(new ToolDefinition("word.reorganize_paragraphs", "Word", "Reorders paragraphs by the supplied order indices.", 2, true, true)
                .WithParameter("order", "string", "Comma-separated paragraph indices in new order (e.g., 3,1,2)", true)
                .WithAlias("reorganize_paragraphs")
                .WithAlias("reorder")
                .WithAlias("word_reorganize_paragraphs")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteReorganizeParagraphs(act.GetParameterString("order") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.normalize_whitespace", "Word", "Normalizes whitespace: trims, collapses multiple spaces, removes duplicate blank paragraphs.", 1, true, true)
                .WithAlias("normalize_whitespace")
                .WithAlias("normalize")
                .WithAlias("word_normalize_whitespace")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteNormalizeWhitespace()));

            Register(new ToolDefinition("word.set_font", "Word", "Sets font family, size, bold, italic, underline, color, highlight for target text or selection.", 1, true, true)
                .WithParameter("target", "string", "Optional target text to locate; empty for selection")
                .WithParameter("font_name", "string", "Font family e.g., Calibri, Arial")
                .WithParameter("font_size", "number", "Font size in points 6-72")
                .WithParameter("bold", "boolean", "true/false")
                .WithParameter("italic", "boolean", "true/false")
                .WithParameter("underline", "boolean", "true/false")
                .WithParameter("color", "string", "Hex #RRGGBB or named color")
                .WithParameter("highlight", "string", "Highlight color or none")
                .WithAlias("set_font")
                .WithAlias("word_set_font")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSetFont(
                    act.GetParameterString("target") ?? act.GetParameterString("target_text"),
                    act.GetParameterString("font_name"), act.GetParameterString("font_size"),
                    act.GetParameterString("bold"), act.GetParameterString("italic"),
                    act.GetParameterString("underline"), act.GetParameterString("color"), act.GetParameterString("highlight"))));

            Register(new ToolDefinition("word.set_paragraph_format", "Word", "Sets paragraph alignment, line spacing, space before/after, indents.", 1, true, true)
                .WithParameter("target", "string", "Optional target text or paragraph index; empty for selection")
                .WithParameter("alignment", "string", "left, center, right, justify")
                .WithParameter("line_spacing", "number", "Line spacing 1.0-3.0")
                .WithParameter("space_before", "number", "Space before in points")
                .WithParameter("space_after", "number", "Space after in points")
                .WithParameter("left_indent", "number", "Left indent in points")
                .WithParameter("first_line_indent", "number", "First line indent in points")
                .WithAlias("set_paragraph_format")
                .WithAlias("word_set_paragraph_format")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSetParagraphFormat(
                    act.GetParameterString("target") ?? act.GetParameterString("target_text"),
                    act.GetParameterString("alignment"), act.GetParameterString("line_spacing"),
                    act.GetParameterString("space_before"), act.GetParameterString("space_after"),
                    act.GetParameterString("left_indent"), act.GetParameterString("first_line_indent"))));

            Register(new ToolDefinition("word.insert_break", "Word", "Inserts a page, column, or section break at the cursor or after target.", 1, true, true)
                .WithParameter("break_type", "string", "page, column, section_next_page, section_continuous", true)
                .WithParameter("target", "string", "Optional target text to locate")
                .WithAlias("insert_break")
                .WithAlias("break")
                .WithAlias("word_insert_break")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertBreak(act.GetParameterString("break_type") ?? act.ContentDisplay, act.GetParameterString("target"))));

            Register(new ToolDefinition("word.set_page_setup", "Word", "Sets page orientation and margins for the active section or whole document.", 1, true, true)
                .WithParameter("orientation", "string", "portrait or landscape")
                .WithParameter("top_margin", "number", "Top margin in inches")
                .WithParameter("bottom_margin", "number", "Bottom margin in inches")
                .WithParameter("left_margin", "number", "Left margin in inches")
                .WithParameter("right_margin", "number", "Right margin in inches")
                .WithAlias("set_page_setup")
                .WithAlias("page_setup")
                .WithAlias("word_set_page_setup")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSetPageSetup(
                    act.GetParameterString("orientation"), act.GetParameterString("top_margin"),
                    act.GetParameterString("bottom_margin"), act.GetParameterString("left_margin"), act.GetParameterString("right_margin"))));

            Register(new ToolDefinition("word.set_header_footer", "Word", "Sets header and/or footer text for the document.", 1, true, true)
                .WithParameter("header", "string", "Header text; empty to keep")
                .WithParameter("footer", "string", "Footer text; empty to keep")
                .WithAlias("set_header_footer")
                .WithAlias("header_footer")
                .WithAlias("word_set_header_footer")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSetHeaderFooter(act.GetParameterString("header"), act.GetParameterString("footer"))));

            Register(new ToolDefinition("word.insert_page_number", "Word", "Inserts page numbers into header or footer.", 1, true, true)
                .WithParameter("alignment", "string", "left, center, right", false, "center")
                .WithParameter("header_footer", "string", "header or footer", false, "footer")
                .WithAlias("insert_page_number")
                .WithAlias("page_number")
                .WithAlias("word_insert_page_number")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertPageNumber(act.GetParameterString("alignment") ?? "center", act.GetParameterString("header_footer") ?? "footer")));

            Register(new ToolDefinition("word.insert_hyperlink", "Word", "Inserts a hyperlink for display text at the cursor or target.", 1, true, true)
                .WithParameter("display_text", "string", "Display text", true)
                .WithParameter("address", "string", "URL or bookmark address", true)
                .WithParameter("target", "string", "Optional target text to replace")
                .WithAlias("insert_hyperlink")
                .WithAlias("hyperlink")
                .WithAlias("word_insert_hyperlink")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertHyperlink(act.GetParameterString("display_text") ?? act.ContentDisplay, act.GetParameterString("address") ?? string.Empty, act.GetParameterString("target"))));

            Register(new ToolDefinition("word.insert_bookmark", "Word", "Inserts a bookmark at the cursor or target text.", 1, true, true)
                .WithParameter("name", "string", "Bookmark name (letters, digits, underscore; start with letter)", true)
                .WithParameter("target", "string", "Optional target text to bookmark")
                .WithAlias("insert_bookmark")
                .WithAlias("bookmark")
                .WithAlias("word_insert_bookmark")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertBookmark(act.GetParameterString("name") ?? string.Empty, act.GetParameterString("target"))));

            Register(new ToolDefinition("word.insert_image", "Word", "Inserts a local image file into the Word document at the cursor.", 1, true, true)
                .WithParameter("image_path", "string", "Absolute local path to image file", true)
                .WithParameter("width", "number", "Width in points (optional)")
                .WithParameter("height", "number", "Height in points (optional)")
                .WithAlias("insert_image")
                .WithAlias("word_insert_image")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertImage(act.GetParameterString("image_path") ?? act.ContentDisplay, act.GetParameterString("width"), act.GetParameterString("height"))));

            Register(new ToolDefinition("word.format_table", "Word", "Applies style, borders, and shading to a Word table by index.", 1, true, true)
                .WithParameter("table_index", "integer", "1-based table index", true)
                .WithParameter("style", "string", "Table style name e.g., Light Grid")
                .WithParameter("borders", "string", "true/false to enable borders")
                .WithParameter("shading", "string", "Hex #RRGGBB shading for header")
                .WithAlias("format_table")
                .WithAlias("word_format_table")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteFormatTable(act.GetParameterInt("table_index"), act.GetParameterString("style"), act.GetParameterString("borders"), act.GetParameterString("shading"))));

            Register(new ToolDefinition("word.list_comments", "Word", "Lists all comments with author, date, and text.", 0, false, false)
                .WithAlias("list_comments")
                .WithAlias("word_list_comments")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteListComments()));

            Register(new ToolDefinition("word.delete_comment", "Word", "Deletes a comment by index or target text.", 1, true, true)
                .WithParameter("comment_index", "integer", "1-based comment index")
                .WithParameter("target_text", "string", "Text snippet of comment to delete")
                .WithAlias("delete_comment")
                .WithAlias("word_delete_comment")
                .WithHandler((ctrl, act) => {
                    int idx = act.GetParameterInt("comment_index");
                    string tgt = act.GetParameterString("target_text") ?? act.ContentDisplay;
                    if (idx > 0) return ((WordController)ctrl).ExecuteDeleteComment(idx);
                    return ((WordController)ctrl).ExecuteDeleteCommentByText(tgt);
                }));

            Register(new ToolDefinition("word.edit_comment", "Word", "Edits a comment's text by index.", 1, true, true)
                .WithParameter("comment_index", "integer", "1-based comment index", true)
                .WithParameter("text", "string", "New comment text", true)
                .WithAlias("edit_comment")
                .WithAlias("word_edit_comment")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteEditComment(act.GetParameterInt("comment_index"), act.GetParameterString("text") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.list_revisions", "Word", "Lists all tracked revisions with author, date, type, and text.", 0, false, false)
                .WithAlias("list_revisions")
                .WithAlias("word_list_revisions")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteListRevisions()));

            Register(new ToolDefinition("word.accept_revision", "Word", "Accepts a specific revision by index.", 1, true, true)
                .WithParameter("revision_index", "integer", "1-based revision index", true)
                .WithAlias("accept_revision")
                .WithAlias("word_accept_revision")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteAcceptRevision(act.GetParameterInt("revision_index"))));

            Register(new ToolDefinition("word.reject_revision", "Word", "Rejects a specific revision by index.", 1, true, true)
                .WithParameter("revision_index", "integer", "1-based revision index", true)
                .WithAlias("reject_revision")
                .WithAlias("word_reject_revision")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteRejectRevision(act.GetParameterInt("revision_index"))));

            Register(new ToolDefinition("word.compare_documents", "Word", "Compares the active document with another local document via Word's native CompareDocuments, producing tracked revisions.", 1, true, true)
                .WithParameter("file_path", "string", "Absolute local path to compare against", true)
                .WithAlias("compare_documents")
                .WithAlias("word_compare_documents")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteCompareDocuments(act.GetParameterString("file_path") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.translate", "Word", "Translates a paragraph or selection to target language with format preservation, side-by-side comment preview, Track Changes, automatic source detection, and paragraph-level review.", 2, true, true)
                .WithParameter("target_language", "string", "Target language code or name (e.g. fr, de, es, zh, auto)", true)
                .WithParameter("source_language", "string", "Source language (optional, auto-detected if omitted)", false)
                .WithParameter("paragraph", "string", "Paragraph index (1-based) or text snippet to locate; empty for selection", false)
                .WithParameter("text", "string", "Translated text content (required, supplied by AI)", true)
                .WithAlias("translate")
                .WithAlias("word_translate")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteTranslate(
                    act.GetParameterString("target_language") ?? act.GetParameterString("target") ?? act.GetParameterString("language") ?? "en",
                    act.GetParameterString("source_language") ?? act.GetParameterString("source") ?? string.Empty,
                    act.GetParameterString("paragraph") ?? act.GetParameterString("target_text") ?? act.GetParameterString("target") ?? string.Empty,
                    act.GetParameterString("text") ?? act.GetParameterString("translated_text") ?? act.GetParameterString("translation") ?? act.ContentDisplay)));

            // --- Word leftovers (verified missing, now implemented locally) ---
            Register(new ToolDefinition("word.insert_toc", "Word", "Inserts a Table of Contents at cursor; update if already present.", 1, true, true)
                .WithParameter("heading_levels", "string", "Optional heading levels e.g. '1-3' (default 1-3)", false)
                .WithAlias("insert_toc")
                .WithAlias("toc")
                .WithAlias("word_insert_toc")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertToc(act.GetParameterString("heading_levels") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.update_toc", "Word", "Updates the first Table of Contents.", 1, true, true)
                .WithAlias("update_toc")
                .WithAlias("word_update_toc")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteUpdateToc()));

            Register(new ToolDefinition("word.export_pdf", "Word", "Exports the active document as PDF to a local path.", 1, true, true)
                .WithParameter("path", "string", "Absolute local .pdf path", true)
                .WithAlias("export_pdf")
                .WithAlias("exportpdf")
                .WithAlias("word_export_pdf")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteExportPdf(act.GetParameterString("path") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.save_as", "Word", "Saves the active document to a local path (docx/pdf).", 1, true, true)
                .WithParameter("path", "string", "Absolute local path", true)
                .WithAlias("save_as")
                .WithAlias("word_save_as")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSaveAs(act.GetParameterString("path") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.toggle_track_changes", "Word", "Enables or disables Track Changes.", 1, true, true)
                .WithParameter("enabled", "string", "true/false or on/off", true)
                .WithAlias("toggle_track_changes")
                .WithAlias("track_changes")
                .WithAlias("word_toggle_track_changes")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteToggleTrackChanges(act.GetParameterString("enabled") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.list_styles", "Word", "Lists available style names (risk 0, read-only).", 0, false, false)
                .WithAlias("list_styles")
                .WithAlias("word_list_styles")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteListStyles()));

            Register(new ToolDefinition("word.set_proofing_language", "Word", "Sets proofing language for selection or paragraph.", 1, true, true)
                .WithParameter("language", "string", "Language id e.g. en-US, fr-FR, or Word language name", true)
                .WithParameter("target", "string", "Optional paragraph index or text snippet", false)
                .WithAlias("set_proofing_language")
                .WithAlias("proofing_language")
                .WithAlias("word_set_proofing_language")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSetProofingLanguage(act.GetParameterString("language") ?? act.ContentDisplay, act.GetParameterString("target"))));

            Register(new ToolDefinition("word.merge_document", "Word", "Inserts another local document at cursor via InsertFile.", 1, true, true)
                .WithParameter("path", "string", "Absolute local .docx/.pdf/.txt path", true)
                .WithAlias("merge_document")
                .WithAlias("merge")
                .WithAlias("word_merge_document")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteMergeDocument(act.GetParameterString("path") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.set_watermark", "Word", "Adds a text watermark (DRAFT/CONFIDENTIAL) via header shapes.", 1, true, true)
                .WithParameter("text", "string", "Watermark text", true)
                .WithParameter("color", "string", "Optional hex color #RRGGBB", false)
                .WithAlias("set_watermark")
                .WithAlias("watermark")
                .WithAlias("word_set_watermark")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteSetWatermark(act.GetParameterString("text") ?? act.ContentDisplay, act.GetParameterString("color"))));

            Register(new ToolDefinition("word.insert_caption", "Word", "Inserts a caption for the nearest table/figure at selection.", 1, true, true)
                .WithParameter("label", "string", "Caption label: Table, Figure, Equation", false)
                .WithParameter("title", "string", "Caption title text", true)
                .WithAlias("insert_caption")
                .WithAlias("caption")
                .WithAlias("word_insert_caption")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertCaption(act.GetParameterString("label") ?? "Figure", act.GetParameterString("title") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.delete", "Word", "Deletes a paragraph, table, or selection by index/text.", 2, true, true)
                .WithParameter("target", "string", "Paragraph index (1-based), 'table:2', or text snippet; empty = selection", false)
                .WithAlias("delete")
                .WithAlias("word_delete")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteDelete(act.GetParameterString("target") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.apply_list", "Word", "Converts paragraphs to bullet or numbered list.", 1, true, true)
                .WithParameter("target", "string", "Optional paragraph index or text snippet; empty = selection", false)
                .WithParameter("list_type", "string", "bullet or number", false)
                .WithAlias("apply_list")
                .WithAlias("word_apply_list")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteApplyList(act.GetParameterString("target"), act.GetParameterString("list_type") ?? act.ContentDisplay)));

            Register(new ToolDefinition("word.readability_stats", "Word", "Returns readability statistics for the document (risk 0).", 0, false, false)
                .WithAlias("readability_stats")
                .WithAlias("readability")
                .WithAlias("word_readability_stats")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteReadabilityStats()));

            // === POWERPOINT ADDITIONAL TOOLS ===
            Register(new ToolDefinition("powerpoint.translate_deck", "PowerPoint", "Translates the entire deck slide-by-slide (proposes per-slide set_shape_text actions).", 2, true, true)
                .WithParameter("target_language", "string", "Target language code", true)
                .WithAlias("translate_deck")
                .WithAlias("powerpoint_translate_deck")
                .WithHandler((ctrl, act) => ((PowerPointController)ctrl).ExecuteTranslateDeck(act.GetParameterString("target_language") ?? act.ContentDisplay)));

            Register(new ToolDefinition("powerpoint.audit_deck", "PowerPoint", "Audits deck for consistency: untitled slides, duplicate titles, hidden slides, missing alt text, font outliers (read-only).", 0, false, false)
                .WithAlias("audit_deck")
                .WithAlias("powerpoint_audit_deck")
                .WithHandler((ctrl, act) => ((PowerPointController)ctrl).ExecuteAuditDeck()));

            Register(new ToolDefinition("powerpoint.audit_alt_text", "PowerPoint", "Lists pictures without alt text (read-only).", 0, false, false)
                .WithAlias("audit_alt_text")
                .WithAlias("powerpoint_audit_alt_text")
                .WithHandler((ctrl, act) => ((PowerPointController)ctrl).ExecuteAuditAltText()));

            Register(new ToolDefinition("powerpoint.set_alt_text", "PowerPoint", "Sets alt text for a picture shape.", 1, true, true)
                .WithParameter("slide", "integer", "Slide index", true)
                .WithParameter("shape", "string", "Shape name or index", true)
                .WithParameter("alt_text", "string", "Alt text", true)
                .WithAlias("set_alt_text")
                .WithAlias("powerpoint_set_alt_text")
                .WithHandler((ctrl, act) => ((PowerPointController)ctrl).ExecuteSetAltText(act.GetParameterInt("slide"), act.GetParameterString("shape"), act.GetParameterString("alt_text") ?? act.ContentDisplay)));
        }

        private static List<string> ParseStringArray(object obj)
        {
            if (obj == null) return null;
            try
            {
                if (obj is Newtonsoft.Json.Linq.JArray)
                {
                    var ja = (Newtonsoft.Json.Linq.JArray)obj;
                    var list = new List<string>();
                    foreach (var t in ja) list.Add(Convert.ToString(t).Trim());
                    return list;
                }
                if (obj is List<object>)
                {
                    var lo = (List<object>)obj;
                    var list = new List<string>();
                    foreach (var t in lo) list.Add(Convert.ToString(t).Trim());
                    return list;
                }
                string s = Convert.ToString(obj);
                if (!string.IsNullOrWhiteSpace(s))
                {
                    // comma-separated fallback
                    var parts = s.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    var list = new List<string>();
                    foreach (var p in parts) list.Add(p.Trim());
                    return list;
                }
            }
            catch { }
            return null;
        }

        private static List<List<string>> Parse2DArray(object obj)
        {
            if (obj == null) return null;
            try
            {
                if (obj is Newtonsoft.Json.Linq.JArray)
                {
                    var ja = (Newtonsoft.Json.Linq.JArray)obj;
                    var result = new List<List<string>>();
                    foreach (var row in ja)
                    {
                        if (row is Newtonsoft.Json.Linq.JArray)
                        {
                            var rja = (Newtonsoft.Json.Linq.JArray)row;
                            var rlist = new List<string>();
                            foreach (var c in rja) rlist.Add(Convert.ToString(c) ?? string.Empty);
                            result.Add(rlist);
                        }
                        else
                        {
                            result.Add(new List<string> { Convert.ToString(row) ?? string.Empty });
                        }
                    }
                    return result;
                }
                if (obj is List<object>)
                {
                    var lo = (List<object>)obj;
                    var result = new List<List<string>>();
                    foreach (var row in lo)
                    {
                        if (row is List<object>)
                        {
                            var rl = (List<object>)row;
                            var rlist = new List<string>();
                            foreach (var c in rl) rlist.Add(Convert.ToString(c) ?? string.Empty);
                            result.Add(rlist);
                        }
                        else
                        {
                            result.Add(new List<string> { Convert.ToString(row) ?? string.Empty });
                        }
                    }
                    return result;
                }
            }
            catch { }
            return null;
        }
    }
}
