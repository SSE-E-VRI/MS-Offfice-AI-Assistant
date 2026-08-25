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
                return "formula, value, filldown, table, create_table, conditional_format, sort, filter, data_validation, chart, pivot_table, named_range, and remove_duplicates";
            }
            if (string.Equals(host, "PowerPoint", StringComparison.OrdinalIgnoreCase))
            {
                return "move_slide, create_section, rename_section, or set_notes";
            }
            if (string.Equals(host, "Word", StringComparison.OrdinalIgnoreCase))
            {
                return "add_comment, insert_table";
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

            Register(new ToolDefinition("excel.conditional_format", "Excel", "Applies conditional formatting rules to a range.", 2, true, true)
                .WithParameter("target", "string", "Target cell or range address", true)
                .WithParameter("rule", "string", "Formatting rule (e.g. 'highlight_gt:50000')", true)
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
                .WithParameter("chart_type", "string", "Chart type (e.g. 'column', 'line', 'bar', 'pie')", false, "column")
                .WithParameter("title", "string", "Chart title")
                .WithAlias("chart")
                .WithAlias("create_chart")
                .WithAlias("excel_create_chart")
                .WithHandler((ctrl, act) => ((ExcelController)ctrl).ExecuteCreateChart(act.TargetDisplay, act.ContentDisplay)));

            Register(new ToolDefinition("excel.create_pivot_table", "Excel", "Creates a Pivot Table summary from a source table.", 2, true, false)
                .WithParameter("target", "string", "Source data range", true)
                .WithParameter("destination", "string", "Target cell for pivot top-left")
                .WithParameter("rows", "string", "Row fields specification")
                .WithParameter("values", "string", "Value fields specification")
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
                .WithParameter("title", "string", "Slide title", true)
                .WithParameter("bullets", "object", "List of bullet points")
                .WithParameter("layout", "string", "Slide layout name")
                .WithAlias("create_slide")
                .WithAlias("slide")
                .WithAlias("powerpoint_create_slide")
                .WithHandler((ctrl, act) => ((PowerPointController)ctrl).ExecuteCreateDeckFromOutline(act.ContentDisplay)));

            Register(new ToolDefinition("powerpoint.insert_image", "PowerPoint", "Inserts an image file onto a slide.", 2, true, true)
                .WithParameter("slide", "integer", "1-based slide index", true)
                .WithParameter("image_path", "string", "Absolute path to the image file", true)
                .WithAlias("insert_image")
                .WithAlias("image")
                .WithAlias("powerpoint_insert_image")
                .WithHandler((ctrl, act) => ((PowerPointController)ctrl).ExecuteInsertImage(act.GetParameterString("image_path") ?? act.ContentDisplay)));

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
                .WithParameter("headers", "object", "Optional header column names")
                .WithAlias("insert_table")
                .WithAlias("table")
                .WithAlias("word_insert_table")
                .WithHandler((ctrl, act) => ((WordController)ctrl).ExecuteInsertTable(act.GetParameterInt("rows") > 0 ? act.GetParameterInt("rows") : 2, act.GetParameterInt("cols") > 0 ? act.GetParameterInt("cols") : 2)));
        }
    }
}
