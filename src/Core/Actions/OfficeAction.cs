using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core.Actions
{
    public enum OfficeActionStatus
    {
        Pending,
        Approved,
        Applying,
        Applied,
        Failed,
        Rejected,
        RolledBack
    }

    public class ActionTarget
    {
        [JsonProperty("sheet")]
        public string Sheet { get; set; }

        [JsonProperty("range")]
        public string Range { get; set; }

        [JsonProperty("slide")]
        public int? Slide { get; set; }

        [JsonProperty("paragraph")]
        public int? Paragraph { get; set; }

        public override string ToString()
        {
            if (Slide.HasValue) return "Slide " + Slide.Value;
            if (Paragraph.HasValue) return "Paragraph " + Paragraph.Value;
            if (!string.IsNullOrEmpty(Sheet) && !string.IsNullOrEmpty(Range)) return Sheet + "!" + Range;
            if (!string.IsNullOrEmpty(Range)) return Range;
            if (!string.IsNullOrEmpty(Sheet)) return Sheet;
            return "Document";
        }
    }

    public class RollbackInfo
    {
        [JsonProperty("is_rollback_possible")]
        public bool IsRollbackPossible { get; set; }

        [JsonProperty("strategy")]
        public string Strategy { get; set; }

        [JsonProperty("failure_reason")]
        public string FailureReason { get; set; }

        [JsonProperty("captured_at")]
        public DateTime? CapturedAt { get; set; }

        [JsonProperty("data")]
        public Dictionary<string, object> Data { get; set; }

        public RollbackInfo()
        {
            Data = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            IsRollbackPossible = true;
        }

        public RollbackInfo(string strategy) : this()
        {
            Strategy = strategy;
            IsRollbackPossible = true;
        }
    }

    public class EvidenceClaim
    {
        [JsonProperty("location")]
        public string Location { get; set; }

        [JsonProperty("extracted_value")]
        public string ExtractedValue { get; set; }

        [JsonProperty("evidence_level")]
        public string EvidenceLevel { get; set; }
    }

    /// <summary>
    /// Unified structured action model implementing SSOT §5.3 across Word, Excel, and PowerPoint.
    /// Acts as the single authoritative action schema for validation, risk gating, UI presentation, verification, and rollback.
    /// </summary>
    public class OfficeAction : INotifyPropertyChanged
    {
        private OfficeActionStatus _status;
        private string _resultText;
        private string _errorMessage;
        private bool _isUndoable;

        [JsonProperty("action_id")]
        public string ActionId { get; set; }

        [JsonProperty("host")]
        public string Host { get; set; }

        [JsonProperty("operation")]
        public string Operation { get; set; }

        [JsonProperty("target")]
        public ActionTarget Target { get; set; }

        [JsonProperty("input")]
        public Dictionary<string, object> Parameters { get; set; }

        [JsonProperty("before_state")]
        public object BeforeState { get; set; }

        [JsonProperty("expected_result")]
        public string ExpectedResult { get; set; }

        [JsonProperty("risk_level")]
        public int RiskLevel { get; set; }

        [JsonProperty("requires_approval")]
        public bool RequiresApproval { get; set; }

        [JsonProperty("rollback_info")]
        public RollbackInfo Rollback { get; set; }

        [JsonProperty("source_reason")]
        public string SourceReason { get; set; }

        [JsonProperty("evidence")]
        public List<EvidenceClaim> Evidence { get; set; }

        // Execution & Presentation state.
        // Serialized so WorkSession round-trips preserve Applied/Failed/RolledBack for
        // PlanExecutor.RollbackAll (RollbackBatch filters on OfficeActionStatus.Applied).
        [JsonProperty("status")]
        public OfficeActionStatus Status
        {
            get { return _status; }
            set
            {
                if (_status != value)
                {
                    _status = value;
                    OnPropertyChanged("Status");
                    OnPropertyChanged("StatusDisplay");
                    OnPropertyChanged("StatusForegroundBrush");
                }
            }
        }

        [JsonIgnore]
        public string ResultText
        {
            get { return _resultText; }
            set
            {
                if (_resultText != value)
                {
                    _resultText = value;
                    OnPropertyChanged("ResultText");
                    OnPropertyChanged("StatusDisplay");
                }
            }
        }

        [JsonIgnore]
        public string ErrorMessage
        {
            get { return _errorMessage; }
            set
            {
                if (_errorMessage != value)
                {
                    _errorMessage = value;
                    OnPropertyChanged("ErrorMessage");
                    OnPropertyChanged("StatusDisplay");
                    OnPropertyChanged("StatusForegroundBrush");
                }
            }
        }

        [JsonIgnore]
        public bool IsUndoable
        {
            get { return _isUndoable; }
            set
            {
                if (_isUndoable != value)
                {
                    _isUndoable = value;
                    OnPropertyChanged("IsUndoable");
                }
            }
        }

        [JsonIgnore]
        public string ActionBadge
        {
            get
            {
                if (string.IsNullOrEmpty(Operation)) return "act";
                int dot = Operation.IndexOf('.');
                string sub = dot >= 0 ? Operation.Substring(dot + 1) : Operation;
                switch (sub.ToLowerInvariant())
                {
                    case "write_formula":
                    case "formula": return "fx";
                    case "write_value":
                    case "value": return "val";
                    case "fill_down":
                    case "filldown": return "fill";
                    case "create_table":
                    case "createtable":
                    case "table": return "tbl";
                    case "conditional_format":
                    case "conditionalformat": return "fmt";
                    case "sort": return "sort";
                    case "filter": return "fltr";
                    case "data_validation":
                    case "datavalidation": return "valid";
                    case "create_chart":
                    case "chart": return "chrt";
                    case "create_pivot_table":
                    case "pivot_table":
                    case "pivottable":
                    case "pivot": return "piv";
                    case "named_range":
                    case "namedrange": return "name";
                    case "remove_duplicates":
                    case "removeduplicates":
                    case "dedupe": return "dedupe";
                    case "add_comment": return "comm";
                    case "create_slide":
                    case "move_slide": return "sld";
                    case "create_section":
                    case "rename_section": return "sec";
                    case "set_notes": return "note";
                    case "delete_slide": return "del";
                    case "duplicate_slide": return "dup";
                    case "hide_slide": return "hide";
                    case "unhide_slide": return "show";
                    case "apply_layout": return "layout";
                    case "set_shape_text": return "text";
                    case "replace_text": return "replace";
                    case "add_table": return "tbl";
                    case "add_chart": return "chart";
                    case "add_shape": return "shape";
                    case "set_font": return "font";
                    case "fit_content": return "fit";
                    case "insert_image": return "img";
                    default: return "act";
                }
            }
        }

        [JsonIgnore]
        public string TypeBadge
        {
            get { return ActionBadge; }
        }

        [JsonIgnore]
        public string PreviewDescription
        {
            get
            {
                if (!string.IsNullOrEmpty(SourceReason)) return SourceReason;
                if (!string.IsNullOrEmpty(ExpectedResult)) return ExpectedResult;
                return string.Format("{0} on {1}", Operation, Target != null ? Target.ToString() : "target");
            }
        }

        [JsonIgnore]
        public string Description
        {
            get { return PreviewDescription; }
        }

        [JsonIgnore]
        public string TargetDisplay
        {
            get { return Target != null ? Target.ToString() : "Document"; }
        }

        [JsonIgnore]
        public string ContentDisplay
        {
            get
            {
                if (Parameters == null || Parameters.Count == 0) return string.Empty;
                if (Parameters.ContainsKey("formula") && Parameters["formula"] != null) return Convert.ToString(Parameters["formula"]);
                if (Parameters.ContainsKey("value") && Parameters["value"] != null) return Convert.ToString(Parameters["value"]);
                if (Parameters.ContainsKey("content") && Parameters["content"] != null) return Convert.ToString(Parameters["content"]);
                if (Parameters.ContainsKey("comment_text") && Parameters["comment_text"] != null) return Convert.ToString(Parameters["comment_text"]);
                if (Parameters.ContainsKey("notes") && Parameters["notes"] != null) return Convert.ToString(Parameters["notes"]);
                if (Parameters.ContainsKey("outline") && Parameters["outline"] != null) return Convert.ToString(Parameters["outline"]);

                var sb = new StringBuilder();
                foreach (var kvp in Parameters)
                {
                    if (kvp.Value == null) continue;
                    if (sb.Length > 0) sb.Append(", ");
                    sb.Append(kvp.Key).Append("=").Append(kvp.Value);
                }
                return sb.ToString();
            }
        }

        [JsonIgnore]
        public bool HasContentDisplay
        {
            get { return !string.IsNullOrWhiteSpace(ContentDisplay); }
        }

        [JsonIgnore]
        public string StatusDisplay
        {
            get
            {
                switch (_status)
                {
                    case OfficeActionStatus.Pending:
                        return string.Empty;
                    case OfficeActionStatus.Approved:
                        return "Approved";
                    case OfficeActionStatus.Applying:
                        return "Applying...";
                    case OfficeActionStatus.Applied:
                        return !string.IsNullOrEmpty(_resultText) ? ("✔ " + _resultText) : "✔ Applied";
                    case OfficeActionStatus.Failed:
                        return !string.IsNullOrEmpty(_errorMessage) ? ("⚠ Error: " + _errorMessage) : "⚠ Error";
                    case OfficeActionStatus.Rejected:
                        return "Rejected";
                    case OfficeActionStatus.RolledBack:
                        return !string.IsNullOrEmpty(_resultText) ? ("↺ " + _resultText) : "↺ Rolled Back";
                    default:
                        return _status.ToString();
                }
            }
        }

        [JsonIgnore]
        public string StatusForegroundBrush
        {
            get
            {
                switch (_status)
                {
                    case OfficeActionStatus.Applied:
                        return "#059669"; // Emerald
                    case OfficeActionStatus.Failed:
                        return "#DC2626"; // Red
                    case OfficeActionStatus.Applying:
                        return "#2563EB"; // Blue
                    case OfficeActionStatus.RolledBack:
                        return "#D97706"; // Amber / Orange
                    default:
                        return "#475569"; // Slate
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public OfficeAction()
        {
            ActionId = Guid.NewGuid().ToString("N");
            Target = new ActionTarget();
            Parameters = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            Evidence = new List<EvidenceClaim>();
            _status = OfficeActionStatus.Pending;
            RequiresApproval = true;
            _isUndoable = true;
        }

        #region Legacy Compatibility Adapters

        /// <summary>
        /// Projects this OfficeAction into a legacy SpreadsheetAction view-model for backward compatibility.
        /// Returns null if the operation is not a recognized spreadsheet action.
        /// </summary>
        public SpreadsheetAction ToSpreadsheetAction()
        {
            SpreadsheetActionType? type = null;
            string op = (Operation ?? string.Empty).ToLowerInvariant();
            if (op.Contains("write_formula") || (op.Contains("formula") && !op.Contains("conditional"))) type = SpreadsheetActionType.Formula;
            else if (op.Contains("fill_down") || op.Contains("filldown")) type = SpreadsheetActionType.FillDown;
            else if (op.Contains("create_table") || op.Contains("createtable")) type = SpreadsheetActionType.CreateTable;
            else if (op.Contains("table") && !op.Contains("pivot")) type = SpreadsheetActionType.Table;
            else if (op.Contains("conditional_format") || op.Contains("conditionalformat")) type = SpreadsheetActionType.ConditionalFormat;
            else if (op.Contains("sort") && !op.Contains("worksheet")) type = SpreadsheetActionType.Sort;
            else if (op.Contains("filter")) type = SpreadsheetActionType.Filter;
            else if (op.Contains("data_validation") || op.Contains("datavalidation")) type = SpreadsheetActionType.DataValidation;
            else if (op.Contains("create_chart") || (op.Contains("chart") && !op.Contains("analysis"))) type = SpreadsheetActionType.Chart;
            else if (op.Contains("pivot")) type = SpreadsheetActionType.PivotTable;
            else if (op.Contains("named_range") || op.Contains("namedrange")) type = SpreadsheetActionType.NamedRange;
            else if (op.Contains("remove_duplicates") || op.Contains("removeduplicates")) type = SpreadsheetActionType.RemoveDuplicates;
            else if (op.Contains("find_replace")) type = SpreadsheetActionType.FindReplace;
            else if (op.Contains("set_case") || op.Contains("setcase")) type = SpreadsheetActionType.SetCase;
            else if (op.Contains("trim_range") || op == "excel.trim") type = SpreadsheetActionType.TrimRange;
            else if (op.Contains("normalize_whitespace") || op.Contains("normalize")) type = SpreadsheetActionType.NormalizeWhitespace;
            else if (op.Contains("text_to_columns") || op.Contains("split_column")) type = SpreadsheetActionType.TextToColumns;
            else if (op.Contains("write_python") || op.Contains("python")) type = SpreadsheetActionType.WritePython;
            else if (op.Contains("apply_theme") || op.Contains("theme")) type = SpreadsheetActionType.ApplyTheme;
            else if (op.Contains("analyze_range") || op.Contains("analyze")) type = SpreadsheetActionType.AnalyzeRange;
            else if (op.Contains("get_formula_details") || op.Contains("explain_formula")) type = SpreadsheetActionType.GetFormulaDetails;
            else if (op.Contains("add_analysis_column")) type = SpreadsheetActionType.AddAnalysisColumn;
            else if (op.Contains("import_worksheet") || op.Contains("import")) type = SpreadsheetActionType.ImportWorksheet;
            else if (op.Contains("create_shape") && !op.Contains("update")) type = SpreadsheetActionType.CreateShape;
            else if (op.Contains("update_shape")) type = SpreadsheetActionType.UpdateShape;
            else if (op.Contains("set_workbook_rule") || op.Contains("set_rule")) type = SpreadsheetActionType.SetWorkbookRule;
            else if (op.Contains("get_workbook_rules") || op.Contains("get_rules")) type = SpreadsheetActionType.GetWorkbookRules;
            else if (op.Contains("clear_workbook_rules") || op.Contains("clear_rules")) type = SpreadsheetActionType.ClearWorkbookRules;
            else if (op.Contains("add_worksheet") || op.Contains("add_sheet")) type = SpreadsheetActionType.AddAnalysisColumn; // fallback mapping will be overridden below
            else if (op.Contains("value") || op.Contains("write_value")) type = SpreadsheetActionType.Value;

            // Handle worksheet lifecycle that doesn't map to legacy enum - map to ImportWorksheet or closest; but better map to dedicated type if exists
            // For add_worksheet etc., we fallback to a generic but ensure not null for newer types: use TryParse via ToolRegistry
            if (!type.HasValue)
            {
                // Attempt to map any remaining excel.* to a known enum via explicit list
                if (op == "excel.add_worksheet" || op == "excel.rename_worksheet" || op == "excel.delete_worksheet" || op == "excel.duplicate_worksheet"
                    || op == "excel.set_tab_color" || op == "excel.insert_rows" || op == "excel.delete_rows" || op == "excel.insert_columns"
                    || op == "excel.delete_columns" || op == "excel.hide_rows" || op == "excel.unhide_rows" || op == "excel.hide_columns"
                    || op == "excel.unhide_columns" || op == "excel.merge_cells" || op == "excel.format_cells" || op == "excel.autofit_columns"
                    || op == "excel.freeze_panes" || op == "excel.add_summary_row" || op == "excel.clear_highlights")
                {
                    // These operations don't have legacy SpreadsheetActionType equivalents; map to a placeholder that round-trips via FromSpreadsheetAction
                    // Use Value as generic carrier but preserve original operation in Description
                    type = SpreadsheetActionType.Value;
                }
            }

            if (!type.HasValue) return null;

            string content = GetParameterString("value") ?? GetParameterString("content") ?? GetParameterString("formula") ?? GetParameterString("rule") ?? GetParameterString("code") ?? GetParameterString("palette") ?? GetParameterString("name") ?? string.Empty;
            // For richer actions, serialize all params if content is still empty
            if (string.IsNullOrWhiteSpace(content) && Parameters != null && Parameters.Count > 0)
            {
                try { content = Newtonsoft.Json.JsonConvert.SerializeObject(Parameters); } catch { content = string.Empty; }
            }
            string target = Target != null ? (Target.Range ?? Target.ToString()) : "A1";
            if (string.IsNullOrWhiteSpace(target) && Target != null && !string.IsNullOrWhiteSpace(Target.Sheet)) target = Target.Sheet;
            if (string.IsNullOrWhiteSpace(target)) target = "A1";

            return new SpreadsheetAction
            {
                Type = type.Value,
                Target = target,
                Content = content,
                Description = PreviewDescription + (type == SpreadsheetActionType.Value && op != "excel.write_value" ? " [compat:" + Operation + "]" : string.Empty),
                Status = ToSpreadsheetActionStatus(Status),
                ErrorMessage = ErrorMessage,
                ResultText = ResultText
            };
        }

        /// <summary>
        /// Creates an OfficeAction from a legacy SpreadsheetAction.
        /// </summary>
        public static OfficeAction FromSpreadsheetAction(SpreadsheetAction sa)
        {
            string opName;
            switch (sa.Type)
            {
                case SpreadsheetActionType.Formula: opName = "excel.write_formula"; break;
                case SpreadsheetActionType.Value: opName = "excel.write_value"; break;
                case SpreadsheetActionType.FillDown: opName = "excel.fill_down"; break;
                case SpreadsheetActionType.Table: opName = "excel.table"; break;
                case SpreadsheetActionType.CreateTable: opName = "excel.create_table"; break;
                case SpreadsheetActionType.ConditionalFormat: opName = "excel.conditional_format"; break;
                case SpreadsheetActionType.Sort: opName = "excel.sort"; break;
                case SpreadsheetActionType.Filter: opName = "excel.filter"; break;
                case SpreadsheetActionType.DataValidation: opName = "excel.data_validation"; break;
                case SpreadsheetActionType.Chart: opName = "excel.create_chart"; break;
                case SpreadsheetActionType.PivotTable: opName = "excel.create_pivot_table"; break;
                case SpreadsheetActionType.NamedRange: opName = "excel.named_range"; break;
                case SpreadsheetActionType.RemoveDuplicates: opName = "excel.remove_duplicates"; break;
                case SpreadsheetActionType.FindReplace: opName = "excel.find_replace"; break;
                case SpreadsheetActionType.SetCase: opName = "excel.set_case"; break;
                case SpreadsheetActionType.TrimRange: opName = "excel.trim_range"; break;
                case SpreadsheetActionType.NormalizeWhitespace: opName = "excel.normalize_whitespace"; break;
                case SpreadsheetActionType.TextToColumns: opName = "excel.text_to_columns"; break;
                case SpreadsheetActionType.WritePython: opName = "excel.write_python"; break;
                case SpreadsheetActionType.ApplyTheme: opName = "excel.apply_theme"; break;
                case SpreadsheetActionType.AnalyzeRange: opName = "excel.analyze_range"; break;
                case SpreadsheetActionType.GetFormulaDetails: opName = "excel.get_formula_details"; break;
                case SpreadsheetActionType.AddAnalysisColumn: opName = "excel.add_analysis_column"; break;
                case SpreadsheetActionType.ImportWorksheet: opName = "excel.import_worksheet"; break;
                case SpreadsheetActionType.CreateShape: opName = "excel.create_shape"; break;
                case SpreadsheetActionType.UpdateShape: opName = "excel.update_shape"; break;
                case SpreadsheetActionType.SetWorkbookRule: opName = "excel.set_workbook_rule"; break;
                case SpreadsheetActionType.GetWorkbookRules: opName = "excel.get_workbook_rules"; break;
                case SpreadsheetActionType.ClearWorkbookRules: opName = "excel.clear_workbook_rules"; break;
                default: opName = "excel." + sa.Type.ToString().ToLowerInvariant(); break;
            }

            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = opName,
                ExpectedResult = sa.Description,
                SourceReason = sa.Description,
                IsUndoable = sa.IsUndoable,
                Status = FromSpreadsheetActionStatus(sa.Status),
                ErrorMessage = sa.ErrorMessage,
                ResultText = sa.ResultText
            };

            if (!string.IsNullOrEmpty(sa.Target))
            {
                if (sa.Target.Contains("!"))
                {
                    var parts = sa.Target.Split('!');
                    action.Target.Sheet = parts[0];
                    action.Target.Range = parts.Length > 1 ? parts[1] : string.Empty;
                }
                else
                {
                    action.Target.Range = sa.Target;
                }
            }

            action.Parameters["content"] = sa.Content;
            if (sa.Type == SpreadsheetActionType.Formula)
            {
                action.Parameters["formula"] = sa.Content;
            }
            else if (sa.Type == SpreadsheetActionType.Value)
            {
                action.Parameters["value"] = sa.Content;
            }

            var tool = ToolRegistry.GetTool(opName, "Excel");
            if (tool != null)
            {
                action.RiskLevel = tool.RiskLevel;
                action.IsUndoable = tool.IsUndoable;
                action.RequiresApproval = tool.RequiresApproval;
            }
            else
            {
                action.RiskLevel = sa.Type == SpreadsheetActionType.RemoveDuplicates ? 3 : (sa.Type == SpreadsheetActionType.Value ? 1 : 2);
                action.IsUndoable = sa.IsUndoable;
            }

            return action;
        }

        /// <summary>
        /// Projects this OfficeAction into a legacy PowerPointAction view-model for backward compatibility.
        /// Returns null if the operation is not a recognized deck action.
        /// </summary>
        public PowerPointAction ToPowerPointAction()
        {
            string op = (Operation ?? string.Empty).ToLowerInvariant();
            string type = null;
            if (op.Contains("create_section")) type = "create_section";
            else if (op.Contains("rename_section")) type = "rename_section";
            else if (op.Contains("set_notes")) type = "set_notes";
            else if (op.Contains("move_slide")) type = "move_slide";
            else if (op.Contains("create_slide")) type = "create_slide";
            else if (op.Contains("insert_image")) type = "insert_image";
            else if (op.Contains("delete_slide")) type = "delete_slide";
            else if (op.Contains("duplicate_slide")) type = "duplicate_slide";
            else if (op.Contains("hide_slide") && !op.Contains("unhide")) type = "hide_slide";
            else if (op.Contains("unhide_slide")) type = "unhide_slide";
            else if (op.Contains("apply_layout") || op.Contains("set_layout")) type = "apply_layout";
            else if (op.Contains("set_shape_text")) type = "set_shape_text";
            else if (op.Contains("replace_text")) type = "replace_text";
            else if (op.Contains("add_table")) type = "add_table";
            else if (op.Contains("add_chart")) type = "add_chart";
            else if (op.Contains("add_shape")) type = "add_shape";
            else if (op.Contains("set_font")) type = "set_font";
            else if (op.Contains("fit_content")) type = "fit_content";

            if (type == null) return null;

            var ppa = new PowerPointAction
            {
                Type = type,
                Source = GetParameterInt("source"),
                Target = GetParameterInt("target"),
                Slide = Target != null && Target.Slide.HasValue ? Target.Slide.Value : GetParameterInt("slide"),
                Section = GetParameterInt("section"),
                Name = GetParameterString("name"),
                Notes = GetParameterString("notes"),
                Layout = GetParameterString("layout"),
                ShapeType = GetParameterString("shape_type") ?? GetParameterString("shape") ?? GetParameterString("type"),
                ImagePath = GetParameterString("image_path") ?? GetParameterString("path"),
                Text = GetParameterString("text") ?? GetParameterString("content"),
                Title = GetParameterString("title"),
                ChartType = GetParameterString("chart_type") ?? GetParameterString("chartType"),
                AltText = GetParameterString("alt_text"),
                Data = GetParameterString("data"),
                Rows = GetParameterInt("rows"),
                Cols = GetParameterInt("cols"),
                FontName = GetParameterString("font_name"),
                FontSize = GetParameterString("font_size"),
                Bold = GetParameterString("bold"),
                Italic = GetParameterString("italic"),
                Color = GetParameterString("color"),
                Status = ToPowerPointActionStatus(Status),
                ErrorMessage = ErrorMessage,
                ResultText = ResultText
            };
            if (ppa.Slide == 0) ppa.Slide = GetParameterInt("index");
            if (string.IsNullOrWhiteSpace(ppa.Name) && !string.IsNullOrWhiteSpace(ppa.Title)) ppa.Name = ppa.Title;
            if (string.IsNullOrWhiteSpace(ppa.Text) && !string.IsNullOrWhiteSpace(GetParameterString("value"))) ppa.Text = GetParameterString("value");
            return ppa;
        }

        /// <summary>
        /// Creates an OfficeAction from a legacy PowerPointAction.
        /// </summary>
        public static OfficeAction FromPowerPointAction(PowerPointAction pa)
        {
            if (pa == null) return null;
            string opName = "powerpoint." + (pa.Type ?? "action").ToLowerInvariant();
            var action = new OfficeAction
            {
                Host = "PowerPoint",
                Operation = opName,
                ExpectedResult = pa.Description,
                SourceReason = pa.Description,
                Status = FromPowerPointActionStatus(pa.Status),
                ErrorMessage = pa.ErrorMessage,
                ResultText = pa.ResultText
            };

            var tool = ToolRegistry.GetTool(opName, "PowerPoint");
            if (tool != null)
            {
                action.RiskLevel = tool.RiskLevel;
                action.IsUndoable = tool.IsUndoable;
                action.RequiresApproval = tool.RequiresApproval;
            }
            else
            {
                action.RiskLevel = string.Equals(pa.Type, "set_notes", StringComparison.OrdinalIgnoreCase) ? 1 : 2;
                action.IsUndoable = true;
            }

            if (pa.Slide > 0) action.Target.Slide = pa.Slide;
            if (pa.Source > 0) action.Parameters["source"] = pa.Source;
            if (pa.Target > 0) action.Parameters["target"] = pa.Target;
            if (pa.Section > 0) action.Parameters["section"] = pa.Section;
            if (!string.IsNullOrEmpty(pa.Name)) action.Parameters["name"] = pa.Name;
            if (!string.IsNullOrEmpty(pa.Notes)) action.Parameters["notes"] = pa.Notes;
            if (!string.IsNullOrWhiteSpace(pa.Layout)) action.Parameters["layout"] = pa.Layout;
            if (!string.IsNullOrWhiteSpace(pa.ShapeType)) action.Parameters["shape_type"] = pa.ShapeType;
            if (!string.IsNullOrWhiteSpace(pa.ImagePath)) action.Parameters["image_path"] = pa.ImagePath;
            if (!string.IsNullOrWhiteSpace(pa.Text)) action.Parameters["text"] = pa.Text;
            if (!string.IsNullOrWhiteSpace(pa.Title)) action.Parameters["title"] = pa.Title;
            if (!string.IsNullOrWhiteSpace(pa.ChartType)) action.Parameters["chart_type"] = pa.ChartType;
            if (!string.IsNullOrWhiteSpace(pa.AltText)) action.Parameters["alt_text"] = pa.AltText;
            if (!string.IsNullOrWhiteSpace(pa.Data)) action.Parameters["data"] = pa.Data;
            if (pa.Rows > 0) action.Parameters["rows"] = pa.Rows;
            if (pa.Cols > 0) action.Parameters["cols"] = pa.Cols;
            if (!string.IsNullOrWhiteSpace(pa.FontName)) action.Parameters["font_name"] = pa.FontName;
            if (!string.IsNullOrWhiteSpace(pa.FontSize)) action.Parameters["font_size"] = pa.FontSize;
            if (!string.IsNullOrWhiteSpace(pa.Bold)) action.Parameters["bold"] = pa.Bold;
            if (!string.IsNullOrWhiteSpace(pa.Italic)) action.Parameters["italic"] = pa.Italic;
            if (!string.IsNullOrWhiteSpace(pa.Color)) action.Parameters["color"] = pa.Color;
            if (pa.ExtraAttributes != null)
                foreach (var kv in pa.ExtraAttributes)
                    if (!action.Parameters.ContainsKey(kv.Key)) action.Parameters[kv.Key] = kv.Value;

            return action;
        }

        public static SpreadsheetActionStatus ToSpreadsheetActionStatus(OfficeActionStatus status)
        {
            switch (status)
            {
                case OfficeActionStatus.Applying:
                    return SpreadsheetActionStatus.Applying;
                case OfficeActionStatus.Applied:
                    return SpreadsheetActionStatus.Applied;
                case OfficeActionStatus.Failed:
                case OfficeActionStatus.Rejected:
                    return SpreadsheetActionStatus.Error;
                case OfficeActionStatus.Pending:
                case OfficeActionStatus.Approved:
                default:
                    return SpreadsheetActionStatus.Pending;
            }
        }

        public static OfficeActionStatus FromSpreadsheetActionStatus(SpreadsheetActionStatus status)
        {
            switch (status)
            {
                case SpreadsheetActionStatus.Applying:
                    return OfficeActionStatus.Applying;
                case SpreadsheetActionStatus.Applied:
                    return OfficeActionStatus.Applied;
                case SpreadsheetActionStatus.Error:
                    return OfficeActionStatus.Failed;
                case SpreadsheetActionStatus.Pending:
                default:
                    return OfficeActionStatus.Pending;
            }
        }

        public static PowerPointActionStatus ToPowerPointActionStatus(OfficeActionStatus status)
        {
            switch (status)
            {
                case OfficeActionStatus.Applying:
                    return PowerPointActionStatus.Applying;
                case OfficeActionStatus.Applied:
                    return PowerPointActionStatus.Applied;
                case OfficeActionStatus.Failed:
                case OfficeActionStatus.Rejected:
                    return PowerPointActionStatus.Error;
                case OfficeActionStatus.Pending:
                case OfficeActionStatus.Approved:
                default:
                    return PowerPointActionStatus.Pending;
            }
        }

        public static OfficeActionStatus FromPowerPointActionStatus(PowerPointActionStatus status)
        {
            switch (status)
            {
                case PowerPointActionStatus.Applying:
                    return OfficeActionStatus.Applying;
                case PowerPointActionStatus.Applied:
                    return OfficeActionStatus.Applied;
                case PowerPointActionStatus.Error:
                    return OfficeActionStatus.Failed;
                case PowerPointActionStatus.Pending:
                default:
                    return OfficeActionStatus.Pending;
            }
        }

        #endregion

        #region Helper Methods

        public string GetParameterString(string key)
        {
            if (Parameters != null && Parameters.ContainsKey(key) && Parameters[key] != null)
            {
                return Convert.ToString(Parameters[key]);
            }
            return null;
        }

        public int GetParameterInt(string key, int defaultValue = 0)
        {
            if (Parameters != null && Parameters.ContainsKey(key) && Parameters[key] != null)
            {
                try
                {
                    return Convert.ToInt32(Parameters[key]);
                }
                catch { }
            }
            return defaultValue;
        }

        #endregion
    }
}
