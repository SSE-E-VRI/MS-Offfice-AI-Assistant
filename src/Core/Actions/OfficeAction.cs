using System;
using System.Collections.Generic;
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
        Rejected
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
        [JsonProperty("strategy")]
        public string Strategy { get; set; }

        [JsonProperty("data")]
        public Dictionary<string, object> Data { get; set; }

        public RollbackInfo()
        {
            Data = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        }

        public RollbackInfo(string strategy) : this()
        {
            Strategy = strategy;
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
    /// Acts as the single authoritative action schema for validation, risk gating, UI presentation, and verification.
    /// </summary>
    public class OfficeAction
    {
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

        // Execution & Presentation state
        [JsonIgnore]
        public OfficeActionStatus Status { get; set; }

        [JsonIgnore]
        public string ResultText { get; set; }

        [JsonIgnore]
        public string ErrorMessage { get; set; }

        [JsonIgnore]
        public bool IsUndoable { get; set; }

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
                    default: return "act";
                }
            }
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

        public OfficeAction()
        {
            ActionId = Guid.NewGuid().ToString("N");
            Target = new ActionTarget();
            Parameters = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            Evidence = new List<EvidenceClaim>();
            Status = OfficeActionStatus.Pending;
            RequiresApproval = true;
            IsUndoable = true;
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
            if (op.Contains("formula")) type = SpreadsheetActionType.Formula;
            else if (op.Contains("fill_down") || op.Contains("filldown")) type = SpreadsheetActionType.FillDown;
            else if (op.Contains("create_table") || op.Contains("createtable")) type = SpreadsheetActionType.CreateTable;
            else if (op.Contains("table")) type = SpreadsheetActionType.Table;
            else if (op.Contains("conditional_format") || op.Contains("conditionalformat")) type = SpreadsheetActionType.ConditionalFormat;
            else if (op.Contains("sort")) type = SpreadsheetActionType.Sort;
            else if (op.Contains("filter")) type = SpreadsheetActionType.Filter;
            else if (op.Contains("data_validation") || op.Contains("datavalidation")) type = SpreadsheetActionType.DataValidation;
            else if (op.Contains("chart")) type = SpreadsheetActionType.Chart;
            else if (op.Contains("pivot")) type = SpreadsheetActionType.PivotTable;
            else if (op.Contains("named_range") || op.Contains("namedrange")) type = SpreadsheetActionType.NamedRange;
            else if (op.Contains("remove_duplicates") || op.Contains("removeduplicates")) type = SpreadsheetActionType.RemoveDuplicates;
            else if (op.Contains("value")) type = SpreadsheetActionType.Value;

            if (!type.HasValue) return null;

            string content = GetParameterString("value") ?? GetParameterString("content") ?? GetParameterString("formula") ?? string.Empty;
            string target = Target != null ? (Target.Range ?? Target.ToString()) : "A1";

            return new SpreadsheetAction
            {
                Type = type.Value,
                Target = target,
                Content = content,
                Description = PreviewDescription,
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
            if (sa == null) return null;
            var action = new OfficeAction
            {
                Host = "Excel",
                Operation = "excel." + sa.Type.ToString().ToLowerInvariant(),
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
                action.RiskLevel = 2;
            }
            else if (sa.Type == SpreadsheetActionType.Value)
            {
                action.Parameters["value"] = sa.Content;
                action.RiskLevel = 1;
            }
            else if (sa.Type == SpreadsheetActionType.RemoveDuplicates)
            {
                action.RiskLevel = 3;
                action.IsUndoable = false;
            }
            else
            {
                action.RiskLevel = 2;
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

            if (type == null) return null;

            return new PowerPointAction
            {
                Type = type,
                Source = GetParameterInt("source"),
                Target = GetParameterInt("target"),
                Slide = Target != null && Target.Slide.HasValue ? Target.Slide.Value : GetParameterInt("slide"),
                Section = GetParameterInt("section"),
                Name = GetParameterString("name"),
                Notes = GetParameterString("notes"),
                Status = ToPowerPointActionStatus(Status),
                ErrorMessage = ErrorMessage,
                ResultText = ResultText
            };
        }

        /// <summary>
        /// Creates an OfficeAction from a legacy PowerPointAction.
        /// </summary>
        public static OfficeAction FromPowerPointAction(PowerPointAction pa)
        {
            if (pa == null) return null;
            var action = new OfficeAction
            {
                Host = "PowerPoint",
                Operation = "powerpoint." + (pa.Type ?? "action").ToLowerInvariant(),
                ExpectedResult = pa.Description,
                SourceReason = pa.Description,
                Status = FromPowerPointActionStatus(pa.Status),
                ErrorMessage = pa.ErrorMessage,
                ResultText = pa.ResultText,
                RiskLevel = string.Equals(pa.Type, "set_notes", StringComparison.OrdinalIgnoreCase) ? 1 : 2
            };

            if (pa.Slide > 0) action.Target.Slide = pa.Slide;
            if (pa.Source > 0) action.Parameters["source"] = pa.Source;
            if (pa.Target > 0) action.Parameters["target"] = pa.Target;
            if (pa.Section > 0) action.Parameters["section"] = pa.Section;
            if (!string.IsNullOrEmpty(pa.Name)) action.Parameters["name"] = pa.Name;
            if (!string.IsNullOrEmpty(pa.Notes)) action.Parameters["notes"] = pa.Notes;

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
