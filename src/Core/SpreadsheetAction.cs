using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace MSOfficeAIAssistant.Core
{
    public enum SpreadsheetActionType
    {
        Formula,
        Value,
        FillDown,
        Table,
        CreateTable,
        ConditionalFormat,
        Sort,
        Filter,
        DataValidation,
        Chart,
        PivotTable,
        NamedRange,
        RemoveDuplicates
    }

    public enum SpreadsheetActionStatus
    {
        Pending,
        Applying,
        Applied,
        Error
    }

    public class SpreadsheetAction : INotifyPropertyChanged
    {
        private string _target;
        private SpreadsheetActionType _type;
        private string _content;
        private string _description;
        private SpreadsheetActionStatus _status = SpreadsheetActionStatus.Pending;
        private string _resultText;
        private string _errorMessage;

        public string Target
        {
            get { return _target; }
            set { _target = value; OnPropertyChanged("Target"); }
        }

        public SpreadsheetActionType Type
        {
            get { return _type; }
            set { _type = value; OnPropertyChanged("Type"); }
        }

        public string Content
        {
            get { return _content; }
            set { _content = value; OnPropertyChanged("Content"); }
        }

        public string Description
        {
            get { return _description; }
            set { _description = value; OnPropertyChanged("Description"); }
        }

        public SpreadsheetActionStatus Status
        {
            get { return _status; }
            set
            {
                _status = value;
                OnPropertyChanged("Status");
                OnPropertyChanged("StatusDisplay");
                OnPropertyChanged("IsPending");
            }
        }

        public bool IsPending
        {
            get { return _status == SpreadsheetActionStatus.Pending; }
        }

        public string ResultText
        {
            get { return _resultText; }
            set { _resultText = value; OnPropertyChanged("ResultText"); OnPropertyChanged("StatusDisplay"); }
        }

        public string ErrorMessage
        {
            get { return _errorMessage; }
            set { _errorMessage = value; OnPropertyChanged("ErrorMessage"); OnPropertyChanged("StatusDisplay"); }
        }

        public string StatusDisplay
        {
            get
            {
                switch (_status)
                {
                    case SpreadsheetActionStatus.Applied:
                        return string.IsNullOrEmpty(_resultText) ? "✓ Applied" : string.Format("✓ Applied ({0})", _resultText);
                    case SpreadsheetActionStatus.Error:
                        return string.IsNullOrEmpty(_errorMessage) ? "⚠ Error" : string.Format("⚠ {0}", _errorMessage);
                    case SpreadsheetActionStatus.Applying:
                        return "Applying...";
                    default:
                        return "Pending";
                }
            }
        }

        public string TypeBadge
        {
            get
            {
                switch (_type)
                {
                    case SpreadsheetActionType.Formula: return "fx";
                    case SpreadsheetActionType.FillDown: return "fill";
                    case SpreadsheetActionType.Table: return "table";
                    case SpreadsheetActionType.CreateTable: return "table+";
                    case SpreadsheetActionType.ConditionalFormat: return "format";
                    case SpreadsheetActionType.Sort: return "sort";
                    case SpreadsheetActionType.Filter: return "filter";
                    case SpreadsheetActionType.DataValidation: return "validate";
                    case SpreadsheetActionType.Chart: return "chart";
                    case SpreadsheetActionType.PivotTable: return "pivot";
                    case SpreadsheetActionType.NamedRange: return "name";
                    case SpreadsheetActionType.RemoveDuplicates: return "dedupe";
                    default: return "val";
                }
            }
        }

        public bool IsUndoable
        {
            get
            {
                switch (_type)
                {
                    case SpreadsheetActionType.RemoveDuplicates:
                    case SpreadsheetActionType.CreateTable:
                    case SpreadsheetActionType.Chart:
                    case SpreadsheetActionType.PivotTable:
                    case SpreadsheetActionType.NamedRange:
                        return false;
                    default:
                        return true;
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }
    }

    public static class SpreadsheetActionParser
    {
        private const int MaxActionsPerResponse = 25;
        private const int MaxActionContentLength = 12000;
        private const int MaxRangeCells = 100000;
        private const int MaxExcelRow = 1048576;
        private const int MaxExcelColumn = 16384;

        // Validates one bounded, standard A1 cell or range. Sheet-qualified references,
        // whole columns/rows, and multi-area ranges are intentionally excluded so an AI
        // response cannot turn a small request into a workbook-wide edit.
        private static readonly Regex CellAddressRegex = new Regex(
            @"^([A-Za-z]{1,3})([1-9][0-9]{0,6})(?::([A-Za-z]{1,3})([1-9][0-9]{0,6}))?$",
            RegexOptions.Compiled);

        public static bool IsSafeTarget(string target)
        {
            int startCol, startRow, endCol, endRow;
            return TryParseRangeExtent(target, out startCol, out startRow, out endCol, out endRow);
        }

        public static bool TryParseRangeExtent(string target, out int startColumn, out int startRow, out int endColumn, out int endRow)
        {
            startColumn = 0;
            startRow = 0;
            endColumn = 0;
            endRow = 0;

            if (string.IsNullOrWhiteSpace(target)) return false;

            Match match = CellAddressRegex.Match(target.Trim());
            if (!match.Success) return false;

            startColumn = ColumnLetterToIndex(match.Groups[1].Value);
            if (!int.TryParse(match.Groups[2].Value, out startRow)) return false;

            endColumn = startColumn;
            endRow = startRow;
            if (match.Groups[3].Success)
            {
                endColumn = ColumnLetterToIndex(match.Groups[3].Value);
                if (!int.TryParse(match.Groups[4].Value, out endRow)) return false;
            }

            return IsSafeRangeBounds(startColumn, startRow, endColumn, endRow);
        }

        public static bool IsSafeRangeBounds(int startColumn, int startRow, int endColumn, int endRow)
        {
            if (startColumn < 1 || endColumn < startColumn || endColumn > MaxExcelColumn ||
                startRow < 1 || endRow < startRow || endRow > MaxExcelRow)
                return false;

            long cells = (long)(endColumn - startColumn + 1) * (long)(endRow - startRow + 1);
            return cells > 0 && cells <= MaxRangeCells;
        }

        public static string IndexToColumnLetter(int column)
        {
            if (column < 1 || column > MaxExcelColumn) return string.Empty;
            string result = string.Empty;
            while (column > 0)
            {
                int modulo = (column - 1) % 26;
                result = Convert.ToChar('A' + modulo) + result;
                column = (column - modulo) / 26;
            }
            return result;
        }

        public static string BuildRangeAddress(int startColumn, int startRow, int endColumn, int endRow)
        {
            string startLetter = IndexToColumnLetter(startColumn);
            if (string.IsNullOrEmpty(startLetter)) return string.Empty;

            if (startColumn == endColumn && startRow == endRow)
            {
                return string.Format("{0}{1}", startLetter, startRow);
            }

            string endLetter = IndexToColumnLetter(endColumn);
            if (string.IsNullOrEmpty(endLetter)) return string.Empty;

            return string.Format("{0}{1}:{2}{3}", startLetter, startRow, endLetter, endRow);
        }

        public static List<SpreadsheetAction> ExtractActions(string text, out string cleanedText)
        {
            var actions = new List<SpreadsheetAction>();
            cleanedText = text;
            if (string.IsNullOrWhiteSpace(text)) return actions;

            var match = Regex.Match(text, @"<excel_actions\b[^>]*>([\s\S]*?)<\/excel_actions>", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return actions;
            }

            string xmlBlock = match.Value;

            if (xmlBlock.Length > 200000)
            {
                Logger.Warn("SpreadsheetActionParser: Action block exceeded the safety limit and was skipped.");
                return actions;
            }

            // Strip the XML block from user-visible conversation text
            cleanedText = text.Replace(xmlBlock, "").Trim();

            try
            {
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreComments = true,
                    IgnoreWhitespace = true
                };

                using (var stringReader = new StringReader(xmlBlock))
                using (var reader = XmlReader.Create(stringReader, settings))
                {
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element && string.Equals(reader.Name, "excel_action", StringComparison.OrdinalIgnoreCase))
                        {
                            string target = reader.GetAttribute("target");
                            string typeStr = reader.GetAttribute("type") ?? "formula";
                            string formula = reader.GetAttribute("formula");
                            string value = reader.GetAttribute("value");
                            string desc = reader.GetAttribute("description") ?? "";

                            string content = !string.IsNullOrEmpty(formula) ? formula : value;
                            if (string.IsNullOrEmpty(content))
                            {
                                content = reader.ReadElementContentAsString();
                            }

                            if (string.IsNullOrWhiteSpace(target))
                                continue;

                            target = target.Trim().ToUpperInvariant();
                            content = (content ?? string.Empty).Trim();

                            // Validate one bounded Excel target address before calling COM.
                            if (!IsSafeTarget(target))
                            {
                                Logger.Warn(string.Format("SpreadsheetActionParser: Unsafe Excel target address '{0}', skipped.", target));
                                continue;
                            }

                            SpreadsheetActionType actionType;
                            if (!TryParseActionType(typeStr, out actionType))
                            {
                                Logger.Warn(string.Format("SpreadsheetActionParser: Unsupported action type '{0}', skipped.", typeStr));
                                continue;
                            }

                            if (RequiresContent(actionType) && string.IsNullOrWhiteSpace(content))
                            {
                                Logger.Warn(string.Format("SpreadsheetActionParser: Action '{0}' requires content and was skipped.", typeStr));
                                continue;
                            }

                            if (content.Length > MaxActionContentLength)
                            {
                                Logger.Warn(string.Format("SpreadsheetActionParser: Action content for '{0}' exceeded the safety limit and was skipped.", target));
                                continue;
                            }

                            if ((actionType == SpreadsheetActionType.Formula || actionType == SpreadsheetActionType.FillDown) &&
                                !string.IsNullOrEmpty(content) && !content.StartsWith("="))
                                content = "=" + content;

                            actions.Add(new SpreadsheetAction
                            {
                                Target = target,
                                Type = actionType,
                                Content = content,
                                Description = Truncate(desc == null ? string.Empty : desc.Trim(), 1000),
                                Status = SpreadsheetActionStatus.Pending
                            });

                            if (actions.Count >= MaxActionsPerResponse)
                            {
                                Logger.Warn("SpreadsheetActionParser: Action count exceeded the safety limit; remaining actions were skipped.");
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("SpreadsheetActionParser failed to parse actions XML: {0}", ex.Message));
            }

            return actions;
        }

        private static bool TryParseActionType(string typeText, out SpreadsheetActionType actionType)
        {
            actionType = SpreadsheetActionType.Value;
            if (string.IsNullOrWhiteSpace(typeText)) return false;

            var tool = ToolRegistry.GetTool(typeText, "Excel");
            if (tool == null)
            {
                return false;
            }

            string op = tool.Name.ToLowerInvariant();
            if (op == "excel.write_formula") { actionType = SpreadsheetActionType.Formula; return true; }
            if (op == "excel.write_value") { actionType = SpreadsheetActionType.Value; return true; }
            if (op == "excel.fill_down") { actionType = SpreadsheetActionType.FillDown; return true; }
            if (op == "excel.table") { actionType = SpreadsheetActionType.Table; return true; }
            if (op == "excel.create_table") { actionType = SpreadsheetActionType.CreateTable; return true; }
            if (op == "excel.conditional_format") { actionType = SpreadsheetActionType.ConditionalFormat; return true; }
            if (op == "excel.sort") { actionType = SpreadsheetActionType.Sort; return true; }
            if (op == "excel.filter") { actionType = SpreadsheetActionType.Filter; return true; }
            if (op == "excel.data_validation") { actionType = SpreadsheetActionType.DataValidation; return true; }
            if (op == "excel.create_chart") { actionType = SpreadsheetActionType.Chart; return true; }
            if (op == "excel.create_pivot_table") { actionType = SpreadsheetActionType.PivotTable; return true; }
            if (op == "excel.named_range") { actionType = SpreadsheetActionType.NamedRange; return true; }
            if (op == "excel.remove_duplicates") { actionType = SpreadsheetActionType.RemoveDuplicates; return true; }

            return false;
        }

        private static bool RequiresContent(SpreadsheetActionType actionType)
        {
            return actionType == SpreadsheetActionType.Formula ||
                   actionType == SpreadsheetActionType.Value ||
                   actionType == SpreadsheetActionType.FillDown ||
                   actionType == SpreadsheetActionType.Table ||
                   actionType == SpreadsheetActionType.Filter ||
                   actionType == SpreadsheetActionType.DataValidation ||
                   actionType == SpreadsheetActionType.PivotTable ||
                   actionType == SpreadsheetActionType.NamedRange;
        }

        private static int ColumnLetterToIndex(string column)
        {
            if (string.IsNullOrWhiteSpace(column)) return 0;
            int result = 0;
            string normalized = column.Trim().ToUpperInvariant();
            for (int i = 0; i < normalized.Length; i++)
            {
                char c = normalized[i];
                if (c < 'A' || c > 'Z') return 0;
                result = result * 26 + (c - 'A' + 1);
            }
            return result;
        }

        private static string Truncate(string value, int maximumLength)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= maximumLength) return value ?? string.Empty;
            return value.Substring(0, maximumLength) + "...";
        }
    }
}
