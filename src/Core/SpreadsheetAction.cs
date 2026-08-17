using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace MistralOfficeAddin.Core
{
    public enum SpreadsheetActionType
    {
        Formula,
        Value,
        FillDown,
        Table
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
                    default: return "val";
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
        // Validates standard Excel cell (e.g. A1, K20) or range (e.g. G2:G27, A1:F50)
        private static readonly Regex CellAddressRegex = new Regex(@"^[A-Za-z]{1,3}[1-9][0-9]{0,6}(?::[A-Za-z]{1,3}[1-9][0-9]{0,6})?$", RegexOptions.Compiled);

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

                            if (string.IsNullOrWhiteSpace(target) || string.IsNullOrWhiteSpace(content))
                                continue;

                            target = target.Trim().ToUpperInvariant();
                            content = content.Trim();

                            // Validate Excel target address
                            if (!CellAddressRegex.IsMatch(target))
                            {
                                Logger.Warn(string.Format("SpreadsheetActionParser: Invalid Excel target address '{0}', skipped.", target));
                                continue;
                            }

                            SpreadsheetActionType actionType;
                            switch (typeStr.ToLowerInvariant())
                            {
                                case "formula":
                                    actionType = SpreadsheetActionType.Formula;
                                    if (!content.StartsWith("=")) content = "=" + content;
                                    break;
                                case "filldown":
                                    actionType = SpreadsheetActionType.FillDown;
                                    if (!content.StartsWith("=")) content = "=" + content;
                                    break;
                                case "table":
                                    actionType = SpreadsheetActionType.Table;
                                    break;
                                default:
                                    actionType = SpreadsheetActionType.Value;
                                    break;
                            }

                            actions.Add(new SpreadsheetAction
                            {
                                Target = target,
                                Type = actionType,
                                Content = content,
                                Description = desc.Trim(),
                                Status = SpreadsheetActionStatus.Pending
                            });
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
    }
}
