using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace MSOfficeAIAssistant.Core
{
    public enum PowerPointActionStatus
    {
        Pending,
        Applying,
        Applied,
        Error
    }

    public class PowerPointAction : INotifyPropertyChanged
    {
        private string _type;
        private int _source;
        private int _target;
        private int _slide;
        private int _section;
        private string _name;
        private string _notes;
        private string _layout;
        private string _shapeType;
        private string _imagePath;
        private string _text;
        private string _altText;
        private string _title;
        private string _chartType;
        private string _data;
        private string _fontName;
        private string _fontSize;
        private string _bold;
        private string _italic;
        private string _color;
        private int _rows;
        private int _cols;
        private PowerPointActionStatus _status = PowerPointActionStatus.Pending;
        private string _resultText;
        private string _errorMessage;

        public string Type
        {
            get { return _type; }
            set { _type = value; OnPropertyChanged("Type"); OnPropertyChanged("TypeBadge"); OnPropertyChanged("Description"); }
        }

        public int Source
        {
            get { return _source; }
            set { _source = value; OnPropertyChanged("Source"); OnPropertyChanged("TargetDisplay"); OnPropertyChanged("Description"); }
        }

        public int Target
        {
            get { return _target; }
            set { _target = value; OnPropertyChanged("Target"); OnPropertyChanged("TargetDisplay"); OnPropertyChanged("Description"); }
        }

        public int Slide
        {
            get { return _slide; }
            set { _slide = value; OnPropertyChanged("Slide"); OnPropertyChanged("TargetDisplay"); OnPropertyChanged("Description"); }
        }

        public int Section
        {
            get { return _section; }
            set { _section = value; OnPropertyChanged("Section"); OnPropertyChanged("TargetDisplay"); OnPropertyChanged("Description"); }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; OnPropertyChanged("Name"); OnPropertyChanged("Description"); OnPropertyChanged("ContentDisplay"); }
        }

        public string Notes
        {
            get { return _notes; }
            set { _notes = value; OnPropertyChanged("Notes"); OnPropertyChanged("Description"); OnPropertyChanged("ContentDisplay"); }
        }

        public string Layout
        {
            get { return _layout; }
            set { _layout = value; OnPropertyChanged("Layout"); OnPropertyChanged("Description"); }
        }

        public string ShapeType
        {
            get { return _shapeType; }
            set { _shapeType = value; OnPropertyChanged("ShapeType"); OnPropertyChanged("Description"); }
        }

        public string ImagePath
        {
            get { return _imagePath; }
            set { _imagePath = value; OnPropertyChanged("ImagePath"); OnPropertyChanged("Description"); }
        }

        public string Text
        {
            get { return _text; }
            set { _text = value; OnPropertyChanged("Text"); OnPropertyChanged("Description"); OnPropertyChanged("ContentDisplay"); }
        }

        public string AltText
        {
            get { return _altText; }
            set { _altText = value; OnPropertyChanged("AltText"); }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; OnPropertyChanged("Title"); OnPropertyChanged("Description"); }
        }

        public string ChartType
        {
            get { return _chartType; }
            set { _chartType = value; OnPropertyChanged("ChartType"); OnPropertyChanged("Description"); }
        }

        public string Data
        {
            get { return _data; }
            set { _data = value; OnPropertyChanged("Data"); }
        }

        public string FontName
        {
            get { return _fontName; }
            set { _fontName = value; OnPropertyChanged("FontName"); }
        }

        public string FontSize
        {
            get { return _fontSize; }
            set { _fontSize = value; OnPropertyChanged("FontSize"); }
        }

        public string Bold
        {
            get { return _bold; }
            set { _bold = value; OnPropertyChanged("Bold"); }
        }

        public string Italic
        {
            get { return _italic; }
            set { _italic = value; OnPropertyChanged("Italic"); }
        }

        public string Color
        {
            get { return _color; }
            set { _color = value; OnPropertyChanged("Color"); }
        }

        public int Rows
        {
            get { return _rows; }
            set { _rows = value; OnPropertyChanged("Rows"); }
        }

        public int Cols
        {
            get { return _cols; }
            set { _cols = value; OnPropertyChanged("Cols"); }
        }

        /// <summary>
        /// Generic catch-all for any extra attributes not mapped above (forward compat).
        /// </summary>
        public Dictionary<string, string> ExtraAttributes { get; set; }

        public PowerPointActionStatus Status
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
            get { return _status == PowerPointActionStatus.Pending; }
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
                    case PowerPointActionStatus.Applied:
                        return string.IsNullOrEmpty(_resultText) ? "✓ Applied" : string.Format("✓ Applied ({0})", _resultText);
                    case PowerPointActionStatus.Error:
                        return string.IsNullOrEmpty(_errorMessage) ? "⚠ Error" : string.Format("⚠ {0}", _errorMessage);
                    case PowerPointActionStatus.Applying:
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
                if (string.Equals(_type, "move_slide", StringComparison.OrdinalIgnoreCase)) return "move";
                if (string.Equals(_type, "create_section", StringComparison.OrdinalIgnoreCase)) return "section+";
                if (string.Equals(_type, "rename_section", StringComparison.OrdinalIgnoreCase)) return "section";
                if (string.Equals(_type, "set_notes", StringComparison.OrdinalIgnoreCase)) return "notes";
                if (string.Equals(_type, "create_slide", StringComparison.OrdinalIgnoreCase)) return "slide+";
                if (string.Equals(_type, "insert_image", StringComparison.OrdinalIgnoreCase)) return "img";
                if (string.Equals(_type, "delete_slide", StringComparison.OrdinalIgnoreCase)) return "del";
                if (string.Equals(_type, "duplicate_slide", StringComparison.OrdinalIgnoreCase)) return "dup";
                if (string.Equals(_type, "hide_slide", StringComparison.OrdinalIgnoreCase)) return "hide";
                if (string.Equals(_type, "unhide_slide", StringComparison.OrdinalIgnoreCase)) return "show";
                if (string.Equals(_type, "apply_layout", StringComparison.OrdinalIgnoreCase)) return "layout";
                if (string.Equals(_type, "set_shape_text", StringComparison.OrdinalIgnoreCase)) return "text";
                if (string.Equals(_type, "replace_text", StringComparison.OrdinalIgnoreCase)) return "replace";
                if (string.Equals(_type, "add_table", StringComparison.OrdinalIgnoreCase)) return "tbl";
                if (string.Equals(_type, "add_chart", StringComparison.OrdinalIgnoreCase)) return "chart";
                if (string.Equals(_type, "add_shape", StringComparison.OrdinalIgnoreCase)) return "shape";
                if (string.Equals(_type, "set_font", StringComparison.OrdinalIgnoreCase)) return "font";
                if (string.Equals(_type, "fit_content", StringComparison.OrdinalIgnoreCase)) return "fit";
                return "action";
            }
        }

        public string TargetDisplay
        {
            get
            {
                if (string.Equals(_type, "move_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0} → {1}", _source, _target);
                if (string.Equals(_type, "create_section", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "rename_section", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Section {0}", _section);
                if (string.Equals(_type, "set_notes", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide);
                if (string.Equals(_type, "create_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "insert_image", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "delete_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "duplicate_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "hide_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "unhide_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "apply_layout", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide);
                if (string.Equals(_type, "set_shape_text", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide);
                if (string.Equals(_type, "add_table", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide);
                if (string.Equals(_type, "add_chart", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide);
                if (string.Equals(_type, "add_shape", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Slide {0}", _slide);
                return "Slide";
            }
        }

        public string Description
        {
            get
            {
                if (string.Equals(_type, "move_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Move slide {0} to position {1}", _source, _target);
                if (string.Equals(_type, "create_section", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Create section '{0}' before slide {1}", _name, _slide > 0 ? _slide : _target);
                if (string.Equals(_type, "rename_section", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Rename section {0} to '{1}'", _section, _name);
                if (string.Equals(_type, "set_notes", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Set speaker notes on slide {0}", _slide);
                if (string.Equals(_type, "create_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Create slide '{0}'", _name ?? "(untitled)");
                if (string.Equals(_type, "insert_image", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Insert image on slide {0}", _slide);
                if (string.Equals(_type, "delete_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Delete slide {0}", _slide);
                if (string.Equals(_type, "duplicate_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Duplicate slide {0}", _slide);
                if (string.Equals(_type, "hide_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Hide slide {0}", _slide);
                if (string.Equals(_type, "unhide_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Unhide slide {0}", _slide);
                if (string.Equals(_type, "apply_layout", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Apply layout to slide {0}", _slide);
                if (string.Equals(_type, "set_shape_text", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Set shape text on slide {0}", _slide);
                if (string.Equals(_type, "replace_text", StringComparison.OrdinalIgnoreCase))
                    return "Replace selected text";
                if (string.Equals(_type, "add_table", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Add table to slide {0}", _slide);
                if (string.Equals(_type, "add_chart", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Add chart to slide {0}", _slide);
                if (string.Equals(_type, "add_shape", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Add shape to slide {0}", _slide);
                if (string.Equals(_type, "set_font", StringComparison.OrdinalIgnoreCase))
                    return "Set font formatting";
                if (string.Equals(_type, "fit_content", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Fit content on slide {0}", _slide);
                return "PowerPoint action";
            }
        }

        public string ContentDisplay
        {
            get
            {
                if (string.Equals(_type, "set_notes", StringComparison.OrdinalIgnoreCase))
                    return _notes ?? string.Empty;
                if (string.Equals(_type, "create_section", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(_type, "rename_section", StringComparison.OrdinalIgnoreCase))
                    return _name ?? string.Empty;
                if (string.Equals(_type, "move_slide", StringComparison.OrdinalIgnoreCase))
                    return string.Format("Position {0} to {1}", _source, _target);
                return string.Empty;
            }
        }

        public bool IsUndoable
        {
            get { return true; }
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
    }

    public class SlideData
    {
        public string Title { get; set; }
        public List<string> Bullets { get; set; }
        public string SpeakerNotes { get; set; }
        public string VisualSuggestion { get; set; }

        public SlideData()
        {
            Bullets = new List<string>();
        }
    }

    public static class PowerPointActionParser
    {
        private static bool IsAllowedActionType(string type)
        {
            if (string.IsNullOrWhiteSpace(type)) return false;
            var tool = ToolRegistry.GetTool(type, "PowerPoint");
            return tool != null;
        }

        public static List<PowerPointAction> ParseStructuredActions(string rawText)
        {
            string dummy;
            return ParseStructuredActions(rawText, out dummy);
        }

        public static List<PowerPointAction> ParseStructuredActions(string rawText, out string cleanedText)
        {
            var actions = new List<PowerPointAction>();
            cleanedText = rawText ?? string.Empty;
            if (string.IsNullOrWhiteSpace(rawText)) return actions;

            var match = Regex.Match(rawText, @"<powerpoint_actions\b[^>]*>([\s\S]*?)<\/powerpoint_actions>", RegexOptions.IgnoreCase);
            if (!match.Success) return actions;

            cleanedText = rawText.Replace(match.Value, "").Trim();

            try
            {
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreComments = true,
                    IgnoreWhitespace = true
                };

                using (var reader = XmlReader.Create(new StringReader(match.Value), settings))
                {
                    while (reader.Read())
                    {
                        if (reader.NodeType != XmlNodeType.Element ||
                            (!string.Equals(reader.Name, "powerpoint_action", StringComparison.OrdinalIgnoreCase) &&
                             !string.Equals(reader.Name, "action", StringComparison.OrdinalIgnoreCase)))
                            continue;

                        string type = (reader.GetAttribute("type") ?? string.Empty).Trim().ToLowerInvariant();
                        if (!IsAllowedActionType(type))
                        {
                            Logger.Warn(string.Format("PowerPointActionParser: Disallowed or unknown action type '{0}' skipped.", type));
                            continue;
                        }

                        int source = ParsePositiveInt(reader.GetAttribute("source"));
                        int target = ParsePositiveInt(reader.GetAttribute("target"));
                        int slide = ParsePositiveInt(reader.GetAttribute("slide"));
                        if (slide == 0) slide = ParsePositiveInt(reader.GetAttribute("index"));
                        int section = ParsePositiveInt(reader.GetAttribute("section"));
                        string name = reader.GetAttribute("name") ?? string.Empty;
                        string notes = reader.GetAttribute("notes") ?? reader.GetAttribute("speaker_notes") ?? string.Empty;
                        string layout = reader.GetAttribute("layout") ?? reader.GetAttribute("layout_name") ?? string.Empty;
                        string shapeType = reader.GetAttribute("shape_type") ?? reader.GetAttribute("shape") ?? reader.GetAttribute("shapeType") ?? string.Empty;
                        string imagePath = reader.GetAttribute("image_path") ?? reader.GetAttribute("image") ?? reader.GetAttribute("path") ?? reader.GetAttribute("file") ?? string.Empty;
                        string text = reader.GetAttribute("text") ?? reader.GetAttribute("content") ?? string.Empty;
                        string title = reader.GetAttribute("title") ?? string.Empty;
                        string chartType = reader.GetAttribute("chart_type") ?? reader.GetAttribute("chartType") ?? reader.GetAttribute("chart") ?? string.Empty;
                        string altText = reader.GetAttribute("alt_text") ?? reader.GetAttribute("alt") ?? string.Empty;
                        string data = reader.GetAttribute("data") ?? string.Empty;
                        int rows = ParsePositiveInt(reader.GetAttribute("rows"));
                        int cols = ParsePositiveInt(reader.GetAttribute("cols"));
                        if (rows == 0) rows = ParsePositiveInt(reader.GetAttribute("row"));
                        if (cols == 0) cols = ParsePositiveInt(reader.GetAttribute("col"));
                        string fontName = reader.GetAttribute("font_name") ?? reader.GetAttribute("font") ?? string.Empty;
                        string fontSize = reader.GetAttribute("font_size") ?? reader.GetAttribute("size") ?? string.Empty;
                        string bold = reader.GetAttribute("bold") ?? string.Empty;
                        string italic = reader.GetAttribute("italic") ?? string.Empty;
                        string color = reader.GetAttribute("color") ?? string.Empty;

                        var extra = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        string[] extraKeys = new[] { "headers", "values", "table", "points", "bullets", "outline", "operation", "target_text" };
                        foreach (var k in extraKeys)
                        {
                            string v = reader.GetAttribute(k);
                            if (!string.IsNullOrWhiteSpace(v)) extra[k] = v;
                        }

                        // inner element text as fallback for data/text/notes where attribute was empty
                        bool isEmpty = reader.IsEmptyElement;
                        string innerText = string.Empty;
                        if (!isEmpty)
                        {
                            try { innerText = reader.ReadElementContentAsString(); } catch { innerText = string.Empty; }
                            innerText = (innerText ?? string.Empty).Trim();
                            if (string.IsNullOrWhiteSpace(text) && !string.IsNullOrWhiteSpace(innerText) && (type == "set_shape_text" || type == "replace_text" || type == "set_notes" || type == "create_slide" || type == "add_shape"))
                                text = innerText;
                            if (string.IsNullOrWhiteSpace(data) && !string.IsNullOrWhiteSpace(innerText) && (type == "add_table" || type == "add_shape" || type == "add_chart"))
                                data = innerText;
                            if (string.IsNullOrWhiteSpace(notes) && !string.IsNullOrWhiteSpace(innerText) && type == "set_notes")
                                notes = innerText;
                        }

                        var action = new PowerPointAction
                        {
                            Type = type,
                            Source = source,
                            Target = target,
                            Slide = slide,
                            Section = section,
                            Name = name,
                            Notes = notes,
                            Layout = layout,
                            ShapeType = shapeType,
                            ImagePath = imagePath,
                            Text = text,
                            Title = title,
                            ChartType = chartType,
                            AltText = altText,
                            Data = data,
                            Rows = rows,
                            Cols = cols,
                            FontName = fontName,
                            FontSize = fontSize,
                            Bold = bold,
                            Italic = italic,
                            Color = color,
                            ExtraAttributes = extra
                        };
                        actions.Add(action);
                        if (!isEmpty)
                        {
                            // ReadElementContentAsString already advanced past end element
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointActionParser failed to parse actions XML: {0}", ex.Message));
            }

            return actions;
        }

        public static List<SlideData> ParseSlideData(string rawText)
        {
            var result = new List<SlideData>();
            if (string.IsNullOrWhiteSpace(rawText)) return result;

            string[] rawBlocks = Regex.Split(rawText, @"(?m)(?=^(?:#{1,3}\s+|Slide\s+\d+[:.]|\s*---\s*$|\s*\*\*\*\s*$|\s*___\s*$))", RegexOptions.IgnoreCase);

            var blocks = new List<string>();
            foreach (var b in rawBlocks)
            {
                if (!string.IsNullOrWhiteSpace(b)) blocks.Add(b.Trim());
            }

            if (blocks.Count == 0)
            {
                blocks.Add(rawText.Trim());
            }

            foreach (var block in blocks)
            {
                var slide = new SlideData();
                var lines = block.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                bool inNotes = false;
                var notesSb = new StringBuilder();

                foreach (var rawLine in lines)
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    if (line == "---" || line == "***" || line == "___") continue;

                    if (line.StartsWith("Visual:", StringComparison.OrdinalIgnoreCase) ||
                        line.StartsWith("Visual suggestion:", StringComparison.OrdinalIgnoreCase) ||
                        line.StartsWith("**Visual:", StringComparison.OrdinalIgnoreCase) ||
                        line.StartsWith("**Visual suggestion:", StringComparison.OrdinalIgnoreCase))
                    {
                        inNotes = false;
                        slide.VisualSuggestion = Regex.Replace(line, @"(?i)^\*?\*?Visual(?:\s+suggestion)?:\*?\*?\s*", "");
                        continue;
                    }

                    if (line.StartsWith("Speaker Notes:", StringComparison.OrdinalIgnoreCase) ||
                        line.StartsWith("**Speaker Notes:**", StringComparison.OrdinalIgnoreCase) ||
                        line.StartsWith("Notes:", StringComparison.OrdinalIgnoreCase))
                    {
                        inNotes = true;
                        string noteContent = Regex.Replace(line, @"(?i)^\*?\*?(?:Speaker\s+)?Notes:\*?\*?\s*", "");
                        if (!string.IsNullOrWhiteSpace(noteContent)) notesSb.AppendLine(noteContent);
                        continue;
                    }

                    if (inNotes)
                    {
                        if (line.StartsWith("#") || line.StartsWith("-") || line.StartsWith("*") || line.StartsWith("•"))
                        {
                            inNotes = false;
                        }
                        else
                        {
                            notesSb.AppendLine(line);
                            continue;
                        }
                    }

                    if (string.IsNullOrEmpty(slide.Title) &&
                        (line.StartsWith("#") ||
                         line.StartsWith("Title:", StringComparison.OrdinalIgnoreCase) ||
                         line.StartsWith("**Title:**", StringComparison.OrdinalIgnoreCase) ||
                         line.StartsWith("**Slide", StringComparison.OrdinalIgnoreCase) ||
                         line.StartsWith("Slide", StringComparison.OrdinalIgnoreCase)))
                    {
                        string titleText = Regex.Replace(line, @"^(?:#{1,4}\s*)?(?:(?:\*?\*?Title:\*?\*?|\*?\*?Slide\s*\d*[:.]?)\s*)*", "", RegexOptions.IgnoreCase);
                        slide.Title = CleanMarkdown(titleText);
                        continue;
                    }

                    if (Regex.IsMatch(line, @"^\*?\*?Content(?:\s*\(.*?\))?:\*?\*?$", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(line, @"^\*?\*?Bullet Points:\*?\*?$", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(line, @"^\*?\*?Design Tip:.*?\*?\*?$", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    string bulletText = Regex.Replace(line, @"^[-*•+]\s*|^\d+[\.)]\s*", "");
                    slide.Bullets.Add(CleanMarkdown(bulletText));
                }

                if (notesSb.Length > 0)
                {
                    slide.SpeakerNotes = notesSb.ToString().Trim();
                }

                if (!string.IsNullOrWhiteSpace(slide.Title) || slide.Bullets.Count > 0)
                {
                    result.Add(slide);
                }
            }

            return result;
        }

        public static string CleanMarkdown(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            string s = input;
            s = Regex.Replace(s, @"<powerpoint_actions>[\s\S]*?</powerpoint_actions>", "", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"```[a-zA-Z]*\n?|```", "");
            s = Regex.Replace(s, @"\*\*([^*]+)\*\*", "$1");
            s = Regex.Replace(s, @"\*([^*]+)\*", "$1");
            s = Regex.Replace(s, @"__([^_]+)__", "$1");
            s = Regex.Replace(s, @"_([^_]+)_", "$1");
            s = Regex.Replace(s, @"`([^`]+)`", "$1");
            s = Regex.Replace(s, @"^#{1,6}\s*", "", RegexOptions.Multiline);
            s = Regex.Replace(s, @"^[-*•+]\s*", "", RegexOptions.Multiline);
            s = s.Replace("**", "").Replace("##", "").Replace("###", "");

            return s.Trim();
        }

        private static int ParsePositiveInt(string value)
        {
            int parsed;
            return int.TryParse(value, out parsed) && parsed > 0 ? parsed : 0;
        }
    }
}
