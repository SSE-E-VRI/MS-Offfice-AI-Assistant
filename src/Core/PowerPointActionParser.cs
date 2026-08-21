using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace MSOfficeAIAssistant.Core
{
    public class PowerPointAction
    {
        public string Type { get; set; }
        public int Source { get; set; }
        public int Target { get; set; }
        public int Slide { get; set; }
        public int Section { get; set; }
        public string Name { get; set; }
        public string Notes { get; set; }
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
        private static readonly HashSet<string> AllowedActionTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "move_slide",
            "create_section",
            "rename_section",
            "set_notes"
        };

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
                        if (reader.NodeType != XmlNodeType.Element || !string.Equals(reader.Name, "powerpoint_action", StringComparison.OrdinalIgnoreCase))
                            continue;

                        string type = (reader.GetAttribute("type") ?? string.Empty).Trim().ToLowerInvariant();
                        if (!AllowedActionTypes.Contains(type))
                        {
                            Logger.Warn(string.Format("PowerPointActionParser: Disallowed or unknown action type '{0}' skipped.", type));
                            continue;
                        }

                        int source = ParsePositiveInt(reader.GetAttribute("source"));
                        int target = ParsePositiveInt(reader.GetAttribute("target"));
                        int slide = ParsePositiveInt(reader.GetAttribute("slide"));
                        int section = ParsePositiveInt(reader.GetAttribute("section"));
                        string name = reader.GetAttribute("name") ?? string.Empty;
                        string notes = reader.GetAttribute("notes") ?? string.Empty;

                        actions.Add(new PowerPointAction
                        {
                            Type = type,
                            Source = source,
                            Target = target,
                            Slide = slide,
                            Section = section,
                            Name = name,
                            Notes = notes
                        });
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
                        line.StartsWith("**Visual:", StringComparison.OrdinalIgnoreCase))
                    {
                        inNotes = false;
                        slide.VisualSuggestion = Regex.Replace(line, @"(?i)^\*?\*?Visual:\*?\*?\s*", "");
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
