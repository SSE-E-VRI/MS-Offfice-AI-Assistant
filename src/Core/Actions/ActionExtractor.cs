using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MSOfficeAIAssistant.Core.Actions
{
    public class ExtractionFailure
    {
        public string RawSnippet { get; set; }
        public string ErrorMessage { get; set; }
        public int LineNumber { get; set; }
        public string FailureType { get; set; }

        public override string ToString()
        {
            return string.Format("[{0}] {1} (Line {2})", FailureType, ErrorMessage, LineNumber);
        }
    }

    public class ExtractionResult
    {
        public List<OfficeAction> Actions { get; set; }
        public string CleanText { get; set; }
        public ExtractionFailure Failure { get; set; }

        public bool HasActions
        {
            get { return Actions != null && Actions.Count > 0; }
        }

        public bool HasFailure
        {
            get { return Failure != null; }
        }

        public ExtractionResult()
        {
            Actions = new List<OfficeAction>();
            CleanText = string.Empty;
        }
    }

    /// <summary>
    /// Unified action extractor accepting native provider tool calls, structured JSON blocks,
    /// and legacy XML formats (<excel_actions>, <powerpoint_actions>).
    /// Emits structured ExtractionResult with typed ExtractionFailure when malformed.
    /// </summary>
    public static class ActionExtractor
    {
        private static readonly Regex OfficeActionsBlockRegex = new Regex(
            @"<office_actions>([\s\S]*?)</office_actions>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ExcelActionsBlockRegex = new Regex(
            @"<excel_actions>[\s\S]*?</excel_actions>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex PowerPointActionsBlockRegex = new Regex(
            @"<powerpoint_actions>[\s\S]*?</powerpoint_actions>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex JsonCodeBlockRegex = new Regex(
            @"```(?:json)?\s*(\[\s*\{[\s\S]*?\}\s*\])\s*```",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Extracts actions from assistant text and optional native tool calls.
        /// </summary>
        public static ExtractionResult Extract(string text, string currentHost = null, IEnumerable<dynamic> nativeToolCalls = null)
        {
            var result = new ExtractionResult();
            if (string.IsNullOrEmpty(text) && nativeToolCalls == null)
            {
                return result;
            }

            string clean = text ?? string.Empty;

            // 1. Process native tool calls if present
            if (nativeToolCalls != null)
            {
                foreach (var tc in nativeToolCalls)
                {
                    if (tc == null) continue;
                    try
                    {
                        string name = GetPropertyValue(tc, "Name") ?? GetPropertyValue(tc, "name");
                        if (string.IsNullOrEmpty(name))
                        {
                            object fn = GetPropertyObj(tc, "function") ?? GetPropertyObj(tc, "Function");
                            if (fn != null)
                            {
                                name = GetPropertyValue(fn, "name") ?? GetPropertyValue(fn, "Name");
                            }
                        }

                        string argsJson = GetPropertyValue(tc, "Arguments") ?? GetPropertyValue(tc, "arguments");
                        if (string.IsNullOrEmpty(argsJson))
                        {
                            object fn = GetPropertyObj(tc, "function") ?? GetPropertyObj(tc, "Function");
                            if (fn != null)
                            {
                                argsJson = GetPropertyValue(fn, "arguments") ?? GetPropertyValue(fn, "Arguments");
                            }
                        }

                        var action = CreateActionFromNativeTool(name, argsJson, currentHost);
                        if (action != null)
                        {
                            result.Actions.Add(action);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Failure = new ExtractionFailure
                        {
                            FailureType = "NativeToolCallParseError",
                            ErrorMessage = ex.Message,
                            RawSnippet = Convert.ToString(tc)
                        };
                    }
                }
            }

            // 2. Check for explicit <office_actions> JSON blocks
            var officeMatch = OfficeActionsBlockRegex.Match(clean);
            if (officeMatch.Success)
            {
                string jsonPayload = officeMatch.Groups[1].Value.Trim();
                try
                {
                    var parsed = JsonConvert.DeserializeObject<List<OfficeAction>>(jsonPayload);
                    if (parsed != null && parsed.Count > 0)
                    {
                        foreach (var act in parsed)
                        {
                            if (string.IsNullOrEmpty(act.Host) && !string.IsNullOrEmpty(currentHost))
                            {
                                act.Host = currentHost;
                            }
                            result.Actions.Add(act);
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Failure = new ExtractionFailure
                    {
                        FailureType = "MalformedJsonOfficeActions",
                        ErrorMessage = ex.Message,
                        RawSnippet = jsonPayload
                    };
                }

                clean = OfficeActionsBlockRegex.Replace(clean, string.Empty).Trim();
            }

            // 3. Check for standalone JSON array codeblocks if no actions found yet
            if (result.Actions.Count == 0 && !result.HasFailure)
            {
                var jsonMatch = JsonCodeBlockRegex.Match(clean);
                if (jsonMatch.Success)
                {
                    string codeJson = jsonMatch.Groups[1].Value.Trim();
                    try
                    {
                        // Test if it contains action fields like "operation" or "host"
                        if (codeJson.Contains("\"operation\"") || codeJson.Contains("\"target\""))
                        {
                            var parsed = JsonConvert.DeserializeObject<List<OfficeAction>>(codeJson);
                            if (parsed != null && parsed.Count > 0)
                            {
                                foreach (var act in parsed)
                                {
                                    if (string.IsNullOrEmpty(act.Host) && !string.IsNullOrEmpty(currentHost))
                                    {
                                        act.Host = currentHost;
                                    }
                                    result.Actions.Add(act);
                                }
                                clean = JsonCodeBlockRegex.Replace(clean, string.Empty).Trim();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Only flag failure if it clearly looked like an action block
                        if (codeJson.Contains("\"operation\""))
                        {
                            result.Failure = new ExtractionFailure
                            {
                                FailureType = "MalformedJsonCodeBlock",
                                ErrorMessage = ex.Message,
                                RawSnippet = codeJson
                            };
                        }
                    }
                }
            }

            // 4. Check for legacy <excel_actions> XML block
            if (clean.Contains("<excel_actions>"))
            {
                try
                {
                    string dummy;
                    var excelActions = SpreadsheetActionParser.ExtractActions(clean, out dummy);
                    if (excelActions != null && excelActions.Count > 0)
                    {
                        foreach (var ea in excelActions)
                        {
                            result.Actions.Add(OfficeAction.FromSpreadsheetAction(ea));
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Failure = new ExtractionFailure
                    {
                        FailureType = "MalformedExcelActionsXml",
                        ErrorMessage = ex.Message,
                        RawSnippet = clean
                    };
                }

                clean = ExcelActionsBlockRegex.Replace(clean, string.Empty).Trim();
            }

            // 5. Check for legacy <powerpoint_actions> XML block
            if (clean.Contains("<powerpoint_actions>"))
            {
                try
                {
                    string dummy;
                    var pptActions = PowerPointActionParser.ParseStructuredActions(clean, out dummy);
                    if (pptActions != null && pptActions.Count > 0)
                    {
                        foreach (var pa in pptActions)
                        {
                            result.Actions.Add(OfficeAction.FromPowerPointAction(pa));
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Failure = new ExtractionFailure
                    {
                        FailureType = "MalformedPowerPointActionsXml",
                        ErrorMessage = ex.Message,
                        RawSnippet = clean
                    };
                }

                clean = PowerPointActionsBlockRegex.Replace(clean, string.Empty).Trim();
            }

            result.CleanText = clean;
            return result;
        }

        private static OfficeAction CreateActionFromNativeTool(string functionName, string argsJson, string defaultHost)
        {
            if (string.IsNullOrWhiteSpace(functionName)) return null;

            string normalizedOp = functionName.Trim();
            string host = defaultHost;
            if (normalizedOp.StartsWith("excel_", StringComparison.OrdinalIgnoreCase))
            {
                normalizedOp = "excel." + normalizedOp.Substring(6);
                host = "Excel";
            }
            else if (normalizedOp.StartsWith("word_", StringComparison.OrdinalIgnoreCase))
            {
                normalizedOp = "word." + normalizedOp.Substring(5);
                host = "Word";
            }
            else if (normalizedOp.StartsWith("powerpoint_", StringComparison.OrdinalIgnoreCase))
            {
                normalizedOp = "powerpoint." + normalizedOp.Substring(11);
                host = "PowerPoint";
            }
            else if (normalizedOp.StartsWith("ppt_", StringComparison.OrdinalIgnoreCase))
            {
                normalizedOp = "powerpoint." + normalizedOp.Substring(4);
                host = "PowerPoint";
            }
            else if (normalizedOp.StartsWith("excel.", StringComparison.OrdinalIgnoreCase))
            {
                host = "Excel";
            }
            else if (normalizedOp.StartsWith("word.", StringComparison.OrdinalIgnoreCase))
            {
                host = "Word";
            }
            else if (normalizedOp.StartsWith("powerpoint.", StringComparison.OrdinalIgnoreCase) ||
                     normalizedOp.StartsWith("ppt.", StringComparison.OrdinalIgnoreCase))
            {
                host = "PowerPoint";
            }

            var action = new OfficeAction
            {
                Host = host,
                Operation = normalizedOp,
                ActionId = Guid.NewGuid().ToString("N")
            };

            if (!string.IsNullOrWhiteSpace(argsJson))
            {
                var dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(argsJson);
                if (dict != null)
                {
                    action.Parameters = dict;
                    if (dict.ContainsKey("target"))
                    {
                        string tgtStr = Convert.ToString(dict["target"]);
                        if (!string.IsNullOrEmpty(tgtStr))
                        {
                            if (tgtStr.Contains("!"))
                            {
                                var parts = tgtStr.Split('!');
                                action.Target.Sheet = parts[0];
                                action.Target.Range = parts.Length > 1 ? parts[1] : string.Empty;
                            }
                            else
                            {
                                action.Target.Range = tgtStr;
                            }
                        }
                    }
                    if (dict.ContainsKey("expected_result"))
                    {
                        action.ExpectedResult = Convert.ToString(dict["expected_result"]);
                    }
                    if (dict.ContainsKey("reason") || dict.ContainsKey("description"))
                    {
                        action.SourceReason = Convert.ToString(dict.ContainsKey("reason") ? dict["reason"] : dict["description"]);
                    }
                }
            }

            return action;
        }

        private static string GetPropertyValue(object obj, string propName)
        {
            if (obj == null || string.IsNullOrEmpty(propName)) return null;
            try
            {
                var prop = obj.GetType().GetProperty(propName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                if (prop != null)
                {
                    object val = prop.GetValue(obj, null);
                    return val != null ? Convert.ToString(val) : null;
                }
            }
            catch { }
            return null;
        }

        private static object GetPropertyObj(object obj, string propName)
        {
            if (obj == null || string.IsNullOrEmpty(propName)) return null;
            try
            {
                var prop = obj.GetType().GetProperty(propName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                if (prop != null)
                {
                    return prop.GetValue(obj, null);
                }
            }
            catch { }
            return null;
        }
    }
}
