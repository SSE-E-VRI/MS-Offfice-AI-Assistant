using System;
using System.Text;

namespace MSOfficeAIAssistant.Core
{
    public enum PromptContextScope
    {
        Selection,
        CurrentFile,
        SelectionAndFile,
        AttachmentsOnly
    }

    /// <summary>
    /// Pure, stateless helper for assembling host-aware system prompts and context-augmented user prompts.
    /// COM-free and directly testable without Office host instances.
    /// </summary>
    public static class PromptAssembler
    {
        public static string BuildHostAwareSystemPrompt(string basePrompt, string hostType)
        {
            string hostContext = "";
            if (string.Equals(hostType, "Excel", StringComparison.OrdinalIgnoreCase))
            {
                hostContext = "\n\nYou are embedded inside Microsoft Excel.\n" +
                    "When generating calculations, formulas, analysis, or spreadsheet changes:\n" +
                    "1. Inspect the provided Worksheet Context with its explicit Column Letters (Col A, Col B, Col C, etc.) and Header names.\n" +
                    "2. When matching category/text columns with prefixes (e.g. '0001-non ferrous items'), use wildcard criteria (e.g. \"*non ferrous*\") in SUMIF/COUNTIF.\n" +
                    "3. Start row-level data formulas at Row 2 (e.g. F2, G2) rather than header Row 1.\n" +
                    "4. Propose only bounded A1 cell/range changes (e.g. B2:B100, not unbounded full columns like B:B). Every change is previewed and requires user confirmation.\n" +
                    "5. ALWAYS return executable spreadsheet actions in a structured <excel_actions> XML block when a native change is requested:\n" +
                    "   <excel_actions>\n" +
                    "     <excel_action target=\"K20\" type=\"formula\" formula=\"=SUMIF(B2:B100, &quot;*non ferrous*&quot;, F2:F100)\" description=\"Total non-ferrous value\" />\n" +
                    "     <excel_action target=\"K21\" type=\"formula\" formula=\"=COUNTIF(E2:E100, 0)\" description=\"Count of zero-quantity items\" />\n" +
                    "     <excel_action target=\"K22\" type=\"formula\" formula=\"=AVERAGEIF(E2:E100, 0, F2:F100)\" description=\"Average value of zero-quantity items\" />\n" +
                    "     <excel_action target=\"G2:G27\" type=\"filldown\" formula=\"=IF(F2&gt;50000, &quot;High Value&quot;, &quot;&quot;)\" description=\"High value flag (&gt;50,000)\" />\n" +
                    "   </excel_actions>\n" +
                    string.Format("6. Supported action types are {0}. ", ToolRegistry.FormatActionTypesList("Excel")) +
                    "Use the value attribute for each action's concise option (for example value=\"descending\", value=\"list:Open,Closed\", or value=\"columns:1,2\") and state any assumptions in the description.\n" +
                    "7. Provide a brief conversational summary above or below the action block without tutorial how-to steps.";
            }
            else if (string.Equals(hostType, "Word", StringComparison.OrdinalIgnoreCase))
            {
                hostContext = "\n\nYou are embedded inside Microsoft Word. The provided document context may include prompt-relevant excerpts, a live outline, and action items. " +
                    "When the user asks to write, edit, rewrite, summarize, translate, or review text, provide polished text directly without tutorial meta-commentary. " +
                    "Preserve factual meaning unless the user requests a substantive change. Use a Markdown table when a table is the clearest result. " +
                    "If sources are attached, cite only supplied sources using [Source: filename, page/section]; do not invent citations. " +
                    "If the message includes [Selected Context] and asks to edit, rewrite, correct, proofread, fix, or polish that text, your entire response must be ONLY the finished replacement text -- " +
                    "no side-by-side original-vs-corrected table, no bulleted list of changes or rationale, no multiple candidate versions, and no headings such as 'Key Improvements' or 'Final Recommendation'. " +
                    "That response is inserted verbatim in place of the selection, so anything else you write lands in the document too. " +
                    "Save that kind of analysis for when the user explicitly asks to review, critique, or explain the text rather than edit it.";
            }
            else if (string.Equals(hostType, "PowerPoint", StringComparison.OrdinalIgnoreCase))
            {
                hostContext = "\n\nYou are embedded inside Microsoft PowerPoint. The supplied context can include the full deck, sections, slide text, and speaker notes. " +
                    "When the user asks for slides or bullet points, provide structured slides with Slide titles, concise bullets, speaker notes, and a one-line Visual suggestion. " +
                    "Preserve the presentation's existing style and structure where possible. " +
                    string.Format("For a requested deck reorganization, use only an optional <powerpoint_actions> block with safe action types {0}; use numbered existing slides and include no other commands.", ToolRegistry.FormatActionTypesList("PowerPoint"));
            }
            return (basePrompt ?? "You are an expert AI assistant embedded inside Microsoft Office.") + hostContext;
        }

        public static bool IncludesSelection(PromptContextScope scope)
        {
            return scope == PromptContextScope.Selection || scope == PromptContextScope.SelectionAndFile;
        }

        public static bool IncludesCurrentFile(PromptContextScope scope)
        {
            return scope == PromptContextScope.CurrentFile || scope == PromptContextScope.SelectionAndFile;
        }

        public static string ComposePromptWithContext(string prompt, PromptContextScope scope, string selectedContext, string currentFileContext)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return prompt ?? string.Empty;
            var builder = new StringBuilder(prompt);

            if (IncludesSelection(scope) && !string.IsNullOrWhiteSpace(selectedContext))
            {
                builder.AppendFormat("\n\n[Selected Context]:\n{0}", selectedContext);
            }

            if (IncludesCurrentFile(scope) && !string.IsNullOrWhiteSpace(currentFileContext))
            {
                builder.AppendFormat("\n\n[Current File Context]:\n{0}", currentFileContext);
            }

            return builder.ToString();
        }

        public static string AppendAttachmentCitationInstruction(string userContent, string textAttachmentContext)
        {
            if (string.IsNullOrWhiteSpace(textAttachmentContext)) return userContent ?? string.Empty;
            string baseContent = userContent ?? string.Empty;
            return string.Format("{0}\n\n{1}\nWhen you rely on an attached source, cite it using [Source: filename, page/section] and do not invent a source.",
                baseContent, textAttachmentContext);
        }

        /// <summary>
        /// Composes an executive briefing deck prompt from topic instructions and/or attached document excerpts.
        /// Formats structured slide outputs with titles, concise bullets, visual suggestions, and speaker notes.
        /// </summary>
        public static string BuildBriefingDeckPrompt(string topicOrInstruction, string documentAttachmentText = null, int targetSlideCount = 5)
        {
            int slideCount = targetSlideCount > 0 ? targetSlideCount : 5;
            var sb = new StringBuilder();
            sb.AppendFormat("Create a concise, executive briefing deck of {0} slides", slideCount);
            if (!string.IsNullOrWhiteSpace(topicOrInstruction))
            {
                sb.AppendFormat(" focusing on: {0}", topicOrInstruction.Trim());
            }
            else
            {
                sb.Append(" based on the attached document");
            }
            sb.Append(".\n\n");
            sb.AppendLine("Structure the response into clear, numbered slide blocks using this exact format:");
            sb.AppendLine("Slide 1: [Executive Title]");
            sb.AppendLine("- [Key takeaway or core thesis]");
            sb.AppendLine("- [Supporting evidence or key context]");
            sb.AppendLine("- [Strategic implication]");
            sb.AppendLine("Visual suggestion: [Clean layout or diagram suggestion]");
            sb.AppendLine("Speaker Notes: [Brief talking points for the presenter]\n");
            sb.AppendLine("Guidelines:");
            sb.AppendLine("1. Extract real facts, metrics, and conclusions from the source material; do not invent details.");
            sb.AppendLine("2. Keep bullet points concise and presentation-ready (1-2 lines each).");
            sb.AppendLine("3. Ensure the narrative flows logically from executive summary to problem/context, key findings, and recommended next steps.");

            if (!string.IsNullOrWhiteSpace(documentAttachmentText))
            {
                sb.AppendFormat("\n[Source Document Excerpts]:\n{0}\n", documentAttachmentText.Trim());
            }

            return sb.ToString().TrimEnd();
        }

        /// <summary>
        /// Appends domain-pack-specific rules to a system prompt.
        /// For "railway" domain: adds railway operational vocabulary guidance.
        /// For "general" or unrecognized: appends nothing (preserves general baseline).
        /// Never throws; handles null systemPrompt gracefully.
        /// </summary>
        public static string AppendDomainPackRules(string systemPrompt, string domainPack)
        {
            string basePrompt = systemPrompt ?? string.Empty;

            if (string.Equals(domainPack, "railway", StringComparison.OrdinalIgnoreCase))
            {
                string railwayRules = "When the user's content relates to railway operations, infrastructure, or maintenance, " +
                    "use accurate railway operational vocabulary. For example, use 'Depot', 'Substation', 'OHE' (Overhead Equipment), " +
                    "'TRD' (Traction Distribution), 'PM/CM' (Preventive/Corrective Maintenance), 'Breakdown', 'DRM' (Divisional Railway Manager), " +
                    "'Sr.DEE' (Senior Divisional Electrical Engineer), 'SSE' (Senior Section Engineer), and 'JE' (Junior Engineer) where appropriate. " +
                    "Maintain the same professional register as the general pack. Do not invent railway-specific facts, terminology, or operational details " +
                    "that are not warranted by the actual content provided.";

                if (string.IsNullOrEmpty(basePrompt))
                {
                    return railwayRules;
                }
                else
                {
                    return basePrompt + "\n\n" + railwayRules;
                }
            }

            // For "general", null, empty, or any unrecognized domain pack: return base unchanged
            return basePrompt;
        }
    }
}
