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
                    "6. Supported action types are formula, value, filldown, table, create_table, conditional_format, sort, filter, data_validation, chart, pivot_table, named_range, and remove_duplicates. " +
                    "Use the value attribute for each action's concise option (for example value=\"descending\", value=\"list:Open,Closed\", or value=\"columns:1,2\") and state any assumptions in the description.\n" +
                    "7. Provide a brief conversational summary above or below the action block without tutorial how-to steps.";
            }
            else if (string.Equals(hostType, "Word", StringComparison.OrdinalIgnoreCase))
            {
                hostContext = "\n\nYou are embedded inside Microsoft Word. The provided document context may include prompt-relevant excerpts, a live outline, and action items. " +
                    "When the user asks to write, edit, rewrite, summarize, translate, or review text, provide polished text directly without tutorial meta-commentary. " +
                    "Preserve factual meaning unless the user requests a substantive change. Use a Markdown table when a table is the clearest result. " +
                    "If sources are attached, cite only supplied sources using [Source: filename, page/section]; do not invent citations.";
            }
            else if (string.Equals(hostType, "PowerPoint", StringComparison.OrdinalIgnoreCase))
            {
                hostContext = "\n\nYou are embedded inside Microsoft PowerPoint. The supplied context can include the full deck, sections, slide text, and speaker notes. " +
                    "When the user asks for slides or bullet points, provide structured slides with Slide titles, concise bullets, speaker notes, and a one-line Visual suggestion. " +
                    "Preserve the presentation's existing style and structure where possible. For a requested deck reorganization, use only an optional <powerpoint_actions> block with safe action types move_slide, create_section, rename_section, or set_notes; use numbered existing slides and include no other commands.";
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
    }
}
