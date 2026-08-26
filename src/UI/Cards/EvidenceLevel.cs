using System;
using System.Text.RegularExpressions;

namespace MSOfficeAIAssistant.UI.Cards
{
    /// <summary>
    /// Evidence level for a Finding, representing the quality/basis of the observation.
    /// Levels are ordered from most to least certain, and are determined by objective signals
    /// (real source citations in the Finding's Content text), never by model self-report claims.
    /// </summary>
    public enum EvidenceLevel
    {
        DirectlyObserved,
        Calculated,
        StrongInference,
        PossibleInference,
        InsufficientEvidence
    }

    /// <summary>
    /// Pure, testable classifier that determines evidence level for a Finding's content.
    /// No WPF dependencies — usable from COM-free test projects.
    ///
    /// Evidence classification priority:
    /// 1. Citation-pattern detection (primary signal): looks for recognizable Phase-A4 provenance tags
    ///    (Word paragraph tags, excerpt labels, Excel cell addresses, PowerPoint slide references).
    ///    If ANY pattern is found in Content, classify as DirectlyObserved immediately.
    ///
    /// 2. Explicit bracketed category tag (secondary signal): if no citation pattern found,
    ///    looks for [Calculated], [Strong Inference], [Possible Inference], or [Insufficient Evidence]
    ///    (case-insensitive, optional surrounding whitespace). Only trusted for structured
    ///    reasoning categories, never for confidence/certainty claims.
    ///
    /// 3. Default: InsufficientEvidence — honest, conservative fallback when neither signal applies.
    /// </summary>
    public static class EvidenceLevelClassifier
    {
        /// <summary>
        /// Classifies a Finding's content into an evidence level based on objective signals:
        /// citation patterns (highest priority) and explicit bracketed tags (secondary).
        /// Null/empty content returns InsufficientEvidence.
        /// </summary>
        public static EvidenceLevel Classify(string findingContent)
        {
            // Null/empty → InsufficientEvidence
            if (string.IsNullOrEmpty(findingContent))
            {
                return EvidenceLevel.InsufficientEvidence;
            }

            // Signal 1 (primary): Citation-pattern detection
            if (ContainsCitationPattern(findingContent))
            {
                return EvidenceLevel.DirectlyObserved;
            }

            // Signal 2 (secondary): Explicit bracketed tag (only if no citation found)
            EvidenceLevel? tagLevel = ExtractBracketedTagLevel(findingContent);
            if (tagLevel.HasValue)
            {
                return tagLevel.Value;
            }

            // Default: InsufficientEvidence (no citation, no explicit tag)
            return EvidenceLevel.InsufficientEvidence;
        }

        /// <summary>
        /// Returns the Finding content with a leading bracketed evidence tag removed (if present).
        /// Case-insensitive; also trims the whitespace/newline that follows the tag.
        /// If no such tag is present, returns the input unchanged.
        /// This is what the Finding card's body text should display — raw tag never shown twice.
        /// </summary>
        public static string StripEvidenceTag(string findingContent)
        {
            if (string.IsNullOrEmpty(findingContent))
            {
                return findingContent;
            }

            string trimmed = findingContent.Trim();

            // Pattern: [tag] at the start (case-insensitive), optional whitespace after
            // Matches [Calculated], [Strong Inference], [Possible Inference], [Insufficient Evidence]
            string pattern = @"^\s*\[(Calculated|Strong Inference|Possible Inference|Insufficient Evidence)\]\s*";
            string result = Regex.Replace(trimmed, pattern, "", RegexOptions.IgnoreCase);

            return result;
        }

        /// <summary>
        /// Returns a short display label for the given evidence level.
        /// </summary>
        public static string GetLabel(EvidenceLevel level)
        {
            switch (level)
            {
                case EvidenceLevel.DirectlyObserved:
                    return "Directly Observed";
                case EvidenceLevel.Calculated:
                    return "Calculated";
                case EvidenceLevel.StrongInference:
                    return "Strong Inference";
                case EvidenceLevel.PossibleInference:
                    return "Possible Inference";
                case EvidenceLevel.InsufficientEvidence:
                    return "Insufficient Evidence";
                default:
                    return "Unknown";
            }
        }

        /// <summary>
        /// Returns a visually distinct icon/glyph for the given evidence level.
        /// Each level has a unique icon so status is never conveyed by color alone.
        /// </summary>
        public static string GetIcon(EvidenceLevel level)
        {
            switch (level)
            {
                case EvidenceLevel.DirectlyObserved:
                    return "✓";
                case EvidenceLevel.Calculated:
                    return "=";
                case EvidenceLevel.StrongInference:
                    return "◆";
                case EvidenceLevel.PossibleInference:
                    return "◇";
                case EvidenceLevel.InsufficientEvidence:
                    return "?";
                default:
                    return "●";
            }
        }

        /// <summary>
        /// Checks if the given content contains any of the recognized Phase-A4 citation patterns.
        /// Returns true if ANY pattern is found; this is the primary, most trustworthy signal.
        /// </summary>
        private static bool ContainsCitationPattern(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return false;
            }

            // Pattern 1: Word paragraph tag, e.g. [¶12]
            if (Regex.IsMatch(content, @"\[¶\d+\]"))
            {
                return true;
            }

            // Pattern 2: Word excerpt label, e.g. ~Paragraph 5
            if (Regex.IsMatch(content, @"~Paragraph\s+\d+"))
            {
                return true;
            }

            // Pattern 3: Excel sheet-qualified cell address, e.g. Sheet1!B7, Budget!C12, or with $ signs
            if (Regex.IsMatch(content, @"[A-Za-z0-9_]+!\$?[A-Z]+\$?\d+"))
            {
                return true;
            }

            // Pattern 4: Bare cell-address-equals-value tag, e.g. B7=1234
            if (Regex.IsMatch(content, @"\b[A-Z]{1,3}\d{1,7}="))
            {
                return true;
            }

            // Pattern 5: PowerPoint slide reference, e.g. Slide 3 of 12
            if (Regex.IsMatch(content, @"Slide\s+\d+\s+of\s+\d+"))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Extracts the evidence level from a leading bracketed tag in the content (if present).
        /// Returns null if no bracketed tag found.
        /// Case-insensitive.
        /// </summary>
        private static EvidenceLevel? ExtractBracketedTagLevel(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return null;
            }

            string trimmed = content.Trim();

            // Try to match leading bracketed tags (case-insensitive)
            Match match = Regex.Match(trimmed, @"^\s*\[([^\]]+)\]", RegexOptions.IgnoreCase);
            if (!match.Success || match.Groups.Count < 2)
            {
                return null;
            }

            string tag = match.Groups[1].Value.Trim();

            // Map tag to EvidenceLevel (case-insensitive comparison)
            if (tag.Equals("Calculated", StringComparison.OrdinalIgnoreCase))
            {
                return EvidenceLevel.Calculated;
            }
            if (tag.Equals("Strong Inference", StringComparison.OrdinalIgnoreCase))
            {
                return EvidenceLevel.StrongInference;
            }
            if (tag.Equals("Possible Inference", StringComparison.OrdinalIgnoreCase))
            {
                return EvidenceLevel.PossibleInference;
            }
            if (tag.Equals("Insufficient Evidence", StringComparison.OrdinalIgnoreCase))
            {
                return EvidenceLevel.InsufficientEvidence;
            }

            // Tag doesn't match known categories → no signal
            return null;
        }
    }
}
