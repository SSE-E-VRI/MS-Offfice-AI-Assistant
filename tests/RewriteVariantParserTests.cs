using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.UI.Helpers;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Covers RewriteVariantParser.Split, which turns a "3 Variants" ribbon response
    /// (RibbonCallback.OnRewriteVariants) into the individual candidates RewriteVariantsWindow
    /// pages through.
    /// </summary>
    public static class RewriteVariantParserTests
    {
        public static void RunAll()
        {
            TestThreeVariantsSeparatedByDelimiter();
            TestNoDelimiterReturnsWholeContentAsSingleVariant();
            TestLeadingAndTrailingDelimitersAreIgnored();
            TestBlankOrNullContentReturnsEmptyList();
            TestCrlfLineEndingsStillSplit();
        }

        private static void TestThreeVariantsSeparatedByDelimiter()
        {
            string content = "First rewrite.\n---VARIANT---\nSecond rewrite.\n---VARIANT---\nThird rewrite.";
            List<string> variants = RewriteVariantParser.Split(content);
            Assert(variants.Count == 3, "expected 3 variants, got " + variants.Count);
            Assert(variants[0] == "First rewrite.", "variant 1 mismatch: " + variants[0]);
            Assert(variants[1] == "Second rewrite.", "variant 2 mismatch: " + variants[1]);
            Assert(variants[2] == "Third rewrite.", "variant 3 mismatch: " + variants[2]);
        }

        private static void TestNoDelimiterReturnsWholeContentAsSingleVariant()
        {
            string content = "Just one rewrite, no delimiter at all.";
            List<string> variants = RewriteVariantParser.Split(content);
            Assert(variants.Count == 1, "expected 1 variant, got " + variants.Count);
            Assert(variants[0] == content, "single variant should be the trimmed whole content");
        }

        private static void TestLeadingAndTrailingDelimitersAreIgnored()
        {
            // Prompt asks for a delimiter before the first alternative too; a stray delimiter
            // after the last one (model didn't follow instructions perfectly) shouldn't produce
            // a blank trailing "variant" either.
            string content = "---VARIANT---\nAlpha\n---VARIANT---\nBeta\n---VARIANT---\n";
            List<string> variants = RewriteVariantParser.Split(content);
            Assert(variants.Count == 2, "expected 2 variants, got " + variants.Count);
            Assert(variants[0] == "Alpha", "variant 1 mismatch: " + variants[0]);
            Assert(variants[1] == "Beta", "variant 2 mismatch: " + variants[1]);
        }

        private static void TestBlankOrNullContentReturnsEmptyList()
        {
            Assert(RewriteVariantParser.Split(null).Count == 0, "null content should yield no variants");
            Assert(RewriteVariantParser.Split("   \n  ").Count == 0, "whitespace-only content should yield no variants");
        }

        private static void TestCrlfLineEndingsStillSplit()
        {
            string content = "Alpha\r\n---VARIANT---\r\nBeta";
            List<string> variants = RewriteVariantParser.Split(content);
            Assert(variants.Count == 2, "expected 2 variants with CRLF endings, got " + variants.Count);
            Assert(variants[0] == "Alpha", "variant 1 mismatch: " + variants[0]);
            Assert(variants[1] == "Beta", "variant 2 mismatch: " + variants[1]);
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion failed: " + message);
            }
        }
    }
}
