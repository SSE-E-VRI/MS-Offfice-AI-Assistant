using System;
using System.Collections.Generic;

namespace MSOfficeAIAssistant.Tests
{
    internal class Program
    {
        // WPF TextBlock/Inlines/TextRange (used by MarkdownHelperTests) require an STA thread.
        [STAThread]
        private static int Main(string[] args)
        {
            Console.WriteLine("=================================================");
            Console.WriteLine("MS Office AI Assistant - Standalone Test Suite   ");
            Console.WriteLine("=================================================");
            Console.WriteLine();

            int passed = 0;
            int failed = 0;
            var suites = new List<Tuple<string, Action>>
            {
                Tuple.Create<string, Action>("GoldenMasterBaselineTests", GoldenMasterBaselineTests.RunAll),
                Tuple.Create<string, Action>("WordDocumentContextBuilderTests", WordDocumentContextBuilderTests.RunAll),
                Tuple.Create<string, Action>("SpreadsheetActionParserTests", SpreadsheetActionParserTests.RunAll),
                Tuple.Create<string, Action>("PowerPointActionParserTests", PowerPointActionParserTests.RunAll),
                Tuple.Create<string, Action>("ActionAuditStoreTests", ActionAuditStoreTests.RunAll),
                Tuple.Create<string, Action>("ChatMessageSerializationTests", ChatMessageSerializationTests.RunAll),
                Tuple.Create<string, Action>("SessionOrchestratorTests", SessionOrchestratorTests.RunAll),
                Tuple.Create<string, Action>("ComResilienceTests", ComResilienceTests.RunAll),
                Tuple.Create<string, Action>("ProviderCapabilitiesTests", ProviderCapabilitiesTests.RunAll),
                Tuple.Create<string, Action>("OfficeHostControllerTests", OfficeHostControllerTests.RunAll),
                Tuple.Create<string, Action>("DocToDeckTests", DocToDeckTests.RunAll),
                Tuple.Create<string, Action>("HostOperationResultTests", HostOperationResultTests.RunAll),
                Tuple.Create<string, Action>("ActionExtractorTests", ActionExtractorTests.RunAll),
                Tuple.Create<string, Action>("ToolRegistryTests", ToolRegistryTests.RunAll),
                Tuple.Create<string, Action>("ActionVerifierTests", ActionVerifierTests.RunAll),
                Tuple.Create<string, Action>("RollbackExecutorTests", RollbackExecutorTests.RunAll),
                Tuple.Create<string, Action>("PlannerTests", PlannerTests.RunAll),
                Tuple.Create<string, Action>("WorkSessionStoreTests", WorkSessionStoreTests.RunAll),
                Tuple.Create<string, Action>("PlanExecutorTests", PlanExecutorTests.RunAll),
                Tuple.Create<string, Action>("CrossHostPlanCoordinatorTests", CrossHostPlanCoordinatorTests.RunAll),
                Tuple.Create<string, Action>("MarkdownHelperTests", MarkdownHelperTests.RunAll),
                Tuple.Create<string, Action>("MarkdownHelperCitationTests", MarkdownHelperCitationTests.RunAll),
                Tuple.Create<string, Action>("XamlLoadTests", XamlLoadTests.RunAll),
                Tuple.Create<string, Action>("AssistantSessionModeTests", AssistantSessionModeTests.RunAll),
                Tuple.Create<string, Action>("AssistantSessionPlanModeTests", AssistantSessionPlanModeTests.RunAll),
                Tuple.Create<string, Action>("ResponseCardCategoryTests", ResponseCardCategoryTests.RunAll),
                Tuple.Create<string, Action>("EvidenceLevelTests", EvidenceLevelTests.RunAll),
                Tuple.Create<string, Action>("AttachmentExtractorProvenanceTests", AttachmentExtractorProvenanceTests.RunAll),
                Tuple.Create<string, Action>("QuickPromptRegistryTests", QuickPromptRegistryTests.RunAll),
                Tuple.Create<string, Action>("SkillRegistryTests", SkillRegistryTests.RunAll),
                Tuple.Create<string, Action>("SkillPickerTests", SkillPickerTests.RunAll),
                Tuple.Create<string, Action>("ConversationStoreSessionTests", ConversationStoreSessionTests.RunAll),
                Tuple.Create<string, Action>("AssistantStatusTests", AssistantStatusTests.RunAll),
                Tuple.Create<string, Action>("ResponseContentCleanerTests", ResponseContentCleanerTests.RunAll),
                Tuple.Create<string, Action>("MarkdownClipboardTests", MarkdownClipboardTests.RunAll),
                Tuple.Create<string, Action>("AuditDisplayConverterTests", AuditDisplayConverterTests.RunAll)
            };

            foreach (var suite in suites)
            {
                Console.Write("Running " + suite.Item1 + "... ");
                try
                {
                    suite.Item2();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("[PASS]");
                    Console.ResetColor();
                    passed++;
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[FAIL]");
                    Console.WriteLine("  Error: " + ex.Message);
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine("  Inner: " + ex.InnerException.Message);
                    }
                    Console.ResetColor();
                    failed++;
                }
            }

            Console.WriteLine();
            Console.WriteLine("=================================================");
            if (failed == 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(string.Format("ALL TEST SUITES PASSED ({0}/{1})", passed, passed + failed));
                Console.ResetColor();
                return 0;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("TEST FAILURES OCCURRED: {0} passed, {1} failed", passed, failed));
                Console.ResetColor();
                return 1;
            }
        }
    }
}
